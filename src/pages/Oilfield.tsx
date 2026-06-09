import React, { useRef, useState, useCallback } from "react";
import * as XLSX from "xlsx";
import { useNavigate } from "react-router-dom";
import Icon from "@/components/ui/icon";
import Chart1HorizontalBars from "@/components/charts/Chart1HorizontalBars";
import Chart2ComboChart from "@/components/charts/Chart2ComboChart";
import Chart1Editor from "@/components/editor/Chart1Editor";
import Chart2Editor from "@/components/editor/Chart2Editor";
import {
  OilfieldData,
  createDefaultOilfield,
  CHART1_INDICATORS,
} from "@/types/oilfield";
import func2url from "../../backend/func2url.json";

const EXPORT_URL = (func2url as Record<string, string>)["export-xlsx"] ?? "";

// Маппинг: название показателя в Chart1 → ключевые слова для поиска в xlsx (строки Проект/Факт)
const INDICATOR_KEYWORDS: Record<string, string[]> = {
  "Добыча нефти":                   ["добыча нефти", "нефть, тыс.т", "добыча нефти, тыс"],
  "Добыча жидкости":                ["добыча жидкости", "жидкость, тыс.т", "добыча жидкости, тыс"],
  "Обводненность (весовая)":        ["обводненность", "обводн"],
  "Дейст. добывающий фонд":         ["фонд доб", "добывающ", "фонд добывающих", "добыв. скв", "фонд добыв"],
  "Дейст. нагнетательный фонд":     ["фонд нагн", "нагнетател", "нагн. скв", "фонд нагнет"],
  "Закачка":                        ["закачка", "закачка рабочего", "закачка, тыс.м"],
  "Приемистость":                   ["приемистость", "приёмистость"],
  "Дебит нефти":                    ["дебит нефти", "дебит нефт"],
  "Дебит жидкости":                 ["дебит жидкости", "дебит жидк"],
};

// Маппинг: поле chart2 → ключевые слова для поиска строки Факт
const CHART2_KEYWORDS: Record<keyof import("@/types/oilfield").Chart2Data, string[]> = {
  years:        [],
  zakachka:     ["закачка"],
  liquid:       ["добыча жидкости", "жидкость"],
  oil:          ["добыча нефти", "нефть"],
  fond_dob:     ["фонд доб", "добывающ"],
  fond_nag:     ["фонд нагн", "нагнетател"],
  compensation: ["компенсация"],
};

function toNum(v: string | number | null): number | null {
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    const n = parseFloat(v.replace(",", ".").replace("%", "").trim());
    return isNaN(n) ? null : n;
  }
  return null;
}

function cellMatchesKeywords(cell: string | number | null, keywords: string[]): boolean {
  if (!cell) return false;
  const low = String(cell).toLowerCase();
  return keywords.some(k => low.includes(k.toLowerCase()));
}

function parseXlsxLocally(file: File): Promise<OilfieldData[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target!.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: "array" });
        const fields: OilfieldData[] = [];

        wb.SheetNames.forEach((sheetName) => {
          const ws = wb.Sheets[sheetName];
          const rows: (string | number | null)[][] = XLSX.utils.sheet_to_json(ws, {
            header: 1,
            defval: null,
          });

          const field = createDefaultOilfield(sheetName);

          // 1. Найти строку с годами (заголовок)
          let headerRowIdx = -1;
          let yearColIdxs: number[] = [];
          let years: number[] = [];

          for (let ri = 0; ri < Math.min(rows.length, 15); ri++) {
            const row = rows[ri];
            const yearCols: { idx: number; year: number }[] = [];
            row.forEach((v, ci) => {
              if (typeof v === "number" && v >= 2000 && v <= 2050) {
                yearCols.push({ idx: ci, year: v });
              }
            });
            if (yearCols.length >= 2) {
              headerRowIdx = ri;
              yearColIdxs = yearCols.map(y => y.idx);
              years = yearCols.map(y => y.year);
              break;
            }
          }

          if (headerRowIdx < 0 || years.length === 0) {
            fields.push(field);
            return;
          }

          field.years = years;
          const dataRows = rows.slice(headerRowIdx + 1);

          // Для каждого показателя ищем строки "Проект" и "Факт"
          // Ищем по контексту: показатель в col0/1, "проект"/"факт" в следующих col0/1

          // Индексируем все строки: для каждой строки смотрим первые 3 ячейки
          // Алгоритм: идём по строкам, если ячейка совпадает с ключевыми словами показателя —
          // ищем следующие строки где написано "проект" или "факт"

          const findProjFact = (keywords: string[]): { proj: (number|null)[]; fact: (number|null)[] } | null => {
            for (let ri = 0; ri < dataRows.length; ri++) {
              const row = dataRows[ri];
              // Проверяем первые 3 колонки на совпадение с ключевым словом показателя
              const cellsToCheck = [row[0], row[1], row[2]].filter(Boolean);
              const isIndicatorRow = cellsToCheck.some(c => cellMatchesKeywords(c, keywords));
              if (!isIndicatorRow) continue;

              // Нашли строку с названием показателя
              // Ищем "проект" и "факт" в диапазоне ±5 строк
              let projRow: (string | number | null)[] | null = null;
              let factRow: (string | number | null)[] | null = null;

              // Сначала проверяем саму строку — может быть "проект" уже в ней
              const allCells = [row[0], row[1], row[2], row[3]].map(c => String(c ?? "").toLowerCase());
              if (allCells.some(c => c.includes("проект"))) projRow = row;
              if (allCells.some(c => c.includes("факт"))) factRow = row;

              // Ищем в соседних строках (вперёд до 6 строк)
              for (let di = 1; di <= 6 && (!projRow || !factRow); di++) {
                const nextRow = dataRows[ri + di];
                if (!nextRow) break;
                const nextCells = [nextRow[0], nextRow[1], nextRow[2], nextRow[3]].map(c => String(c ?? "").toLowerCase());
                // Если строка уже относится к другому показателю — стоп
                const isOtherIndicator = Object.values(INDICATOR_KEYWORDS).some(kws =>
                  !keywords.some(k => kws.includes(k)) && nextCells.some(c => kws.some(kw => c.includes(kw.toLowerCase())))
                );
                if (isOtherIndicator && di > 1) break;
                if (!projRow && nextCells.some(c => c.includes("проект"))) projRow = nextRow;
                if (!factRow && nextCells.some(c => c.includes("факт") && !c.includes("проект"))) factRow = nextRow;
              }

              if (!projRow && !factRow) continue;

              const extractValues = (r: (string | number | null)[]): (number|null)[] =>
                yearColIdxs.map(ci => toNum(r[ci]));

              return {
                proj: projRow ? extractValues(projRow) : years.map(() => null),
                fact: factRow ? extractValues(factRow) : years.map(() => null),
              };
            }
            return null;
          };

          // Chart1: (Факт - Проект) / |Проект| * 100
          const chart1Rows = CHART1_INDICATORS.map((indicator) => {
            const keywords = INDICATOR_KEYWORDS[indicator] ?? [indicator.toLowerCase()];
            const pf = findProjFact(keywords);
            if (!pf) return { indicator, values: years.map(() => null as number | null) };

            const values = years.map((_, yi) => {
              const proj = pf.proj[yi];
              const fact = pf.fact[yi];
              if (proj === null || fact === null || proj === 0) return null;
              return Math.round((fact - proj) / Math.abs(proj) * 1000) / 10; // 1 знак
            });
            return { indicator, values };
          });
          field.chart1Rows = chart1Rows;

          // Chart2: берём Факт-строки тех же показателей
          const getFactValues = (keywords: string[]): number[] => {
            const pf = findProjFact(keywords);
            if (!pf) return years.map(() => 0);
            return pf.fact.map(v => v ?? 0);
          };

          field.chart2Data = {
            years,
            zakachka:     getFactValues(CHART2_KEYWORDS.zakachka),
            liquid:       getFactValues(CHART2_KEYWORDS.liquid),
            oil:          getFactValues(CHART2_KEYWORDS.oil),
            fond_dob:     getFactValues(CHART2_KEYWORDS.fond_dob),
            fond_nag:     getFactValues(CHART2_KEYWORDS.fond_nag),
            compensation: getFactValues(CHART2_KEYWORDS.compensation),
          };

          fields.push(field);
        });

        resolve(fields);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

export default function Oilfield() {
  const navigate = useNavigate();
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [fields, setFields] = useState<OilfieldData[]>([
    createDefaultOilfield("Месторождение 1"),
  ]);
  const [selectedIdx, setSelectedIdx] = useState(0);
  const [activeTab, setActiveTab] = useState("chart1");
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [fileName, setFileName] = useState<string | null>(null);
  const [parseLog, setParseLog] = useState<string | null>(null);

  const selected = fields[selectedIdx];

  const updateField = useCallback(
    (updated: OilfieldData) => {
      setFields((prev) => prev.map((f, i) => (i === selectedIdx ? updated : f)));
    },
    [selectedIdx]
  );

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setLoading(true);
    setFileName(file.name);
    setParseLog(null);
    try {
      const parsed = await parseXlsxLocally(file);
      if (parsed.length > 0) {
        setFields(parsed);
        setSelectedIdx(0);
        // Проверяем сколько данных нашлось
        const firstField = parsed[0];
        const filledCells = firstField.chart1Rows.reduce((acc, row) =>
          acc + row.values.filter(v => v !== null).length, 0);
        const totalCells = firstField.chart1Rows.length * firstField.years.length;
        if (filledCells === 0) {
          setParseLog(`Файл загружен, но данные Проект/Факт не найдены. Проверь структуру таблицы — нужны строки «Проект» и «Факт» для каждого показателя.`);
        } else if (filledCells < totalCells * 0.5) {
          setParseLog(`Загружено: ${filledCells} из ${totalCells} ячеек. Часть показателей не распознана.`);
        }
      }
    } catch {
      setParseLog("Ошибка чтения файла. Проверьте формат xlsx.");
    } finally {
      setLoading(false);
      e.target.value = "";
    }
  };

  const addField = () => {
    const newField = createDefaultOilfield(`Месторождение ${fields.length + 1}`);
    setFields((prev) => [...prev, newField]);
    setSelectedIdx(fields.length);
  };

  const removeField = (idx: number) => {
    if (fields.length <= 1) return;
    setFields((prev) => prev.filter((_, i) => i !== idx));
    setSelectedIdx(Math.max(0, idx - 1));
  };

  const handleExport = async () => {
    if (!EXPORT_URL) {
      alert("Сервис экспорта недоступен. Попробуйте позже.");
      return;
    }
    setExporting(true);
    try {
      const payload = {
        field_name: selected.name,
        years: selected.years,
        chart1_rows: selected.chart1Rows.map((r) => r.values),
        chart2_data: selected.chart2Data,
      };
      const resp = await fetch(EXPORT_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const json = await resp.json();
      if (json.file) {
        const byteStr = atob(json.file);
        const bytes = new Uint8Array(byteStr.length);
        for (let i = 0; i < byteStr.length; i++) bytes[i] = byteStr.charCodeAt(i);
        const blob = new Blob([bytes], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = json.filename ?? `${selected.name}_report.xlsx`;
        a.click();
        URL.revokeObjectURL(url);
      }
    } catch {
      alert("Ошибка экспорта. Попробуйте ещё раз.");
    } finally {
      setExporting(false);
    }
  };

  return (
    <div className="min-h-screen flex flex-col" style={{ background: "#f5f5f5", fontFamily: "Calibri, Arial, sans-serif" }}>
      {/* Шапка */}
      <div className="border-b border-gray-400 flex-shrink-0" style={{ background: "#217346" }}>
        <div className="flex items-center gap-2 px-3 py-1.5">
          <button
            onClick={() => navigate("/")}
            className="flex items-center gap-1.5 text-white/80 hover:text-white text-xs transition-colors px-2 py-1 rounded hover:bg-white/10"
          >
            <Icon name="ArrowLeft" size={13} />
            DataMind
          </button>
          <div className="w-px h-4 bg-white/30" />
          <div className="flex items-center gap-1.5">
            <Icon name="BarChart2" size={16} className="text-white" />
            <span className="text-white font-semibold text-sm">
              Анализ разработки месторождений
            </span>
          </div>
          <div className="flex-1" />

          <label className="cursor-pointer">
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls"
              className="hidden"
              onChange={handleFileUpload}
            />
            <div className="flex items-center gap-1.5 px-3 py-1 rounded text-xs font-medium text-white border border-white/30 hover:bg-white/10 transition cursor-pointer">
              <Icon name="FolderOpen" size={13} />
              {loading ? "Загрузка..." : fileName ? fileName : "Открыть xlsx"}
            </div>
          </label>

          <button
            onClick={handleExport}
            disabled={exporting}
            className="flex items-center gap-1.5 px-3 py-1 rounded text-xs font-semibold border border-white/30 hover:bg-white/10 transition disabled:opacity-50"
            style={{ color: "#fff" }}
          >
            <Icon name="Download" size={13} />
            {exporting ? "Сохранение..." : "Скачать xlsx с графиками"}
          </button>
        </div>

        {/* Вкладки */}
        <div className="flex items-end gap-0.5 px-3 pt-1">
          {[
            { value: "chart1", label: "График 1: Изменения %" },
            { value: "chart2", label: "График 2: Динамика" },
          ].map((tab) => (
            <button
              key={tab.value}
              onClick={() => setActiveTab(tab.value)}
              className="px-4 py-1 text-xs rounded-t-sm font-medium transition-colors"
              style={{
                background: activeTab === tab.value ? "#fff" : "rgba(255,255,255,0.15)",
                color: activeTab === tab.value ? "#217346" : "#fff",
                borderTop: "1px solid",
                borderLeft: "1px solid",
                borderRight: "1px solid",
                borderColor: activeTab === tab.value ? "#ccc" : "transparent",
                marginBottom: activeTab === tab.value ? "-1px" : 0,
              }}
            >
              {tab.label}
            </button>
          ))}
        </div>
      </div>

      {parseLog && (
        <div className="flex items-center gap-2 px-4 py-2 text-xs border-b border-yellow-300/40" style={{ background: "rgba(234,179,8,0.1)", color: "#854d0e" }}>
          <Icon name="AlertCircle" size={13} />
          {parseLog}
          <button onClick={() => setParseLog(null)} className="ml-auto opacity-60 hover:opacity-100"><Icon name="X" size={12} /></button>
        </div>
      )}

      <div className="flex flex-1 overflow-hidden">
        {/* Боковая панель */}
        <div
          className="flex-shrink-0 border-r border-gray-300 flex flex-col"
          style={{ width: 190, background: "#fff", minWidth: 150 }}
        >
          <div
            className="text-xs font-semibold px-3 py-2 border-b border-gray-200"
            style={{ background: "#e8f5e9", color: "#1b5e20" }}
          >
            Месторождения / объекты
          </div>
          <div className="flex-1 overflow-y-auto">
            {fields.map((f, idx) => (
              <div
                key={idx}
                className={`group flex items-center gap-1 px-2 py-2 cursor-pointer text-xs border-b border-gray-100 transition ${
                  idx === selectedIdx
                    ? "bg-green-100 border-l-2 border-l-green-600"
                    : "hover:bg-green-50 text-gray-700"
                }`}
                onClick={() => setSelectedIdx(idx)}
              >
                <Icon name="Layers" size={11} />
                <input
                  className="flex-1 bg-transparent outline-none text-xs min-w-0 cursor-pointer"
                  value={f.name}
                  onClick={(e) => { e.stopPropagation(); setSelectedIdx(idx); }}
                  onChange={(e) =>
                    setFields((prev) =>
                      prev.map((ff, ii) => ii === idx ? { ...ff, name: e.target.value } : ff)
                    )
                  }
                  style={{
                    color: idx === selectedIdx ? "#1b5e20" : "#333",
                    fontWeight: idx === selectedIdx ? 600 : 400,
                  }}
                />
                {fields.length > 1 && (
                  <button
                    className="opacity-0 group-hover:opacity-100 text-red-400 hover:text-red-600 text-xs"
                    onClick={(e) => { e.stopPropagation(); removeField(idx); }}
                  >
                    ×
                  </button>
                )}
              </div>
            ))}
          </div>
          <button
            onClick={addField}
            className="text-xs text-green-700 hover:bg-green-50 px-3 py-2 border-t border-gray-200 flex items-center gap-1.5 transition"
          >
            <Icon name="Plus" size={12} />
            Добавить объект
          </button>
        </div>

        {/* Рабочая область */}
        <div className="flex-1 overflow-auto" style={{ background: "#fafafa" }}>
          <div
            className="flex items-center gap-2 px-3 py-1 border-b border-gray-300 flex-shrink-0 sticky top-0 z-20"
            style={{ background: "#fff" }}
          >
            <div
              className="text-xs font-semibold px-2 py-0.5 rounded border border-gray-300 text-gray-600 min-w-[140px]"
              style={{ fontFamily: "Calibri, Arial, sans-serif" }}
            >
              {selected.name}
            </div>
            <div className="w-px h-4 bg-gray-300" />
            <div className="text-xs text-gray-500 flex items-center gap-1">
              <Icon name="Info" size={11} />
              {activeTab === "chart1"
                ? `Изменения показателей по годам (%), лет: ${selected.years.length}`
                : `Динамика разработки, лет: ${selected.chart2Data.years.length}`}
            </div>
          </div>

          <div className="p-4 flex flex-col gap-4">
            {activeTab === "chart1" && (
              <>
                <div className="bg-white border border-gray-400 shadow-sm">
                  <div
                    className="px-3 py-1 border-b border-gray-300 text-xs font-semibold flex items-center gap-2"
                    style={{ background: "#e8f5e9", color: "#1b5e20" }}
                  >
                    <Icon name="Table2" size={12} />
                    Таблица данных — изменения показателей, %
                  </div>
                  <div className="p-2">
                    <Chart1Editor
                      years={selected.years}
                      rows={selected.chart1Rows}
                      onChange={(rows) => updateField({ ...selected, chart1Rows: rows })}
                      onYearsChange={(years) => updateField({ ...selected, years })}
                    />
                  </div>
                </div>

                <div className="bg-white border border-gray-400 shadow-sm">
                  <div
                    className="px-3 py-1 border-b border-gray-300 text-xs font-semibold flex items-center gap-2"
                    style={{ background: "#e8f5e9", color: "#1b5e20" }}
                  >
                    <Icon name="BarChart2" size={12} />
                    Предпросмотр — как будет в Excel
                  </div>
                  <div className="p-4">
                    <Chart1HorizontalBars
                      years={selected.years}
                      rows={selected.chart1Rows}
                    />
                  </div>
                </div>
              </>
            )}

            {activeTab === "chart2" && (
              <>
                <div className="bg-white border border-gray-400 shadow-sm">
                  <div
                    className="px-3 py-1 border-b border-gray-300 text-xs font-semibold flex items-center gap-2"
                    style={{ background: "#e8f5e9", color: "#1b5e20" }}
                  >
                    <Icon name="Table2" size={12} />
                    Таблица данных — динамика показателей (Факт)
                  </div>
                  <div className="p-2">
                    <Chart2Editor
                      data={selected.chart2Data}
                      onChange={(chart2Data) => updateField({ ...selected, chart2Data })}
                    />
                  </div>
                </div>

                <div className="bg-white border border-gray-400 shadow-sm">
                  <div
                    className="px-3 py-1 border-b border-gray-300 text-xs font-semibold flex items-center gap-2"
                    style={{ background: "#e8f5e9", color: "#1b5e20" }}
                  >
                    <Icon name="TrendingUp" size={12} />
                    Предпросмотр — как будет в Excel
                  </div>
                  <div className="p-4">
                    <Chart2ComboChart data={selected.chart2Data} />
                  </div>
                </div>
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
