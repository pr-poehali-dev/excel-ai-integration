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

// Ключевые слова для поиска показателей в xlsx
const INDICATOR_KEYWORDS: Record<string, string[]> = {
  "Добыча нефти":               ["добыча нефти", "нефть, тыс", "добычи нефти"],
  "Добыча жидкости":            ["добыча жидкости", "жидкость, тыс", "добычи жидкости"],
  "Обводненность (весовая)":    ["обводненность", "обводн", "весовая"],
  "Дейст. добывающий фонд":     ["фонд добыв", "добывающих скв", "действующий фонд доб", "фонд доб"],
  "Дейст. нагнетательный фонд": ["фонд нагн", "нагнетательных скв", "действующий фонд наг"],
  "Закачка":                    ["закачка", "закачки"],
  "Приемистость":               ["приемистость", "приёмистость", "приимистость"],
  "Дебит нефти":                ["дебит нефти", "дебит нефт"],
  "Дебит жидкости":             ["дебит жидкости", "дебит жидк"],
};

const CHART2_KEYWORDS: Record<string, string[]> = {
  zakachka:     ["закачка", "закачки"],
  liquid:       ["добыча жидкости", "жидкость, тыс"],
  oil:          ["добыча нефти", "нефть, тыс"],
  fond_dob:     ["фонд добыв", "добывающих скв"],
  fond_nag:     ["фонд нагн", "нагнетательных скв"],
  compensation: ["компенсация"],
};

function toNum(v: string | number | null | undefined): number | null {
  if (typeof v === "number") return isFinite(v) ? v : null;
  if (typeof v === "string") {
    const n = parseFloat(v.replace(",", ".").replace("%", "").replace(/\s/g, "").trim());
    return isNaN(n) ? null : n;
  }
  return null;
}

function rowText(row: (string | number | null)[]): string {
  return row.map(c => String(c ?? "").toLowerCase()).join(" | ");
}

function matchesAny(text: string, keywords: string[]): boolean {
  return keywords.some(k => text.includes(k.toLowerCase()));
}

// Универсальный парсер: читает всю таблицу, ищет строки с "проект" и "факт"
// рядом с названием показателя
function parseXlsxLocally(file: File): Promise<{ fields: OilfieldData[]; log: string[] }> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target!.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: "array" });
        const fields: OilfieldData[] = [];
        const log: string[] = [];

        wb.SheetNames.forEach((sheetName) => {
          const ws = wb.Sheets[sheetName];
          const rows: (string | number | null)[][] = XLSX.utils.sheet_to_json(ws, {
            header: 1,
            defval: null,
          });

          log.push(`Лист: ${sheetName}, строк: ${rows.length}`);

          const field = createDefaultOilfield(sheetName);

          // 1. Найти строку с годами
          let headerRowIdx = -1;
          let yearColIdxs: number[] = [];
          let years: number[] = [];

          for (let ri = 0; ri < Math.min(rows.length, 20); ri++) {
            const row = rows[ri];
            const yearCols: { idx: number; year: number }[] = [];
            row.forEach((v, ci) => {
              const n = typeof v === "number" ? v : toNum(v);
              if (n !== null && n >= 2000 && n <= 2050) {
                yearCols.push({ idx: ci, year: n });
              }
            });
            if (yearCols.length >= 2) {
              headerRowIdx = ri;
              yearColIdxs = yearCols.map(y => y.idx);
              years = yearCols.map(y => y.year);
              log.push(`  Годы найдены в строке ${ri}: ${years.join(", ")} (колонки: ${yearColIdxs.join(",")})`);
              break;
            }
          }

          if (headerRowIdx < 0 || years.length === 0) {
            log.push(`  ОШИБКА: строка с годами не найдена`);
            fields.push(field);
            return;
          }

          field.years = years;

          // 2. Строим индекс: для каждой строки — её текст и значения по колонкам годов
          interface RowInfo {
            ri: number;
            text: string;       // весь текст строки в нижнем регистре
            firstCells: string; // первые 4 ячейки
            values: (number | null)[];
          }

          const rowInfos: RowInfo[] = rows.slice(headerRowIdx + 1).map((row, ri) => ({
            ri,
            text: rowText(row),
            firstCells: row.slice(0, 6).map(c => String(c ?? "").toLowerCase()).join(" "),
            values: yearColIdxs.map(ci => toNum(row[ci])),
          }));

          // 3. Для показателя ищем его строку, потом ищем строки проект/факт рядом
          const findProjFact = (
            indicatorName: string,
            keywords: string[]
          ): { proj: (number|null)[]; fact: (number|null)[] } | null => {

            // Ищем строку, где упоминается показатель
            // Это может быть сама строка Проект или строка-заголовок показателя
            for (let i = 0; i < rowInfos.length; i++) {
              const r = rowInfos[i];
              if (!matchesAny(r.firstCells, keywords)) continue;

              // Нашли строку с упоминанием показателя
              // Смотрим её и следующие 8 строк на наличие "проект" и "факт"
              let projRow: RowInfo | null = null;
              let factRow: RowInfo | null = null;

              const window = rowInfos.slice(i, i + 9);
              for (const wr of window) {
                if (!projRow && wr.text.includes("проект") && !wr.text.includes("факт")) {
                  projRow = wr;
                }
                if (!factRow && wr.text.includes("факт") && !wr.text.includes("прогноз")) {
                  factRow = wr;
                }
                if (projRow && factRow) break;
              }

              // Если не нашли оба — пробуем обратное: сама строка является проектом или фактом
              if (!projRow && r.text.includes("проект")) projRow = r;
              if (!factRow && r.text.includes("факт")) factRow = r;

              if (!projRow && !factRow) continue;

              log.push(`  ${indicatorName}: proj=${projRow ? `строка ${projRow.ri}` : "нет"} fact=${factRow ? `строка ${factRow.ri}` : "нет"}`);
              if (projRow) log.push(`    proj vals: ${projRow.values.join(", ")}`);
              if (factRow) log.push(`    fact vals: ${factRow.values.join(", ")}`);

              return {
                proj: projRow ? projRow.values : years.map(() => null),
                fact: factRow ? factRow.values : years.map(() => null),
              };
            }

            log.push(`  ${indicatorName}: НЕ НАЙДЕНО (ключи: ${keywords.join(", ")})`);
            return null;
          };

          // 4. Chart1: отклонения %
          const chart1Rows = CHART1_INDICATORS.map((indicator) => {
            const keywords = INDICATOR_KEYWORDS[indicator] ?? [indicator.toLowerCase()];
            const pf = findProjFact(indicator, keywords);
            if (!pf) return { indicator, values: years.map(() => null as number | null) };

            const values = years.map((_, yi) => {
              const proj = pf.proj[yi];
              const fact = pf.fact[yi];
              if (proj === null || fact === null || proj === 0) return null;
              return Math.round((fact - proj) / Math.abs(proj) * 1000) / 10;
            });
            return { indicator, values };
          });
          field.chart1Rows = chart1Rows;

          // 5. Chart2: значения Факт
          const chart2Keys = ["zakachka", "liquid", "oil", "fond_dob", "fond_nag", "compensation"] as const;
          const chart2 = { years, zakachka: [] as number[], liquid: [] as number[], oil: [] as number[], fond_dob: [] as number[], fond_nag: [] as number[], compensation: [] as number[] };

          for (const key of chart2Keys) {
            const keywords = CHART2_KEYWORDS[key] ?? [];
            const pf = findProjFact(key, keywords);
            chart2[key] = pf ? pf.fact.map(v => v ?? 0) : years.map(() => 0);
          }
          field.chart2Data = chart2;

          fields.push(field);
        });

        resolve({ fields, log });
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
  const [notice, setNotice] = useState<{ type: "warn" | "ok" | "err"; text: string } | null>(null);
  const [showLog, setShowLog] = useState(false);
  const [parseLog, setParseLog] = useState<string[]>([]);

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
    setNotice(null);
    try {
      const { fields: parsed, log } = await parseXlsxLocally(file);
      setParseLog(log);
      if (parsed.length > 0) {
        setFields(parsed);
        setSelectedIdx(0);
        const firstField = parsed[0];
        const filled = firstField.chart1Rows.reduce((a, r) => a + r.values.filter(v => v !== null).length, 0);
        const total = firstField.chart1Rows.length * firstField.years.length;
        if (filled === 0) {
          setNotice({ type: "warn", text: `Данные не распознаны. Нажми «Лог» чтобы увидеть что прочиталось.` });
        } else if (filled < total * 0.5) {
          setNotice({ type: "warn", text: `Распознано ${filled} из ${total} ячеек. Часть показателей не найдена.` });
        } else {
          setNotice({ type: "ok", text: `Загружено: ${parsed.length} лист(ов), ${filled}/${total} ячеек заполнено.` });
        }
      }
    } catch {
      setNotice({ type: "err", text: "Ошибка чтения файла. Проверьте формат xlsx." });
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
    if (!EXPORT_URL) { alert("Сервис экспорта недоступен."); return; }
    setExporting(true);
    try {
      const payload = {
        field_name: selected.name,
        years: selected.years,
        chart1_rows: selected.chart1Rows.map((r) => r.values),
        chart2_data: selected.chart2Data,
      };
      const resp = await fetch(EXPORT_URL, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
      const json = await resp.json();
      if (json.file) {
        const byteStr = atob(json.file);
        const bytes = new Uint8Array(byteStr.length);
        for (let i = 0; i < byteStr.length; i++) bytes[i] = byteStr.charCodeAt(i);
        const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = json.filename ?? `${selected.name}_report.xlsx`;
        a.click();
        URL.revokeObjectURL(url);
      }
    } catch { alert("Ошибка экспорта."); }
    finally { setExporting(false); }
  };

  const noticeColors = {
    warn: { bg: "rgba(234,179,8,0.1)", color: "#854d0e", border: "border-yellow-300/40" },
    ok:   { bg: "rgba(34,197,94,0.1)", color: "#166534", border: "border-green-300/40" },
    err:  { bg: "rgba(220,38,38,0.1)", color: "#991b1b", border: "border-red-300/40" },
  };

  return (
    <div className="min-h-screen flex flex-col" style={{ background: "#f5f5f5", fontFamily: "Calibri, Arial, sans-serif" }}>
      {/* Шапка */}
      <div className="border-b border-gray-400 flex-shrink-0" style={{ background: "#217346" }}>
        <div className="flex items-center gap-2 px-3 py-1.5 flex-wrap">
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
            <span className="text-white font-semibold text-sm">Анализ разработки месторождений</span>
          </div>
          <div className="flex-1" />

          <label className="cursor-pointer">
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleFileUpload} />
            <div className="flex items-center gap-1.5 px-3 py-1 rounded text-xs font-medium text-white border border-white/30 hover:bg-white/10 transition cursor-pointer">
              <Icon name="FolderOpen" size={13} />
              {loading ? "Загрузка..." : fileName ?? "Открыть xlsx"}
            </div>
          </label>

          {parseLog.length > 0 && (
            <button
              onClick={() => setShowLog(v => !v)}
              className="flex items-center gap-1 px-2 py-1 rounded text-xs border transition"
              style={{ borderColor: "rgba(255,255,255,0.3)", color: showLog ? "#34d399" : "rgba(255,255,255,0.6)" }}
            >
              <Icon name="Terminal" size={12} />
              Лог
            </button>
          )}

          <button
            onClick={handleExport}
            disabled={exporting}
            className="flex items-center gap-1.5 px-3 py-1 rounded text-xs font-semibold border border-white/30 hover:bg-white/10 transition disabled:opacity-50"
            style={{ color: "#fff" }}
          >
            <Icon name="Download" size={13} />
            {exporting ? "Сохранение..." : "Скачать xlsx"}
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
                borderTop: "1px solid", borderLeft: "1px solid", borderRight: "1px solid",
                borderColor: activeTab === tab.value ? "#ccc" : "transparent",
                marginBottom: activeTab === tab.value ? "-1px" : 0,
              }}
            >
              {tab.label}
            </button>
          ))}
        </div>
      </div>

      {/* Лог-панель */}
      {showLog && parseLog.length > 0 && (
        <div className="border-b border-gray-300 bg-gray-900 text-green-400 text-xs font-mono px-4 py-2 max-h-48 overflow-y-auto">
          {parseLog.map((line, i) => <div key={i}>{line}</div>)}
        </div>
      )}

      {/* Уведомление */}
      {notice && (() => {
        const c = noticeColors[notice.type];
        return (
          <div className={`flex items-center gap-2 px-4 py-2 text-xs border-b ${c.border}`} style={{ background: c.bg, color: c.color }}>
            <Icon name={notice.type === "ok" ? "CheckCircle" : "AlertCircle"} size={13} />
            {notice.text}
            <button onClick={() => setNotice(null)} className="ml-auto opacity-60 hover:opacity-100"><Icon name="X" size={12} /></button>
          </div>
        );
      })()}

      <div className="flex flex-1 overflow-hidden">
        {/* Боковая панель */}
        <div className="flex-shrink-0 border-r border-gray-300 flex flex-col" style={{ width: 190, background: "#fff" }}>
          <div className="text-xs font-semibold px-3 py-2 border-b border-gray-200" style={{ background: "#e8f5e9", color: "#1b5e20" }}>
            Месторождения / объекты
          </div>
          <div className="flex-1 overflow-y-auto">
            {fields.map((f, idx) => (
              <div
                key={idx}
                className={`group flex items-center gap-1 px-2 py-2 cursor-pointer text-xs border-b border-gray-100 transition ${
                  idx === selectedIdx ? "bg-green-100 border-l-2 border-l-green-600" : "hover:bg-green-50 text-gray-700"
                }`}
                onClick={() => setSelectedIdx(idx)}
              >
                <Icon name="Layers" size={11} />
                <input
                  className="flex-1 bg-transparent outline-none text-xs min-w-0 cursor-pointer"
                  value={f.name}
                  onClick={(e) => { e.stopPropagation(); setSelectedIdx(idx); }}
                  onChange={(e) => setFields(prev => prev.map((ff, ii) => ii === idx ? { ...ff, name: e.target.value } : ff))}
                  style={{ color: idx === selectedIdx ? "#1b5e20" : "#333", fontWeight: idx === selectedIdx ? 600 : 400 }}
                />
                {fields.length > 1 && (
                  <button
                    className="opacity-0 group-hover:opacity-100 text-red-400 hover:text-red-600 text-xs"
                    onClick={(e) => { e.stopPropagation(); removeField(idx); }}
                  >×</button>
                )}
              </div>
            ))}
          </div>
          <button onClick={addField} className="text-xs text-green-700 hover:bg-green-50 px-3 py-2 border-t border-gray-200 flex items-center gap-1.5 transition">
            <Icon name="Plus" size={12} />
            Добавить объект
          </button>
        </div>

        {/* Рабочая область */}
        <div className="flex-1 overflow-auto" style={{ background: "#fafafa" }}>
          <div className="flex items-center gap-2 px-3 py-1 border-b border-gray-300 sticky top-0 z-20" style={{ background: "#fff" }}>
            <div className="text-xs font-semibold px-2 py-0.5 rounded border border-gray-300 text-gray-600 min-w-[140px]">
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
                  <div className="px-3 py-1 border-b border-gray-300 text-xs font-semibold flex items-center gap-2" style={{ background: "#e8f5e9", color: "#1b5e20" }}>
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
                  <div className="px-3 py-1 border-b border-gray-300 text-xs font-semibold flex items-center gap-2" style={{ background: "#e8f5e9", color: "#1b5e20" }}>
                    <Icon name="BarChart2" size={12} />
                    Предпросмотр — как будет в Excel
                  </div>
                  <div className="p-4">
                    <Chart1HorizontalBars years={selected.years} rows={selected.chart1Rows} />
                  </div>
                </div>
              </>
            )}

            {activeTab === "chart2" && (
              <>
                <div className="bg-white border border-gray-400 shadow-sm">
                  <div className="px-3 py-1 border-b border-gray-300 text-xs font-semibold flex items-center gap-2" style={{ background: "#e8f5e9", color: "#1b5e20" }}>
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
                  <div className="px-3 py-1 border-b border-gray-300 text-xs font-semibold flex items-center gap-2" style={{ background: "#e8f5e9", color: "#1b5e20" }}>
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
