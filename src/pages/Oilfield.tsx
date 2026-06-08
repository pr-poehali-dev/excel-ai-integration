import React, { useRef, useState, useCallback } from "react";
import * as XLSX from "xlsx";
import { useNavigate } from "react-router-dom";
import Icon from "@/components/ui/icon";
import { Button } from "@/components/ui/button";
import { Tabs, TabsList, TabsTrigger, TabsContent } from "@/components/ui/tabs";
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

          let headerRowIdx = -1;
          let yearsFound: number[] = [];
          for (let ri = 0; ri < Math.min(rows.length, 10); ri++) {
            const row = rows[ri];
            const numCols = row.filter(
              (v) => typeof v === "number" && v >= 2000 && v <= 2050
            );
            if (numCols.length >= 2) {
              headerRowIdx = ri;
              yearsFound = row
                .filter((v): v is number => typeof v === "number" && v >= 2000 && v <= 2050)
                .slice(0, 10);
              break;
            }
          }

          if (headerRowIdx >= 0 && yearsFound.length > 0) {
            field.years = yearsFound;
            const dataRows = rows.slice(headerRowIdx + 1);
            const chart1Rows = CHART1_INDICATORS.map((indicator, iIdx) => {
              const dataRow = dataRows[iIdx];
              const values: (number | null)[] = yearsFound.map((_, yIdx) => {
                const cell = dataRow?.[yIdx + 1];
                if (typeof cell === "number") return cell;
                if (typeof cell === "string") {
                  const parsed = parseFloat(cell.replace(",", ".").replace("%", ""));
                  return isNaN(parsed) ? null : parsed;
                }
                return null;
              });
              return { indicator, values };
            });
            field.chart1Rows = chart1Rows;
          }

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
  const [aiError, setAiError] = useState<string | null>(null);
  const [pasteOpen, setPasteOpen] = useState(false);
  const [pasteText, setPasteText] = useState("");

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
    try {
      const parsed = await parseXlsxLocally(file);
      if (parsed.length > 0) {
        setFields(parsed);
        setSelectedIdx(0);
      }
    } catch {
      alert("Ошибка чтения файла. Проверьте формат xlsx.");
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

  // Извлекает все JSON-объекты из текста и возвращает самый "полезный"
  const extractJsonFromText = (text: string): Record<string, unknown> | null => {
    const candidates: Record<string, unknown>[] = [];
    let i = 0;
    while (i < text.length) {
      if (text[i] !== '{') { i++; continue; }
      // Находим закрывающую скобку с учётом вложенности
      let depth = 0;
      let j = i;
      while (j < text.length) {
        if (text[j] === '{') depth++;
        else if (text[j] === '}') depth--;
        if (depth === 0) break;
        j++;
      }
      if (depth === 0) {
        const candidate = text.slice(i, j + 1);
        try {
          const obj = JSON.parse(candidate) as Record<string, unknown>;
          candidates.push(obj);
        } catch { /* не JSON */ }
      }
      i++;
    }
    if (candidates.length === 0) return null;
    // Предпочитаем объект с годами или analysis/rows, иначе самый большой
    return (
      candidates.find(o => Object.keys(o).some(k => /^\d{4}$/.test(k))) ??
      candidates.find(o => o.analysis) ??
      candidates.find(o => o.rows) ??
      candidates.reduce((a, b) => JSON.stringify(a).length >= JSON.stringify(b).length ? a : b)
    );
  };

  // Применяет данные из ответа ИИ к таблице
  const applyAiJson = (raw: string, currentField: typeof fields[0], currentIdx: number) => {
    type AnyObj = Record<string, unknown>;

    const parsed = extractJsonFromText(raw);
    if (!parsed) throw new Error("Не найден JSON с данными в тексте");

    console.log("[APPLY AI] keys:", Object.keys(parsed));

    // Формат 1: { rows: [{indicator, values}] }
    if (parsed.rows && Array.isArray(parsed.rows)) {
      const rows = parsed.rows as { indicator: string; values: (number | null)[] }[];
      const updatedRows = currentField.chart1Rows.map((origRow, i) => {
        const aiRow = rows.find(r => r.indicator === origRow.indicator) ?? rows[i];
        if (!aiRow) return origRow;
        const values = currentField.years.map((_, yi) => {
          const v = aiRow.values[yi];
          if (v === null || v === undefined) return origRow.values[yi];
          const num = Number(v);
          return isNaN(num) ? origRow.values[yi] : num;
        });
        return { ...origRow, values };
      });
      setFields(prev => prev.map((f, i) => i === currentIdx ? { ...f, chart1Rows: updatedRows } : f));
      return;
    }

    // Формат 2: { analysis: { "2021": {...} } } или { "2021": {...} }
    const analysisRoot = (parsed.analysis ?? parsed) as Record<string, unknown>;
    const yearKeys = Object.keys(analysisRoot).filter(k => /^\d{4}$/.test(k)).sort();
    if (yearKeys.length === 0) throw new Error(`Годы не найдены. Ключи в JSON: ${Object.keys(parsed).join(", ")}`);

    const aiYears = yearKeys.map(Number);
    const analysis = analysisRoot as Record<string, Record<string, AnyObj>>;

    const PCT_FIELDS = ["откл%", "откл_%", "отклонение_%", "отклонение%", "процент", "percent", "pct", "delta_%", "delta%"];

    const extractPct = (cell: AnyObj): number | null => {
      for (const key of Object.keys(cell)) {
        if (PCT_FIELDS.includes(key.toLowerCase())) {
          const v = Number(cell[key]);
          if (!isNaN(v)) return Math.round(v * 10) / 10;
        }
      }
      // Берём первое числовое поле — обычно это и есть отклонение
      for (const key of Object.keys(cell)) {
        if (typeof cell[key] === "number") return Math.round((cell[key] as number) * 10) / 10;
      }
      return null;
    };

    let filledCount = 0;
    const updatedRows = currentField.chart1Rows.map((origRow) => {
      const origLow = origRow.indicator.toLowerCase();
      const values = aiYears.map((yr, yi) => {
        const yearData = analysis[String(yr)];
        if (!yearData) return origRow.values[yi] ?? null;

        const indicatorKey =
          Object.keys(yearData).find(k => k.toLowerCase() === origLow) ??
          Object.keys(yearData).find(k => k.toLowerCase().includes(origLow.slice(0, 6))) ??
          Object.keys(yearData).find(k => origLow.includes(k.toLowerCase().slice(0, 6)));

        if (!indicatorKey) return origRow.values[yi] ?? null;
        const pct = extractPct(yearData[indicatorKey] as AnyObj);
        if (pct !== null) filledCount++;
        return pct ?? (origRow.values[yi] ?? null);
      });
      return { ...origRow, values };
    });

    console.log("[APPLY AI] filled cells:", filledCount, "years:", aiYears);
    if (filledCount === 0) throw new Error("JSON найден, но не удалось извлечь числовые отклонения. Убедись что в ответе ИИ есть блок с годами и процентами.");

    setFields(prev => prev.map((f, i) =>
      i === currentIdx ? { ...f, years: aiYears, chart1Rows: updatedRows } : f
    ));
  };



  return (
    <div className="min-h-screen flex flex-col" style={{ background: "#f5f5f5", fontFamily: "Calibri, Arial, sans-serif" }}>
      {/* Шапка в стиле Excel */}
      <div
        className="border-b border-gray-400 flex-shrink-0"
        style={{ background: "#217346" }}
      >
        {/* Ribbon верхний */}
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

          {/* Файл */}
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

          {activeTab === "chart1" && (
            <button
              onClick={() => { setPasteText(""); setPasteOpen(true); }}
              className="flex items-center gap-1.5 px-3 py-1 rounded text-xs font-semibold border transition"
              style={{ background: "rgba(52,211,153,0.2)", borderColor: "rgba(52,211,153,0.6)", color: "#34d399" }}
            >
              <Icon name="Sparkles" size={13} />
              Применить анализ ИИ
            </button>
          )}

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

        {/* Вкладки листов */}
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

      {aiError && (
        <div className="flex items-center gap-2 px-4 py-2 text-xs border-b border-red-300/40" style={{ background: "rgba(220,38,38,0.1)", color: "#fca5a5" }}>
          <Icon name="AlertCircle" size={13} />
          {aiError}
          <button onClick={() => setAiError(null)} className="ml-auto opacity-60 hover:opacity-100"><Icon name="X" size={12} /></button>
        </div>
      )}

      <div className="flex flex-1 overflow-hidden">
        {/* Боковая панель месторождений */}
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
                  onClick={(e) => {
                    e.stopPropagation();
                    setSelectedIdx(idx);
                  }}
                  onChange={(e) =>
                    setFields((prev) =>
                      prev.map((ff, ii) =>
                        ii === idx ? { ...ff, name: e.target.value } : ff
                      )
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
                    onClick={(e) => {
                      e.stopPropagation();
                      removeField(idx);
                    }}
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
          {/* Строка формул как в Excel */}
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
                {/* Таблица */}
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

                {/* График */}
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
                    Таблица данных — динамика показателей
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

      {/* Модалка вставки ответа ИИ */}
      {pasteOpen && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center"
          style={{ background: "rgba(0,0,0,0.5)" }}
          onClick={(e) => { if (e.target === e.currentTarget) setPasteOpen(false); }}
        >
          <div className="bg-white rounded-lg shadow-xl flex flex-col" style={{ width: 560, maxHeight: "80vh" }}>
            <div className="flex items-center justify-between px-4 py-3 border-b border-gray-200">
              <div className="flex items-center gap-2 text-sm font-semibold text-gray-800">
                <Icon name="ClipboardPaste" size={15} />
                Вставить ответ ИИ
              </div>
              <button onClick={() => setPasteOpen(false)} className="text-gray-400 hover:text-gray-600">
                <Icon name="X" size={16} />
              </button>
            </div>
            <div className="px-4 py-2.5 text-xs border-b border-gray-100" style={{ background: "#f0f7f0", color: "#1b5e20" }}>
              <b>Как использовать:</b> в чате DataMind попроси ИИ проанализировать таблицу проект/факт →
              выдели весь ответ ИИ (Ctrl+A в поле чата) → скопируй (Ctrl+C) → вставь сюда (Ctrl+V) → нажми «Применить»
            </div>
            <textarea
              autoFocus
              className="flex-1 p-4 text-xs font-mono resize-none outline-none text-gray-700"
              style={{ minHeight: 280 }}
              placeholder={"Вставь сюда полный текст ответа ИИ...\n\nПрограмма автоматически найдёт в нём годы и процентные отклонения по каждому показателю и обновит таблицу."}
              value={pasteText}
              onChange={e => setPasteText(e.target.value)}
            />
            {pasteText.length > 0 && (() => {
              const found = extractJsonFromText(pasteText);
              const yearKeys = found ? Object.keys(found.analysis as object || found).filter(k => /^\d{4}$/.test(k)) : [];
              return (
                <div className="px-4 py-2 text-xs border-t border-gray-100" style={{
                  background: found && (yearKeys.length > 0 || (found as Record<string,unknown>).rows) ? "#f0fff4" : "#fff8f0",
                  color: found && (yearKeys.length > 0 || (found as Record<string,unknown>).rows) ? "#276749" : "#c05621"
                }}>
                  {!found && "⚠ JSON не найден — убедись что вставил полный ответ ИИ включая фигурные скобки"}
                  {found && yearKeys.length > 0 && `✓ Найдены данные за ${yearKeys.length} лет: ${yearKeys.join(", ")}`}
                  {found && yearKeys.length === 0 && (found as Record<string,unknown>).rows && "✓ Найден формат rows — готово к применению"}
                  {found && yearKeys.length === 0 && !(found as Record<string,unknown>).rows && `⚠ JSON найден, но годы не обнаружены (ключи: ${Object.keys(found).join(", ")})`}
                </div>
              );
            })()}
            <div className="flex items-center gap-2 px-4 py-3 border-t border-gray-200 justify-end">
              <span className="text-xs text-gray-400 flex-1">
                {pasteText.length > 0 ? `${pasteText.length} символов` : ""}
              </span>
              <button
                onClick={() => setPasteOpen(false)}
                className="px-4 py-1.5 text-xs rounded border border-gray-300 text-gray-600 hover:bg-gray-50"
              >
                Отмена
              </button>
              <button
                disabled={pasteText.trim().length === 0}
                onClick={() => {
                  try {
                    applyAiJson(pasteText, fields[selectedIdx], selectedIdx);
                    setAiError(null);
                    setPasteOpen(false);
                  } catch (e) {
                    setAiError(e instanceof Error ? e.message : "Ошибка разбора");
                    setPasteOpen(false);
                  }
                }}
                className="px-4 py-1.5 text-xs rounded font-semibold disabled:opacity-40"
                style={{ background: "#217346", color: "#fff" }}
              >
                Применить к таблице
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}