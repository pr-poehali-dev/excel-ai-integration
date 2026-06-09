import React, { useState } from "react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
  ReferenceLine, Cell, ResponsiveContainer, LabelList,
} from "recharts";
import { Chart1Row } from "@/types/oilfield";

interface Props {
  years: number[];
  rows: Chart1Row[];
  labels?: string[];          // кастомные подписи показателей
  onLabelsChange?: (labels: string[]) => void;
  fontSize?: number;
  onFontSizeChange?: (size: number) => void;
  chartTitle?: string;
  onChartTitleChange?: (t: string) => void;
}

const POSITIVE_COLOR = "#4472C4";
const NEGATIVE_COLOR = "#808080";

// SVG-тик с переносом строк и кликом для редактирования
const CustomYTick = (props: {
  x?: number; y?: number; payload?: { value: string };
  fontSize: number;
  label: string;
  onEdit: () => void;
}) => {
  const { x = 0, y = 0, fontSize, label, onEdit } = props;
  if (!label) return null;

  const words = label.split(" ");
  const maxChars = Math.max(14, Math.floor(160 / (fontSize * 0.65)));
  const lines: string[] = [];
  let cur = "";
  for (const w of words) {
    if (cur.length + w.length + 1 > maxChars && cur.length > 0) {
      lines.push(cur); cur = w;
    } else {
      cur = cur ? cur + " " + w : w;
    }
  }
  if (cur) lines.push(cur);

  const lineH = fontSize + 2;
  const totalH = (lines.length - 1) * lineH;

  return (
    <g style={{ cursor: "pointer" }} onClick={onEdit}>
      <rect x={x - 164} y={y - totalH / 2 - lineH / 2 - 2} width={158} height={totalH + lineH + 4} fill="transparent" />
      {lines.map((line, i) => (
        <text key={i} x={x - 6} y={y - totalH / 2 + i * lineH}
          textAnchor="end" dominantBaseline="middle"
          style={{ fontSize, fill: "#444", fontFamily: "Calibri, Arial, sans-serif" }}
        >{line}</text>
      ))}
    </g>
  );
};

// Inline-редактор подписи
const LabelEditor = ({ value, onSave, onCancel }: { value: string; onSave: (v: string) => void; onCancel: () => void }) => {
  const [val, setVal] = useState(value);
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center" style={{ background: "rgba(0,0,0,0.3)" }}
      onClick={onCancel}>
      <div className="bg-white rounded-lg shadow-xl p-4 flex flex-col gap-3" style={{ minWidth: 320 }}
        onClick={e => e.stopPropagation()}>
        <div className="text-sm font-semibold text-gray-700">Редактировать подпись</div>
        <input
          autoFocus className="border border-gray-300 rounded px-3 py-1.5 text-sm outline-none focus:border-green-500"
          value={val} onChange={e => setVal(e.target.value)}
          onKeyDown={e => { if (e.key === "Enter") onSave(val); if (e.key === "Escape") onCancel(); }}
        />
        <div className="flex gap-2 justify-end">
          <button onClick={onCancel} className="px-3 py-1 text-xs rounded border border-gray-300 text-gray-600 hover:bg-gray-50">Отмена</button>
          <button onClick={() => onSave(val)} className="px-3 py-1 text-xs rounded font-semibold text-white" style={{ background: "#217346" }}>Сохранить</button>
        </div>
      </div>
    </div>
  );
};

export default function Chart1HorizontalBars({
  years, rows,
  labels: labelsProp,
  onLabelsChange,
  fontSize: fontSizeProp = 11,
  onFontSizeChange,
  chartTitle: chartTitleProp,
  onChartTitleChange,
}: Props) {
  const [editingIdx, setEditingIdx] = useState<number | null>(null);
  const [editingTitle, setEditingTitle] = useState(false);
  const [localLabels, setLocalLabels] = useState<string[]>([]);
  const [localFontSize, setLocalFontSize] = useState(fontSizeProp);
  const [localTitle, setLocalTitle] = useState("Анализ изменений показателей по годам, %");

  const labels = labelsProp ?? (localLabels.length === rows.length ? localLabels : rows.map(r => r.indicator));
  const fontSize = fontSizeProp ?? localFontSize;
  const chartTitle = chartTitleProp ?? localTitle;

  const setLabel = (idx: number, val: string) => {
    const next = labels.map((l, i) => i === idx ? val : l);
    if (onLabelsChange) onLabelsChange(next);
    else setLocalLabels(next);
  };

  const setFontSize = (v: number) => {
    if (onFontSizeChange) onFontSizeChange(v);
    else setLocalFontSize(v);
  };

  const setTitle = (v: string) => {
    if (onChartTitleChange) onChartTitleChange(v);
    else setLocalTitle(v);
  };

  // Ширина колонки подписей зависит от шрифта
  const LABEL_WIDTH = Math.max(120, fontSize * 14);
  const BAR_HEIGHT = Math.max(16, fontSize + 4);
  const CHART_HEIGHT = rows.length * (BAR_HEIGHT + 14) + 30;
  const BAR_SIZE = BAR_HEIGHT - 2;

  return (
    <div className="bg-white rounded border border-gray-300 p-3 select-none">
      {/* Заголовок — кликабельный */}
      {editingTitle ? (
        <LabelEditor value={chartTitle} onSave={v => { setTitle(v); setEditingTitle(false); }} onCancel={() => setEditingTitle(false)} />
      ) : null}
      <div
        className="font-bold text-center text-gray-800 mb-1 cursor-pointer hover:bg-yellow-50 rounded px-1 transition-colors"
        style={{ fontSize: fontSize + 2 }}
        title="Нажми чтобы изменить название"
        onClick={() => setEditingTitle(true)}
      >
        {chartTitle}
      </div>

      {/* Панель настройки шрифта */}
      <div className="flex items-center gap-2 mb-2 justify-end">
        <span className="text-xs text-gray-400">Шрифт:</span>
        <button onClick={() => setFontSize(Math.max(8, fontSize - 1))}
          className="w-5 h-5 flex items-center justify-center text-xs rounded border border-gray-300 hover:bg-gray-100 font-bold">−</button>
        <span className="text-xs text-gray-600 w-5 text-center">{fontSize}</span>
        <button onClick={() => setFontSize(Math.min(20, fontSize + 1))}
          className="w-5 h-5 flex items-center justify-center text-xs rounded border border-gray-300 hover:bg-gray-100 font-bold">+</button>
        <span className="text-xs text-gray-400 ml-2">← Нажми на подпись чтобы изменить</span>
      </div>

      <div className="flex overflow-x-auto">
        {/* Первый год — с кликабельными подписями YAxis */}
        {years.length > 0 && (() => {
          const yIdx = 0;
          const chartData = rows.map((row, ri) => ({
            name: labels[ri] ?? row.indicator,
            value: row.values[yIdx] ?? 0,
          }));
          return (
            <div key={years[yIdx]} className="flex-shrink-0" style={{ width: LABEL_WIDTH + 120 }}>
              <div className="font-bold text-gray-700 mb-1 underline"
                style={{ fontSize, textAlign: "center", paddingLeft: LABEL_WIDTH }}>
                {years[yIdx]}
              </div>
              <ResponsiveContainer width="100%" height={CHART_HEIGHT}>
                <BarChart data={chartData} layout="vertical"
                  margin={{ top: 2, right: 30, left: LABEL_WIDTH, bottom: 2 }}
                  barSize={BAR_SIZE}>
                  <CartesianGrid strokeDasharray="2 2" horizontal={false} stroke="#e0e0e0" />
                  <XAxis type="number" domain={["auto", "auto"]}
                    tick={{ fontSize: fontSize - 2, fill: "#555" }} tickLine={false} axisLine={{ stroke: "#ccc" }} />
                  <YAxis type="category" dataKey="name" width={LABEL_WIDTH} axisLine={false} tickLine={false}
                    tick={(props) => {
                      const idx = chartData.findIndex(d => d.name === props.payload?.value);
                      return <CustomYTick {...props} fontSize={fontSize} label={labels[idx] ?? props.payload?.value ?? ""}
                        onEdit={() => setEditingIdx(idx)} />;
                    }}
                  />
                  <Tooltip formatter={(v: number) => [`${v}%`, ""]}
                    contentStyle={{ fontSize, background: "#fff", border: "1px solid #ccc" }} />
                  <ReferenceLine x={0} stroke="#999" strokeWidth={1} />
                  <Bar dataKey="value" radius={1}>
                    <LabelList dataKey="value" position="right"
                      style={{ fontSize: fontSize - 1, fill: "#333" }} formatter={(v: number) => `${v}%`} />
                    {chartData.map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={entry.value >= 0 ? POSITIVE_COLOR : NEGATIVE_COLOR} />
                    ))}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          );
        })()}

        {/* Остальные годы — без подписей */}
        {years.slice(1).map((year, i) => {
          const yIdx = i + 1;
          const chartData = rows.map((row, ri) => ({
            name: labels[ri] ?? row.indicator,
            value: row.values[yIdx] ?? 0,
          }));
          return (
            <div key={year} className="flex-shrink-0" style={{ width: 120 }}>
              <div className="font-bold text-center text-gray-700 mb-1 underline" style={{ fontSize }}>{year}</div>
              <ResponsiveContainer width="100%" height={CHART_HEIGHT}>
                <BarChart data={chartData} layout="vertical"
                  margin={{ top: 2, right: 30, left: 4, bottom: 2 }} barSize={BAR_SIZE}>
                  <CartesianGrid strokeDasharray="2 2" horizontal={false} stroke="#e0e0e0" />
                  <XAxis type="number" domain={["auto", "auto"]}
                    tick={{ fontSize: fontSize - 2, fill: "#555" }} tickLine={false} axisLine={{ stroke: "#ccc" }} />
                  <YAxis type="category" dataKey="name" width={0} tick={false} axisLine={false} tickLine={false} />
                  <Tooltip formatter={(v: number) => [`${v}%`, ""]}
                    contentStyle={{ fontSize, background: "#fff", border: "1px solid #ccc" }} />
                  <ReferenceLine x={0} stroke="#999" strokeWidth={1} />
                  <Bar dataKey="value" radius={1}>
                    <LabelList dataKey="value" position="right"
                      style={{ fontSize: fontSize - 1, fill: "#333" }} formatter={(v: number) => `${v}%`} />
                    {chartData.map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={entry.value >= 0 ? POSITIVE_COLOR : NEGATIVE_COLOR} />
                    ))}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          );
        })}
      </div>

      <div className="flex gap-4 mt-2 justify-center">
        <div className="flex items-center gap-1 text-gray-600" style={{ fontSize: fontSize - 1 }}>
          <div className="w-4 h-3 rounded-sm" style={{ background: POSITIVE_COLOR }} />Рост (≥0%)
        </div>
        <div className="flex items-center gap-1 text-gray-600" style={{ fontSize: fontSize - 1 }}>
          <div className="w-4 h-3 rounded-sm" style={{ background: NEGATIVE_COLOR }} />Снижение (&lt;0%)
        </div>
      </div>

      {/* Редактор подписи */}
      {editingIdx !== null && (
        <LabelEditor
          value={labels[editingIdx] ?? ""}
          onSave={v => { setLabel(editingIdx, v); setEditingIdx(null); }}
          onCancel={() => setEditingIdx(null)}
        />
      )}
    </div>
  );
}
