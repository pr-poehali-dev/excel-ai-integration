import React, { useState, useRef, useCallback, useLayoutEffect } from "react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
  ReferenceLine, Cell, LabelList, ResponsiveContainer,
} from "recharts";
import { Chart1Row } from "@/types/oilfield";

interface Props {
  years: number[];
  rows: Chart1Row[];
}

const POSITIVE_COLOR = "#4472C4";
const NEGATIVE_COLOR = "#808080";

function EditDialog({ value, onSave, onCancel }: { value: string; onSave: (v: string) => void; onCancel: () => void }) {
  const [val, setVal] = useState(value);
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center" style={{ background: "rgba(0,0,0,0.35)" }} onClick={onCancel}>
      <div className="bg-white rounded-xl shadow-2xl p-5 flex flex-col gap-3 w-80" onClick={e => e.stopPropagation()}>
        <div className="text-sm font-semibold text-gray-700">Изменить текст</div>
        <input autoFocus className="border border-gray-300 rounded px-3 py-2 text-sm outline-none focus:border-blue-400"
          value={val} onChange={e => setVal(e.target.value)}
          onKeyDown={e => { if (e.key === "Enter") onSave(val); if (e.key === "Escape") onCancel(); }} />
        <div className="flex gap-2 justify-end">
          <button onClick={onCancel} className="px-3 py-1 text-xs rounded border border-gray-200 text-gray-500 hover:bg-gray-50">Отмена</button>
          <button onClick={() => onSave(val)} className="px-4 py-1 text-xs rounded font-bold text-white bg-blue-600 hover:bg-blue-700">ОК</button>
        </div>
      </div>
    </div>
  );
}

// Draggable — позиция относительно родителя; по умолчанию центрируется, после drag — абсолютно
function DraggableLabel({ text, fontSize, onEdit }: { text: string; fontSize: number; onEdit: () => void }) {
  const [offset, setOffset] = useState({ x: 0, y: 4 }); // смещение от центра
  const dragging = useRef(false);
  const mouseStart = useRef({ x: 0, y: 0 });
  const offsetStart = useRef({ x: 0, y: 0 });

  const onMouseDown = (e: React.MouseEvent) => {
    dragging.current = true;
    mouseStart.current = { x: e.clientX, y: e.clientY };
    offsetStart.current = { ...offset };
    e.preventDefault();
    e.stopPropagation();
  };

  useLayoutEffect(() => {
    const onMove = (e: MouseEvent) => {
      if (!dragging.current) return;
      setOffset({
        x: offsetStart.current.x + (e.clientX - mouseStart.current.x),
        y: offsetStart.current.y + (e.clientY - mouseStart.current.y),
      });
    };
    const onUp = () => { dragging.current = false; };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    return () => { window.removeEventListener("mousemove", onMove); window.removeEventListener("mouseup", onUp); };
  }, []);

  return (
    <div
      style={{
        position: "absolute",
        left: `calc(50% + ${offset.x}px)`,
        top: offset.y,
        transform: "translateX(-50%)",
        cursor: "move",
        userSelect: "none",
        zIndex: 10,
        background: "rgba(255,255,255,0.95)",
        border: "1px dashed #aaa",
        borderRadius: 4,
        padding: "2px 10px",
        fontSize: fontSize + 2,
        fontWeight: 700,
        color: "#222",
        boxShadow: "0 1px 4px rgba(0,0,0,0.10)",
        whiteSpace: "nowrap",
      }}
      onMouseDown={onMouseDown}
      onDoubleClick={onEdit}
      title="Тяни чтобы переместить · двойной клик — изменить"
    >
      {text}
    </div>
  );
}

export default function Chart1HorizontalBars({ years, rows }: Props) {
  const [fontSize, setFontSize] = useState(11);
  const [labels, setLabels] = useState<string[]>(() => rows.map(r => r.indicator));
  const [chartTitle, setChartTitle] = useState("Анализ изменений показателей по годам, %");
  const [editingIdx, setEditingIdx] = useState<number | null>(null); // -1 = заголовок

  // Сброс labels при смене файла (изменение кол-ва строк)
  const prevLen = useRef(rows.length);
  if (rows.length !== prevLen.current) {
    prevLen.current = rows.length;
     
    setLabels(rows.map(r => r.indicator));
  }

  const saveLabel = useCallback((idx: number, val: string) => {
    setLabels(prev => prev.map((l, i) => i === idx ? val : l));
  }, []);

  const rowH = Math.max(20, fontSize * 2 + 6);
  const chartH = rows.length * rowH + 24;
  const barSize = Math.max(10, rowH - 6);
  const labelColW = Math.max(140, fontSize * 13);

  const renderChart = (yIdx: number) => {
    const data = rows.map((row, ri) => ({ idx: ri, value: row.values[yIdx] ?? 0 }));
    return (
      <BarChart data={data} layout="vertical"
        margin={{ top: 0, right: 34, left: 0, bottom: 0 }} barSize={barSize}>
        <CartesianGrid strokeDasharray="2 2" horizontal={false} stroke="#e8e8e8" />
        <XAxis type="number" domain={["auto", "auto"]}
          tick={{ fontSize: Math.max(8, fontSize - 2), fill: "#666" }}
          tickLine={false} axisLine={{ stroke: "#ccc" }} height={20} />
        <YAxis type="category" dataKey="idx" width={0} tick={false} axisLine={false} tickLine={false} />
        <Tooltip
          formatter={(v: number, _n: string, p) => [`${v}%`, labels[p.payload?.idx] ?? ""]}
          contentStyle={{ fontSize: Math.max(9, fontSize - 1), background: "#fff", border: "1px solid #ccc" }}
        />
        <ReferenceLine x={0} stroke="#bbb" strokeWidth={1} />
        <Bar dataKey="value" radius={1}>
          <LabelList dataKey="value" position="right"
            style={{ fontSize: Math.max(8, fontSize - 1), fill: "#333" }}
            formatter={(v: number) => `${v}%`} />
          {data.map((entry, i) => (
            <Cell key={i} fill={entry.value >= 0 ? POSITIVE_COLOR : NEGATIVE_COLOR} />
          ))}
        </Bar>
      </BarChart>
    );
  };

  return (
    <div className="bg-white rounded border border-gray-300 p-3 relative" style={{ userSelect: "none" }}>

      {/* Draggable заголовок */}
      <DraggableLabel text={chartTitle} fontSize={fontSize} onEdit={() => setEditingIdx(-1)} />

      {/* Кнопки шрифта и подсказка */}
      <div className="flex items-center gap-1.5 justify-end" style={{ marginTop: fontSize + 14 }}>
        <span style={{ fontSize: 10, color: "#aaa" }}>Шрифт:</span>
        <button onClick={() => setFontSize(f => Math.max(8, f - 1))}
          className="w-5 h-5 flex items-center justify-center rounded border border-gray-300 hover:bg-gray-100 text-xs font-bold">−</button>
        <span style={{ fontSize: 11, color: "#555", minWidth: 16, textAlign: "center" }}>{fontSize}</span>
        <button onClick={() => setFontSize(f => Math.min(22, f + 1))}
          className="w-5 h-5 flex items-center justify-center rounded border border-gray-300 hover:bg-gray-100 text-xs font-bold">+</button>
        <span style={{ fontSize: 10, color: "#ccc", marginLeft: 6 }}>клик на подпись — изменить</span>
      </div>

      {/* Главная область: подписи HTML слева + графики справа */}
      <div className="flex mt-1" style={{ overflowX: "auto" }}>

        {/* HTML-колонка подписей — никогда не обрезается */}
        <div className="flex-shrink-0" style={{ width: labelColW }}>
          {/* Пустое место под XAxis каждого графика */}
          <div style={{ height: 20 }} />
          {rows.map((row, ri) => {
            const lbl = labels[ri] ?? row.indicator;
            return (
              <div
                key={ri}
                style={{
                  height: rowH,
                  display: "flex", alignItems: "center", justifyContent: "flex-end",
                  paddingRight: 8, paddingLeft: 4,
                  fontSize, color: "#333", fontFamily: "Calibri, Arial, sans-serif",
                  lineHeight: 1.25, textAlign: "right",
                  cursor: "pointer", borderRadius: 3,
                  wordBreak: "break-word",
                }}
                title="Кликни чтобы изменить подпись"
                onClick={() => setEditingIdx(ri)}
                onMouseEnter={e => (e.currentTarget.style.background = "#fffbe6")}
                onMouseLeave={e => (e.currentTarget.style.background = "transparent")}
              >
                {lbl}
              </div>
            );
          })}
        </div>

        {/* Графики по годам */}
        {years.map((year, yIdx) => (
          <div key={year} className="flex-shrink-0" style={{ width: 125 }}>
            <div style={{
              height: 20, fontSize, fontWeight: 700, color: "#333",
              textAlign: "center", textDecoration: "underline",
              display: "flex", alignItems: "center", justifyContent: "center",
            }}>
              {year}
            </div>
            <div style={{ width: 125, height: chartH - 20 }}>
              <ResponsiveContainer width="100%" height="100%">
                {renderChart(yIdx)}
              </ResponsiveContainer>
            </div>
          </div>
        ))}
      </div>

      {/* Легенда */}
      <div className="flex gap-5 mt-2 justify-center">
        {[{ color: POSITIVE_COLOR, label: "Рост (≥0%)" }, { color: NEGATIVE_COLOR, label: "Снижение (<0%)" }].map(({ color, label }) => (
          <div key={label} className="flex items-center gap-1.5" style={{ fontSize: Math.max(9, fontSize - 1), color: "#555" }}>
            <div style={{ width: 16, height: 10, background: color, borderRadius: 2 }} />
            {label}
          </div>
        ))}
      </div>

      {/* Диалоги */}
      {editingIdx === -1 && (
        <EditDialog value={chartTitle} onSave={v => { setChartTitle(v); setEditingIdx(null); }} onCancel={() => setEditingIdx(null)} />
      )}
      {editingIdx !== null && editingIdx >= 0 && (
        <EditDialog
          value={labels[editingIdx] ?? rows[editingIdx]?.indicator ?? ""}
          onSave={v => { saveLabel(editingIdx, v); setEditingIdx(null); }}
          onCancel={() => setEditingIdx(null)}
        />
      )}
    </div>
  );
}
