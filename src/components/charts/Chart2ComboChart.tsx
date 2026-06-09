import React, { useState } from "react";
import {
  ComposedChart, Area, Bar, Line,
  XAxis, YAxis, CartesianGrid, Tooltip,
  ResponsiveContainer, ReferenceLine,
} from "recharts";
import { Chart2Data } from "@/types/oilfield";

interface Props {
  data: Chart2Data;
  fontSize?: number;
  onFontSizeChange?: (v: number) => void;
  chartTitle?: string;
  onChartTitleChange?: (v: string) => void;
  leftAxisLabel?: string;
  onLeftAxisLabelChange?: (v: string) => void;
  rightAxisLabel?: string;
  onRightAxisLabelChange?: (v: string) => void;
}

// Мини-редактор текста
const InlineEditor = ({ value, onSave, onCancel }: { value: string; onSave: (v: string) => void; onCancel: () => void }) => {
  const [val, setVal] = useState(value);
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center" style={{ background: "rgba(0,0,0,0.3)" }}
      onClick={onCancel}>
      <div className="bg-white rounded-lg shadow-xl p-4 flex flex-col gap-3" style={{ minWidth: 320 }}
        onClick={e => e.stopPropagation()}>
        <div className="text-sm font-semibold text-gray-700">Редактировать текст</div>
        <input autoFocus
          className="border border-gray-300 rounded px-3 py-1.5 text-sm outline-none focus:border-green-500"
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

export default function Chart2ComboChart({
  data,
  fontSize: fsProp,
  onFontSizeChange,
  chartTitle: titleProp,
  onChartTitleChange,
  leftAxisLabel: leftProp,
  onLeftAxisLabelChange,
  rightAxisLabel: rightProp,
  onRightAxisLabelChange,
}: Props) {
  const [localFs, setLocalFs] = useState(12);
  const [localTitle, setLocalTitle] = useState("Динамика показателей разработки");
  const [localLeft, setLocalLeft] = useState("Добыча, тыс.т; Закачка, тыс.м³");
  const [localRight, setLocalRight] = useState("Фонд скв., шт; Компенсация, %");
  const [editing, setEditing] = useState<null | "title" | "left" | "right">(null);

  const fs = fsProp ?? localFs;
  const title = titleProp ?? localTitle;
  const leftLabel = leftProp ?? localLeft;
  const rightLabel = rightProp ?? localRight;

  const setFs = (v: number) => { if (onFontSizeChange) onFontSizeChange(v); else setLocalFs(v); };
  const setTitle = (v: string) => { if (onChartTitleChange) onChartTitleChange(v); else setLocalTitle(v); };
  const setLeft = (v: string) => { if (onLeftAxisLabelChange) onLeftAxisLabelChange(v); else setLocalLeft(v); };
  const setRight = (v: string) => { if (onRightAxisLabelChange) onRightAxisLabelChange(v); else setLocalRight(v); };

  const chartData = data.years.map((year, i) => ({
    year: String(year),
    zakachka: data.zakachka[i] ?? 0,
    liquid: data.liquid[i] ?? 0,
    oil: data.oil[i] ?? 0,
    fond_dob: data.fond_dob[i] ?? 0,
    fond_nag: data.fond_nag[i] ?? 0,
    compensation: data.compensation[i] ?? 0,
  }));

  const SERIES = [
    { key: "zakachka", label: "Закачка, тыс.м³",          color: "#FFFF00", stroke: "#b8b800", type: "area" as const },
    { key: "liquid",   label: "Добыча жидкости, тыс.т",    color: "#ADD8E6", stroke: "#6cb4cc", type: "area" as const },
    { key: "oil",      label: "Добыча нефти, тыс.т",       color: "#C0504D", stroke: "#a03030", type: "area" as const },
    { key: "fond_dob", label: "Фонд добыв. скв., ед",      color: "#8B4513", stroke: "#5a2d0c", type: "bar"  as const },
    { key: "fond_nag", label: "Фонд нагн. скв., ед",       color: "none",    stroke: "#00BFFF", type: "bar"  as const },
    { key: "compensation", label: "Компенсация тек., %",   color: "#00008B", stroke: "#00008B", type: "line" as const },
  ];

  const leftMargin = 20 + fs * 5;
  const rightMargin = 20 + fs * 5;

  return (
    <div className="bg-white rounded border border-gray-300 p-3 select-none">

      {/* Редактор */}
      {editing === "title" && <InlineEditor value={title} onSave={v => { setTitle(v); setEditing(null); }} onCancel={() => setEditing(null)} />}
      {editing === "left"  && <InlineEditor value={leftLabel} onSave={v => { setLeft(v); setEditing(null); }} onCancel={() => setEditing(null)} />}
      {editing === "right" && <InlineEditor value={rightLabel} onSave={v => { setRight(v); setEditing(null); }} onCancel={() => setEditing(null)} />}

      {/* Заголовок */}
      <div className="font-bold text-center text-gray-800 mb-1 cursor-pointer hover:bg-yellow-50 rounded px-1 transition-colors"
        style={{ fontSize: fs + 2 }} title="Нажми чтобы изменить" onClick={() => setEditing("title")}>
        {title}
      </div>

      {/* Настройка шрифта */}
      <div className="flex items-center gap-2 mb-2 justify-end flex-wrap">
        <span className="text-xs text-gray-400">Шрифт:</span>
        <button onClick={() => setFs(Math.max(8, fs - 1))} className="w-5 h-5 flex items-center justify-center text-xs rounded border border-gray-300 hover:bg-gray-100 font-bold">−</button>
        <span className="text-xs text-gray-600 w-5 text-center">{fs}</span>
        <button onClick={() => setFs(Math.min(22, fs + 1))} className="w-5 h-5 flex items-center justify-center text-xs rounded border border-gray-300 hover:bg-gray-100 font-bold">+</button>
        <span className="text-xs text-gray-400 ml-2">← Нажми на название оси чтобы изменить</span>
      </div>

      {/* Подписи осей — кликабельные */}
      <div className="flex justify-between mb-1 px-1">
        <span className="text-xs cursor-pointer hover:bg-yellow-50 rounded px-1 transition-colors"
          style={{ fontSize: fs - 1, color: "#555" }} title="Нажми чтобы изменить"
          onClick={() => setEditing("left")}>
          ← {leftLabel}
        </span>
        <span className="text-xs cursor-pointer hover:bg-yellow-50 rounded px-1 transition-colors"
          style={{ fontSize: fs - 1, color: "#555" }} title="Нажми чтобы изменить"
          onClick={() => setEditing("right")}>
          {rightLabel} →
        </span>
      </div>

      <ResponsiveContainer width="100%" height={320}>
        <ComposedChart data={chartData} margin={{ top: 8, right: rightMargin, left: leftMargin, bottom: 4 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="#e0e0e0" />
          <XAxis dataKey="year" tick={{ fontSize: fs, fill: "#555" }} axisLine={{ stroke: "#ccc" }} tickLine={false} />

          {/* Левая ось */}
          <YAxis yAxisId="left" orientation="left"
            tick={{ fontSize: fs - 1, fill: "#555" }} axisLine={{ stroke: "#ccc" }} tickLine={false} />

          {/* Правая ось */}
          <YAxis yAxisId="right" orientation="right"
            tick={{ fontSize: fs - 1, fill: "#555" }} axisLine={{ stroke: "#ccc" }} tickLine={false} />

          <Tooltip contentStyle={{ fontSize: fs, background: "#fff", border: "1px solid #ccc" }} />
          <ReferenceLine yAxisId="left" y={0} stroke="#aaa" strokeWidth={1} />

          {/* Закачка — отдельная Area без stackId */}
          <Area yAxisId="left" type="linear" dataKey="zakachka" name="Закачка, тыс.м³"
            fill="#FFFF00" stroke="#b8b800" strokeWidth={1} fillOpacity={0.7} />

          {/* Жидкость и нефть — стек */}
          <Area yAxisId="left" type="linear" dataKey="liquid" name="Добыча жидкости, тыс.т"
            fill="#ADD8E6" stroke="#6cb4cc" strokeWidth={1} fillOpacity={0.65} stackId="oil" />
          <Area yAxisId="left" type="linear" dataKey="oil" name="Добыча нефти, тыс.т"
            fill="#C0504D" stroke="#a03030" strokeWidth={1} fillOpacity={0.85} stackId="oil" />

          {/* Фонды */}
          <Bar yAxisId="right" dataKey="fond_dob" name="Фонд добыв. скв., ед"
            fill="#8B4513" stroke="#5a2d0c" strokeWidth={1} barSize={18} fillOpacity={0.9} />
          <Bar yAxisId="right" dataKey="fond_nag" name="Фонд нагн. скв., ед"
            fill="transparent" stroke="#00BFFF" strokeWidth={2} barSize={26} fillOpacity={0} />

          {/* Компенсация */}
          <Line yAxisId="right" type="linear" dataKey="compensation" name="Компенсация тек., %"
            stroke="#00008B" strokeWidth={2} dot={{ fill: "#00008B", r: 3 }} activeDot={{ r: 5 }} />
        </ComposedChart>
      </ResponsiveContainer>

      {/* Легенда */}
      <div className="flex flex-wrap gap-x-4 gap-y-1 mt-2 justify-center">
        {SERIES.map(s => (
          <div key={s.key} className="flex items-center gap-1" style={{ fontSize: fs - 1, color: "#555" }}>
            <div className="w-5 h-3 rounded-sm border"
              style={{
                background: s.color === "none" ? "transparent" : s.color,
                borderColor: s.stroke,
                borderWidth: s.color === "none" ? 2 : 1,
              }} />
            {s.label}
          </div>
        ))}
      </div>
    </div>
  );
}
