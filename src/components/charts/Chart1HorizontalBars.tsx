import React from "react";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ReferenceLine,
  Cell,
  ResponsiveContainer,
  LabelList,
} from "recharts";
import { Chart1Row } from "@/types/oilfield";

interface Props {
  years: number[];
  rows: Chart1Row[];
}

const POSITIVE_COLOR = "#4472C4";
const NEGATIVE_COLOR = "#808080";
const CHART_HEIGHT = 290;
const LABEL_WIDTH = 165;

// Кастомный тик YAxis — подпись показателя встроена в ось, выровнена с баром
const CustomYTick = (props: { x?: number; y?: number; payload?: { value: string } }) => {
  const { x = 0, y = 0, payload } = props;
  if (!payload) return null;
  return (
    <foreignObject x={x - LABEL_WIDTH} y={y - 13} width={LABEL_WIDTH - 4} height={26}>
      <div
        xmlns="http://www.w3.org/1999/xhtml"
        style={{
          fontSize: 11,
          color: "#444",
          fontFamily: "Calibri, Arial, sans-serif",
          lineHeight: 1.2,
          textAlign: "right",
          paddingRight: 6,
          height: "100%",
          display: "flex",
          alignItems: "center",
          justifyContent: "flex-end",
          overflow: "hidden",
          whiteSpace: "nowrap",
        }}
        title={payload.value}
      >
        {payload.value}
      </div>
    </foreignObject>
  );
};

export default function Chart1HorizontalBars({ years, rows }: Props) {
  return (
    <div className="bg-white rounded border border-gray-300 p-3">
      <div className="text-sm font-bold text-center text-gray-800 mb-3">
        Анализ изменений показателей по годам, %
      </div>

      <div className="flex overflow-x-auto">

        {/* Первый год — с подписями через YAxis */}
        {years.length > 0 && (() => {
          const yIdx = 0;
          const chartData = rows.map(row => ({
            name: row.indicator,
            value: row.values[yIdx] ?? 0,
          }));
          return (
            <div key={years[yIdx]} className="flex-shrink-0" style={{ width: LABEL_WIDTH + 130 }}>
              <div
                className="text-xs font-bold text-gray-700 mb-1 underline"
                style={{ textAlign: "center", paddingLeft: LABEL_WIDTH }}
              >
                {years[yIdx]}
              </div>
              <ResponsiveContainer width="100%" height={CHART_HEIGHT}>
                <BarChart
                  data={chartData}
                  layout="vertical"
                  margin={{ top: 0, right: 28, left: LABEL_WIDTH, bottom: 0 }}
                  barSize={14}
                >
                  <CartesianGrid strokeDasharray="2 2" horizontal={false} stroke="#e0e0e0" />
                  <XAxis type="number" domain={["auto", "auto"]} tick={{ fontSize: 9, fill: "#555" }} tickLine={false} axisLine={{ stroke: "#ccc" }} />
                  <YAxis
                    type="category"
                    dataKey="name"
                    width={LABEL_WIDTH}
                    tick={<CustomYTick />}
                    axisLine={false}
                    tickLine={false}
                  />
                  <Tooltip formatter={(v: number) => [`${v}%`, ""]} contentStyle={{ fontSize: 11, background: "#fff", border: "1px solid #ccc" }} />
                  <ReferenceLine x={0} stroke="#999" strokeWidth={1} />
                  <Bar dataKey="value" radius={1}>
                    <LabelList dataKey="value" position="right" style={{ fontSize: 9, fill: "#333" }} formatter={(v: number) => `${v}%`} />
                    {chartData.map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={entry.value >= 0 ? POSITIVE_COLOR : NEGATIVE_COLOR} />
                    ))}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          );
        })()}

        {/* Остальные годы — без подписей YAxis */}
        {years.slice(1).map((year, i) => {
          const yIdx = i + 1;
          const chartData = rows.map(row => ({
            name: row.indicator,
            value: row.values[yIdx] ?? 0,
          }));
          return (
            <div key={year} className="flex-shrink-0" style={{ width: 138 }}>
              <div className="text-xs font-bold text-center text-gray-700 mb-1 underline">{year}</div>
              <ResponsiveContainer width="100%" height={CHART_HEIGHT}>
                <BarChart data={chartData} layout="vertical" margin={{ top: 0, right: 28, left: 4, bottom: 0 }} barSize={14}>
                  <CartesianGrid strokeDasharray="2 2" horizontal={false} stroke="#e0e0e0" />
                  <XAxis type="number" domain={["auto", "auto"]} tick={{ fontSize: 9, fill: "#555" }} tickLine={false} axisLine={{ stroke: "#ccc" }} />
                  <YAxis type="category" dataKey="name" width={0} tick={false} axisLine={false} tickLine={false} />
                  <Tooltip formatter={(v: number) => [`${v}%`, ""]} contentStyle={{ fontSize: 11, background: "#fff", border: "1px solid #ccc" }} />
                  <ReferenceLine x={0} stroke="#999" strokeWidth={1} />
                  <Bar dataKey="value" radius={1}>
                    <LabelList dataKey="value" position="right" style={{ fontSize: 9, fill: "#333" }} formatter={(v: number) => `${v}%`} />
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
        <div className="flex items-center gap-1 text-xs text-gray-600">
          <div className="w-4 h-3 rounded-sm" style={{ background: POSITIVE_COLOR }} />
          Рост (≥0%)
        </div>
        <div className="flex items-center gap-1 text-xs text-gray-600">
          <div className="w-4 h-3 rounded-sm" style={{ background: NEGATIVE_COLOR }} />
          Снижение (&lt;0%)
        </div>
      </div>
    </div>
  );
}
