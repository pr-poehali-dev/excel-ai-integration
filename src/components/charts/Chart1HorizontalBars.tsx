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

export default function Chart1HorizontalBars({ years, rows }: Props) {
  return (
    <div className="bg-white rounded border border-gray-300 p-3">
      <div className="text-sm font-bold text-center text-gray-800 mb-3">
        Анализ изменений показателей по годам, %
      </div>
      <div className="flex gap-1 overflow-x-auto">
        {years.map((year, yIdx) => {
          const chartData = rows.map((row) => ({
            name: row.indicator,
            value: row.values[yIdx] ?? 0,
          }));

          return (
            <div key={year} className="flex-shrink-0" style={{ width: 160 }}>
              <div className="text-xs font-bold text-center text-gray-700 mb-1 underline">
                {year}
              </div>
              <ResponsiveContainer width="100%" height={260}>
                <BarChart
                  data={chartData}
                  layout="vertical"
                  margin={{ top: 0, right: 30, left: 4, bottom: 0 }}
                  barSize={14}
                >
                  <CartesianGrid
                    strokeDasharray="2 2"
                    horizontal={false}
                    stroke="#e0e0e0"
                  />
                  <XAxis
                    type="number"
                    domain={["auto", "auto"]}
                    tick={{ fontSize: 9, fill: "#555" }}
                    tickLine={false}
                    axisLine={{ stroke: "#ccc" }}
                  />
                  <YAxis
                    type="category"
                    dataKey="name"
                    width={0}
                    tick={false}
                    axisLine={false}
                    tickLine={false}
                  />
                  <Tooltip
                    formatter={(v: number) => [`${v}%`, ""]}
                    contentStyle={{
                      fontSize: 11,
                      background: "#fff",
                      border: "1px solid #ccc",
                    }}
                  />
                  <ReferenceLine x={0} stroke="#999" strokeWidth={1} />
                  <Bar dataKey="value" radius={1}>
                    <LabelList
                      dataKey="value"
                      position="right"
                      style={{ fontSize: 9, fill: "#333" }}
                      formatter={(v: number) => `${v}%`}
                    />
                    {chartData.map((entry, index) => (
                      <Cell
                        key={`cell-${index}`}
                        fill={entry.value >= 0 ? POSITIVE_COLOR : NEGATIVE_COLOR}
                      />
                    ))}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          );
        })}

        {/* Легенда — названия показателей */}
        <div
          className="flex-shrink-0 flex flex-col justify-start pt-6"
          style={{ width: 180, paddingTop: 28 }}
        >
          {rows.map((row, i) => (
            <div
              key={i}
              className="text-xs text-gray-700 font-medium"
              style={{ height: 260 / rows.length, display: "flex", alignItems: "center" }}
            >
              {row.indicator}
            </div>
          ))}
        </div>
      </div>

      {/* Легенда цветов */}
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
