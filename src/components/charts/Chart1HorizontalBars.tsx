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

// Высота одного бара + промежуток между барами в Recharts при barSize=14
const BAR_SIZE = 14;
const CHART_HEIGHT = 280;

export default function Chart1HorizontalBars({ years, rows }: Props) {
  const n = rows.length;

  // Recharts раскидывает n баров равномерно по высоте.
  // Высота на один элемент = CHART_HEIGHT / n
  const itemHeight = CHART_HEIGHT / n;

  return (
    <div className="bg-white rounded border border-gray-300 p-3">
      <div className="text-sm font-bold text-center text-gray-800 mb-3">
        Анализ изменений показателей по годам, %
      </div>

      <div className="flex gap-0 overflow-x-auto items-stretch">

        {/* Подписи показателей — выровнены по барам */}
        <div
          className="flex-shrink-0 flex flex-col"
          style={{ width: 170, paddingTop: 18 /* отступ под ось X */ }}
        >
          {rows.map((row, i) => (
            <div
              key={i}
              className="text-xs text-gray-700 font-medium flex items-center pr-2"
              style={{
                height: itemHeight,
                lineHeight: 1.25,
              }}
            >
              {row.indicator}
            </div>
          ))}
        </div>

        {/* Графики по годам */}
        {years.map((year, yIdx) => {
          const chartData = rows.map((row) => ({
            name: row.indicator,
            value: row.values[yIdx] ?? 0,
          }));

          return (
            <div key={year} className="flex-shrink-0" style={{ width: 155 }}>
              <div className="text-xs font-bold text-center text-gray-700 mb-1 underline">
                {year}
              </div>
              <ResponsiveContainer width="100%" height={CHART_HEIGHT}>
                <BarChart
                  data={chartData}
                  layout="vertical"
                  margin={{ top: 0, right: 28, left: 4, bottom: 0 }}
                  barSize={BAR_SIZE}
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
