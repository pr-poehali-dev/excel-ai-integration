import React from "react";
import {
  ComposedChart,
  Area,
  Bar,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
} from "recharts";
import { Chart2Data } from "@/types/oilfield";

interface Props {
  data: Chart2Data;
}

export default function Chart2ComboChart({ data }: Props) {
  const chartData = data.years.map((year, i) => ({
    year: String(year),
    zakachka: data.zakachka[i] ?? 0,
    liquid: data.liquid[i] ?? 0,
    oil: data.oil[i] ?? 0,
    fond_dob: data.fond_dob[i] ?? 0,
    fond_nag: data.fond_nag[i] ?? 0,
    compensation: data.compensation[i] ?? 0,
  }));

  const LEGEND_ITEMS = [
    { color: "#FFFF00", label: "Закачка, тыс.м³", border: "#ccc" },
    { color: "#ADD8E6", label: "Добыча жидкости, тыс.т", border: "#999" },
    { color: "#C0504D", label: "Добыча нефти, тыс.т", border: "#a00" },
    { color: "#8B4513", label: "Фонд добывающих скв., ед", border: "#5a2d0c" },
    { color: "#00BFFF", label: "Фонд нагнетательных скв., ед", border: "#0080cc" },
    { color: "#00008B", label: "Компенсация тек., %", border: "#00008B" },
  ];

  return (
    <div className="bg-white rounded border border-gray-300 p-3">
      <div className="text-sm font-bold text-center text-gray-800 mb-3">
        Динамика показателей разработки
      </div>

      <ResponsiveContainer width="100%" height={340}>
        <ComposedChart
          data={chartData}
          margin={{ top: 10, right: 60, left: 10, bottom: 10 }}
        >
          <CartesianGrid strokeDasharray="3 3" stroke="#e0e0e0" />
          <XAxis
            dataKey="year"
            tick={{ fontSize: 11, fill: "#555" }}
            axisLine={{ stroke: "#ccc" }}
            tickLine={false}
          />
          {/* Левая ось — добыча/закачка */}
          <YAxis
            yAxisId="left"
            orientation="left"
            tick={{ fontSize: 10, fill: "#555" }}
            axisLine={{ stroke: "#ccc" }}
            tickLine={false}
            label={{
              value: "Добыча, тыс.т; Закачка, тыс.м³",
              angle: -90,
              position: "insideLeft",
              style: { fontSize: 9, fill: "#777" },
              dx: -4,
            }}
          />
          {/* Правая ось — фонды и компенсация */}
          <YAxis
            yAxisId="right"
            orientation="right"
            tick={{ fontSize: 10, fill: "#555" }}
            axisLine={{ stroke: "#ccc" }}
            tickLine={false}
            label={{
              value: "Фонд скв., шт; Компенсация, %",
              angle: 90,
              position: "insideRight",
              style: { fontSize: 9, fill: "#777" },
              dx: 14,
            }}
          />
          <Tooltip
            contentStyle={{ fontSize: 11, background: "#fff", border: "1px solid #ccc" }}
          />

          {/* Площади */}
          <Area
            yAxisId="left"
            type="linear"
            dataKey="zakachka"
            name="Закачка, тыс.м³"
            fill="#FFFF00"
            stroke="#cccc00"
            strokeWidth={1}
            fillOpacity={0.85}
            stackId="a"
          />
          <Area
            yAxisId="left"
            type="linear"
            dataKey="liquid"
            name="Добыча жидкости, тыс.т"
            fill="#ADD8E6"
            stroke="#87CEEB"
            strokeWidth={1}
            fillOpacity={0.7}
            stackId="b"
          />
          <Area
            yAxisId="left"
            type="linear"
            dataKey="oil"
            name="Добыча нефти, тыс.т"
            fill="#C0504D"
            stroke="#a03030"
            strokeWidth={1}
            fillOpacity={0.85}
            stackId="b"
          />

          {/* Столбцы фондов (правая ось) */}
          <Bar
            yAxisId="right"
            dataKey="fond_dob"
            name="Фонд добыв. скв., ед"
            fill="#8B4513"
            stroke="#5a2d0c"
            strokeWidth={1}
            barSize={20}
            fillOpacity={0.9}
          />
          <Bar
            yAxisId="right"
            dataKey="fond_nag"
            name="Фонд нагн. скв., ед"
            fill="none"
            stroke="#00BFFF"
            strokeWidth={2}
            barSize={28}
            fillOpacity={0}
          />

          {/* Линия компенсации */}
          <Line
            yAxisId="right"
            type="linear"
            dataKey="compensation"
            name="Компенсация тек., %"
            stroke="#00008B"
            strokeWidth={2}
            dot={{ fill: "#00008B", r: 4 }}
            activeDot={{ r: 6 }}
          />
        </ComposedChart>
      </ResponsiveContainer>

      {/* Кастомная легенда */}
      <div className="flex flex-wrap gap-x-4 gap-y-1 mt-2 justify-center">
        {LEGEND_ITEMS.map((item) => (
          <div key={item.label} className="flex items-center gap-1 text-xs text-gray-600">
            <div
              className="w-5 h-3 rounded-sm border"
              style={{
                background: item.label.includes("нагн") ? "transparent" : item.color,
                borderColor: item.border,
                borderWidth: item.label.includes("нагн") ? 2 : 1,
              }}
            />
            {item.label}
          </div>
        ))}
      </div>
    </div>
  );
}
