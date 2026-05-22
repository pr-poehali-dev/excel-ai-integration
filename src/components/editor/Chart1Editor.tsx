import React from "react";
import { Chart1Row } from "@/types/oilfield";

interface Props {
  years: number[];
  rows: Chart1Row[];
  onChange: (rows: Chart1Row[]) => void;
  onYearsChange: (years: number[]) => void;
}

export default function Chart1Editor({ years, rows, onChange, onYearsChange }: Props) {
  const handleCell = (rowIdx: number, colIdx: number, val: string) => {
    const num = val === "" ? null : Number(val.replace(",", "."));
    const updated = rows.map((r, ri) => {
      if (ri !== rowIdx) return r;
      const newVals = [...r.values];
      newVals[colIdx] = isNaN(num as number) ? null : num;
      return { ...r, values: newVals };
    });
    onChange(updated);
  };

  const handleYear = (idx: number, val: string) => {
    const yr = parseInt(val);
    if (!isNaN(yr)) {
      const updated = [...years];
      updated[idx] = yr;
      onYearsChange(updated);
    }
  };

  const addYear = () => {
    const lastYear = years[years.length - 1] ?? 2024;
    onYearsChange([...years, lastYear + 1]);
    onChange(rows.map((r) => ({ ...r, values: [...r.values, 0] })));
  };

  const removeYear = (idx: number) => {
    if (years.length <= 1) return;
    onYearsChange(years.filter((_, i) => i !== idx));
    onChange(rows.map((r) => ({ ...r, values: r.values.filter((_, i) => i !== idx) })));
  };

  return (
    <div className="overflow-x-auto">
      <table className="border-collapse text-xs" style={{ minWidth: 500 }}>
        <thead>
          <tr>
            <th
              className="border border-gray-400 bg-gray-100 text-gray-700 text-left px-2 py-1 font-semibold sticky left-0 z-10"
              style={{ minWidth: 180, background: "#f0f0f0" }}
            >
              Показатель
            </th>
            {years.map((yr, idx) => (
              <th
                key={idx}
                className="border border-gray-400 bg-gray-100 text-center font-semibold"
                style={{ minWidth: 80 }}
              >
                <div className="flex items-center justify-center gap-1 px-1 py-0.5">
                  <input
                    type="number"
                    value={yr}
                    onChange={(e) => handleYear(idx, e.target.value)}
                    className="w-14 text-center text-xs border border-gray-300 rounded focus:outline-none focus:border-blue-400 bg-white text-gray-800"
                    style={{ MozAppearance: "textfield" }}
                  />
                  <button
                    onClick={() => removeYear(idx)}
                    className="text-red-400 hover:text-red-600 text-xs leading-none"
                    title="Удалить год"
                  >
                    ×
                  </button>
                </div>
              </th>
            ))}
            <th className="border border-gray-400 bg-gray-100 px-2">
              <button
                onClick={addYear}
                className="text-blue-600 hover:text-blue-800 text-xs font-bold"
                title="Добавить год"
              >
                + год
              </button>
            </th>
          </tr>
        </thead>
        <tbody>
          {rows.map((row, rIdx) => (
            <tr
              key={rIdx}
              className={rIdx % 2 === 0 ? "bg-white" : "bg-gray-50"}
            >
              <td
                className="border border-gray-300 px-2 py-0.5 text-gray-800 font-medium sticky left-0 z-10"
                style={{ background: rIdx % 2 === 0 ? "#fff" : "#f9f9f9" }}
              >
                {row.indicator}
              </td>
              {years.map((_, cIdx) => {
                const val = row.values[cIdx];
                const isNeg = typeof val === "number" && val < 0;
                const isPos = typeof val === "number" && val > 0;
                return (
                  <td
                    key={cIdx}
                    className="border border-gray-300 p-0"
                    style={{ minWidth: 80 }}
                  >
                    <input
                      type="number"
                      value={val ?? ""}
                      onChange={(e) => handleCell(rIdx, cIdx, e.target.value)}
                      className="w-full text-center text-xs px-1 py-0.5 focus:outline-none focus:bg-blue-50 border-0"
                      style={{
                        background: isNeg
                          ? "rgba(128,128,128,0.12)"
                          : isPos
                          ? "rgba(68,114,196,0.10)"
                          : "transparent",
                        color: isNeg ? "#c00" : isPos ? "#1a3a8a" : "#333",
                        fontWeight: val !== 0 ? "600" : "400",
                        MozAppearance: "textfield",
                      }}
                    />
                  </td>
                );
              })}
              <td className="border border-gray-300" />
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
