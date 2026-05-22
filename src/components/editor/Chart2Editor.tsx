import React from "react";
import { Chart2Data } from "@/types/oilfield";

interface Props {
  data: Chart2Data;
  onChange: (data: Chart2Data) => void;
}

const FIELDS: { key: keyof Omit<Chart2Data, "years">; label: string }[] = [
  { key: "zakachka", label: "Закачка, тыс.м³" },
  { key: "liquid", label: "Добыча жидкости, тыс.т" },
  { key: "oil", label: "Добыча нефти, тыс.т" },
  { key: "fond_dob", label: "Фонд добыв. скв., ед" },
  { key: "fond_nag", label: "Фонд нагн. скв., ед" },
  { key: "compensation", label: "Компенсация тек., %" },
];

export default function Chart2Editor({ data, onChange }: Props) {
  const handleCell = (field: keyof Omit<Chart2Data, "years">, idx: number, val: string) => {
    const num = val === "" ? 0 : Number(val.replace(",", "."));
    const arr = [...(data[field] as number[])];
    arr[idx] = isNaN(num) ? 0 : num;
    onChange({ ...data, [field]: arr });
  };

  const handleYear = (idx: number, val: string) => {
    const yr = parseInt(val);
    if (!isNaN(yr)) {
      const updated = [...data.years];
      updated[idx] = yr;
      onChange({ ...data, years: updated });
    }
  };

  const addYear = () => {
    const lastYear = data.years[data.years.length - 1] ?? 2025;
    const updated: Chart2Data = {
      years: [...data.years, lastYear + 1],
      zakachka: [...data.zakachka, 0],
      liquid: [...data.liquid, 0],
      oil: [...data.oil, 0],
      fond_dob: [...data.fond_dob, 0],
      fond_nag: [...data.fond_nag, 0],
      compensation: [...data.compensation, 0],
    };
    onChange(updated);
  };

  const removeYear = (idx: number) => {
    if (data.years.length <= 1) return;
    const remove = (arr: number[]) => arr.filter((_, i) => i !== idx);
    onChange({
      years: data.years.filter((_, i) => i !== idx),
      zakachka: remove(data.zakachka),
      liquid: remove(data.liquid),
      oil: remove(data.oil),
      fond_dob: remove(data.fond_dob),
      fond_nag: remove(data.fond_nag),
      compensation: remove(data.compensation),
    });
  };

  return (
    <div className="overflow-x-auto">
      <table className="border-collapse text-xs" style={{ minWidth: 400 }}>
        <thead>
          <tr>
            <th
              className="border border-gray-400 bg-gray-100 text-left px-2 py-1 font-semibold sticky left-0 z-10"
              style={{ minWidth: 180, background: "#f0f0f0" }}
            >
              Показатель
            </th>
            {data.years.map((yr, idx) => (
              <th
                key={idx}
                className="border border-gray-400 bg-gray-100 text-center font-semibold"
                style={{ minWidth: 90 }}
              >
                <div className="flex items-center justify-center gap-1 px-1 py-0.5">
                  <input
                    type="number"
                    value={yr}
                    onChange={(e) => handleYear(idx, e.target.value)}
                    className="w-16 text-center text-xs border border-gray-300 rounded focus:outline-none focus:border-blue-400 bg-white text-gray-800"
                    style={{ MozAppearance: "textfield" }}
                  />
                  <button
                    onClick={() => removeYear(idx)}
                    className="text-red-400 hover:text-red-600 text-xs"
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
              >
                + год
              </button>
            </th>
          </tr>
        </thead>
        <tbody>
          {FIELDS.map(({ key, label }, rIdx) => (
            <tr key={key} className={rIdx % 2 === 0 ? "bg-white" : "bg-gray-50"}>
              <td
                className="border border-gray-300 px-2 py-0.5 text-gray-800 font-medium sticky left-0 z-10"
                style={{ background: rIdx % 2 === 0 ? "#fff" : "#f9f9f9" }}
              >
                {label}
              </td>
              {data.years.map((_, cIdx) => {
                const val = (data[key] as number[])[cIdx] ?? 0;
                return (
                  <td key={cIdx} className="border border-gray-300 p-0" style={{ minWidth: 90 }}>
                    <input
                      type="number"
                      value={val}
                      onChange={(e) => handleCell(key, cIdx, e.target.value)}
                      className="w-full text-center text-xs px-1 py-0.5 focus:outline-none focus:bg-blue-50 border-0 bg-transparent text-gray-800"
                      style={{ MozAppearance: "textfield" }}
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
