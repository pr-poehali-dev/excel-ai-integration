export interface Chart1Row {
  indicator: string;
  values: (number | null)[];
}

export interface Chart2Data {
  years: number[];
  zakachka: number[];
  liquid: number[];
  oil: number[];
  fond_dob: number[];
  fond_nag: number[];
  compensation: number[];
}

export interface OilfieldData {
  name: string;
  years: number[];
  chart1Rows: Chart1Row[];
  chart2Data: Chart2Data;
}

export const CHART1_INDICATORS = [
  "Добыча нефти",
  "Добыча жидкости",
  "Обводненность (весовая)",
  "Дейст. добывающий фонд",
  "Дейст. нагнетательный фонд",
  "Закачка",
  "Приемистость",
  "Дебит нефти",
  "Дебит жидкости",
];

export const DEFAULT_CHART1_VALUES: number[][] = [
  [-5, -22, 9, 2, 10],
  [-2, -20, -2, 4, 12],
  [1, 1, -4, 1, 1],
  [-6, 0, 7, -3, -1],
  [12, -15, 12, 11, 24],
  [6, -10, 0, 7, 10],
  [3, 12, 6, -3, -8],
  [3, -10, 20, -6, 2],
  [5, -8, 8, -3, 4],
];

export const DEFAULT_YEARS = [2020, 2021, 2022, 2023, 2024];

export const DEFAULT_CHART2: Chart2Data = {
  years: [2022, 2023, 2024, 2025],
  zakachka: [550, 1050, 1600, 2000],
  liquid: [0, 0, 0, 2100],
  oil: [500, 1000, 1550, 1700],
  fond_dob: [10, 20, 20, 20],
  fond_nag: [10, 28, 30, 55],
  compensation: [45, 90, 108, 108],
};

export function createDefaultOilfield(name: string): OilfieldData {
  return {
    name,
    years: [...DEFAULT_YEARS],
    chart1Rows: CHART1_INDICATORS.map((indicator, i) => ({
      indicator,
      values: [...DEFAULT_CHART1_VALUES[i]],
    })),
    chart2Data: { ...DEFAULT_CHART2 },
  };
}
