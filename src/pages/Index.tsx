import { useState, useRef, useCallback, useEffect } from "react";
import { useNavigate } from "react-router-dom";
import * as XLSX from "xlsx";
import { PieChart, Pie, Cell, Tooltip, Legend, ResponsiveContainer, BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid } from "recharts";
import Icon from "@/components/ui/icon";
import mammoth from "mammoth";
import {
  saveSession, loadSession, listSessions, deleteSession,
  saveKnowledge, loadKnowledge, formatKnowledgeForAI, getKnowledgePdfImages,
  OIL_KNOWLEDGE_TEMPLATES, KNOWLEDGE_CATEGORIES,
  type SavedSession, type KnowledgeEntry, type KnowledgeCategory, type KnowledgeSourceType,
} from "@/lib/session-db";

// ─── Types ────────────────────────────────────────────────────────────────────

type CellValue = string | number | boolean | null;
type ChartType = "pie" | "bar" | "line";

interface CellBorder {
  style?: string; // thin | medium | thick | dashed | dotted | double
  color?: string;
}

interface CellStyle {
  // Fill
  bgColor?: string;
  // Font
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  // Alignment
  hAlign?: "left" | "center" | "right" | "general";
  vAlign?: "top" | "middle" | "bottom";
  wrapText?: boolean;
  // Borders
  borderTop?: CellBorder;
  borderBottom?: CellBorder;
  borderLeft?: CellBorder;
  borderRight?: CellBorder;
  // Number format
  numFmt?: string;
}

interface CellData {
  v: CellValue;       // raw value
  w?: string;         // formatted text (from Excel)
  s?: CellStyle;
}

interface MergeRange {
  r1: number; c1: number;
  r2: number; c2: number;
}

interface SheetData {
  name: string;
  cells: CellData[][];   // [row][col]
  merges: MergeRange[];
  colWidths: number[];
  rowHeights: number[];
  chartType?: ChartType;
}

interface ExcelFile {
  id: string;
  name: string;
  role: "main" | "reference" | null;
  sheets: SheetData[];
  activeSheet: number;
  workbook: XLSX.WorkBook;
  isDirty: boolean;
}

// Текстовые документы (PDF, DOCX)
type DocFileType = "pdf" | "docx";
interface DocFile {
  id: string;
  name: string;
  type: DocFileType;
  role: "report" | "protocol" | "database" | null;
  text: string;
  html?: string;            // HTML для DOCX
  pageCount?: number;
  pageImageUrls?: string[]; // PDF: URL PNG-изображений страниц (для ИИ)
  loading?: boolean;        // PDF: идёт конвертация
  pageFrom?: number;        // диапазон страниц для ИИ (1-based)
  pageTo?: number;
  buffer?: ArrayBuffer;     // оригинальный файл для сохранения в сессии
}

// Промпты с галочками
interface PromptPreset {
  id: string;
  label: string;
  text: string;
  enabled: boolean;
}

const DEFAULT_PROMPTS: PromptPreset[] = [
  {
    id: "role",
    label: "Роль: разработчик месторождений",
    enabled: true,
    text: `Ты разработчик нефтяных и газовых месторождений. Выполняешь проекты пробной эксплуатации, проекты на разработку месторождений и дополнения к ним (ПТД). За основу берёшь действующий проект ПТД, актуализируешь данные и анализ на основе обновлённой базы данных.`,
  },
  {
    id: "tables_db",
    label: "Таблицы из базы данных",
    enabled: false,
    text: `Твоя задача: брать базу данных которую тебе предоставят, проверять и обновлять таблицы протокола ЦКР. Объект разработки — это сумма всех пластов, входящих в него (см. Таблицу «База данных» или раздел VII Принципиальные положения документа ПТД).`,
  },
  {
    id: "tables_protocol",
    label: "Таблицы протокола ЦКР",
    enabled: false,
    text: `На основе таблиц протокола и базы данных актуализируй таблицы протокола ЦКР. Эталоном являются таблицы протокола — при любых расхождениях с текстом отчёта сообщай, выделяя красным. Сначала утверждаем данные в таблицах ЦКР, затем правим текст отчёта.`,
  },
  {
    id: "update_report",
    label: "Актуализация отчёта (Главы 3, 5, 6)",
    enabled: false,
    text: `Актуализируй текст и таблицы отчёта в Word (Главы 3, 5, 6). В Word сохраняй все стили как в исходнике. Твоя задача — только описание, сохраняя стиль и подачу исходника, актуализируя только числовые значения. Все изменения закрашивай в Word зелёным цветом, сохраняя основные шрифты. По запросу проверяй главы с цифрами по всему отчёту на основе данных из таблиц протокола и базы данных.`,
  },
  {
    id: "excel_formulas",
    label: "Формулы Excel 2010",
    enabled: false,
    text: `Все формулы пиши для Excel 2010. Примеры: =СУММ(B5:B10); =ВПР(); =ПРОМЕЖУТОЧНЫЕ.ИТОГИ(9;). Формулы указывай в синтаксисе русского Excel 2010.`,
  },
  {
    id: "protocol_update",
    label: "Обновление протоколов ЦКР",
    enabled: false,
    text: `Проверяй и обновляй Протоколы ЦКР: описывай подразделы согласно образцу, редактируй данные на основе обновлений таблиц протокола.`,
  },
];

interface ChatImage {
  dataUrl: string;
  name: string;
}

interface ChatMessage {
  role: "user" | "ai";
  text: string;
  ts: string;
  refs?: string[];
  images?: ChatImage[];
  chartData?: { name: string; value: number }[];
  chartTitle?: string;
  ask_user?: string;        // вопрос ИИ для уточнения — показываем интерактивную карточку
  pendingPrompt?: string;   // исходный промпт пользователя, ждёт уточнения
  docText?: string;         // готовый текст документа — показываем прямо в чате с кнопкой копировать
  docName?: string;         // имя документа для заголовка блока
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function colLetter(n: number) {
  let s = "";
  while (n >= 0) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

function getTime() {
  return new Date().toLocaleTimeString("ru-RU", { hour: "2-digit", minute: "2-digit" });
}

function uid() {
  return Math.random().toString(36).slice(2, 8);
}

// Convert ARGB "FF112233" → "#112233"
function argbToHex(argb?: string): string | undefined {
  if (!argb) return undefined;
  const hex = argb.replace(/^FF/i, "");
  if (hex.length === 6) return `#${hex}`;
  if (hex.length === 8) return `#${hex.slice(2)}`;
  return undefined;
}

function parseBorder(b?: XLSX.Border): CellBorder | undefined {
  if (!b || !b.style) return undefined;
  return {
    style: b.style,
    color: argbToHex(b.color?.argb),
  };
}

function borderCss(b?: CellBorder, fallback = "1px solid rgba(255,255,255,0.07)"): string {
  if (!b?.style) return fallback;
  const w = b.style === "thin" ? "1px" : b.style === "medium" ? "2px" : b.style === "thick" ? "3px" : "1px";
  const t = b.style === "dashed" ? "dashed" : b.style === "dotted" ? "dotted" : b.style === "double" ? "double" : "solid";
  const c = b.color ?? "rgba(255,255,255,0.25)";
  return `${w} ${t} ${c}`;
}

function parseStyle(cell: XLSX.CellObject): CellStyle | undefined {
  const s = (cell as { s?: Record<string, unknown> }).s;
  if (!s) return undefined;

  const fill = s.fgColor as { argb?: string } | undefined;
  const patternType = (s.patType ?? s.patternType) as string | undefined;
  const font = s.font as {
    bold?: boolean; italic?: boolean; underline?: boolean;
    sz?: number; color?: { argb?: string };
  } | undefined;
  const alignment = s.alignment as {
    horizontal?: string; vertical?: string; wrapText?: boolean;
  } | undefined;
  const border = s.border as {
    top?: XLSX.Border; bottom?: XLSX.Border;
    left?: XLSX.Border; right?: XLSX.Border;
  } | undefined;
  const numFmt = s.numFmt as string | undefined;

  const bg = fill?.argb && patternType !== "none" ? argbToHex(fill.argb) : undefined;

  const style: CellStyle = {};
  if (bg && bg !== "#000000" && bg !== "#FFFFFF00") style.bgColor = bg;
  if (font?.bold) style.bold = true;
  if (font?.italic) style.italic = true;
  if (font?.underline) style.underline = true;
  if (font?.sz && font.sz !== 11) style.fontSize = font.sz;
  if (font?.color?.argb) {
    const c = argbToHex(font.color.argb);
    if (c && c !== "#000000") style.color = c;
  }
  if (alignment?.horizontal && alignment.horizontal !== "general") {
    style.hAlign = alignment.horizontal as CellStyle["hAlign"];
  }
  if (alignment?.vertical) {
    const m: Record<string, CellStyle["vAlign"]> = { top: "top", center: "middle", bottom: "bottom" };
    style.vAlign = m[alignment.vertical] ?? "middle";
  }
  if (alignment?.wrapText) style.wrapText = true;
  if (border?.top) style.borderTop = parseBorder(border.top);
  if (border?.bottom) style.borderBottom = parseBorder(border.bottom);
  if (border?.left) style.borderLeft = parseBorder(border.left);
  if (border?.right) style.borderRight = parseBorder(border.right);
  if (numFmt) style.numFmt = numFmt;

  return Object.keys(style).length ? style : undefined;
}

function parseWorkbook(wb: XLSX.WorkBook, name: string, id: string): ExcelFile {
  const sheets: SheetData[] = wb.SheetNames.map((sn) => {
    const ws = wb.Sheets[sn];

    // Determine actual range
    const ref = ws["!ref"];
    const range = ref ? XLSX.utils.decode_range(ref) : { s: { r: 0, c: 0 }, e: { r: 49, c: 19 } };
    const ROWS = Math.max(range.e.r + 1 + 10, 50);
    const COLS = Math.max(range.e.c + 1 + 2, 20);

    // Build cells grid
    const cells: CellData[][] = Array.from({ length: ROWS }, () =>
      Array.from({ length: COLS }, () => ({ v: null }))
    );

    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr] as XLSX.CellObject | undefined;
        if (!cell) continue;
        const cd: CellData = { v: cell.v as CellValue ?? null };
        if (cell.w && cell.w !== String(cell.v)) cd.w = cell.w;
        const st = parseStyle(cell);
        if (st) cd.s = st;
        cells[r][c] = cd;
      }
    }

    // Merges
    const merges: MergeRange[] = (ws["!merges"] ?? []).map((m: XLSX.Range) => ({
      r1: m.s.r, c1: m.s.c, r2: m.e.r, c2: m.e.c,
    }));

    // Col widths (Excel stores in chars ~7px each)
    const colInfos = ws["!cols"] as { wpx?: number; wch?: number; width?: number }[] | undefined;
    const colWidths = Array.from({ length: COLS }, (_, ci) => {
      const ci2 = colInfos?.[ci];
      if (ci2?.wpx) return Math.max(ci2.wpx, 40);
      if (ci2?.wch) return Math.max(Math.round(ci2.wch * 7), 40);
      return 100;
    });

    // Row heights
    const rowInfos = ws["!rows"] as { hpx?: number; hpt?: number }[] | undefined;
    const rowHeights = Array.from({ length: ROWS }, (_, ri) => {
      const ri2 = rowInfos?.[ri];
      if (ri2?.hpx) return Math.max(ri2.hpx, 18);
      return 22;
    });

    return { name: sn, cells, merges, colWidths, rowHeights };
  });

  return { id, name, role: null, sheets, activeSheet: 0, workbook: wb, isDirty: false };
}

function buildWorkbook(file: ExcelFile): XLSX.WorkBook {
  // Клонируем оригинальный workbook — сохраняем все стили, объединения, форматы
  const origWb = file.workbook;
  const wb: XLSX.WorkBook = {
    SheetNames: [...origWb.SheetNames],
    Sheets: {},
    Props: origWb.Props,
  };

  file.sheets.forEach((sh) => {
    const origWs = origWb.Sheets[sh.name];
    if (!origWs) {
      // Новый лист (созданный ИИ) — просто aoa
      const aoa = sh.cells.map(row => row.map(c => c.v));
      wb.Sheets[sh.name] = XLSX.utils.aoa_to_sheet(aoa);
      if (!wb.SheetNames.includes(sh.name)) wb.SheetNames.push(sh.name);
      return;
    }

    // Глубокое копирование оригинального листа (все стили сохраняются)
    const ws: XLSX.WorkSheet = {};
    for (const key of Object.keys(origWs)) {
      const val = origWs[key];
      if (key.startsWith("!")) {
        ws[key] = val; // метаданные: !ref, !merges, !cols, !rows — копируем как есть
      } else {
        // Копируем ячейку — сохраняем все поля включая .s (стиль)
        ws[key] = { ...val };
      }
    }

    // Записываем только изменённые значения, не трогая стили
    sh.cells.forEach((row, r) => {
      row.forEach((cd, c) => {
        const addr = XLSX.utils.encode_cell({ r, c });
        const origCell = origWs[addr] as XLSX.CellObject | undefined;
        const origVal = origCell?.v ?? null;
        if (cd.v === origVal) return; // не изменилось — пропускаем

        if (cd.v === null) {
          delete ws[addr];
        } else {
          const t = typeof cd.v === "number" ? "n" : typeof cd.v === "boolean" ? "b" : "s";
          ws[addr] = { ...(origCell ?? {}), v: cd.v, t } as XLSX.CellObject;
        }
      });
    });

    // Обновляем !ref под реальный размер данных
    const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
    const maxR = Math.max(range.e.r, sh.cells.length - 1);
    const maxC = Math.max(range.e.c, (sh.cells[0]?.length ?? 1) - 1);
    ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });

    wb.Sheets[sh.name] = ws;
  });

  return wb;
}

// Полный текст листа с номерами строк (для активного листа)
// Excel-адрес ячейки: (0,0) → "A1", (12, 13) → "N13"  (строки 1-based как в Excel)
function cellAddr(row: number, col: number): string {
  return `${colLetter(col)}${row + 1}`;
}

// Реальная ширина листа
function getRealColCount(sh: SheetData): number {
  let maxC = 0;
  sh.cells.forEach(row => {
    row.forEach((c, ci) => { if (c.v !== null) maxC = Math.max(maxC, ci + 1); });
  });
  sh.merges.forEach(m => { maxC = Math.max(maxC, m.c2 + 1); });
  return Math.min(maxC + 1, 60);
}

// Если ячейка входит в объединение — возвращает координаты хозяина (верхний левый угол).
// Если сама хозяин или не в объединении — null.
function getMergeOwner(sh: SheetData, r: number, c: number): { r: number; c: number } | null {
  for (const m of sh.merges) {
    if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) {
      if (r === m.r1 && c === m.c1) return null;
      return { r: m.r1, c: m.c1 };
    }
  }
  return null;
}

// Последняя строка зоны заголовков.
// Алгоритм:
//   1. Строки-разделители (объединение шириной >= 80% maxCols) не считаются заголовками столбцов.
//   2. Берём нижнюю границу последнего "настоящего" заголовочного объединения
//      (ширина > 1 col или высота > 1 row, но ширина < 80% maxCols).
//   3. Если нет объединений — первая строка где >50% непустых ячеек числовые, минус 1.
function detectLastHeaderRow(sh: SheetData): number {
  const maxC = getRealColCount(sh);
  const FULL_ROW_THRESHOLD = 0.8; // объединение шириной ≥ 80% = строка-разделитель, не заголовок

  let lastHeaderMergeRow = -1;
  for (const m of sh.merges) {
    const spanCols = m.c2 - m.c1 + 1;
    if (spanCols / maxC >= FULL_ROW_THRESHOLD) continue; // строка-разделитель — пропускаем
    lastHeaderMergeRow = Math.max(lastHeaderMergeRow, m.r2);
  }
  if (lastHeaderMergeRow >= 0) return lastHeaderMergeRow;

  // Fallback: первая строка где >50% непустых ячеек числовые
  for (let r = 0; r < Math.min(sh.cells.length, 30); r++) {
    const row = sh.cells[r];
    if (!row) continue;
    let nums = 0, total = 0;
    for (let c = 0; c < maxC; c++) {
      const v = row[c]?.v;
      if (v !== null && v !== undefined && v !== "") { total++; if (typeof v === "number") nums++; }
    }
    if (total > 2 && nums / total > 0.5) return r - 1;
  }
  return 2;
}

// Возвращает смысловую цепочку заголовков для столбца c, читая строки 0..lastHR сверху вниз.
// Каждый уровень — уникальный текст из объединения или одиночной ячейки.
// Числа и пустые ячейки пропускаются.
function getColHeaderChain(sh: SheetData, c: number, lastHR: number): string[] {
  const seen = new Set<string>();
  const parts: string[] = [];
  for (let r = 0; r <= lastHR && r < sh.cells.length; r++) {
    const owner = getMergeOwner(sh, r, c);
    const cell = owner ? sh.cells[owner.r]?.[owner.c] : sh.cells[r]?.[c];
    const val = (cell?.w ?? (cell?.v != null ? String(cell.v) : "")).trim();
    if (!val) continue;
    // Годы (2020-2040) включаем — это важная часть заголовка столбца
    const isYear = typeof cell?.v === "number" && cell.v >= 2020 && cell.v <= 2040;
    // Номера столбцов (1, 2, 3...) и прочие числа — пропускаем
    if (!isYear && typeof cell?.v === "number") continue;
    if (!isYear && /^\d+$/.test(val)) continue;
    if (!seen.has(val)) { seen.add(val); parts.push(val); }
  }
  return parts;
}

// Строит полный контекст листа для ИИ.
// Ключевая идея: каждая ячейка данных сопровождается ПОЛНОЙ смысловой цепочкой заголовков.
// Пример: P12=1943 [Текущие запасы нефти / извлекаемые / A+B1]
function buildSheetContext(sh: SheetData, maxDataRows = 250): string {
  const maxCols = getRealColCount(sh);
  // Лимит токенов: если лист большой — переходим на компактный TSV (заголовки 1 раз, данные без тегов)
  const totalDataRows = sh.cells.filter(r => r.some(c => c.v !== null)).length;
  const COMPACT_THRESHOLD = 80; // строк — при превышении используем TSV
  if (totalDataRows > COMPACT_THRESHOLD) {
    const lines: string[] = [];
    const lastHR = detectLastHeaderRow(sh);

    // Строим карту: для каждой ячейки заголовка — реальное значение с учётом объединений
    // Объединённые ячейки: значение владельца копируется во все дочерние ячейки
    const resolveHeaderCell = (r: number, c: number): string => {
      const owner = getMergeOwner(sh, r, c);
      const cell = owner ? sh.cells[owner.r]?.[owner.c] : sh.cells[r]?.[c];
      const v = cell?.v;
      return v !== null && v !== undefined ? String(v) : "";
    };

    // Строим цепочку заголовков для каждого столбца (как getColHeaderChain но для TSV)
    // Формат первой строки: A=col0 | B=col1 | ... (чтобы ИИ знал букву столбца)
    const colLetters = Array.from({ length: maxCols }, (_, c) => `${colLetter(c)}=col${c}`);
    lines.push("СТОЛБЦЫ: " + colLetters.join(" | "));

    // Заголовки — каждая строка, объединения раскрыты
    lines.push("ЗАГОЛОВКИ:");
    for (let r = 0; r <= lastHR; r++) {
      const row = sh.cells[r];
      if (!row && !sh.merges.some(m => m.r1 <= r && r <= m.r2)) continue;
      const cells = Array.from({ length: maxCols }, (_, c) => resolveHeaderCell(r, c));
      // Пропускаем полностью пустые строки заголовков
      if (cells.every(v => !v)) continue;
      lines.push(cells.join("\t"));
    }

    // Сводная таблица заголовков столбцов (цепочка сверху вниз для каждого столбца)
    lines.push("ЦЕПОЧКИ ЗАГОЛОВКОВ (столбец → полный путь):");
    for (let c = 0; c < maxCols; c++) {
      const chain = getColHeaderChain(sh, c, lastHR);
      if (chain.length > 0) lines.push(`  ${colLetter(c)}(col=${c}) = "${chain.join(" / ")}"`);
    }

    lines.push(`ДАННЫЕ (формат: "Excel_строка(row=индекс_0based)\\tA\\tB\\tC...")`);
    lines.push(`ВАЖНО: для cell_styles/cell_updates используй row=индекс из скобок, col=из ЦЕПОЧЕК ЗАГОЛОВКОВ`);
    let written = 0;
    for (let r = lastHR + 1; r < sh.cells.length && written < maxDataRows; r++) {
      const row = sh.cells[r];
      if (!row || !row.some(c => c.v !== null)) continue;
      const cells = Array.from({ length: maxCols }, (_, c) => {
        const v = row[c]?.v;
        return v !== null && v !== undefined ? String(v) : "";
      });
      lines.push(`${r + 1}(row=${r})\t` + cells.join("\t"));
      written++;
    }
    if (written >= maxDataRows) lines.push(`[... ещё ${totalDataRows - written} строк не передано]`);
    return lines.join("\n");
  }
  const lastHR = detectLastHeaderRow(sh);
  const firstDataRow = lastHR + 1;
  const lines: string[] = [];

  // Кэш цепочек заголовков для каждого столбца
  const colChains: Map<number, string[]> = new Map();
  for (let c = 0; c < maxCols; c++) {
    const chain = getColHeaderChain(sh, c, lastHR);
    if (chain.length > 0) colChains.set(c, chain);
  }

  // 1. ОБЪЕДИНЁННЫЕ ЯЧЕЙКИ — все диапазоны с текстом в формате Excel
  lines.push("── ОБЪЕДИНЁННЫЕ ЯЧЕЙКИ (все диапазоны) ──");
  for (const m of sh.merges) {
    const ownerCell = sh.cells[m.r1]?.[m.c1];
    const val = (ownerCell?.w ?? (ownerCell?.v != null ? String(ownerCell.v) : "")).trim();
    if (!val) continue;
    const addr = (m.r1 === m.r2 && m.c1 === m.c2)
      ? cellAddr(m.r1, m.c1)
      : `${cellAddr(m.r1, m.c1)}:${cellAddr(m.r2, m.c2)}`;
    lines.push(`  ${addr} = "${val}"`);
  }
  lines.push("");

  // 2. СМЫСЛОВЫЕ ЗАГОЛОВКИ КАЖДОГО СТОЛБЦА
  // Это критически важно: для P = "Текущие запасы нефти / извлекаемые / A+B1"
  lines.push("── ЗАГОЛОВКИ СТОЛБЦОВ ──");
  lines.push(`(строки заголовков: 1..${lastHR + 1} | данные начинаются со строки ${firstDataRow + 1})`);
  lines.push(`(col=позиция 0-based: A=0 B=1 C=2 D=3 E=4 F=5 G=6 H=7 I=8 J=9 K=10 L=11 M=12 N=13 O=14 P=15 Q=16 R=17 S=18 T=19)`);
  for (let c = 0; c < maxCols; c++) {
    const chain = colChains.get(c);
    if (chain && chain.length > 0) {
      lines.push(`  ${colLetter(c)}(col=${c}) = "${chain.join(" / ")}"`);
    }
  }
  lines.push("");

  // 3. ДАННЫЕ — каждая ячейка с адресом И смысловым заголовком столбца
  // Формат: A12="K1br пл.АС 5"  P12=1943[Текущие запасы нефти/извлекаемые/A+B1]
  lines.push(`── ДАННЫЕ (строки ${firstDataRow + 1} и далее) ──`);
  let rowsWritten = 0;
  const MAX_DATA_ROWS = maxDataRows;

  for (let r = firstDataRow; r < sh.cells.length && rowsWritten < MAX_DATA_ROWS; r++) {
    const row = sh.cells[r];
    if (!row || !row.some(c => c.v !== null)) continue;

    // Проверяем: строка-разделитель (одно широкое объединение на всю ширину)?
    // Такие строки тоже выводим, но помечаем как разделитель
    const maxC = getRealColCount(sh);
    const isGroupHeader = sh.merges.some(m =>
      m.r1 === r && m.r2 === r && (m.c2 - m.c1 + 1) / maxC >= 0.8
    );

    const cells: string[] = [];
    for (let c = 0; c < maxCols; c++) {
      const owner = getMergeOwner(sh, r, c);
      if (owner) continue;
      const cell = row[c] ?? { v: null };
      if (cell.v === null) continue;
      const val = typeof cell.v === "number" ? String(cell.v) : (cell.w ?? String(cell.v));
      const chain = colChains.get(c);
      // Для строк-разделителей заголовок столбца не добавляем (там текст на всю строку)
      const tag = (!isGroupHeader && chain && chain.length > 0)
        ? `[${chain.join(" / ")}]`
        : "";
      cells.push(`${cellAddr(r, c)}=${val}${tag}`);
    }

    if (cells.length > 0) {
      const prefix = isGroupHeader ? "  [РАЗДЕЛ] " : "  ";
      lines.push(`${prefix}строка ${r + 1}(row=${r}): ${cells.join("  ")}`);
      rowsWritten++;
    }
  }
  if (rowsWritten >= MAX_DATA_ROWS) lines.push(`  [... показано ${MAX_DATA_ROWS} строк из ${sh.cells.length}]`);

  return lines.join("\n");
}

// Контекст листа для ИИ с опциональным лимитом строк
function sheetToText(sh: SheetData, full = false, maxRows = 500): string {
  if (full) return buildSheetContext(sh, maxRows);
  const MAX_ROWS = 5;
  const numCols = getRealColCount(sh);
  const dataRows: string[] = [];
  sh.cells.forEach((row, ri) => {
    if (!row.some(c => c.v !== null)) return;
    const vals = Array.from({ length: numCols }, (_, ci) => {
      const c = row[ci] ?? { v: null };
      if (typeof c.v === "number") return String(c.v);
      return c.w ?? (c.v !== null ? String(c.v) : "");
    }).join("\t");
    dataRows.push(`${ri + 1}\t${vals}`);
  });
  const header = "row\t" + Array.from({ length: numCols }, (_, i) => colLetter(i)).join("\t");
  const sliced = dataRows.slice(0, MAX_ROWS);
  const suffix = dataRows.length > MAX_ROWS ? `\n[...ещё ${dataRows.length - MAX_ROWS} строк]` : "";
  return [header, ...sliced].join("\n") + suffix;
}

// Алиас для кнопки "Контекст ИИ" в шапке
function buildMergeContext(sh: SheetData): string {
  return buildSheetContext(sh);
}

// ─── AI Settings ─────────────────────────────────────────────────────────────

const EXCEL_CHART_URL = "https://functions.poehali.dev/8376b365-14fc-4551-9fc7-0798d13ac4e6";
const AI_EXCEL_URL = "https://functions.poehali.dev/0ca54b20-aea0-424c-8e74-fa9b445c95ba";

const PRESET_PROVIDERS = [
  { label: "RouterAI (рекомендуется)", baseUrl: "https://routerai.ru/api/v1" },
  { label: "OpenAI (прямой)", baseUrl: "https://api.openai.com/v1" },
  { label: "ProxyAPI.ru", baseUrl: "https://api.proxyapi.ru/openai/v1" },
  { label: "Свой endpoint", baseUrl: "" },
];

const PRESET_MODELS = [
  { label: "DeepSeek V3 Chat", value: "deepseek/deepseek-chat" },
  { label: "DeepSeek R1", value: "deepseek/deepseek-r1" },
  { label: "GPT-4o mini", value: "openai/gpt-4o-mini" },
  { label: "GPT-4o", value: "openai/gpt-4o" },
  { label: "Claude 3.5 Haiku", value: "anthropic/claude-3-5-haiku" },
  { label: "Gemini 2.0 Flash", value: "google/gemini-2.0-flash-001" },
];

type ReasoningEffort = "none" | "low" | "medium" | "high";

interface AiSettings {
  apiKey: string;
  baseUrl: string;
  model: string;
  customModel: string;
  reasoningEffort: ReasoningEffort;
}

function loadSettings(): AiSettings {
  try {
    const s = localStorage.getItem("datamind_ai_settings");
    if (s) return { reasoningEffort: "none", ...JSON.parse(s) };
  } catch (e) { void e; }
  return { apiKey: "", baseUrl: "https://routerai.ru/api/v1", model: "deepseek/deepseek-chat", customModel: "", reasoningEffort: "none" };
}

function saveSettings(s: AiSettings) {
  localStorage.setItem("datamind_ai_settings", JSON.stringify(s));
}

function loadPrompts(): PromptPreset[] {
  try {
    const s = localStorage.getItem("datamind_prompts");
    if (s) {
      const saved = JSON.parse(s) as PromptPreset[];
      // Мержим сохранённые с дефолтными (на случай новых промптов)
      return DEFAULT_PROMPTS.map(def => {
        const found = saved.find(p => p.id === def.id);
        return found ? { ...def, enabled: found.enabled, text: found.text } : def;
      });
    }
  } catch (e) { void e; }
  return DEFAULT_PROMPTS;
}

function savePrompts(prompts: PromptPreset[]) {
  localStorage.setItem("datamind_prompts", JSON.stringify(prompts));
}

// ─── Settings Modal ───────────────────────────────────────────────────────────

interface SettingsModalProps {
  settings: AiSettings;
  onChange: (s: AiSettings) => void;
  onClose: () => void;
}

function SettingsModal({ settings, onChange, onClose }: SettingsModalProps) {
  const [local, setLocal] = useState<AiSettings>(settings);
  const [showKey, setShowKey] = useState(false);
  const [liveModels, setLiveModels] = useState<{ id: string; name: string }[]>([]);
  const [loadingModels, setLoadingModels] = useState(false);

  const update = (patch: Partial<AiSettings>) => setLocal((p) => ({ ...p, ...patch }));

  const fetchModels = async () => {
    if (!local.baseUrl) return;
    setLoadingModels(true);
    try {
      const r = await fetch(`${local.baseUrl.replace(/\/$/, "")}/models`, {
        headers: local.apiKey ? { Authorization: `Bearer ${local.apiKey}` } : {},
      });
      const json = await r.json();
      setLiveModels(
        (json.data as { id: string; name?: string }[])
          .map(m => ({ id: m.id, name: m.name || m.id }))
          .sort((a, b) => a.id.localeCompare(b.id))
      );
    } catch { void 0; }
    finally { setLoadingModels(false); }
  };

  const effectiveModel = local.model === "__custom__" ? local.customModel : local.model;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: "rgba(0,0,0,0.7)" }}
      onClick={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="w-full max-w-md rounded-2xl border border-border/60 p-6 animate-scale-in" style={{ background: "hsl(220,14%,9%)" }}>
        <div className="flex items-center justify-between mb-5">
          <div className="flex items-center gap-2">
            <div className="w-7 h-7 rounded-lg bg-secondary flex items-center justify-center">
              <Icon name="Settings2" size={14} className="text-primary" />
            </div>
            <span className="font-semibold text-foreground">Настройки ИИ</span>
          </div>
          <button onClick={onClose} className="text-muted-foreground hover:text-foreground transition-colors">
            <Icon name="X" size={16} />
          </button>
        </div>

        <div className="space-y-4">
          <div>
            <label className="text-xs text-muted-foreground mb-1.5 block font-medium">Провайдер</label>
            <div className="grid grid-cols-2 gap-1.5">
              {PRESET_PROVIDERS.map((p) => (
                <button key={p.label} onClick={() => { if (p.baseUrl) update({ baseUrl: p.baseUrl }); }}
                  className={`px-3 py-2 rounded-lg text-xs text-left transition-all border ${local.baseUrl === p.baseUrl && p.baseUrl ? "border-primary/50 bg-primary/10 text-primary" : "border-border/40 text-muted-foreground hover:text-foreground hover:border-border"}`}
                  style={{ background: local.baseUrl === p.baseUrl && p.baseUrl ? undefined : "rgba(255,255,255,0.02)" }}>
                  {p.label}
                </button>
              ))}
            </div>
            <input value={local.baseUrl} onChange={(e) => update({ baseUrl: e.target.value })}
              placeholder="https://your-provider.com/v1"
              className="mt-2 w-full px-3 py-2 rounded-lg border border-border/50 text-xs text-foreground placeholder:text-muted-foreground focus:outline-none focus:border-primary/50 transition-all font-mono"
              style={{ background: "rgba(255,255,255,0.03)" }} />
          </div>

          <div>
            <label className="text-xs text-muted-foreground mb-1.5 block font-medium">API ключ</label>
            <div className="relative">
              <input type={showKey ? "text" : "password"} value={local.apiKey}
                onChange={(e) => update({ apiKey: e.target.value })}
                placeholder="sk-..."
                className="w-full px-3 py-2 pr-8 rounded-lg border border-border/50 text-xs text-foreground placeholder:text-muted-foreground focus:outline-none focus:border-primary/50 font-mono transition-all"
                style={{ background: "rgba(255,255,255,0.03)" }} />
              <button onClick={() => setShowKey(p => !p)} className="absolute right-2 top-1/2 -translate-y-1/2 text-muted-foreground hover:text-foreground transition-colors">
                <Icon name={showKey ? "EyeOff" : "Eye"} size={14} />
              </button>
            </div>
          </div>

          <div>
            <div className="flex items-center justify-between mb-1.5">
              <label className="text-xs text-muted-foreground font-medium">Модель</label>
              <button onClick={fetchModels} disabled={loadingModels || !local.baseUrl}
                className="text-[10px] text-primary hover:opacity-80 transition-opacity disabled:opacity-40">
                {loadingModels ? "Загрузка..." : "Загрузить список"}
              </button>
            </div>
            <div className="space-y-1 max-h-36 overflow-y-auto scrollbar-thin">
              {[...PRESET_MODELS, ...(liveModels.filter(m => !PRESET_MODELS.find(p => p.value === m.id)).map(m => ({ label: m.name, value: m.id })))].map(m => (
                <button key={m.value} onClick={() => update({ model: m.value })}
                  className={`w-full text-left px-3 py-2 rounded-lg text-xs transition-all border ${local.model === m.value ? "border-primary/50 bg-primary/10 text-primary" : "border-border/40 text-muted-foreground hover:text-foreground"}`}
                  style={{ background: local.model === m.value ? undefined : "rgba(255,255,255,0.02)" }}>
                  {m.label}
                </button>
              ))}
              <button onClick={() => update({ model: "__custom__" })}
                className={`w-full text-left px-3 py-2 rounded-lg text-xs transition-all border ${local.model === "__custom__" ? "border-primary/50 bg-primary/10 text-primary" : "border-border/40 text-muted-foreground hover:text-foreground"}`}
                style={{ background: local.model === "__custom__" ? undefined : "rgba(255,255,255,0.02)" }}>
                Своя модель...
              </button>
            </div>
            {local.model === "__custom__" && (
              <input value={local.customModel} onChange={(e) => update({ customModel: e.target.value })}
                placeholder="model-name"
                className="mt-2 w-full px-3 py-2 rounded-lg border border-border/50 text-xs text-foreground placeholder:text-muted-foreground focus:outline-none focus:border-primary/50 font-mono"
                style={{ background: "rgba(255,255,255,0.03)" }} />
            )}
          </div>

          {/* Reasoning effort */}
          <div>
            <label className="text-xs text-muted-foreground mb-1.5 block font-medium">Время размышления</label>
            <div className="grid grid-cols-4 gap-1.5">
              {([
                { value: "none",   label: "Выкл",    desc: "Быстрее всего, без раздумий" },
                { value: "low",    label: "Низкое",  desc: "Лёгкие задачи" },
                { value: "medium", label: "Среднее", desc: "Баланс скорости и качества" },
                { value: "high",   label: "Высокое", desc: "Сложный анализ, медленнее" },
              ] as { value: ReasoningEffort; label: string; desc: string }[]).map(opt => (
                <button key={opt.value} onClick={() => update({ reasoningEffort: opt.value })}
                  title={opt.desc}
                  className={`px-2 py-2 rounded-lg text-xs font-medium text-center transition-all border ${
                    local.reasoningEffort === opt.value
                      ? "border-primary/50 bg-primary/10 text-primary"
                      : "border-border/40 text-muted-foreground hover:text-foreground hover:border-border"
                  }`}
                  style={{ background: local.reasoningEffort === opt.value ? undefined : "rgba(255,255,255,0.02)" }}>
                  {opt.label}
                </button>
              ))}
            </div>
            <p className="text-[10px] text-muted-foreground mt-1.5">
              {local.reasoningEffort === "none"   && "Модель отвечает без предварительного анализа — быстро, но может ошибаться на сложных задачах."}
              {local.reasoningEffort === "low"    && "Краткое обдумывание. Подходит для простых запросов — суммы, фильтрация, переименование."}
              {local.reasoningEffort === "medium" && "Оптимальный режим для большинства задач — анализ данных, поиск по условию, покраска ячеек."}
              {local.reasoningEffort === "high"   && "Глубокий анализ. Нужен для сложной логики, большого количества условий. Ответ занимает больше времени."}
            </p>
          </div>
        </div>

        <div className="flex justify-between mt-6 gap-2">
          <p className="text-[10px] text-muted-foreground self-center">Модель: {effectiveModel || "—"}</p>
          <button onClick={() => { saveSettings(local); onChange(local); onClose(); }} className="btn-primary px-4 py-2 rounded-lg text-xs font-semibold">
            Сохранить
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── Detect chart type ────────────────────────────────────────────────────────

function detectChartType(sheetName: string): ChartType {
  const n = sheetName.toLowerCase();
  if (n.includes("бар") || n.includes("bar") || n.includes("столб")) return "bar";
  if (n.includes("линия") || n.includes("line") || n.includes("динам")) return "line";
  return "pie";
}

// ─── AI call (direct — no backend, no timeout) ───────────────────────────────

const SYSTEM_PROMPT = `Ты — профессиональный аналитик данных и эксперт по Excel, специализирующийся на нефтяной отрасли.
Тебе передают: содержимое Excel-файлов и всех их листов, задание пользователя, и (опционально) изображения-образцы.

КРИТИЧЕСКИ ВАЖНО: ты ВСЕГДА отвечаешь ТОЛЬКО валидным JSON. Никакого текста вне JSON. Никаких \`\`\`json\`\`\` обёрток.

═══════════════════════════════════════════════
ПОЛНЫЙ ФОРМАТ ОТВЕТА (используй только нужные поля):
{
  "text": "Объяснение что сделано или что уточнить",
  "ask_user": "Вопрос к пользователю если данных недостаточно",
  "new_sheet": { "file_index": 0, "sheet_name": "Название листа", "data": [["Заголовок1","Заголовок2",...],[...]] },
  "cell_updates": [{ "file_index": 0, "sheet_name": "Лист1", "changes": [{"row":10,"col":1,"value":76924,"formula":"=SUM(B5:B10)"}] }],
  "cell_styles": [{ "file_index": 0, "sheet_name": "Лист1", "changes": [{"row":10,"col":1,"bgColor":"FFFFFF00"}] }],
  "doc_update": { "doc_name": "Описание проект факт", "text": "Полный обновлённый текст документа целиком" }
}
═══════════════════════════════════════════════

━━━ РЕЖИМ 1: СОЗДАНИЕ НОВОЙ ТАБЛИЦЫ / ГРАФИКА (new_sheet) ━━━

Используй new_sheet когда задание: "создай таблицу", "сделай сводную", "построй график", "подготовь лист по образцу", "сформируй отчёт", "обнови данные в таблице", "заполни по шаблону".

АЛГОРИТМ:
1. Определи источник данных — какой лист/файл содержит нужные данные (по названию листа в задании).
2. Найди нужные строки и столбцы в данных источника, используя заголовки столбцов как смысловые метки.
3. Если передан образец (изображение) — воспроизведи его структуру: заголовки, строки, порядок столбцов.
4. Если образца нет — построй логичную таблицу с понятными заголовками.
5. Верни new_sheet с полными данными: первая строка — заголовки, остальные — данные.
6. Числа передавай как числа (не строки). Пустые ячейки — null.
7. Название листа (sheet_name) — короткое и понятное, на русском.

ПРАВИЛА new_sheet:
- new_sheet заменяет/создаёт лист целиком — не используй для точечных правок.
- file_index — индекс файла из контекста (0 = первый файл, 1 = второй, ...).
- Максимум 500 строк данных.
- Если нужно несколько таблиц — верни несколько new_sheet в массиве (поле "new_sheets": [...]).

━━━ РЕЖИМ 2: ТОЧЕЧНОЕ ОБНОВЛЕНИЕ ЯЧЕЕК (cell_updates) ━━━

Используй cell_updates когда: "пропиши формулы", "вставь суммы", "обнови ячейку", "добавь строку итогов".
- row и col — 0-based (строка Excel 1 = row 0, столбец A = col 0)
- Если есть formula — также укажи value = вычисленный результат.

━━━ РЕЖИМ 3: СТИЛИ (cell_styles) ━━━

Используй для покраски/жирности ячеек.
bgColor AARRGGBB: жёлтый=FFFFFF00, оранжевый=FFFFA500, зелёный=FF92D050, красный=FFFF0000, синий=FF4472C4

━━━ РЕЖИМ 4: ОБНОВЛЕНИЕ ТЕКСТОВОГО ДОКУМЕНТА (doc_update) ━━━

Используй doc_update когда задание: "обнови текст", "перепиши описание", "замени период", "обнови аналитику в документе", "выдай текстом в файле".

АЛГОРИТМ:
1. Найди документ по имени (частичное совпадение с doc_name).
2. Возьми оригинальный текст документа из контекста (он передаётся как =Документ «имя»=).
3. Обнови только нужные части текста — цифры, периоды, выводы.
4. Верни ПОЛНЫЙ текст документа целиком в поле "text" внутри doc_update.

Формат:
{"text": "Обновил текст документа «Описание»", "doc_update": {"doc_name": "Описание проект факт", "text": "...полный обновлённый текст..."}}

ПРАВИЛА:
- doc_name — часть имени файла (регистр не важен).
- text внутри doc_update — ВЕСЬ текст документа, не фрагмент.
- Сохраняй структуру и стиль оригинала, меняй только данные.

━━━ ЗАПРОС УТОЧНЕНИЙ (ask_user) ━━━

Используй ask_user когда данных НЕДОСТАТОЧНО для выполнения задания:
- Не найден нужный лист или файл
- Непонятно какие строки/объекты включить в таблицу
- Нет образца а структура таблицы неоднозначна
- Нужно уточнить единицы измерения, год, категорию запасов

Формат: {"text": "Не хватает данных", "ask_user": "Уточни: какой год данных использовать — 2023 или 2024?"}
После получения ответа — выполни задание.
НЕ спрашивай о том что уже очевидно из контекста. Задавай ОДИН конкретный вопрос.

━━━ КАК ЧИТАТЬ КОНТЕКСТ ЛИСТОВ ━━━

Данные каждого листа передаются в секциях:

1. ОБЪЕДИНЁННЫЕ ЯЧЕЙКИ — диапазоны заголовков.
2. ЗАГОЛОВКИ СТОЛБЦОВ — смысловая цепочка для каждого столбца (читай сверху вниз).
   Пример: K = "Балансовые запасы / извлекаемые / B2"
   Буква столбца (K, L...) НЕ означает смысл — только цепочка.
3. ДАННЫЕ — ячейки: адрес + значение + [цепочка столбца].
   Пример: L15=0.281[КИН/A+B1]
   [РАЗДЕЛ] строки — разделители, не данные.

Листы в контексте помечены:
- [АКТИВНЫЙ ЛИСТ] — тот что открыт у пользователя, обычно целевой для записи результата
- [лист: "Название"] — остальные листы того же файла = источники данных
- Файлы с ролью [ОБРАЗЕЦ] — берёшь из них структуру/формат

АЛГОРИТМ поиска нужного столбца (САМОЕ ВАЖНОЕ):
ШАГ 1. Найди нужный лист по имени из задания.
ШАГ 2. Смотри секцию "ЦЕПОЧКИ ЗАГОЛОВКОВ" — там каждый столбец описан как: J(col=9) = "2023 / Проект"
ШАГ 3. Найди столбец по ПОЛНОЙ цепочке — год + подзаголовок. Например "2023 / Проект" → col=9 → буква J.
ШАГ 4. Найди нужные строки по значению в столбце A (названия показателей).
ШАГ 5. row и col — всегда 0-based. Строка Excel 1 = row 0. Столбец A = col 0, J = col 9, N = col 13.

КАК ЧИТАТЬ TSV-ФОРМАТ (большие таблицы):
- Строка "СТОЛБЦЫ:" — буква и номер каждого столбца: A=col0, B=col1, ...
- Строки "ЗАГОЛОВКИ:" — несколько строк, объединённые ячейки раскрыты (значение повторяется)
- Строки "ЦЕПОЧКИ ЗАГОЛОВКОВ:" — итоговый путь каждого столбца: J(col=9) = "2023 / Проект"
- Строки "ДАННЫЕ:" — первая колонка = номер строки Excel, остальные = значения столбцов A, B, C...
  Пример строки данных: "28\tДействующий фонд\t\t224\t\t263\t\t282\t\t316"
  Здесь строка Excel 28, значение в col=9 (J) = 316 — это "2023 / Проект"

━━━ КРИТИЧЕСКИЕ ПРАВИЛА ━━━
- ВСЕГДА используй "ЦЕПОЧКИ ЗАГОЛОВКОВ" чтобы найти нужный столбец. Никогда не угадывай букву.
- row: в данных каждая строка помечена явно: "28(row=27)" — используй ТОЛЬКО число в скобках (row=27), не 28.
- col: используй число после "col=" в цепочке. J(col=9) → col=9. L(col=11) → col=11.
- ПРИМЕР: строка "28(row=27): J28=316[2023/Проект]" → row=27, col=9.
- Числа в data — числами, не строками.
- Если данных нет в переданном контексте — спроси через ask_user, не выдумывай.`;

// Вспомогательный компонент: карточка с вопросом ИИ
// (используется ниже в JSX)



type CellStyleChange = { row: number; col: number; bgColor?: string; fontColor?: string; bold?: boolean };
type CellStyleMutation = { fileId: string; sheetName: string; changes: CellStyleChange[] };

function DocTextBlock({ text, name }: { text: string; name?: string }) {
  const [copied, setCopied] = useState(false);
  return (
    <div className="mt-3 rounded-xl border border-primary/30 overflow-hidden" style={{ background: "rgba(52,211,153,0.04)" }}>
      <div className="flex items-center justify-between px-3 py-2 border-b border-primary/20">
        <span className="text-[10px] font-semibold text-primary flex items-center gap-1.5">
          <Icon name="FileText" size={11} />
          {name || "Текст документа"}
        </span>
        <button
          className="text-[10px] px-2 py-0.5 rounded-md border border-primary/40 text-primary hover:bg-primary/10 transition-colors flex items-center gap-1"
          onClick={() => { navigator.clipboard.writeText(text); setCopied(true); setTimeout(() => setCopied(false), 2000); }}
        >
          <Icon name={copied ? "Check" : "Copy"} size={10} />
          {copied ? "Скопировано!" : "Скопировать"}
        </button>
      </div>
      <pre className="px-3 py-2.5 text-[11px] leading-relaxed text-foreground/90 whitespace-pre-wrap font-sans max-h-96 overflow-y-auto">{text}</pre>
    </div>
  );
}

function extractJson(raw: string): Record<string, unknown> {
  // 1. Убираем markdown-блоки (```json ... ``` в любом месте)
  let s = raw.trim();
  s = s.replace(/```(?:json)?\s*([\s\S]*?)```/g, "$1").trim();
  // 2. Прямой парс
  try { return JSON.parse(s); } catch { /* fall through */ }
  // 3. Ищем ПОСЛЕДНИЙ полный JSON-объект (жадно, от первой { до последней })
  const start = s.indexOf("{");
  const end = s.lastIndexOf("}");
  if (start !== -1 && end > start) {
    try { return JSON.parse(s.slice(start, end + 1)); } catch { /* fall through */ }
  }
  // 4. Пытаемся найти любой JSON-объект регуляркой (на случай вложенных)
  const m = s.match(/\{[\s\S]*\}/);
  if (m) { try { return JSON.parse(m[0]); } catch { /* fall through */ } }
  // 5. Ничего не нашли — возвращаем текст как есть, но помечаем что это не JSON
  return { text: raw || "ИИ вернул пустой ответ", _raw_fallback: true };
}

type CellValueChange = { row: number; col: number; value: CellValue; formula?: string };
type CellValueMutation = { fileId: string; sheetName: string; changes: CellValueChange[] };

async function callAi(
  prompt: string,
  files: ExcelFile[],
  settings: AiSettings,
  images: ChatImage[],
  docs: DocFile[],
  activePrompts: PromptPreset[],
  history: ChatMessage[],
  knowledgeEntries: KnowledgeEntry[]
): Promise<{
  text: string;
  ask_user?: string;
  mutations?: { fileId: string; sheetName: string; data: CellValue[][] }[];
  styleMutations?: CellStyleMutation[];
  valueMutations?: CellValueMutation[];
  docUpdates?: { docId: string; docName: string; text: string }[];
}> {
  const selectedModel = settings.model === "__custom__" ? settings.customModel : settings.model;
  // Модели без поддержки vision — при наличии картинок автоматически переключаем на gpt-4o-mini
  const VISION_UNSUPPORTED = ["deepseek/deepseek-chat", "deepseek/deepseek-r1"];
  const effectiveModel = (images.length > 0 && VISION_UNSUPPORTED.includes(selectedModel))
    ? "openai/gpt-4o-mini"
    : selectedModel;
  const baseUrl = settings.baseUrl.replace(/\/$/, "");

  // Строим контекст файлов — умный режим экономии токенов:
  // • Активный лист → до 200 строк (крупные листы — компактный TSV, мелкие — подробный формат)
  // • Упомянутый в запросе → до 120 строк
  // • Остальные листы → только 5 строк превью
  const ACTIVE_MAX_ROWS = 200;
  const MENTIONED_MAX_ROWS = 120;

  // Определяем какие листы упомянуты в тексте запроса
  const promptLower = (prompt || "").toLowerCase();
  const isMentioned = (sheetName: string) =>
    promptLower.includes(sheetName.toLowerCase()) ||
    promptLower.includes(sheetName.toLowerCase().slice(0, 8));

  const contextParts: string[] = [];
  files.forEach((f, fi) => {
    const roleLabel = f.role === "main" ? " [ОСНОВНОЙ]" : f.role === "reference" ? " [ОБРАЗЕЦ]" : "";
    contextParts.push(`=== Excel-файл ${fi} «${f.name}»${roleLabel} ===`);
    contextParts.push(`Листы: ${f.sheets.map((s, si) => `«${s.name}»${si === f.activeSheet ? " (активный)" : ""}`).join(", ")}`);
    f.sheets.forEach((sh, si) => {
      const isActive = si === f.activeSheet;
      const mentioned = isMentioned(sh.name);
      const totalRows = sh.cells.filter(r => r.some(c => c.v !== null)).length;
      const realCols = getRealColCount(sh);

      if (isActive) {
        // Активный лист — полный контекст
        contextParts.push(`--- Лист «${sh.name}» [АКТИВНЫЙ — сюда записывать результат, строк: ${totalRows}, столбцов: ${realCols}] ---`);
        contextParts.push(sheetToText(sh, true, ACTIVE_MAX_ROWS));
      } else if (mentioned) {
        // Лист упомянут в запросе — тоже полный контекст
        contextParts.push(`--- Лист «${sh.name}» [источник данных, строк: ${totalRows}] ---`);
        contextParts.push(sheetToText(sh, true, MENTIONED_MAX_ROWS));
      } else {
        // Остальные листы — только заголовки столбцов (без данных строк)
        contextParts.push(`--- Лист «${sh.name}» [не активный, строк: ${totalRows} — данные не переданы, упомяни имя листа в запросе чтобы получить доступ] ---`);
        contextParts.push(sheetToText(sh, false)); // краткий TSV-превью, 5 строк
      }
    });
  });

  // Добавляем DOCX-документы (текст/HTML)
  const MAX_DOC_CHARS = 30000;
  docs.filter(d => d.type === "docx").forEach((doc, di) => {
    const roleLabel = doc.role ? ` [${doc.role.toUpperCase()}]` : "";
    contextParts.push(`=== Word-документ ${di} «${doc.name}»${roleLabel} ===`);
    const text = doc.text.length > MAX_DOC_CHARS
      ? doc.text.slice(0, MAX_DOC_CHARS) + `\n[... документ обрезан, показано ${MAX_DOC_CHARS} из ${doc.text.length} символов]`
      : doc.text;
    contextParts.push(text);
  });

  // PDF-страницы — собираем URL с учётом диапазона
  const pdfImageUrls: string[] = [];
  docs.filter(d => d.type === "pdf" && d.pageImageUrls && d.pageImageUrls.length > 0).forEach((doc) => {
    const roleLabel = doc.role ? ` [${doc.role.toUpperCase()}]` : "";
    const allUrls = doc.pageImageUrls ?? [];
    const from = Math.max(1, doc.pageFrom ?? 1);
    const to = Math.min(allUrls.length, doc.pageTo ?? allUrls.length);
    const selectedUrls = allUrls.slice(from - 1, to);
    const rangeNote = (doc.pageFrom || doc.pageTo) ? ` стр.${from}–${to}` : ` все ${allUrls.length} стр.`;
    contextParts.push(`=== PDF «${doc.name}»${roleLabel} (${doc.pageCount} стр.,${rangeNote} в контексте) ===`);
    pdfImageUrls.push(...selectedUrls);
  });

  const fullContext = contextParts.join("\n");
  (window as Window & { __datamind_last_context?: string }).__datamind_last_context = fullContext;

  // База знаний + активные промпты → системный промпт
  const knowledgeText = formatKnowledgeForAI(knowledgeEntries);
  const activePromptsText = activePrompts.filter(p => p.enabled).map(p => p.text).join("\n\n");
  const effectiveSystemPrompt = [knowledgeText, activePromptsText, SYSTEM_PROMPT]
    .filter(Boolean).join("\n\n---\n\n");

  const textBlock = `ДАННЫЕ ФАЙЛОВ:\n${fullContext}\n\nЗАДАНИЕ: ${prompt || "(см. изображения)"}

⚠️ ОБЯЗАТЕЛЬНО: твой ответ должен быть ТОЛЬКО валидным JSON-объектом. Никакого текста до или после. Никаких \`\`\`json\`\`\` обёрток. Начни ответ с символа { и заверши символом }.`;

  // Собираем все изображения: скриншоты + PDF рабочих документов + PDF из базы знаний
  type ContentPart = { type: "text"; text: string } | { type: "image_url"; image_url: { url: string } };
  const allImageParts: ContentPart[] = [];
  for (const img of images) {
    allImageParts.push({ type: "image_url", image_url: { url: img.dataUrl } });
  }
  for (const url of pdfImageUrls) {
    allImageParts.push({ type: "image_url", image_url: { url } });
  }
  // PDF из базы знаний
  for (const kbPdf of getKnowledgePdfImages(knowledgeEntries)) {
    for (const url of kbPdf.urls) {
      allImageParts.push({ type: "image_url", image_url: { url } });
    }
  }

  let userContent: string | ContentPart[];
  if (allImageParts.length > 0) {
    userContent = [{ type: "text", text: textBlock } as ContentPart, ...allImageParts];
  } else {
    userContent = textBlock;
  }

  // История переписки — последние 6 пар (user+ai), без images чтобы не раздувать контекст
  const historyMessages: { role: "user" | "assistant"; content: string }[] = [];
  const recentHistory = history.slice(-12).filter(m => m.role === "user" || m.role === "ai");
  for (const m of recentHistory) {
    historyMessages.push({ role: m.role === "ai" ? "assistant" : "user", content: m.text });
  }

  // Если есть PDF-картинки — форсируем vision-модель
  const hasPdfImages = pdfImageUrls.length > 0;
  const VISION_UNSUPPORTED_FOR_PDF = ["deepseek/deepseek-chat", "deepseek/deepseek-r1"];
  const finalModel = (hasPdfImages && VISION_UNSUPPORTED_FOR_PDF.includes(effectiveModel))
    ? "openai/gpt-4o"
    : effectiveModel;

  // Модели поддерживающие принудительный JSON-режим (response_format)
  // Qwen, Llama и ряд других НЕ поддерживают — передавать им не нужно (вернут ошибку 400)
  const JSON_MODE_MODELS = [
    "gpt-4o", "gpt-4o-mini", "gpt-4-turbo", "gpt-4",
    "openai/gpt-4o", "openai/gpt-4o-mini",
    "deepseek/deepseek-chat", "deepseek/deepseek-r1",
  ];
  const supportsJsonMode = JSON_MODE_MODELS.some(m => finalModel.includes(m.split("/").pop()!))
    && allImageParts.length === 0; // vision + json_object несовместимы

  const resp = await fetch(`${baseUrl}/chat/completions`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${settings.apiKey}`,
    },
    body: JSON.stringify({
      model: finalModel,
      messages: [
        { role: "system", content: effectiveSystemPrompt },
        ...historyMessages,
        { role: "user", content: userContent },
      ],
      max_tokens: 4000,
      temperature: 0.1,
      ...(supportsJsonMode ? { response_format: { type: "json_object" } } : {}),
      ...(settings.reasoningEffort !== "none"
        ? { reasoning_effort: settings.reasoningEffort }
        : {}),
    }),
  });

  if (!resp.ok) {
    const err = await resp.json().catch(() => ({})) as { error?: { message?: string } };
    const errMsg = err.error?.message || "";
    if (resp.status === 503 || resp.status === 502 || resp.status === 504) {
      throw new Error(`503 Сервер ИИ временно недоступен — подожди минуту и попробуй снова`);
    }
    if (resp.status === 401) throw new Error(`401 Неверный API ключ`);
    if (resp.status === 429) throw new Error(`429 Превышен лимит запросов — подожди немного`);
    if (resp.status === 400 && errMsg.includes("response_format")) {
      // Модель не поддерживает json_object — это нормально, не бросаем ошибку
      // Это не должно доходить сюда, но на случай если достигло
      throw new Error(`Модель не поддерживает режим JSON — попробуй GPT-4o или DeepSeek`);
    }
    throw new Error(errMsg || `HTTP ${resp.status}`);
  }

  const json = await resp.json() as { choices: { message: { content: string } }[] };
  const raw = json.choices?.[0]?.message?.content ?? "{}";
  console.log("[AI RAW]", raw.slice(0, 2000));
  const result = extractJson(raw);
  if (result.cell_styles) console.log("[AI CELL_STYLES]", JSON.stringify(result.cell_styles));
  if (result.cell_updates) console.log("[AI CELL_UPDATES]", JSON.stringify(result.cell_updates));

  // Модель ответила текстом вместо JSON — показываем понятную ошибку
  if (result._raw_fallback) {
    const hint = "⚠️ Модель ответила текстом вместо JSON. Попробуй: 1) Переформулировать запрос, 2) Сменить модель на GPT-4o или DeepSeek, 3) Уточнить задание.";
    return { text: hint };
  }

  // ask_user — ИИ запрашивает уточнение перед выполнением
  if (result.ask_user) {
    return {
      text: (result.text as string) || "Нужно уточнение",
      ask_user: result.ask_user as string,
    };
  }

  const mutations: { fileId: string; sheetName: string; data: CellValue[][] }[] = [];

  // Один лист (new_sheet)
  if (result.new_sheet) {
    const ns = result.new_sheet as { file_index?: number; sheet_name: string; data: CellValue[][] };
    const targetFile = files[ns.file_index ?? 0] ?? files[0];
    if (targetFile) mutations.push({ fileId: targetFile.id, sheetName: ns.sheet_name, data: ns.data });
  }
  // Несколько листов (new_sheets — массив)
  if (Array.isArray(result.new_sheets)) {
    for (const ns of result.new_sheets as { file_index?: number; sheet_name: string; data: CellValue[][] }[]) {
      const targetFile = files[ns.file_index ?? 0] ?? files[0];
      if (targetFile && ns.sheet_name && Array.isArray(ns.data)) {
        mutations.push({ fileId: targetFile.id, sheetName: ns.sheet_name, data: ns.data });
      }
    }
  }

  const styleMutations: CellStyleMutation[] = [];
  if (Array.isArray(result.cell_styles)) {
    for (const cs of result.cell_styles as { file_index?: number; sheet_name: string; changes: CellStyleChange[] }[]) {
      const targetFile = files[cs.file_index ?? 0] ?? files[0];
      if (targetFile && cs.sheet_name && Array.isArray(cs.changes)) {
        // Находим лист чтобы знать реальный размер — для проверки row/col
        const targetSheet = targetFile.sheets.find(s => s.name === cs.sheet_name);
        const maxRows = targetSheet?.cells.length ?? 9999;
        const maxCols = targetSheet?.cells[0]?.length ?? 9999;

        const safeChanges = cs.changes
          .filter(ch => {
            const r = ch.row;
            const c = ch.col;
            // Отбрасываем явно невалидные (отрицательные)
            return typeof r === "number" && typeof c === "number" && r >= 0 && c >= 0;
          })
          .map(ch => {
            let { row, col } = ch;
            // Если ИИ передал 1-based (row > maxRows но row-1 < maxRows) — корректируем
            if (row >= maxRows && row - 1 < maxRows) row = row - 1;
            if (col >= maxCols && col - 1 < maxCols) col = col - 1;
            return { ...ch, row, col };
          });

        if (safeChanges.length > 0) {
          styleMutations.push({ fileId: targetFile.id, sheetName: cs.sheet_name, changes: safeChanges });
        }
      }
    }
  }

  // cell_updates — точечное изменение значений ячеек без замены всего листа
  const valueMutations: CellValueMutation[] = [];
  if (Array.isArray(result.cell_updates)) {
    for (const cu of result.cell_updates as { file_index?: number; sheet_name: string; changes: CellValueChange[] }[]) {
      const targetFile = files[cu.file_index ?? 0] ?? files[0];
      if (targetFile && cu.sheet_name && Array.isArray(cu.changes)) {
        const safeChanges = cu.changes.filter(ch =>
          typeof ch.row === "number" && typeof ch.col === "number" && ch.row >= 0 && ch.col >= 0
        );
        if (safeChanges.length > 0) {
          valueMutations.push({ fileId: targetFile.id, sheetName: cu.sheet_name, changes: safeChanges });
        }
      }
    }
  }

  // doc_update — обновление текстового документа
  const docUpdates: { docId: string; docName: string; text: string }[] = [];
  if (result.doc_update) {
    const du = result.doc_update as { doc_name: string; text: string };
    if (du.doc_name && du.text) {
      const nameLower = du.doc_name.toLowerCase();
      const targetDoc = docs.find(d => d.name.toLowerCase().includes(nameLower));
      if (targetDoc) {
        docUpdates.push({ docId: targetDoc.id, docName: targetDoc.name, text: du.text });
      }
    }
  }

  return {
    text: (result.text as string) || "Готово!",
    mutations: mutations.length ? mutations : undefined,
    styleMutations: styleMutations.length ? styleMutations : undefined,
    valueMutations: valueMutations.length ? valueMutations : undefined,
    docUpdates: docUpdates.length ? docUpdates : undefined,
  };
}

// ─── CopySheetButton ──────────────────────────────────────────────────────────
// Копирует активный лист как TSV — вставляется прямо в Excel, Google Sheets, Word

function CopySheetButton({ sheet }: { sheet: SheetData }) {
  const [state, setState] = useState<"idle" | "ok" | "err">("idle");

  const handleCopy = async () => {
    try {
      const numCols = getRealColCount(sheet);
      const rows: string[] = [];
      sheet.cells.forEach(row => {
        // Пропускаем полностью пустые строки
        if (!row.some(c => c.v !== null)) return;
        const cells = Array.from({ length: numCols }, (_, ci) => {
          const v = row[ci]?.v;
          if (v === null || v === undefined) return "";
          const s = typeof v === "number" ? String(v) : String(v);
          // Если значение содержит таб/перенос — оборачиваем в кавычки
          return s.includes("\t") || s.includes("\n") ? `"${s.replace(/"/g, '""')}"` : s;
        });
        rows.push(cells.join("\t"));
      });
      const tsv = rows.join("\n");
      await navigator.clipboard.writeText(tsv);
      setState("ok");
      setTimeout(() => setState("idle"), 2000);
    } catch {
      setState("err");
      setTimeout(() => setState("idle"), 2000);
    }
  };

  return (
    <button
      onClick={handleCopy}
      title="Скопировать лист — вставляется в Excel, Google Sheets, Word"
      className={`flex items-center gap-1 px-2 py-1 rounded-lg text-[11px] border transition-all ${
        state === "ok"
          ? "text-primary border-primary/40 bg-primary/10"
          : state === "err"
          ? "text-red-400 border-red-400/30 bg-red-400/10"
          : "text-muted-foreground border-transparent hover:text-foreground hover:bg-secondary hover:border-border/40"
      }`}>
      <Icon
        name={state === "ok" ? "CheckCheck" : state === "err" ? "X" : "Copy"}
        size={12}
      />
      <span className="hidden md:inline">
        {state === "ok" ? "Скопировано!" : state === "err" ? "Ошибка" : "Скопировать"}
      </span>
    </button>
  );
}

// ─── AI Question Card ──────────────────────────────────────────────────────────
// Карточка с вопросом ИИ — вставляется в чат, пользователь отвечает прямо в ней

// ─── Cell Editor ──────────────────────────────────────────────────────────────

interface CellEditorProps {
  value: CellValue;
  onCommit: (v: CellValue) => void;
  onCancel: () => void;
}
function CellEditor({ value, onCommit, onCancel }: CellEditorProps) {
  const [v, setV] = useState(String(value ?? ""));
  const ref = useRef<HTMLInputElement>(null);
  useEffect(() => { ref.current?.focus(); ref.current?.select(); }, []);
  return (
    <input ref={ref} value={v} onChange={(e) => setV(e.target.value)}
      onKeyDown={(e) => {
        if (e.key === "Enter") onCommit(v === "" ? null : isNaN(Number(v)) ? v : Number(v));
        if (e.key === "Escape") onCancel();
      }}
      onBlur={() => onCommit(v === "" ? null : isNaN(Number(v)) ? v : Number(v))}
      className="absolute inset-0 w-full h-full px-1 text-xs bg-[rgba(52,211,153,0.12)] border border-primary outline-none text-foreground font-mono z-10"
      style={{ fontFamily: "'IBM Plex Mono', monospace" }}
    />
  );
}

// ─── Spreadsheet renderer ─────────────────────────────────────────────────────

interface SpreadsheetProps {
  sheet: SheetData;
  selected: { row: number; col: number } | null;
  editing: { row: number; col: number } | null;
  onSelect: (r: number, c: number) => void;
  onEdit: (r: number, c: number) => void;
  onCommit: (r: number, c: number, v: CellValue) => void;
  onCancelEdit: () => void;
  tableRef: React.RefObject<HTMLDivElement>;
  onDrop: (files: FileList) => void;
}

function Spreadsheet({ sheet, selected, editing, onSelect, onEdit, onCommit, onCancelEdit, tableRef, onDrop }: SpreadsheetProps) {
  const VISIBLE_ROWS = Math.min(sheet.cells.length, 200);
  const VISIBLE_COLS = Math.min(sheet.cells[0]?.length ?? 0, 60);

  // Build merge lookup: "r,c" → MergeRange (only for top-left cell)
  // Also build skip set for cells that are covered by a merge
  const mergeMap = new Map<string, MergeRange>();
  const skipSet = new Set<string>();
  for (const m of sheet.merges) {
    mergeMap.set(`${m.r1},${m.c1}`, m);
    for (let r = m.r1; r <= m.r2; r++) {
      for (let c = m.c1; c <= m.c2; c++) {
        if (r !== m.r1 || c !== m.c1) skipSet.add(`${r},${c}`);
      }
    }
  }

  return (
    <div ref={tableRef} className="flex-1 overflow-auto scrollbar-thin" style={{ background: "hsl(220,16%,5.5%)" }}
      onDragOver={(e) => e.preventDefault()}
      onDrop={(e) => { e.preventDefault(); onDrop(e.dataTransfer.files); }}>
      <table className="border-collapse" style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: "12px", tableLayout: "fixed" }}>
        <thead>
          <tr>
            <th className="sticky left-0 top-0 z-30 text-center text-[11px] text-muted-foreground font-normal select-none"
              style={{ background: "hsl(220,14%,8%)", minWidth: 40, width: 40, borderBottom: "1px solid rgba(255,255,255,0.1)", borderRight: "1px solid rgba(255,255,255,0.07)" }} />
            {Array.from({ length: VISIBLE_COLS }, (_, ci) => (
              <th key={ci} className="sticky top-0 z-20 text-center text-[11px] text-muted-foreground font-medium select-none px-1 py-1"
                style={{ background: "hsl(220,14%,8%)", width: sheet.colWidths[ci] ?? 100, minWidth: sheet.colWidths[ci] ?? 40, borderBottom: "1px solid rgba(255,255,255,0.1)", borderRight: "1px solid rgba(255,255,255,0.07)" }}>
                {colLetter(ci)}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {Array.from({ length: VISIBLE_ROWS }, (_, ri) => (
            <tr key={ri}>
              <td className="sticky left-0 z-10 text-center text-[11px] text-muted-foreground select-none px-1"
                style={{ background: "hsl(220,14%,8%)", width: 40, height: sheet.rowHeights[ri] ?? 22, borderBottom: "1px solid rgba(255,255,255,0.05)", borderRight: "1px solid rgba(255,255,255,0.07)" }}>
                {ri + 1}
              </td>
              {Array.from({ length: VISIBLE_COLS }, (_, ci) => {
                const key = `${ri},${ci}`;
                if (skipSet.has(key)) return null;

                const merge = mergeMap.get(key);
                const colSpan = merge ? merge.c2 - merge.c1 + 1 : 1;
                const rowSpan = merge ? merge.r2 - merge.r1 + 1 : 1;

                const cell = sheet.cells[ri]?.[ci] ?? { v: null };
                const st = cell.s;
                const isSelected = selected?.row === ri && selected?.col === ci;
                const isEditing = editing?.row === ri && editing?.col === ci;
                const displayVal = cell.w ?? (cell.v !== null && cell.v !== undefined ? String(cell.v) : "");

                const cellStyle: React.CSSProperties = {
                  height: sheet.rowHeights[ri] ?? 22,
                  minWidth: sheet.colWidths[ci] ?? 40,
                  position: "relative",
                  padding: "0 4px",
                  cursor: "cell",
                  userSelect: "none",
                  verticalAlign: st?.vAlign ?? "middle",
                  textAlign: st?.hAlign === "general"
                    ? (typeof cell.v === "number" ? "right" : "left")
                    : (st?.hAlign ?? (typeof cell.v === "number" ? "right" : "left")),
                  whiteSpace: st?.wrapText ? "pre-wrap" : "nowrap",
                  overflow: "hidden",
                  backgroundColor: isSelected ? "rgba(52,211,153,0.1)" : (st?.bgColor ?? "transparent"),
                  color: st?.color ?? (typeof cell.v === "number" ? "hsl(158,60%,75%)" : "hsl(210,20%,88%)"),
                  fontWeight: st?.bold ? "bold" : undefined,
                  fontStyle: st?.italic ? "italic" : undefined,
                  textDecoration: st?.underline ? "underline" : undefined,
                  fontSize: st?.fontSize ? `${st.fontSize}pt` : "12px",
                  borderTop: borderCss(st?.borderTop),
                  borderBottom: borderCss(st?.borderBottom),
                  borderLeft: borderCss(st?.borderLeft),
                  borderRight: borderCss(st?.borderRight),
                  outline: isSelected ? "1px solid hsl(158,64%,52%)" : undefined,
                  outlineOffset: "-1px",
                  zIndex: isSelected ? 10 : undefined,
                };

                return (
                  <td key={ci} colSpan={colSpan} rowSpan={rowSpan}
                    style={cellStyle}
                    onClick={() => { onSelect(ri, ci); }}
                    onDoubleClick={() => { onSelect(ri, ci); onEdit(ri, ci); }}>
                    {isEditing ? (
                      <CellEditor value={cell.v} onCommit={(v) => onCommit(ri, ci, v)} onCancel={onCancelEdit} />
                    ) : (
                      <span style={{ display: "block", overflow: "hidden", textOverflow: "ellipsis" }}>
                        {displayVal}
                      </span>
                    )}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// ─── Chart View ───────────────────────────────────────────────────────────────

const CHART_COLORS = [
  "hsl(158,64%,52%)", "hsl(43,96%,58%)", "hsl(213,100%,65%)",
  "hsl(280,65%,65%)", "hsl(10,80%,60%)", "hsl(170,60%,50%)",
  "hsl(330,70%,60%)", "hsl(55,90%,55%)",
];

interface ChartViewProps {
  sheet: SheetData;
  activeFileId: string | null;
  onChangeType: (t: ChartType | undefined) => void;
}

function ChartView({ sheet, onChangeType }: ChartViewProps) {
  const dataRows = sheet.cells.filter(r => r[0]?.v != null && r[1]?.v != null);
  const rows = dataRows.slice(1);
  const chartData = rows
    .map(r => ({ name: String(r[0].v), value: Number(r[1].v) }))
    .filter(r => !isNaN(r.value));

  const ttStyle = { background: "hsl(220,14%,10%)", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 8, fontSize: 12 };
  const axTick = { fill: "hsl(215,14%,55%)", fontSize: 11 };

  return (
    <div className="flex-1 flex flex-col items-center p-8 overflow-auto" style={{ background: "hsl(220,16%,5.5%)" }}>
      <div className="w-full max-w-3xl">
        <div className="flex items-center justify-between mb-6">
          <div>
            <h2 className="text-base font-semibold text-foreground">{sheet.name}</h2>
            <p className="text-xs text-muted-foreground mt-0.5">{chartData.length} значений</p>
          </div>
          <div className="flex gap-1 flex-wrap">
            {(["pie", "bar", "line"] as ChartType[]).map(t => (
              <button key={t} onClick={() => onChangeType(t)}
                className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-all ${sheet.chartType === t ? "btn-primary" : "bg-secondary text-muted-foreground hover:text-foreground"}`}>
                {t === "pie" ? "Круговой" : t === "bar" ? "Столбчатый" : "Линейный"}
              </button>
            ))}
            <button onClick={() => onChangeType(undefined)}
              className="px-3 py-1.5 rounded-lg text-xs text-muted-foreground hover:text-foreground bg-secondary transition-all ml-2">
              Таблица
            </button>
          </div>
        </div>
        <ResponsiveContainer width="100%" height={420}>
          {sheet.chartType === "pie" ? (
            <PieChart>
              <Pie data={chartData} cx="50%" cy="50%" outerRadius={160} innerRadius={60} paddingAngle={3} dataKey="value"
                label={({ name, percent }) => `${name} ${(percent * 100).toFixed(1)}%`} labelLine>
                {chartData.map((_, ci) => <Cell key={ci} fill={CHART_COLORS[ci % CHART_COLORS.length]} />)}
              </Pie>
              <Tooltip contentStyle={ttStyle} />
              <Legend wrapperStyle={{ fontSize: 12, color: "hsl(215,14%,65%)" }} />
            </PieChart>
          ) : sheet.chartType === "bar" ? (
            <BarChart data={chartData} margin={{ top: 10, right: 20, left: 0, bottom: 40 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.05)" />
              <XAxis dataKey="name" tick={axTick} angle={-30} textAnchor="end" />
              <YAxis tick={axTick} />
              <Tooltip contentStyle={ttStyle} />
              <Bar dataKey="value" radius={[4, 4, 0, 0]}>
                {chartData.map((_, ci) => <Cell key={ci} fill={CHART_COLORS[ci % CHART_COLORS.length]} />)}
              </Bar>
            </BarChart>
          ) : (
            <LineChart data={chartData} margin={{ top: 10, right: 20, left: 0, bottom: 40 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.05)" />
              <XAxis dataKey="name" tick={axTick} angle={-30} textAnchor="end" />
              <YAxis tick={axTick} />
              <Tooltip contentStyle={ttStyle} />
              <Line type="monotone" dataKey="value" stroke="hsl(158,64%,52%)" strokeWidth={2} dot={{ fill: "hsl(158,64%,52%)", r: 4 }} />
            </LineChart>
          )}
        </ResponsiveContainer>
      </div>
    </div>
  );
}

// ─── HMR state preservation ───────────────────────────────────────────────────
// При горячей замене модуля (HMR) React-состояние сбрасывается.
// Сохраняем файлы и чат в window чтобы восстановить после перезагрузки модуля.

declare global {
  interface Window {
    __datamind_files?: ExcelFile[];
    __datamind_activeFileId?: string | null;
    __datamind_messages?: ChatMessage[];
  }
}

// ─── Main Component ───────────────────────────────────────────────────────────

export default function Index() {
  const [files, setFiles] = useState<ExcelFile[]>(() => window.__datamind_files ?? []);
  const [activeFileId, setActiveFileId] = useState<string | null>(() => window.__datamind_activeFileId ?? null);
  const [editing, setEditing] = useState<{ row: number; col: number } | null>(null);
  const [selected, setSelected] = useState<{ row: number; col: number } | null>(null);
  const [messages, setMessages] = useState<ChatMessage[]>(() => window.__datamind_messages ?? [{
    role: "ai",
    text: "Загрузи Excel, Word или PDF файлы. Я вижу всё содержимое и могу выполнять задания — анализировать, сверять данные, актуализировать таблицы и текст.\n\nМожно прикрепить **скриншот графика** — я пойму как его воспроизвести на твоих данных.",
    ts: getTime(),
  }]);
  const [aiInput, setAiInput] = useState("");
  const [pendingImages, setPendingImages] = useState<ChatImage[]>([]);
  const [aiThinking, setAiThinking] = useState(false);
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [promptsOpen, setPromptsOpen] = useState(false);
  const navigate = useNavigate();
  const [aiSettings, setAiSettings] = useState<AiSettings>(loadSettings);
  const [, setAiError] = useState<string | null>(null);
  const [isListening, setIsListening] = useState(false);
  const [docs, setDocs] = useState<DocFile[]>([]);
  const [activeDocId, setActiveDocId] = useState<string | null>(null);
  const [prompts, setPrompts] = useState<PromptPreset[]>(loadPrompts);
  const [editingPrompt, setEditingPrompt] = useState<string | null>(null);
  const docInputRef = useRef<HTMLInputElement>(null);
  const knowledgeTxtRef = useRef<HTMLInputElement>(null);
  const knowledgeFileRef = useRef<HTMLInputElement>(null); // Excel/PDF/DOCX в базу знаний

  // База знаний
  const [knowledge, setKnowledge] = useState<KnowledgeEntry[]>([]);
  const [knowledgeOpen, setKnowledgeOpen] = useState(false);
  const [editingKbId, setEditingKbId] = useState<string | null>(null);
  const [kbCategoryFilter, setKbCategoryFilter] = useState<KnowledgeCategory | "all">("all");
  const [kbLoadingId, setKbLoadingId] = useState<string | null>(null);

  // Сессия
  const [sessionOpen, setSessionOpen] = useState(false);
  const [savedSessions, setSavedSessions] = useState<Pick<SavedSession, "id" | "name" | "savedAt">[]>([]);
  const [sessionSaving, setSessionSaving] = useState(false);

  // Загружаем базу знаний из IndexedDB при старте (с миграцией старых записей)
  useEffect(() => {
    loadKnowledge().then(entries => {
      const migrated = entries.map(e => ({
        ...e,
        category: e.category ?? "custom",
        sourceType: e.sourceType ?? "text",
      } as KnowledgeEntry));
      setKnowledge(migrated);
    }).catch(() => {});
  }, []);

  // Автосохранение сессии каждые 60 секунд
  useEffect(() => {
    const timer = setInterval(() => {
      if (files.length === 0 && docs.length === 0) return;
      doSaveSession("current", "Автосохранение").catch(() => {});
    }, 60_000);
    return () => clearInterval(timer);
   
  }, [files, docs, messages]);

  // Сохраняем в window при каждом изменении
  useEffect(() => { window.__datamind_files = files; }, [files]);
  useEffect(() => { window.__datamind_activeFileId = activeFileId; }, [activeFileId]);
  useEffect(() => { window.__datamind_messages = messages; }, [messages]);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const imageInputRef = useRef<HTMLInputElement>(null);
  const chatEndRef = useRef<HTMLDivElement>(null);
  const tableRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const recognitionRef = useRef<any>(null);

  // ── Session save/load ──
  const doSaveSession = useCallback(async (id: string, name: string) => {
    setSessionSaving(true);
    try {
      const excelFiles = await Promise.all(files.map(async (f) => {
        const wb = f.workbook;
        const buf = XLSX.write(wb, { bookType: "xlsx", type: "array", cellStyles: true }) as ArrayBuffer;
        return {
          id: f.id, name: f.name, role: f.role,
          activeSheet: f.activeSheet, buffer: buf,
        };
      }));

      const docFiles = docs.map(d => ({
        id: d.id, name: d.name, type: d.type, role: d.role,
        text: d.text, html: d.html, pageCount: d.pageCount,
        pageImageUrls: d.pageImageUrls,
        buffer: d.buffer,
      }));

      const session: SavedSession = {
        id, name, savedAt: Date.now(),
        excelFiles, docFiles,
        knowledgeBase: knowledge,
        messages: messages.slice(-30).map(m => ({ role: m.role, text: m.text, ts: m.ts })),
      };
      await saveSession(session);
    } finally {
      setSessionSaving(false);
    }
   
  }, [files, docs, knowledge, messages]);

  const doLoadSession = useCallback(async (id: string) => {
    const session = await loadSession(id);
    if (!session) return;

    // Восстанавливаем Excel
    const excelFiles = session.excelFiles.map((sf) => {
      const wb = XLSX.read(sf.buffer, { type: "array", cellStyles: true });
      const sheets = wb.SheetNames.map(name => {
        const ws = wb.Sheets[name];
        const raw = XLSX.utils.sheet_to_json<(string | number | null)[]>(ws, { header: 1, defval: null }) as (string | number | null)[][];
        const ROWS = Math.max(raw.length + 20, 50);
        const COLS = Math.max(...raw.map(r => r.length), 10);
        const cells = Array.from({ length: ROWS }, (_, ri) =>
          Array.from({ length: COLS }, (_, ci) => ({ v: raw[ri]?.[ci] ?? null }))
        );
        return { name, cells, merges: [], colWidths: Array(COLS).fill(100), rowHeights: Array(ROWS).fill(22) };
      });
      return { id: sf.id, name: sf.name, role: sf.role, sheets, activeSheet: sf.activeSheet, workbook: wb, isDirty: false };
    });

    setFiles(excelFiles);
    if (excelFiles.length) setActiveFileId(excelFiles[0].id);

    // Восстанавливаем документы
    const docFiles: DocFile[] = session.docFiles.map(df => ({
      id: df.id, name: df.name, type: df.type, role: df.role,
      text: df.text, html: df.html, pageCount: df.pageCount,
      pageImageUrls: df.pageImageUrls, buffer: df.buffer,
    }));
    setDocs(docFiles);

    // Восстанавливаем базу знаний
    if (session.knowledgeBase?.length) {
      setKnowledge(session.knowledgeBase);
      await saveKnowledge(session.knowledgeBase);
    }

    // Восстанавливаем последние сообщения
    if (session.messages?.length) {
      setMessages(session.messages.map(m => ({ ...m, refs: undefined, images: undefined })));
    }

    setSessionOpen(false);
    setMessages(prev => [...prev, {
      role: "ai", text: `Сессия **«${session.name}»** восстановлена: ${excelFiles.length} Excel, ${docFiles.length} документов.`, ts: getTime(),
    }]);
   
  }, []);

  // ── Voice input ──
  const toggleVoice = useCallback(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const SpeechRecognition = (window as any).SpeechRecognition || (window as any).webkitSpeechRecognition;
    if (!SpeechRecognition) {
      alert("Ваш браузер не поддерживает голосовой ввод. Используйте Chrome или Edge.");
      return;
    }

    if (isListening) {
      recognitionRef.current?.stop();
      setIsListening(false);
      return;
    }

    const recognition = new SpeechRecognition();
    recognitionRef.current = recognition;
    recognition.lang = "ru-RU";
    recognition.continuous = false;
    recognition.interimResults = false;

    recognition.onstart = () => setIsListening(true);
    recognition.onend = () => setIsListening(false);
    recognition.onerror = () => setIsListening(false);
    recognition.onresult = (e: { results: { transcript: string }[][] }) => {
      const transcript = e.results[0]?.[0]?.transcript ?? "";
      if (transcript) {
        setAiInput(prev => prev ? `${prev} ${transcript}` : transcript);
        textareaRef.current?.focus();
      }
    };

    recognition.start();
  }, [isListening]);

  const activeFile = files.find((f) => f.id === activeFileId) ?? null;
  const activeSheet = activeFile ? activeFile.sheets[activeFile.activeSheet] : null;

  // ── Load Excel file ──
  const loadFile = useCallback((file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target!.result as ArrayBuffer);
      const wb = XLSX.read(data, { type: "array", cellStyles: true, cellDates: true, cellNF: true });
      const id = uid();
      const ef = parseWorkbook(wb, file.name, id);
      setFiles((prev) => [...prev, ef]);
      setActiveFileId(id);
      setMessages((prev) => [...prev, {
        role: "ai",
        text: `Файл **«${file.name}»** загружен: ${ef.sheets.length} лист(а/ов), первый лист — ${ef.sheets[0]?.cells?.filter(r => r.some(c => c.v !== null)).length ?? 0} строк данных.`,
        ts: getTime(),
        refs: [file.name],
      }]);
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const handleFiles = useCallback((fileList: FileList) => {
    Array.from(fileList).forEach((f) => {
      if (f.name.match(/\.(xlsx|xls|csv)$/i)) loadFile(f);
      else if (f.name.match(/\.docx$/i)) loadDocx(f);
      else if (f.name.match(/\.pdf$/i)) loadPdf(f);
    });
  }, [loadFile]); // loadDocx и loadPdf добавятся ниже

  // ── Load DOCX ──
  const loadDocx = useCallback((file: File) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      const buf = e.target!.result as ArrayBuffer;
      const result = await mammoth.convertToHtml({ arrayBuffer: buf });
      const textResult = await mammoth.extractRawText({ arrayBuffer: buf });
      const id = crypto.randomUUID();
      const doc: DocFile = {
        id, name: file.name, type: "docx", role: null,
        text: textResult.value,
        html: result.value,
        buffer: buf,
      };
      setDocs(prev => [...prev, doc]);
      setActiveDocId(id);
      setMessages(prev => [...prev, {
        role: "ai",
        text: `Документ **«${file.name}»** загружен: ~${Math.round(textResult.value.length / 1000)} тыс. символов.`,
        ts: getTime(), refs: [file.name],
      }]);
    };
    reader.readAsArrayBuffer(file);
  }, []);

  // ── Load PDF → бэкенд рендерит страницы в PNG, модель видит как картинки ──
  const loadPdf = useCallback(async (file: File) => {
    const id = crypto.randomUUID();
    // Сразу показываем заглушку с индикатором загрузки
    setDocs(prev => [...prev, {
      id, name: file.name, type: "pdf", role: null,
      text: "", pageCount: undefined, pageImageUrls: [], loading: true,
    }]);
    setActiveDocId(id);
    setMessages(prev => [...prev, {
      role: "ai",
      text: `Конвертирую **«${file.name}»** в изображения страниц...`,
      ts: getTime(),
    }]);

    try {
      const buf = await file.arrayBuffer();
      const bytes = new Uint8Array(buf);
      let binary = "";
      for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
      const pdf_b64 = btoa(binary);

      const resp = await fetch(AI_EXCEL_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "pdf_to_images", pdf_b64, dpi: 150, max_pages: 80 }),
      });
      const result = await resp.json() as { page_urls: string[]; total_pages: number; rendered_pages: number };

      setDocs(prev => prev.map(d => d.id !== id ? d : {
        ...d,
        loading: false,
        pageCount: result.total_pages,
        pageImageUrls: result.page_urls,
        text: `PDF «${file.name}»: ${result.total_pages} страниц`,
      }));
      setMessages(prev => [...prev, {
        role: "ai",
        text: `PDF **«${file.name}»** готов: ${result.rendered_pages} стр. из ${result.total_pages}. Задай вопрос — ИИ увидит каждую страницу как изображение.`,
        ts: getTime(), refs: [file.name],
      }]);
    } catch (e) {
      setDocs(prev => prev.map(d => d.id !== id ? d : { ...d, loading: false, text: "Ошибка конвертации" }));
      setMessages(prev => [...prev, { role: "ai", text: `⚠️ Не удалось загрузить PDF: ${String(e)}`, ts: getTime() }]);
    }
  }, []);

  // ── Load image for AI ──
  const loadImage = useCallback((file: File) => {
    if (!file.type.startsWith("image/")) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      const dataUrl = e.target!.result as string;
      setPendingImages(prev => [...prev, { dataUrl, name: file.name }]);
    };
    reader.readAsDataURL(file);
  }, []);

  // ── Paste images from clipboard ──
  useEffect(() => {
    const onPaste = (e: ClipboardEvent) => {
      const items = e.clipboardData?.items;
      if (!items) return;
      for (const item of Array.from(items)) {
        if (item.type.startsWith("image/")) {
          const file = item.getAsFile();
          if (file) loadImage(file);
        }
      }
    };
    window.addEventListener("paste", onPaste);
    return () => window.removeEventListener("paste", onPaste);
  }, [loadImage]);

  // ── Cell update ──
  const updateCell = useCallback((row: number, col: number, value: CellValue) => {
    if (!activeFile) return;
    setFiles((prev) => prev.map((f) => {
      if (f.id !== activeFile.id) return f;
      const sheets = f.sheets.map((sh, si) => {
        if (si !== f.activeSheet) return sh;
        const cells = sh.cells.map((r, ri) =>
          ri === row ? r.map((c, ci) => ci === col ? { ...c, v: value, w: undefined } : c) : r
        );
        return { ...sh, cells };
      });
      return { ...f, sheets, isDirty: true };
    }));
  }, [activeFile]);

  // ── Save file ──
  const saveFile = useCallback(async (fileId: string) => {
    const f = files.find((ff) => ff.id === fileId);
    if (!f) return;

    const chartSheets = f.sheets.filter(sh => sh.chartType);

    if (chartSheets.length > 0) {
      // Есть листы с графиками — сначала собираем полный xlsx со всеми стилями,
      // затем отправляем его на бэкенд который добавляет диаграмму не трогая остальное
      try {
        const wb = buildWorkbook(f);
        // Сериализуем в Uint8Array и кодируем в base64
        const wbOut = XLSX.write(wb, { bookType: "xlsx", type: "array", cellStyles: true }) as ArrayBuffer;
        const bytes = new Uint8Array(wbOut);
        let binary = "";
        for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
        const xlsxB64 = btoa(binary);

        const firstChart = chartSheets[0];
        const resp = await fetch(EXCEL_CHART_URL, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            xlsx_b64: xlsxB64,
            chart_sheet: firstChart.name,
            chart_type: firstChart.chartType ?? "pie",
            chart_title: firstChart.name,
          }),
        });

        if (!resp.ok) throw new Error(`Ошибка ${resp.status}`);
        const { xlsx_b64: resultB64 } = await resp.json();

        const resBinary = atob(resultB64);
        const resBytes = new Uint8Array(resBinary.length);
        for (let i = 0; i < resBinary.length; i++) resBytes[i] = resBinary.charCodeAt(i);
        const blob = new Blob([resBytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = f.name;
        a.click();
        URL.revokeObjectURL(url);
      } catch {
        // Fallback — сохраняем без графика но со стилями
        const wb = buildWorkbook(f);
        XLSX.writeFile(wb, f.name, { bookType: "xlsx", cellStyles: true });
      }
    } else {
      // Обычное сохранение — стили сохраняются из оригинального workbook
      const wb = buildWorkbook(f);
      XLSX.writeFile(wb, f.name, { bookType: "xlsx", cellStyles: true });
    }

    setFiles((prev) => prev.map((ff) => ff.id === fileId ? { ...ff, isDirty: false } : ff));
  }, [files]);

  // ── AI send ──
  const handleAiSend = async () => {
    const text = aiInput.trim();
    if ((!text && pendingImages.length === 0) || aiThinking) return;
    if (!aiSettings.apiKey) { setSettingsOpen(true); return; }

    const imgs = [...pendingImages];
    setMessages((prev) => [...prev, { role: "user", text, ts: getTime(), images: imgs.length ? imgs : undefined }]);
    setAiInput("");
    setPendingImages([]);
    setAiThinking(true);
    setAiError(null);

    try {
      const result = await callAi(text, files, aiSettings, imgs, docs, prompts, messages, knowledge);

      // ИИ запрашивает уточнение — показываем карточку с вопросом
      if (result.ask_user) {
        setMessages((prev) => [...prev, {
          role: "ai",
          text: result.text,
          ts: getTime(),
          ask_user: result.ask_user,
          pendingPrompt: text,
        }]);
        return;
      }

      let chartData: { name: string; value: number }[] | undefined;
      let chartTitle: string | undefined;

      if (result.mutations && result.mutations.length > 0) {
        const firstMut = result.mutations[0];
        const rows = firstMut.data;
        const dataRows = rows.slice(1).filter(r => r[0] != null && r[1] != null);
        if (dataRows.length > 0 && dataRows.every(r => !isNaN(Number(r[1])))) {
          chartData = dataRows.map(r => ({ name: String(r[0]), value: Number(r[1]) }));
          chartTitle = firstMut.sheetName;
        }

        setFiles((prev) => {
          let next = [...prev];
          result.mutations!.forEach(({ fileId, sheetName, data }) => {
            const COLS = Math.max(...data.map(r => r.length), 2);
            const ROWS = Math.max(data.length + 10, 50);
            const isChart = chartData != null && fileId === firstMut.fileId && sheetName === firstMut.sheetName;
            const ct: ChartType | undefined = isChart ? detectChartType(sheetName) : undefined;

            next = next.map((f) => {
              if (f.id !== fileId) return f;
              const cells: CellData[][] = Array.from({ length: ROWS }, (_, r) =>
                Array.from({ length: COLS }, (_, c) => ({ v: data[r]?.[c] ?? null }))
              );
              const newSheet: SheetData = {
                name: sheetName, cells, merges: [],
                colWidths: Array(COLS).fill(120),
                rowHeights: Array(ROWS).fill(22),
                chartType: ct,
              };
              const existing = f.sheets.findIndex(s => s.name === sheetName);
              let sheets = [...f.sheets];
              if (existing >= 0) sheets[existing] = newSheet;
              else sheets = [...sheets, newSheet];
              return { ...f, sheets, activeSheet: sheets.length - 1, isDirty: true };
            });
          });
          return next;
        });
        setActiveFileId(firstMut.fileId);
      }

      // Применяем стили к ячейкам — пишем в cells[] и в оригинальный workbook
      if (result.styleMutations && result.styleMutations.length > 0) {
        setFiles((prev) => prev.map((f) => {
          const sm = result.styleMutations!.filter(s => s.fileId === f.id);
          if (!sm.length) return f;

          // Обновляем оригинальный workbook (чтобы сохранились при Ctrl+S)
          const wb = f.workbook;
          sm.forEach(({ sheetName, changes }) => {
            const ws = wb.Sheets[sheetName];
            if (!ws) return;
            changes.forEach(({ row, col, bgColor, fontColor, bold }) => {
              const addr = XLSX.utils.encode_cell({ r: row, c: col });
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              let cell = ws[addr] as any;
              if (!cell) {
                // Создаём пустую ячейку если её нет — чтобы покрасить даже пустые
                cell = { t: "z", v: null };
                ws[addr] = cell;
                // Расширяем !ref если нужно
                const ref = XLSX.utils.decode_range(ws["!ref"] || "A1");
                if (row > ref.e.r) ref.e.r = row;
                if (col > ref.e.c) ref.e.c = col;
                ws["!ref"] = XLSX.utils.encode_range(ref);
              }
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const s: Record<string, any> = { ...(cell.s ?? {}) };
              if (bgColor) {
                s.fgColor = { argb: bgColor };
                s.patType = "solid";
                s.fill = { type: "pattern", patternType: "solid", fgColor: { argb: bgColor } };
              }
              if (fontColor) {
                s.font = { ...(s.font ?? {}), color: { argb: fontColor } };
              }
              if (bold !== undefined) {
                s.font = { ...(s.font ?? {}), bold };
              }
              cell.s = s;
            });
          });

          // Обновляем cells[] для отображения
          const sheets = f.sheets.map((sh) => {
            const sm2 = sm.find(s => s.sheetName === sh.name);
            if (!sm2) return sh;
            const cells = sh.cells.map(row => [...row]);
            sm2.changes.forEach(({ row, col, bgColor, fontColor, bold }) => {
              // Расширяем массив если ИИ указал строку за пределами
              if (row >= cells.length) {
                const colCount = cells[0]?.length ?? 30;
                while (cells.length <= row) cells.push(Array.from({ length: colCount }, () => ({ v: null })));
              }
              if (col >= cells[row].length) {
                while (cells[row].length <= col) cells[row].push({ v: null });
              }
              const prev = cells[row][col] ?? { v: null };
              const newStyle: CellStyle = { ...(prev.s ?? {}) };
              if (bgColor) {
                // Поддерживаем форматы: AARRGGBB (8 символов) и RRGGBB (6 символов)
                const hex = bgColor.length === 8 ? bgColor.slice(2) : bgColor.length === 6 ? bgColor : bgColor.replace(/^#/, "");
                newStyle.bgColor = `#${hex}`;
              }
              if (fontColor) {
                const hex = fontColor.length === 8 ? fontColor.slice(2) : fontColor.length === 6 ? fontColor : fontColor.replace(/^#/, "");
                newStyle.color = `#${hex}`;
              }
              if (bold !== undefined) newStyle.bold = bold;
              cells[row][col] = { ...prev, s: newStyle };
            });
            return { ...sh, cells };
          });

          return { ...f, sheets, isDirty: true };
        }));

        // Переключаемся на нужный лист
        const firstSm = result.styleMutations[0];
        setActiveFileId(firstSm.fileId);
        setFiles((prev) => prev.map((f) => {
          if (f.id !== firstSm.fileId) return f;
          const si = f.sheets.findIndex(s => s.name === firstSm.sheetName);
          return si >= 0 ? { ...f, activeSheet: si } : f;
        }));
      }

      // Применяем точечные изменения значений ячеек (без замены листа)
      if (result.valueMutations && result.valueMutations.length > 0) {
        setFiles((prev) => prev.map((f) => {
          const vm = result.valueMutations!.filter(m => m.fileId === f.id);
          if (!vm.length) return f;

          const wb = f.workbook;
          vm.forEach(({ sheetName, changes }) => {
            const ws = wb.Sheets[sheetName];
            if (!ws) return;
            changes.forEach(({ row, col, value, formula }) => {
              const addr = XLSX.utils.encode_cell({ r: row, c: col });
              const displayVal = formula ? value : value;
              const cellType = typeof value === "number" ? "n" : "s";
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const existing = ws[addr] as any;
              ws[addr] = {
                ...(existing ?? {}),
                t: value == null ? "z" : cellType,
                v: displayVal,
                w: value != null ? String(value) : undefined,
                ...(formula ? { f: formula.replace(/^=/, "") } : {}),
              };
              const ref = XLSX.utils.decode_range(ws["!ref"] || "A1");
              if (row > ref.e.r) ref.e.r = row;
              if (col > ref.e.c) ref.e.c = col;
              ws["!ref"] = XLSX.utils.encode_range(ref);
            });
          });

          const sheets = f.sheets.map((sh) => {
            const vm2 = vm.find(m => m.sheetName === sh.name);
            if (!vm2) return sh;
            const cells = sh.cells.map(row => [...row]);
            vm2.changes.forEach(({ row, col, value, formula }) => {
              if (row >= cells.length) {
                const colCount = cells[0]?.length ?? 30;
                while (cells.length <= row) cells.push(Array.from({ length: colCount }, () => ({ v: null })));
              }
              if (col >= cells[row].length) {
                while (cells[row].length <= col) cells[row].push({ v: null });
              }
              const prev = cells[row][col] ?? { v: null };
              cells[row][col] = {
                ...prev,
                v: value,
                w: formula ? formula : (value != null ? String(value) : undefined),
              };
            });
            return { ...sh, cells };
          });

          return { ...f, sheets, isDirty: true };
        }));

        const firstVm = result.valueMutations[0];
        setActiveFileId(firstVm.fileId);
        setFiles((prev) => prev.map((f) => {
          if (f.id !== firstVm.fileId) return f;
          const si = f.sheets.findIndex(s => s.name === firstVm.sheetName);
          return si >= 0 ? { ...f, activeSheet: si } : f;
        }));
      }

      // doc_update — показываем текст прямо в чате + обновляем в state
      if (result.docUpdates && result.docUpdates.length > 0) {
        const upd = result.docUpdates[0];
        // Обновляем текст документа в state чтобы при открытии тоже был новый
        setDocs((prev) => prev.map((d) => {
          const u = result.docUpdates!.find(x => x.docId === d.id);
          if (!u) return d;
          return { ...d, text: u.text, html: undefined };
        }));
        // Показываем текст прямо в чате с кнопкой копировать
        setMessages((prev) => [...prev, {
          role: "ai",
          text: result.text,
          ts: getTime(),
          docText: upd.text,
          docName: upd.docName,
        }]);
      } else {
        setMessages((prev) => [...prev, { role: "ai", text: result.text, ts: getTime(), chartData, chartTitle }]);
      }
    } catch (err) {
      const raw = err instanceof Error ? err.message : "Неизвестная ошибка";
      let msg = raw;
      if (raw.includes("Failed to fetch") || raw.includes("NetworkError") || raw.includes("fetch")) {
        msg = "Не удалось подключиться к серверу ИИ. Возможные причины:\n• Провайдер временно недоступен (503/504) — подожди минуту и попробуй снова\n• Неверный Base URL в настройках\n• Нет интернета";
      } else if (raw.includes("401") || raw.includes("Unauthorized")) {
        msg = "Неверный API ключ — проверь настройки (кнопка «Нет ключа» в хедере)";
      } else if (raw.includes("429") || raw.includes("rate limit")) {
        msg = "Превышен лимит запросов к модели — подожди немного или смени провайдера";
      } else if (raw.includes("503") || raw.includes("Service Unavailable")) {
        msg = "Сервер ИИ временно недоступен (503) — попробуй через минуту или смени провайдера в настройках";
      } else if (raw.includes("timeout") || raw.includes("timed out")) {
        msg = "Модель не успела ответить (таймаут) — попробуй более быструю модель или сократи запрос";
      }
      setAiError(msg);
      setMessages((prev) => [...prev, { role: "ai", text: `⚠️ ${msg}`, ts: getTime() }]);
    } finally {
      setAiThinking(false);
      setTimeout(() => chatEndRef.current?.scrollIntoView({ behavior: "smooth" }), 80);
    }
  };

  const renderMarkdown = (text: string) =>
    text
      .replace(/\*\*(.*?)\*\*/g, '<strong style="color:hsl(158,64%,60%)">$1</strong>')
      .replace(/•/g, '<span style="color:hsl(43,96%,60%)">•</span>')
      .replace(/\n/g, "<br/>");

  // ── Keyboard navigation ──
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      // Не перехватываем если фокус в input/textarea/select
      const tag = (e.target as HTMLElement)?.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT") return;
      if (editing || !selected || !activeSheet) return;
      const { row, col } = selected;
      const maxRow = activeSheet.cells.length - 1;
      const maxCol = (activeSheet.cells[0]?.length ?? 1) - 1;
      if (e.key === "ArrowDown") { e.preventDefault(); setSelected({ row: Math.min(row + 1, maxRow), col }); }
      if (e.key === "ArrowUp") { e.preventDefault(); setSelected({ row: Math.max(row - 1, 0), col }); }
      if (e.key === "ArrowRight") { e.preventDefault(); setSelected({ row, col: Math.min(col + 1, maxCol) }); }
      if (e.key === "ArrowLeft") { e.preventDefault(); setSelected({ row, col: Math.max(col - 1, 0) }); }
      if (e.key === "Enter" || e.key === "F2") { e.preventDefault(); setEditing(selected); }
      if (e.key === "Delete" || e.key === "Backspace") updateCell(row, col, null);
      if (e.key.length === 1 && !e.ctrlKey && !e.metaKey) setEditing(selected);
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [editing, selected, activeSheet, updateCell]);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key === "s" && activeFileId) {
        e.preventDefault();
        saveFile(activeFileId);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [activeFileId, saveFile]);

  const effectiveModelLabel = PRESET_MODELS.find(m => m.value === aiSettings.model)?.label ?? aiSettings.model;

  return (
    <div className="h-screen flex flex-col overflow-hidden">
      {settingsOpen && (
        <SettingsModal settings={aiSettings} onChange={setAiSettings} onClose={() => setSettingsOpen(false)} />
      )}

      {/* ── Header ── */}
      <header className="border-b border-border/60 px-4 py-3 flex items-center gap-3 flex-shrink-0"
        style={{ background: "hsla(220,16%,6%,0.97)" }}>
        <div className="flex items-center gap-2 flex-shrink-0">
          <div className="w-7 h-7 rounded-lg btn-primary flex items-center justify-center">
            <Icon name="Sparkles" size={14} />
          </div>
          <span className="font-semibold text-sm text-foreground tracking-tight hidden sm:block">DataMind</span>
        </div>
        <div className="h-5 w-px bg-border/60 hidden sm:block" />

        {/* File tabs */}
        <div className="flex items-center gap-1 flex-1 overflow-x-auto scrollbar-thin min-w-0">
          {files.map((f) => (
            <div key={f.id} role="button" tabIndex={0}
              onClick={() => { setActiveFileId(f.id); setActiveDocId(null); setEditing(null); setSelected(null); }}
              onKeyDown={(e) => { if (e.key === "Enter") { setActiveFileId(f.id); setActiveDocId(null); } }}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium flex-shrink-0 transition-all cursor-pointer ${f.id === activeFileId ? "bg-primary/10 text-primary border border-primary/30" : "text-muted-foreground hover:text-foreground hover:bg-secondary"}`}>
              <Icon name="FileSpreadsheet" size={13} />
              <span className="max-w-[120px] truncate">{f.name}</span>
              {f.isDirty && <span className="w-1.5 h-1.5 rounded-full bg-amber-400 flex-shrink-0" />}
              {f.role && (
                <span className={`text-[10px] px-1 rounded ${f.role === "main" ? "tag-emerald" : "tag-blue"}`}>
                  {f.role === "main" ? "осн." : "обр."}
                </span>
              )}
              <button onClick={(e) => {
                e.stopPropagation();
                setFiles(prev => prev.filter(ff => ff.id !== f.id));
                if (activeFileId === f.id) setActiveFileId(files.find(ff => ff.id !== f.id)?.id ?? null);
              }} className="ml-0.5 opacity-40 hover:opacity-100 transition-opacity">
                <Icon name="X" size={11} />
              </button>
            </div>
          ))}
          {/* Doc tabs (PDF / DOCX) */}
          {docs.map((d) => (
            <div key={d.id} role="button" tabIndex={0}
              onClick={() => { setActiveDocId(d.id); setActiveFileId(null); }}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium flex-shrink-0 transition-all cursor-pointer ${d.id === activeDocId ? "bg-blue-500/15 text-blue-400 border border-blue-500/30" : "text-muted-foreground hover:text-foreground hover:bg-secondary"}`}>
              <Icon name={d.type === "pdf" ? "FileText" : "FileEdit"} size={13} />
              <span className="max-w-[120px] truncate">{d.name}</span>
              {d.role && <span className="text-[10px] px-1 rounded tag-blue">{d.role}</span>}
              <button onClick={(e) => {
                e.stopPropagation();
                setDocs(prev => prev.filter(dd => dd.id !== d.id));
                if (activeDocId === d.id) setActiveDocId(null);
              }} className="ml-0.5 opacity-40 hover:opacity-100 transition-opacity">
                <Icon name="X" size={11} />
              </button>
            </div>
          ))}
          <button onClick={() => docInputRef.current?.click()}
            className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs text-muted-foreground hover:text-foreground hover:bg-secondary transition-all flex-shrink-0"
            title="Открыть Excel, Word или PDF">
            <Icon name="Plus" size={13} />
            <span className="hidden sm:inline">Открыть</span>
          </button>
        </div>

        {/* Right actions */}
        <div className="flex items-center gap-2 flex-shrink-0">
          {activeFile && (
            <>
              <select value={activeFile.role ?? ""}
                onChange={(e) => setFiles(prev => prev.map(f => f.id === activeFileId ? { ...f, role: (e.target.value as "main" | "reference") || null } : f))}
                className="text-xs bg-secondary border border-border/60 text-foreground rounded-lg px-2 py-1.5 outline-none hidden sm:block">
                <option value="">— роль —</option>
                <option value="main">Основной</option>
                <option value="reference">Образец</option>
              </select>
              {activeSheet && (
                <button
                  onClick={() => {
                    const sh = activeSheet;
                    const ctx = buildSheetContext(sh);
                    const w = window.open("", "_blank", "width=1000,height=750,scrollbars=yes");
                    if (w) {
                      w.document.write(`<html><head><title>Контекст ИИ — ${activeFile.name}</title><style>body{font:13px/1.6 monospace;white-space:pre;padding:20px;background:#0f1117;color:#e2e8f0}h2{color:#34d399;margin-bottom:8px}</style></head><body><h2>${activeFile.name} → ${activeSheet.name}</h2>${ctx.replace(/</g,"&lt;")}</body></html>`);
                      w.document.close();
                    }
                  }}
                  className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs bg-primary/10 text-primary border border-primary/30 hover:bg-primary/20 transition-all"
                  title="Показать что видит ИИ">
                  <Icon name="Table2" size={13} />
                  <span className="hidden sm:inline">Контекст ИИ</span>
                </button>
              )}
              <button onClick={() => saveFile(activeFile.id)}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold transition-all ${activeFile.isDirty ? "btn-primary" : "bg-secondary text-muted-foreground"}`}>
                <Icon name="Save" size={13} />
                <span className="hidden sm:inline">Сохранить</span>
              </button>
            </>
          )}
          <button onClick={() => setSettingsOpen(true)}
            className={`flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs transition-all ${aiSettings.apiKey ? "text-muted-foreground hover:text-foreground hover:bg-secondary" : "text-amber-400 border border-amber-400/30 hover:bg-amber-400/10"}`}
            title="Настройки ИИ">
            <Icon name="Settings2" size={14} />
            <span className="hidden lg:inline max-w-[100px] truncate">
              {aiSettings.apiKey ? effectiveModelLabel : "Нет ключа"}
            </span>
          </button>
          <button onClick={() => setKnowledgeOpen(true)}
            className={`flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs transition-all border ${knowledge.some(k => k.enabled) ? "text-amber-400 border-amber-400/40 bg-amber-400/10" : "text-muted-foreground border-border/40 hover:text-foreground hover:bg-secondary"}`}
            title="База знаний проекта">
            <Icon name="BookOpen" size={14} />
            <span className="hidden sm:inline">База знаний</span>
            {knowledge.filter(k => k.enabled).length > 0 && (
              <span className="w-4 h-4 rounded-full bg-amber-400 text-[9px] text-black flex items-center justify-center font-bold">
                {knowledge.filter(k => k.enabled).length}
              </span>
            )}
          </button>
          <button onClick={() => setPromptsOpen(p => !p)}
            className={`flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs transition-all border ${prompts.some(p => p.enabled) ? "text-primary border-primary/40 bg-primary/10" : "text-muted-foreground border-border/40 hover:text-foreground hover:bg-secondary"}`}
            title="Управление промптами">
            <Icon name="ListChecks" size={14} />
            <span className="hidden sm:inline">Промпты</span>
            {prompts.some(p => p.enabled) && (
              <span className="w-4 h-4 rounded-full bg-primary text-[9px] text-primary-foreground flex items-center justify-center font-bold">
                {prompts.filter(p => p.enabled).length}
              </span>
            )}
          </button>
          <button onClick={async () => {
            const list = await listSessions();
            setSavedSessions(list);
            setSessionOpen(true);
          }}
            className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs text-muted-foreground hover:text-foreground hover:bg-secondary transition-all border border-border/40"
            title="Сохранить / открыть проект">
            {sessionSaving
              ? <Icon name="Loader2" size={14} className="spinner" />
              : <Icon name="FolderOpen" size={14} />}
            <span className="hidden sm:inline">Проект</span>
          </button>
          <button onClick={() => navigate("/oilfield")}
            className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs text-muted-foreground hover:text-foreground hover:bg-secondary transition-all border border-border/40"
            title="Анализ месторождений">
            <Icon name="BarChart2" size={14} />
            <span className="hidden sm:inline">Месторождения</span>
          </button>
          <button onClick={() => setSidebarOpen(p => !p)}
            className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs text-muted-foreground hover:text-foreground hover:bg-secondary transition-all">
            <Icon name="MessageSquare" size={14} />
          </button>
        </div>
      </header>

      {/* ── Main ── */}
      <div className="flex flex-1 overflow-hidden">

        {/* ── Table area ── */}
        <div className="flex flex-col flex-1 overflow-hidden min-w-0">

          {activeFile && activeSheet && (
            <>
              {/* Formula bar + кнопки копирования */}
              <div className="flex items-center gap-2 px-3 py-1.5 border-b border-border/40 flex-shrink-0"
                style={{ background: "rgba(255,255,255,0.01)" }}>
                <span className="text-xs font-mono text-muted-foreground w-12 text-center flex-shrink-0">
                  {selected ? `${colLetter(selected.col)}${selected.row + 1}` : "—"}
                </span>
                <div className="h-4 w-px bg-border/60" />
                <span className="text-xs font-mono text-foreground flex-1 truncate">
                  {selected ? String(activeSheet.cells[selected.row]?.[selected.col]?.v ?? "") : ""}
                </span>

                {/* ── Кнопки копирования / экспорта листа ── */}
                <div className="flex items-center gap-1 flex-shrink-0 ml-2">
                  {/* Копировать лист как TSV (вставляется в Excel/Sheets) */}
                  <CopySheetButton sheet={activeSheet} />
                  {/* Скачать только этот лист как xlsx */}
                  <button
                    title="Скачать этот лист как отдельный Excel-файл"
                    onClick={() => {
                      const wb = XLSX.utils.book_new();
                      const numCols = getRealColCount(activeSheet);
                      const data = activeSheet.cells
                        .filter(row => row.some(c => c.v !== null))
                        .map(row => Array.from({ length: numCols }, (_, ci) => row[ci]?.v ?? null));
                      const ws = XLSX.utils.aoa_to_sheet(data);
                      XLSX.utils.book_append_sheet(wb, ws, activeSheet.name);
                      XLSX.writeFile(wb, `${activeSheet.name}.xlsx`, { bookType: "xlsx" });
                    }}
                    className="flex items-center gap-1 px-2 py-1 rounded-lg text-[11px] text-muted-foreground hover:text-foreground hover:bg-secondary border border-transparent hover:border-border/40 transition-all">
                    <Icon name="Download" size={12} />
                    <span className="hidden md:inline">Скачать лист</span>
                  </button>
                </div>
              </div>

              {/* Sheet tabs */}
              <div className="flex items-center border-b border-border/40 px-2 gap-0.5 flex-shrink-0 overflow-x-auto scrollbar-thin"
                style={{ background: "rgba(255,255,255,0.01)" }}>
                {activeFile.sheets.map((sh, si) => (
                  <button key={si}
                    onClick={() => setFiles(prev => prev.map(f => f.id === activeFileId ? { ...f, activeSheet: si } : f))}
                    className={`px-3 py-2 text-xs font-medium border-b-2 -mb-px transition-all whitespace-nowrap ${si === activeFile.activeSheet ? "border-primary text-primary" : "border-transparent text-muted-foreground hover:text-foreground"}`}>
                    {sh.chartType && <span className="mr-1 opacity-60">◉</span>}
                    {sh.name}
                  </button>
                ))}
              </div>
            </>
          )}

          {/* Drop zone / empty state */}
          {files.length === 0 && docs.length === 0 && (
            <div className="flex-1 flex flex-col items-center justify-center gap-6 p-8 cursor-pointer"
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => { e.preventDefault(); handleFiles(e.dataTransfer.files); }}
              onClick={() => docInputRef.current?.click()}>
              <div className="w-20 h-20 rounded-3xl bg-secondary flex items-center justify-center"
                style={{ boxShadow: "0 0 60px rgba(52,211,153,0.08)" }}>
                <Icon name="Files" size={40} className="text-primary" />
              </div>
              <div className="text-center">
                <p className="text-foreground font-semibold text-lg mb-1">Перетащи файлы сюда</p>
                <p className="text-muted-foreground text-sm">Excel · Word · PDF</p>
              </div>
              <div className="flex gap-3 text-[11px] text-muted-foreground">
                <span className="flex items-center gap-1"><Icon name="FileSpreadsheet" size={12} /> .xlsx .xls .csv</span>
                <span className="flex items-center gap-1"><Icon name="FileEdit" size={12} /> .docx</span>
                <span className="flex items-center gap-1"><Icon name="FileText" size={12} /> .pdf</span>
              </div>
            </div>
          )}

          {/* Document viewer (PDF / DOCX) */}
          {activeDocId && (() => {
            const doc = docs.find(d => d.id === activeDocId);
            if (!doc) return null;
            return (
              <div className="flex-1 overflow-auto p-4 scrollbar-thin flex flex-col gap-3">
                {/* Header */}
                <div className="flex items-center gap-3 pb-3 border-b border-border/40 flex-shrink-0">
                  <Icon name={doc.type === "pdf" ? "FileText" : "FileEdit"} size={16} className="text-blue-400" />
                  <span className="text-sm font-medium text-foreground">{doc.name}</span>
                  {doc.loading && <Icon name="Loader2" size={13} className="text-primary spinner" />}
                  {doc.pageCount && !doc.loading && <span className="text-xs text-muted-foreground">{doc.pageCount} стр.</span>}
                  {doc.type === "pdf" && doc.pageImageUrls && !doc.loading && (
                    <span className="text-[10px] px-2 py-0.5 rounded-full bg-primary/15 text-primary">
                      {doc.pageImageUrls.length} стр. → ИИ видит как картинки
                    </span>
                  )}
                  <select value={doc.role ?? ""} onChange={(e) => setDocs(prev => prev.map(d => d.id === doc.id ? { ...d, role: (e.target.value as DocFile["role"]) || null } : d))}
                    className="ml-auto text-xs bg-secondary border border-border/60 text-foreground rounded-lg px-2 py-1 outline-none">
                    <option value="">— роль —</option>
                    <option value="report">Отчёт</option>
                    <option value="protocol">Протокол ЦКР</option>
                    <option value="database">База данных</option>
                  </select>
                </div>

                {/* Loading state */}
                {doc.loading && (
                  <div className="flex-1 flex flex-col items-center justify-center gap-3 text-muted-foreground">
                    <Icon name="Loader2" size={32} className="text-primary spinner" />
                    <p className="text-sm">Конвертирую страницы PDF в изображения...</p>
                    <p className="text-xs">ИИ увидит документ точно как человек — с таблицами, графиками и форматированием</p>
                  </div>
                )}

                {/* PDF: диапазон страниц для ИИ */}
                {doc.type === "pdf" && !doc.loading && doc.pageImageUrls && doc.pageImageUrls.length > 0 && (
                  <div className="flex-shrink-0 flex items-center gap-3 px-1 py-2 rounded-xl border border-border/40"
                    style={{ background: "rgba(255,255,255,0.02)" }}>
                    <Icon name="SlidersHorizontal" size={14} className="text-muted-foreground flex-shrink-0" />
                    <span className="text-xs text-muted-foreground whitespace-nowrap">Страницы для ИИ:</span>
                    <div className="flex items-center gap-2">
                      <span className="text-xs text-muted-foreground">с</span>
                      <input type="number" min={1} max={doc.pageImageUrls.length}
                        value={doc.pageFrom ?? ""}
                        placeholder="1"
                        onChange={(e) => setDocs(prev => prev.map(d => d.id === doc.id
                          ? { ...d, pageFrom: e.target.value ? Math.max(1, parseInt(e.target.value)) : undefined }
                          : d))}
                        className="w-16 text-xs bg-secondary border border-border/60 text-foreground rounded-lg px-2 py-1 outline-none focus:border-primary/50 text-center" />
                      <span className="text-xs text-muted-foreground">по</span>
                      <input type="number" min={1} max={doc.pageImageUrls.length}
                        value={doc.pageTo ?? ""}
                        placeholder={String(doc.pageImageUrls.length)}
                        onChange={(e) => setDocs(prev => prev.map(d => d.id === doc.id
                          ? { ...d, pageTo: e.target.value ? Math.min(doc.pageImageUrls!.length, parseInt(e.target.value)) : undefined }
                          : d))}
                        className="w-16 text-xs bg-secondary border border-border/60 text-foreground rounded-lg px-2 py-1 outline-none focus:border-primary/50 text-center" />
                      <span className="text-xs text-muted-foreground">из {doc.pageImageUrls.length}</span>
                    </div>
                    {(doc.pageFrom || doc.pageTo) && (
                      <button onClick={() => setDocs(prev => prev.map(d => d.id === doc.id ? { ...d, pageFrom: undefined, pageTo: undefined } : d))}
                        className="text-xs text-muted-foreground hover:text-foreground ml-auto flex items-center gap-1">
                        <Icon name="X" size={11} /> сброс
                      </button>
                    )}
                  </div>
                )}

                {/* PDF: grid of page images */}
                {doc.type === "pdf" && !doc.loading && doc.pageImageUrls && doc.pageImageUrls.length > 0 && (() => {
                  const from = Math.max(1, doc.pageFrom ?? 1);
                  const to = Math.min(doc.pageImageUrls.length, doc.pageTo ?? doc.pageImageUrls.length);
                  return (
                    <div className="grid grid-cols-2 lg:grid-cols-3 gap-3">
                      {doc.pageImageUrls.map((url, pi) => {
                        const pageNum = pi + 1;
                        const inRange = pageNum >= from && pageNum <= to;
                        return (
                          <div key={pi}
                            className={`relative rounded-lg overflow-hidden border cursor-pointer transition-all ${inRange ? "border-primary/50 ring-1 ring-primary/20" : "border-border/30 opacity-40"}`}
                            onClick={() => window.open(url, "_blank")}>
                            <img src={url} alt={`Стр. ${pageNum}`} className="w-full object-contain bg-white" loading="lazy" />
                            <div className={`absolute bottom-0 left-0 right-0 px-2 py-1 text-[10px] text-center font-medium ${inRange ? "text-primary" : "text-white/60"}`}
                              style={{ background: "rgba(0,0,0,0.6)" }}>
                              стр. {pageNum}{inRange ? " ✓" : ""}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  );
                })()}

                {/* DOCX: HTML content */}
                {doc.type === "docx" && doc.html && (
                  <div className="prose prose-invert prose-sm max-w-none text-xs leading-relaxed"
                    style={{ color: "hsl(215,14%,75%)" }}
                    dangerouslySetInnerHTML={{ __html: doc.html }} />
                )}
              </div>
            );
          })()}

          {/* Chart view */}
          {activeFile && activeSheet && activeSheet.chartType && (
            <ChartView
              sheet={activeSheet}
              activeFileId={activeFileId}
              onChangeType={(t) => {
                setFiles(prev => prev.map(f => f.id !== activeFileId ? f : {
                  ...f, sheets: f.sheets.map((s, si) => si !== f.activeSheet ? s : { ...s, chartType: t }),
                }));
              }}
            />
          )}

          {/* Spreadsheet */}
          {activeFile && activeSheet && !activeSheet.chartType && (
            <Spreadsheet
              sheet={activeSheet}
              selected={selected}
              editing={editing}
              onSelect={(r, c) => { setSelected({ row: r, col: c }); setEditing(null); }}
              onEdit={(r, c) => setEditing({ row: r, col: c })}
              onCommit={(r, c, v) => { updateCell(r, c, v); setEditing(null); setSelected({ row: r, col: c + 1 }); }}
              onCancelEdit={() => setEditing(null)}
              tableRef={tableRef}
              onDrop={handleFiles}
            />
          )}
        </div>

        {/* ── AI Sidebar ── */}
        {sidebarOpen && (
          <div className="w-80 lg:w-96 flex flex-col border-l border-border/60 flex-shrink-0" style={{ background: "hsl(220,14%,7.5%)" }}>
            {/* Sidebar header */}
            <div className="px-4 py-3 border-b border-border/40 flex items-center gap-2 flex-shrink-0">
              <div className="w-7 h-7 rounded-full btn-primary flex items-center justify-center">
                <Icon name="Bot" size={14} />
              </div>
              <div className="flex-1 min-w-0">
                <p className="text-xs font-semibold text-foreground">ИИ-аналитик</p>
                <p className="text-[11px] text-muted-foreground truncate">
                  {aiSettings.apiKey ? effectiveModelLabel : "⚠ Ключ не настроен"}
                </p>
              </div>
              {aiThinking
                ? <Icon name="Loader2" size={14} className="text-primary spinner flex-shrink-0" />
                : <button onClick={() => setSettingsOpen(true)} className="text-muted-foreground hover:text-foreground transition-colors flex-shrink-0">
                  <Icon name="Settings2" size={14} />
                </button>}
            </div>

            {!aiSettings.apiKey && (
              <button onClick={() => setSettingsOpen(true)}
                className="mx-3 mt-2 p-2.5 rounded-lg border border-amber-400/30 text-left text-[11px] text-amber-400 hover:bg-amber-400/5 transition-all animate-fade-in"
                style={{ background: "rgba(251,191,36,0.04)" }}>
                <span className="font-medium">Укажи API ключ</span> — нажми для настройки провайдера и модели
              </button>
            )}

            {(files.length > 0 || docs.length > 0) && (
              <div className="px-3 py-2 border-b border-border/30 flex flex-wrap gap-1.5">
                {files.map((f) => (
                  <div key={f.id} className="flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] tag-emerald">
                    <Icon name="FileSpreadsheet" size={10} />
                    <span className="max-w-[80px] truncate">{f.name}</span>
                  </div>
                ))}
                {docs.map((d) => (
                  <div key={d.id} className="flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px]"
                    style={{ background: "rgba(59,130,246,0.15)", color: "#60a5fa" }}>
                    <Icon name={d.type === "pdf" ? "FileText" : "FileEdit"} size={10} />
                    <span className="max-w-[80px] truncate">{d.name}</span>
                  </div>
                ))}
                {prompts.some(p => p.enabled) && (
                  <div className="flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] bg-primary/15 text-primary cursor-pointer"
                    onClick={() => setPromptsOpen(true)}>
                    <Icon name="ListChecks" size={10} />
                    <span>{prompts.filter(p => p.enabled).length} промпт(а)</span>
                  </div>
                )}
              </div>
            )}

            {/* Messages */}
            <div className="flex-1 overflow-y-auto scrollbar-thin p-3 space-y-3">
              {messages.map((msg, i) => (
                <div key={i} className={`flex gap-2 animate-fade-in ${msg.role === "user" ? "flex-row-reverse" : ""}`}>
                  {msg.role === "ai" && (
                    <div className="w-6 h-6 rounded-full btn-primary flex items-center justify-center flex-shrink-0 mt-0.5">
                      <Icon name="Sparkles" size={11} />
                    </div>
                  )}
                  <div className={`max-w-[86%] rounded-xl px-3 py-2 text-xs leading-relaxed ${msg.role === "ai" ? "rounded-tl-sm border border-border/40" : "btn-amber rounded-tr-sm"}`}
                    style={msg.role === "ai" ? { background: "rgba(255,255,255,0.03)" } : {}}>
                    {/* Images in message */}
                    {msg.images && msg.images.length > 0 && (
                      <div className="flex flex-wrap gap-1.5 mb-2">
                        {msg.images.map((img, ii) => (
                          <div key={ii} className="relative group">
                            <img src={img.dataUrl} alt={img.name}
                              className="max-h-28 max-w-[180px] rounded-lg object-cover border border-border/40 cursor-pointer"
                              onClick={() => window.open(img.dataUrl, "_blank")} />
                            <span className="absolute bottom-1 left-1 text-[9px] bg-black/60 text-white px-1 rounded opacity-0 group-hover:opacity-100 transition-opacity truncate max-w-[120px]">
                              {img.name}
                            </span>
                          </div>
                        ))}
                      </div>
                    )}
                    <p dangerouslySetInnerHTML={{ __html: renderMarkdown(msg.text) }} />

                    {/* ── docText: готовый текст документа прямо в чате ── */}
                    {msg.docText && <DocTextBlock text={msg.docText} name={msg.docName} />}

                    {/* ── ask_user: карточка с вопросом ИИ ── */}
                    {msg.ask_user && msg.role === "ai" && (() => {
                      const isAnswered = messages.slice(i + 1).some(m => m.role === "user");
                      return !isAnswered ? (
                        <div className="mt-2.5 rounded-xl border border-primary/30 overflow-hidden"
                          style={{ background: "rgba(52,211,153,0.05)" }}>
                          <div className="flex items-start gap-2 px-3 py-2.5">
                            <Icon name="HelpCircle" size={14} className="text-primary flex-shrink-0 mt-0.5" />
                            <p className="text-[11px] text-foreground leading-relaxed flex-1">{msg.ask_user}</p>
                          </div>
                          <div className="flex gap-1.5 px-3 pb-2.5">
                            <input
                              autoFocus
                              placeholder="Ответ..."
                              className="flex-1 text-xs bg-background/60 border border-border/60 text-foreground rounded-lg px-3 py-1.5 outline-none focus:border-primary/50 placeholder:text-muted-foreground"
                              onKeyDown={(e) => {
                                if (e.key === "Enter" && e.currentTarget.value.trim()) {
                                  const answer = e.currentTarget.value.trim();
                                  // Объединяем исходный промпт + ответ на вопрос
                                  const combined = `${msg.pendingPrompt || ""}\n\nУточнение: ${answer}`;
                                  setAiInput(combined);
                                  // Убираем карточку вопроса, помечая сообщение как отвеченное
                                  setMessages(prev => prev.map((m, mi) =>
                                    mi === i ? { ...m, ask_user: undefined } : m
                                  ));
                                  setTimeout(() => handleAiSend(), 50);
                                }
                              }}
                            />
                            <button
                              className="px-3 py-1.5 rounded-lg text-xs btn-primary flex-shrink-0"
                              onClick={(e) => {
                                const input = e.currentTarget.previousElementSibling as HTMLInputElement;
                                const answer = input?.value?.trim();
                                if (!answer) return;
                                const combined = `${msg.pendingPrompt || ""}\n\nУточнение: ${answer}`;
                                setAiInput(combined);
                                setMessages(prev => prev.map((m, mi) =>
                                  mi === i ? { ...m, ask_user: undefined } : m
                                ));
                                setTimeout(() => handleAiSend(), 50);
                              }}>
                              Отправить
                            </button>
                          </div>
                        </div>
                      ) : null;
                    })()}

                    {/* Inline chart in AI message */}
                    {msg.chartData && msg.chartData.length > 0 && (
                      <div className="mt-3 -mx-1">
                        {msg.chartTitle && <p className="text-[10px] text-muted-foreground mb-1 font-medium">{msg.chartTitle}</p>}
                        <ResponsiveContainer width="100%" height={200}>
                          <PieChart>
                            <Pie data={msg.chartData} cx="50%" cy="50%" innerRadius={40} outerRadius={75} paddingAngle={3} dataKey="value">
                              {msg.chartData.map((_, ci) => (
                                <Cell key={ci} fill={CHART_COLORS[ci % CHART_COLORS.length]} />
                              ))}
                            </Pie>
                            <Tooltip contentStyle={{ background: "hsl(220,14%,10%)", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 8, fontSize: 11 }} />
                            <Legend wrapperStyle={{ fontSize: 10, color: "hsl(215,14%,55%)" }} iconType="circle" iconSize={8} />
                          </PieChart>
                        </ResponsiveContainer>
                      </div>
                    )}
                    <p className={`text-[10px] mt-1 ${msg.role === "ai" ? "text-muted-foreground" : "opacity-50"}`}>{msg.ts}</p>
                  </div>
                </div>
              ))}
              {aiThinking && (
                <div className="flex gap-2 animate-fade-in">
                  <div className="w-6 h-6 rounded-full btn-primary flex items-center justify-center flex-shrink-0">
                    <Icon name="Sparkles" size={11} />
                  </div>
                  <div className="px-3 py-2 rounded-xl rounded-tl-sm border border-border/40" style={{ background: "rgba(255,255,255,0.03)" }}>
                    <div className="flex gap-1 items-center h-3">
                      {[0, 150, 300].map((d) => (
                        <span key={d} className="w-1.5 h-1.5 rounded-full bg-primary animate-bounce" style={{ animationDelay: `${d}ms` }} />
                      ))}
                    </div>
                  </div>
                </div>
              )}
              <div ref={chatEndRef} />
            </div>

            {/* Quick prompts */}
            {files.length > 0 && messages.length <= 3 && (
              <div className="px-3 pb-2">
                <p className="text-[10px] text-muted-foreground mb-1.5">Примеры заданий:</p>
                <div className="space-y-1">
                  {[
                    "Посчитай итоги по всем числовым столбцам",
                    "Удали дубликаты строк",
                    "Возьми формат из образца и примени к основному файлу",
                  ].map((q) => (
                    <button key={q} onClick={() => setAiInput(q)}
                      className="w-full text-left text-[11px] px-2.5 py-1.5 rounded-lg border border-border/50 text-muted-foreground hover:text-foreground hover:border-primary/30 transition-all"
                      style={{ background: "rgba(255,255,255,0.02)" }}>
                      {q}
                    </button>
                  ))}
                </div>
              </div>
            )}

            {/* Pending images preview */}
            {pendingImages.length > 0 && (
              <div className="px-3 pb-1 flex flex-wrap gap-1.5">
                {pendingImages.map((img, i) => (
                  <div key={i} className="relative group">
                    <img src={img.dataUrl} alt={img.name} className="h-14 w-14 rounded-lg object-cover border border-primary/40" />
                    <button onClick={() => setPendingImages(p => p.filter((_, ii) => ii !== i))}
                      className="absolute -top-1 -right-1 w-4 h-4 rounded-full bg-background border border-border flex items-center justify-center opacity-0 group-hover:opacity-100 transition-opacity">
                      <Icon name="X" size={8} />
                    </button>
                  </div>
                ))}
              </div>
            )}

            {/* Input */}
            <div className="p-3 border-t border-border/40 flex-shrink-0">
              <div className="flex gap-1.5">
                {/* Image attach button */}
                <button onClick={() => imageInputRef.current?.click()}
                  title="Прикрепить картинку (или Ctrl+V)"
                  className="w-8 h-8 self-end rounded-lg bg-secondary text-muted-foreground hover:text-foreground flex items-center justify-center flex-shrink-0 transition-all">
                  <Icon name="ImagePlus" size={14} />
                </button>
                {/* Voice input button */}
                <button onClick={toggleVoice}
                  title={isListening ? "Остановить запись" : "Голосовой ввод"}
                  className={`w-8 h-8 self-end rounded-lg flex items-center justify-center flex-shrink-0 transition-all ${
                    isListening
                      ? "bg-red-500/20 text-red-400 border border-red-500/40 animate-pulse"
                      : "bg-secondary text-muted-foreground hover:text-foreground"
                  }`}>
                  <Icon name={isListening ? "MicOff" : "Mic"} size={14} />
                </button>
                <textarea ref={textareaRef} value={aiInput}
                  onChange={(e) => setAiInput(e.target.value)}
                  onKeyDown={(e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleAiSend(); } }}
                  placeholder={isListening ? "🎤 Говорите..." : (files.length === 0 && docs.length === 0 ? "Сначала загрузи файл..." : "Задание для ИИ...")}
                  disabled={aiThinking}
                  rows={2}
                  className={`flex-1 px-3 py-2 rounded-lg border text-xs text-foreground placeholder:text-muted-foreground focus:outline-none resize-none transition-all scrollbar-thin ${
                    isListening ? "border-red-500/40 bg-red-500/5" : "border-border/50 focus:border-primary/40"
                  }`}
                  style={{ background: isListening ? undefined : "rgba(255,255,255,0.03)" }}
                />
                <button onClick={handleAiSend}
                  disabled={(!aiInput.trim() && pendingImages.length === 0) || aiThinking}
                  className="w-8 h-8 self-end rounded-lg btn-primary flex items-center justify-center flex-shrink-0 disabled:opacity-30 disabled:cursor-not-allowed">
                  <Icon name="Send" size={13} />
                </button>
              </div>
              <p className="text-[10px] text-muted-foreground mt-1.5 text-center">
                Enter — отправить · Ctrl+V — скриншот · 🎤 — голос
              </p>
            </div>
          </div>
        )}
      </div>

      <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" multiple className="hidden"
        onChange={(e) => e.target.files && handleFiles(e.target.files)} />
      <input ref={docInputRef} type="file" accept=".xlsx,.xls,.csv,.docx,.pdf" multiple className="hidden"
        onChange={(e) => { if (e.target.files) handleFiles(e.target.files); e.target.value = ""; }} />
      <input ref={imageInputRef} type="file" accept="image/*" multiple className="hidden"
        onChange={(e) => { if (e.target.files) Array.from(e.target.files).forEach(loadImage); e.target.value = ""; }} />
      <input ref={knowledgeTxtRef} type="file" accept=".txt,.md" multiple className="hidden"
        onChange={async (e) => {
          if (!e.target.files) return;
          const newEntries: KnowledgeEntry[] = [];
          for (const f of Array.from(e.target.files)) {
            const text = await f.text();
            newEntries.push({
              id: crypto.randomUUID(), title: f.name.replace(/\.[^.]+$/, ""),
              content: text, category: "custom", sourceType: "text",
              enabled: true, updatedAt: Date.now(), fileName: f.name,
            });
          }
          const updated = [...knowledge, ...newEntries];
          setKnowledge(updated); await saveKnowledge(updated);
          e.target.value = "";
        }} />

      {/* Файлы в базу знаний: Excel / PDF / DOCX */}
      <input ref={knowledgeFileRef} type="file" accept=".xlsx,.xls,.csv,.docx,.pdf" multiple className="hidden"
        onChange={async (e) => {
          if (!e.target.files) return;
          for (const f of Array.from(e.target.files)) {
            const id = crypto.randomUUID();
            const title = f.name.replace(/\.[^.]+$/, "");
            const isPdf = f.name.match(/\.pdf$/i);
            const isDocx = f.name.match(/\.docx$/i);
            const isExcel = f.name.match(/\.(xlsx|xls|csv)$/i);
            const sourceType: KnowledgeSourceType = isPdf ? "pdf" : isDocx ? "docx" : "excel";

            if (isPdf) {
              // PDF → конвертируем через бэкенд
              setKbLoadingId(id);
              const placeholder: KnowledgeEntry = {
                id, title, content: "Загрузка...", category: "docs",
                sourceType: "pdf", enabled: false, updatedAt: Date.now(), fileName: f.name,
              };
              const updatedPlaceholder = [...knowledge, placeholder];
              setKnowledge(updatedPlaceholder);

              const buf = await f.arrayBuffer();
              const bytes = new Uint8Array(buf);
              let binary = "";
              for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
              const pdf_b64 = btoa(binary);
              const resp = await fetch(AI_EXCEL_URL, {
                method: "POST", headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ action: "pdf_to_images", pdf_b64, dpi: 150, max_pages: 80 }),
              });
              const result = await resp.json() as { page_urls: string[]; total_pages: number };
              const entry: KnowledgeEntry = {
                id, title, content: `PDF «${f.name}»: ${result.total_pages} стр. Передаётся ИИ как изображения страниц.`,
                category: "docs", sourceType: "pdf", enabled: true, updatedAt: Date.now(),
                fileName: f.name, pageImageUrls: result.page_urls, pageCount: result.total_pages,
              };
              const updated = knowledge.filter(k => k.id !== id).concat(entry);
              setKnowledge(updated); await saveKnowledge(updated);
              setKbLoadingId(null);

            } else if (isDocx) {
              const buf = await f.arrayBuffer();
              const textResult = await mammoth.extractRawText({ arrayBuffer: buf });
              const entry: KnowledgeEntry = {
                id, title, content: textResult.value, category: "docs",
                sourceType: "docx", enabled: true, updatedAt: Date.now(), fileName: f.name,
              };
              const updated = [...knowledge, entry];
              setKnowledge(updated); await saveKnowledge(updated);

            } else if (isExcel) {
              const buf = await f.arrayBuffer();
              const wb = XLSX.read(buf, { type: "array" });
              const lines: string[] = [`Excel-файл: ${f.name}`];
              wb.SheetNames.forEach(shName => {
                const ws = wb.Sheets[shName];
                const rows = XLSX.utils.sheet_to_json<string[]>(ws, { header: 1, defval: "" }) as string[][];
                lines.push(`\n--- Лист «${shName}» ---`);
                rows.slice(0, 200).forEach(row => lines.push(row.join("\t")));
                if (rows.length > 200) lines.push(`[... ещё ${rows.length - 200} строк]`);
              });
              const entry: KnowledgeEntry = {
                id, title, content: lines.join("\n"), category: "tables",
                sourceType: "excel", enabled: true, updatedAt: Date.now(), fileName: f.name,
              };
              const updated = [...knowledge, entry];
              setKnowledge(updated); await saveKnowledge(updated);
            }
          }
          e.target.value = "";
        }} />

      {/* ── Prompts Modal ── */}
      {promptsOpen && (
        <div className="fixed inset-0 z-50 flex items-start justify-center pt-16 px-4" style={{ background: "rgba(0,0,0,0.7)" }}
          onClick={(e) => { if (e.target === e.currentTarget) setPromptsOpen(false); }}>
          <div className="w-full max-w-2xl rounded-2xl border border-border/60 flex flex-col max-h-[80vh]"
            style={{ background: "hsl(220,14%,8%)" }}>
            <div className="px-5 py-4 border-b border-border/40 flex items-center gap-3">
              <Icon name="ListChecks" size={16} className="text-primary" />
              <h2 className="font-semibold text-sm text-foreground flex-1">Управление промптами</h2>
              <p className="text-xs text-muted-foreground">Галочкой выбери что подключить к ИИ</p>
              <button onClick={() => setPromptsOpen(false)} className="text-muted-foreground hover:text-foreground transition-colors ml-2">
                <Icon name="X" size={16} />
              </button>
            </div>

            <div className="flex-1 overflow-y-auto scrollbar-thin p-4 space-y-2">
              {prompts.map((p) => (
                <div key={p.id} className={`rounded-xl border transition-all ${p.enabled ? "border-primary/40 bg-primary/5" : "border-border/40"}`}>
                  <div className="flex items-start gap-3 p-3">
                    <button onClick={() => {
                      const updated = prompts.map(pp => pp.id === p.id ? { ...pp, enabled: !pp.enabled } : pp);
                      setPrompts(updated);
                      savePrompts(updated);
                    }} className={`w-5 h-5 rounded flex items-center justify-center flex-shrink-0 mt-0.5 border-2 transition-all ${p.enabled ? "bg-primary border-primary" : "border-border/60 hover:border-primary/50"}`}>
                      {p.enabled && <Icon name="Check" size={11} className="text-primary-foreground" />}
                    </button>
                    <div className="flex-1 min-w-0">
                      <p className="text-xs font-semibold text-foreground mb-1">{p.label}</p>
                      {editingPrompt === p.id ? (
                        <textarea
                          defaultValue={p.text}
                          onBlur={(e) => {
                            const updated = prompts.map(pp => pp.id === p.id ? { ...pp, text: e.target.value } : pp);
                            setPrompts(updated);
                            savePrompts(updated);
                            setEditingPrompt(null);
                          }}
                          autoFocus
                          rows={5}
                          className="w-full text-xs text-foreground bg-background/50 border border-primary/30 rounded-lg px-3 py-2 resize-none outline-none focus:border-primary/60 font-mono"
                        />
                      ) : (
                        <p className="text-[11px] text-muted-foreground leading-relaxed line-clamp-3">{p.text}</p>
                      )}
                    </div>
                    <button onClick={() => setEditingPrompt(editingPrompt === p.id ? null : p.id)}
                      className="text-muted-foreground hover:text-foreground transition-colors flex-shrink-0 mt-0.5"
                      title="Редактировать текст промпта">
                      <Icon name={editingPrompt === p.id ? "Check" : "Pencil"} size={13} />
                    </button>
                  </div>
                </div>
              ))}

              {/* Добавить свой промпт */}
              <button onClick={() => {
                const newP: PromptPreset = {
                  id: crypto.randomUUID(),
                  label: "Новый промпт",
                  text: "",
                  enabled: true,
                };
                const updated = [...prompts, newP];
                setPrompts(updated);
                savePrompts(updated);
                setEditingPrompt(newP.id);
              }} className="w-full py-2.5 rounded-xl border border-dashed border-border/40 text-xs text-muted-foreground hover:text-foreground hover:border-primary/30 transition-all flex items-center justify-center gap-2">
                <Icon name="Plus" size={13} />
                Добавить свой промпт
              </button>
            </div>

            <div className="px-5 py-3 border-t border-border/40 flex items-center justify-between">
              <p className="text-xs text-muted-foreground">
                {prompts.filter(p => p.enabled).length > 0
                  ? `Подключено: ${prompts.filter(p => p.enabled).map(p => p.label).join(", ")}`
                  : "Ни один промпт не подключён — ИИ работает в базовом режиме"}
              </p>
              <button onClick={() => setPromptsOpen(false)}
                className="px-4 py-1.5 rounded-lg text-xs btn-primary">
                Готово
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── Knowledge Base Modal ── */}
      {knowledgeOpen && (() => {
        const filtered = kbCategoryFilter === "all"
          ? knowledge
          : knowledge.filter(k => k.category === kbCategoryFilter);
        const activeCount = knowledge.filter(k => k.enabled).length;
        const catCounts = Object.fromEntries(
          (Object.keys(KNOWLEDGE_CATEGORIES) as KnowledgeCategory[]).map(c => [c, knowledge.filter(k => k.category === c).length])
        );

        // Иконка источника
        const srcIcon = (t: KnowledgeSourceType) =>
          t === "pdf" ? "FileText" : t === "docx" ? "FileEdit" : t === "excel" ? "FileSpreadsheet" : "AlignLeft";

        return (
          <div className="fixed inset-0 z-50 flex items-start justify-center pt-6 px-3 pb-3"
            style={{ background: "rgba(0,0,0,0.8)" }}
            onClick={(e) => { if (e.target === e.currentTarget) setKnowledgeOpen(false); }}>
            <div className="w-full max-w-3xl rounded-2xl border border-amber-400/25 flex flex-col"
              style={{ background: "hsl(220,14%,8%)", maxHeight: "calc(100vh - 48px)" }}>

              {/* ── Header ── */}
              <div className="px-5 py-4 border-b border-border/40 flex items-center gap-3 flex-shrink-0">
                <Icon name="BookOpen" size={16} className="text-amber-400" />
                <div className="flex-1">
                  <h2 className="font-semibold text-sm text-foreground">База знаний проекта</h2>
                  <p className="text-[11px] text-muted-foreground">Галочкой включаешь что ИИ знает. Меняй состав под каждый проект.</p>
                </div>
                <div className="flex items-center gap-2 flex-shrink-0">
                  {knowledge.length === 0 && (
                    <button onClick={async () => {
                      const entries: KnowledgeEntry[] = OIL_KNOWLEDGE_TEMPLATES.map(t => ({
                        ...t, id: crypto.randomUUID(), updatedAt: Date.now(),
                      }));
                      const updated = [...knowledge, ...entries];
                      setKnowledge(updated); await saveKnowledge(updated);
                    }} className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs bg-amber-400/15 text-amber-400 border border-amber-400/30 hover:bg-amber-400/25 transition-all">
                      <Icon name="Sparkles" size={12} /> Загрузить шаблоны нефтяной отрасли
                    </button>
                  )}
                  <button onClick={() => setKnowledgeOpen(false)} className="text-muted-foreground hover:text-foreground p-1">
                    <Icon name="X" size={16} />
                  </button>
                </div>
              </div>

              {/* ── Кнопки загрузки ── */}
              <div className="px-4 py-3 border-b border-border/30 flex flex-wrap gap-2 flex-shrink-0"
                style={{ background: "rgba(255,255,255,0.01)" }}>
                <button onClick={() => {
                  const newK: KnowledgeEntry = {
                    id: crypto.randomUUID(), title: "Новое правило", content: "",
                    category: "custom", sourceType: "text", enabled: true, updatedAt: Date.now(),
                  };
                  setKnowledge(prev => { const u = [...prev, newK]; saveKnowledge(u); return u; });
                  setEditingKbId(newK.id);
                }} className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs border border-border/50 text-muted-foreground hover:text-foreground hover:border-amber-400/40 transition-all">
                  <Icon name="Plus" size={12} /> Написать правило
                </button>
                <button onClick={() => knowledgeTxtRef.current?.click()}
                  className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs border border-border/50 text-muted-foreground hover:text-foreground hover:border-amber-400/40 transition-all">
                  <Icon name="FileText" size={12} /> .txt / .md
                </button>
                <button onClick={() => knowledgeFileRef.current?.click()}
                  className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs border border-border/50 text-muted-foreground hover:text-foreground hover:border-amber-400/40 transition-all">
                  <Icon name="Upload" size={12} /> Excel / Word / PDF
                  {kbLoadingId && <Icon name="Loader2" size={11} className="spinner text-amber-400" />}
                </button>
                {knowledge.length > 0 && (
                  <button onClick={async () => {
                    const entries: KnowledgeEntry[] = OIL_KNOWLEDGE_TEMPLATES.map(t => ({
                      ...t, id: crypto.randomUUID(), updatedAt: Date.now(),
                    }));
                    const existTitles = new Set(knowledge.map(k => k.title));
                    const toAdd = entries.filter(e => !existTitles.has(e.title));
                    if (!toAdd.length) return;
                    const updated = [...knowledge, ...toAdd];
                    setKnowledge(updated); await saveKnowledge(updated);
                  }} className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs border border-amber-400/20 text-amber-400/70 hover:text-amber-400 hover:border-amber-400/40 transition-all ml-auto">
                    <Icon name="Sparkles" size={12} /> + шаблоны
                  </button>
                )}
              </div>

              {/* ── Категории-фильтры ── */}
              <div className="px-4 py-2 border-b border-border/20 flex gap-1.5 overflow-x-auto scrollbar-thin flex-shrink-0">
                <button onClick={() => setKbCategoryFilter("all")}
                  className={`flex items-center gap-1 px-2.5 py-1 rounded-lg text-[11px] whitespace-nowrap transition-all flex-shrink-0 ${kbCategoryFilter === "all" ? "bg-amber-400/20 text-amber-300 border border-amber-400/30" : "text-muted-foreground hover:text-foreground border border-transparent hover:border-border/40"}`}>
                  Все ({knowledge.length})
                </button>
                {(Object.entries(KNOWLEDGE_CATEGORIES) as [KnowledgeCategory, typeof KNOWLEDGE_CATEGORIES[KnowledgeCategory]][]).map(([cat, meta]) =>
                  catCounts[cat] > 0 && (
                    <button key={cat} onClick={() => setKbCategoryFilter(cat)}
                      className={`flex items-center gap-1 px-2.5 py-1 rounded-lg text-[11px] whitespace-nowrap transition-all flex-shrink-0 ${kbCategoryFilter === cat ? "bg-amber-400/20 text-amber-300 border border-amber-400/30" : "text-muted-foreground hover:text-foreground border border-transparent hover:border-border/40"}`}>
                      <Icon name={meta.icon as Parameters<typeof Icon>[0]["name"]} size={11} className={meta.color} />
                      {meta.label} ({catCounts[cat]})
                    </button>
                  )
                )}
              </div>

              {/* ── Список записей ── */}
              <div className="flex-1 overflow-y-auto scrollbar-thin p-3 space-y-1.5 min-h-0">
                {filtered.length === 0 && (
                  <div className="text-center py-10 text-muted-foreground">
                    <Icon name="BookOpen" size={28} className="mx-auto mb-2 opacity-20" />
                    <p className="text-sm">
                      {knowledge.length === 0
                        ? "База знаний пуста — нажми «Загрузить шаблоны»"
                        : "В этой категории пусто"}
                    </p>
                  </div>
                )}
                {filtered.map((k) => {
                  const catMeta = KNOWLEDGE_CATEGORIES[k.category ?? "custom"];
                  const isLoading = kbLoadingId === k.id;
                  return (
                    <div key={k.id}
                      className={`rounded-xl border transition-all ${k.enabled ? "border-amber-400/35 bg-amber-400/5" : "border-border/30 opacity-60 hover:opacity-80"}`}>
                      <div className="flex items-start gap-2.5 p-3">
                        {/* Чекбокс */}
                        <button onClick={async () => {
                          const updated = knowledge.map(kk => kk.id === k.id ? { ...kk, enabled: !kk.enabled } : kk);
                          setKnowledge(updated); await saveKnowledge(updated);
                        }} className={`w-5 h-5 rounded flex items-center justify-center flex-shrink-0 mt-0.5 border-2 transition-all ${k.enabled ? "bg-amber-400 border-amber-400" : "border-border/60 hover:border-amber-400/60"}`}>
                          {k.enabled && <Icon name="Check" size={10} className="text-black" />}
                        </button>

                        <div className="flex-1 min-w-0">
                          {/* Заголовок + мета */}
                          <div className="flex items-center gap-1.5 mb-0.5 flex-wrap">
                            <Icon name={srcIcon(k.sourceType ?? "text") as Parameters<typeof Icon>[0]["name"]} size={11} className="text-muted-foreground flex-shrink-0" />
                            {editingKbId === k.id ? (
                              <input defaultValue={k.title} autoFocus
                                onBlur={async (e) => {
                                  const updated = knowledge.map(kk => kk.id === k.id ? { ...kk, title: e.target.value } : kk);
                                  setKnowledge(updated); await saveKnowledge(updated);
                                }}
                                className="flex-1 text-xs font-semibold text-foreground bg-background/50 border border-amber-400/40 rounded px-2 py-0.5 outline-none" />
                            ) : (
                              <span className="text-xs font-semibold text-amber-200 truncate">{k.title}</span>
                            )}
                            <span className={`text-[10px] px-1.5 py-0.5 rounded-full border flex-shrink-0 ${catMeta.color} border-current/20 bg-current/5`}
                              style={{ opacity: 0.8 }}>{catMeta.label}</span>
                            {k.fileName && (
                              <span className="text-[10px] text-muted-foreground truncate max-w-[120px]">{k.fileName}</span>
                            )}
                            {isLoading && <Icon name="Loader2" size={11} className="spinner text-amber-400 flex-shrink-0" />}
                          </div>

                          {/* Категория (select при редактировании) */}
                          {editingKbId === k.id && (
                            <select value={k.category ?? "custom"}
                              onChange={async (e) => {
                                const updated = knowledge.map(kk => kk.id === k.id ? { ...kk, category: e.target.value as KnowledgeCategory } : kk);
                                setKnowledge(updated); await saveKnowledge(updated);
                              }}
                              className="mb-1.5 text-[11px] bg-secondary border border-border/60 text-foreground rounded px-2 py-0.5 outline-none">
                              {(Object.entries(KNOWLEDGE_CATEGORIES) as [KnowledgeCategory, typeof KNOWLEDGE_CATEGORIES[KnowledgeCategory]][]).map(([cat, meta]) => (
                                <option key={cat} value={cat}>{meta.label}</option>
                              ))}
                            </select>
                          )}

                          {/* Содержимое */}
                          {k.sourceType === "pdf" && k.pageImageUrls?.length ? (
                            <div className="space-y-1.5">
                              <p className="text-[11px] text-muted-foreground">
                                {k.pageImageUrls.length} стр. → ИИ видит как изображения
                              </p>
                              {/* Диапазон страниц */}
                              <div className="flex items-center gap-2 flex-wrap">
                                <span className="text-[11px] text-muted-foreground">Страницы для ИИ:</span>
                                <input type="number" min={1} max={k.pageImageUrls.length}
                                  value={k.pageFrom ?? ""} placeholder="1"
                                  onChange={async (e) => {
                                    const updated = knowledge.map(kk => kk.id === k.id
                                      ? { ...kk, pageFrom: e.target.value ? parseInt(e.target.value) : undefined } : kk);
                                    setKnowledge(updated); await saveKnowledge(updated);
                                  }}
                                  className="w-14 text-[11px] bg-secondary border border-border/60 text-foreground rounded px-1.5 py-0.5 outline-none text-center" />
                                <span className="text-[11px] text-muted-foreground">—</span>
                                <input type="number" min={1} max={k.pageImageUrls.length}
                                  value={k.pageTo ?? ""} placeholder={String(k.pageImageUrls.length)}
                                  onChange={async (e) => {
                                    const updated = knowledge.map(kk => kk.id === k.id
                                      ? { ...kk, pageTo: e.target.value ? parseInt(e.target.value) : undefined } : kk);
                                    setKnowledge(updated); await saveKnowledge(updated);
                                  }}
                                  className="w-14 text-[11px] bg-secondary border border-border/60 text-foreground rounded px-1.5 py-0.5 outline-none text-center" />
                                <span className="text-[11px] text-muted-foreground">из {k.pageImageUrls.length}</span>
                              </div>
                              {/* Миниатюры первых 4 страниц */}
                              <div className="flex gap-1.5 mt-1">
                                {k.pageImageUrls.slice(0, 4).map((url, pi) => (
                                  <img key={pi} src={url} alt={`стр.${pi+1}`}
                                    className="h-14 w-10 object-cover rounded border border-border/40 bg-white cursor-pointer"
                                    onClick={() => window.open(url, "_blank")} />
                                ))}
                                {k.pageImageUrls.length > 4 && (
                                  <div className="h-14 w-10 rounded border border-border/40 bg-secondary flex items-center justify-center text-[10px] text-muted-foreground">
                                    +{k.pageImageUrls.length - 4}
                                  </div>
                                )}
                              </div>
                            </div>
                          ) : editingKbId === k.id ? (
                            <textarea defaultValue={k.content} rows={5}
                              onBlur={async (e) => {
                                const updated = knowledge.map(kk => kk.id === k.id
                                  ? { ...kk, content: e.target.value, updatedAt: Date.now() } : kk);
                                setKnowledge(updated); await saveKnowledge(updated);
                                setEditingKbId(null);
                              }}
                              className="w-full mt-1 text-[11px] text-foreground bg-background/60 border border-amber-400/30 rounded-lg px-3 py-2 resize-none outline-none focus:border-amber-400/60 font-mono leading-relaxed" />
                          ) : (
                            <p className="text-[11px] text-muted-foreground leading-relaxed line-clamp-2 whitespace-pre-wrap mt-0.5">{k.content}</p>
                          )}
                        </div>

                        {/* Действия */}
                        <div className="flex flex-col gap-1 flex-shrink-0 mt-0.5">
                          {k.sourceType !== "pdf" && (
                            <button onClick={() => setEditingKbId(editingKbId === k.id ? null : k.id)}
                              className="w-6 h-6 flex items-center justify-center rounded text-muted-foreground hover:text-foreground hover:bg-secondary transition-all"
                              title="Редактировать">
                              <Icon name={editingKbId === k.id ? "Check" : "Pencil"} size={12} />
                            </button>
                          )}
                          <button onClick={async () => {
                            const updated = knowledge.filter(kk => kk.id !== k.id);
                            setKnowledge(updated); await saveKnowledge(updated);
                          }} className="w-6 h-6 flex items-center justify-center rounded text-muted-foreground hover:text-red-400 hover:bg-red-400/10 transition-all"
                            title="Удалить">
                            <Icon name="Trash2" size={12} />
                          </button>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>

              {/* ── Footer ── */}
              <div className="px-5 py-3 border-t border-border/40 flex items-center justify-between gap-3 flex-shrink-0">
                <div className="flex items-center gap-2 min-w-0">
                  <span className={`text-[11px] font-medium ${activeCount > 0 ? "text-amber-400" : "text-muted-foreground"}`}>
                    {activeCount > 0 ? `Активно ${activeCount} из ${knowledge.length}` : "Ничего не активно"}
                  </span>
                  {activeCount > 0 && (
                    <span className="text-[11px] text-muted-foreground truncate hidden sm:block">
                      — {knowledge.filter(k => k.enabled).map(k => k.title).join(", ")}
                    </span>
                  )}
                </div>
                <div className="flex gap-2 flex-shrink-0">
                  <button onClick={async () => {
                    const updated = knowledge.map(k => ({ ...k, enabled: true }));
                    setKnowledge(updated); await saveKnowledge(updated);
                  }} className="px-3 py-1.5 rounded-lg text-xs text-muted-foreground hover:text-foreground border border-border/40 hover:border-border/70 transition-all">
                    Все вкл.
                  </button>
                  <button onClick={async () => {
                    const updated = knowledge.map(k => ({ ...k, enabled: false }));
                    setKnowledge(updated); await saveKnowledge(updated);
                  }} className="px-3 py-1.5 rounded-lg text-xs text-muted-foreground hover:text-foreground border border-border/40 hover:border-border/70 transition-all">
                    Все выкл.
                  </button>
                  <button onClick={() => setKnowledgeOpen(false)} className="px-4 py-1.5 rounded-lg text-xs btn-primary">
                    Готово
                  </button>
                </div>
              </div>
            </div>
          </div>
        );
      })()}

      {/* ── Session Modal ── */}
      {sessionOpen && (
        <div className="fixed inset-0 z-50 flex items-start justify-center pt-12 px-4" style={{ background: "rgba(0,0,0,0.75)" }}
          onClick={(e) => { if (e.target === e.currentTarget) setSessionOpen(false); }}>
          <div className="w-full max-w-lg rounded-2xl border border-border/60 flex flex-col max-h-[80vh]"
            style={{ background: "hsl(220,14%,8%)" }}>

            <div className="px-5 py-4 border-b border-border/40 flex items-center gap-3">
              <Icon name="FolderOpen" size={16} className="text-primary" />
              <div className="flex-1">
                <h2 className="font-semibold text-sm text-foreground">Проект</h2>
                <p className="text-[11px] text-muted-foreground">Сохранение и восстановление рабочей сессии</p>
              </div>
              <button onClick={() => setSessionOpen(false)} className="text-muted-foreground hover:text-foreground">
                <Icon name="X" size={16} />
              </button>
            </div>

            <div className="flex-1 overflow-y-auto scrollbar-thin p-4 space-y-3">
              {/* Сохранить текущее */}
              <div className="rounded-xl border border-primary/30 p-4 space-y-2" style={{ background: "rgba(52,211,153,0.04)" }}>
                <p className="text-xs font-semibold text-primary">Сохранить текущую сессию</p>
                <p className="text-[11px] text-muted-foreground">
                  Файлы: {files.length} Excel, {docs.length} документов · {messages.length} сообщений
                </p>
                <div className="flex gap-2">
                  <input id="session-name-input" type="text" placeholder="Название сессии..."
                    defaultValue={`Сессия ${new Date().toLocaleDateString("ru")}`}
                    className="flex-1 text-xs bg-secondary border border-border/60 text-foreground rounded-lg px-3 py-1.5 outline-none focus:border-primary/50" />
                  <button onClick={async () => {
                    const input = document.getElementById("session-name-input") as HTMLInputElement;
                    const name = input?.value || "Сессия";
                    const id = `session_${Date.now()}`;
                    await doSaveSession(id, name);
                    const list = await listSessions();
                    setSavedSessions(list);
                  }} disabled={sessionSaving}
                    className="px-3 py-1.5 rounded-lg text-xs btn-primary disabled:opacity-50 flex items-center gap-1.5">
                    {sessionSaving ? <Icon name="Loader2" size={12} className="spinner" /> : <Icon name="Save" size={12} />}
                    Сохранить
                  </button>
                </div>
              </div>

              {/* Список сессий */}
              {savedSessions.length === 0 ? (
                <p className="text-xs text-muted-foreground text-center py-4">Нет сохранённых сессий</p>
              ) : (
                <div className="space-y-2">
                  <p className="text-xs text-muted-foreground font-medium">Сохранённые сессии:</p>
                  {savedSessions.map((s) => (
                    <div key={s.id} className="flex items-center gap-3 px-3 py-2.5 rounded-xl border border-border/40 hover:border-primary/30 transition-all"
                      style={{ background: "rgba(255,255,255,0.02)" }}>
                      <Icon name="Clock" size={14} className="text-muted-foreground flex-shrink-0" />
                      <div className="flex-1 min-w-0">
                        <p className="text-xs font-medium text-foreground truncate">{s.name}</p>
                        <p className="text-[11px] text-muted-foreground">
                          {new Date(s.savedAt).toLocaleString("ru", { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" })}
                        </p>
                      </div>
                      <button onClick={() => doLoadSession(s.id)}
                        className="px-2.5 py-1 rounded-lg text-[11px] btn-primary flex-shrink-0">
                        Открыть
                      </button>
                      <button onClick={async () => {
                        await deleteSession(s.id);
                        setSavedSessions(prev => prev.filter(ss => ss.id !== s.id));
                      }} className="text-muted-foreground hover:text-red-400 transition-colors flex-shrink-0">
                        <Icon name="Trash2" size={13} />
                      </button>
                    </div>
                  ))}
                </div>
              )}
            </div>

            <div className="px-5 py-3 border-t border-border/40 text-right">
              <button onClick={() => setSessionOpen(false)} className="px-4 py-1.5 rounded-lg text-xs btn-primary">Закрыть</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}