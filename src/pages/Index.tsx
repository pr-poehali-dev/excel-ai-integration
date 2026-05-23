import { useState, useRef, useCallback, useEffect } from "react";
import { useNavigate } from "react-router-dom";
import * as XLSX from "xlsx";
import { PieChart, Pie, Cell, Tooltip, Legend, ResponsiveContainer, BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid } from "recharts";
import Icon from "@/components/ui/icon";

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
// Возвращает значение ячейки-хозяина объединения (верхняя левая)
function getMergeOwner(sh: SheetData, r: number, c: number): { r: number; c: number } | null {
  for (const m of sh.merges) {
    if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) {
      if (r === m.r1 && c === m.c1) return null; // сама хозяйская ячейка
      return { r: m.r1, c: m.c1 };
    }
  }
  return null;
}

// Строит карту: для каждого столбца — стек заголовков сверху (из объединённых строк-заголовков)
function buildMergeContext(sh: SheetData, maxCols: number): string {
  if (!sh.merges.length) return "";

  const lines: string[] = ["ОБЪЕДИНЁННЫЕ ЯЧЕЙКИ И ИЕРАРХИЯ ЗАГОЛОВКОВ:"];

  // Собираем все объединения, у которых значение непустое
  for (const m of sh.merges) {
    const ownerCell = sh.cells[m.r1]?.[m.c1];
    const val = ownerCell?.w ?? (ownerCell?.v != null ? String(ownerCell.v) : "");
    if (!val.trim()) continue;

    const r1l = m.r1, r2l = m.r2, c1l = m.c1, c2l = m.c2;
    const colRange = c1l === c2l
      ? colLetter(c1l)
      : `${colLetter(c1l)}-${colLetter(c2l)}`;
    const rowRange = r1l === r2l ? `строка ${r1l}` : `строки ${r1l}-${r2l}`;
    lines.push(`  «${val}» → ${rowRange}, столбцы ${colRange} (col ${c1l}..${c2l})`);
  }

  // Строим для каждого столбца список заголовков сверху вниз
  const MAX_HEADER_ROWS = 20;
  const colHeaders: Record<number, string[]> = {};
  for (let c = 0; c < maxCols; c++) {
    const stack: string[] = [];
    for (let r = 0; r < MAX_HEADER_ROWS && r < sh.cells.length; r++) {
      // Определяем хозяина (если ячейка вторична — берём хозяина)
      const owner = getMergeOwner(sh, r, c);
      const cell = owner
        ? sh.cells[owner.r]?.[owner.c]
        : sh.cells[r]?.[c];
      const val = cell?.w ?? (cell?.v != null ? String(cell.v) : "");
      if (val.trim()) stack.push(val.trim());
    }
    if (stack.length > 1) colHeaders[c] = stack;
  }

  if (Object.keys(colHeaders).length) {
    lines.push("\nИЕРАРХИЯ ЗАГОЛОВКОВ ПО СТОЛБЦАМ (только столбцы с многоуровневыми заголовками):");
    for (const [c, stack] of Object.entries(colHeaders)) {
      lines.push(`  col ${c} (${colLetter(Number(c))}): ${stack.join(" → ")}`);
    }
  }

  return lines.join("\n");
}

function sheetToText(sh: SheetData, full = false): string {
  const MAX_ROWS = full ? 300 : 5;
  const MAX_COLS = 26; // до колонки Z
  const MAX_CELL_LEN = 40;

  const numCols = Math.min(sh.cells[0]?.length ?? 0, MAX_COLS);

  // Строки данных — для вторичных ячеек объединения показываем значение хозяина со знаком ↑ или ←
  const dataRows: string[] = [];
  sh.cells.forEach((row, ri) => {
    if (!row.some(c => c.v !== null)) return;
    const vals = row.slice(0, MAX_COLS).map((c, ci) => {
      const owner = getMergeOwner(sh, ri, ci);
      if (owner) {
        // Вторичная ячейка: показываем откуда берётся значение
        const ownerCell = sh.cells[owner.r]?.[owner.c];
        const ownerVal = ownerCell?.w ?? (ownerCell?.v != null ? String(ownerCell.v) : "");
        if (!ownerVal.trim()) return "↑merged";
        const short = ownerVal.length > 20 ? ownerVal.slice(0, 20) + "…" : ownerVal;
        return `[merged:«${short}»]`;
      }
      const s = c.w ?? (c.v !== null ? String(c.v) : "");
      return s.length > MAX_CELL_LEN ? s.slice(0, MAX_CELL_LEN) + "…" : s;
    }).join("\t");
    dataRows.push(`${ri}\t${vals}`);
  });

  const totalDataRows = dataRows.length;
  const sliced = dataRows.slice(0, MAX_ROWS);

  const header = "ROW\t" + Array.from({ length: numCols }, (_, i) => colLetter(i)).join("\t");
  const suffix = !full && totalDataRows > MAX_ROWS ? `\n[...ещё ${totalDataRows - MAX_ROWS} строк не показано]` : "";

  const mergeCtx = full ? "\n\n" + buildMergeContext(sh, numCols) : "";

  return [header, ...sliced].join("\n") + suffix + mergeCtx;
}

// ─── AI Settings ─────────────────────────────────────────────────────────────

const EXCEL_CHART_URL = "https://functions.poehali.dev/8376b365-14fc-4551-9fc7-0798d13ac4e6";

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

const SYSTEM_PROMPT = `Ты — профессиональный аналитик данных и эксперт по Excel.
Тебе передают содержимое одного или нескольких Excel-файлов в виде текста (TSV), задание пользователя, и (опционально) изображения — скриншоты графиков, примеры оформления.

КРИТИЧЕСКИ ВАЖНО: ты ВСЕГДА отвечаешь ТОЛЬКО валидным JSON и НИЧЕМ КРОМЕ JSON. Никакого текста до или после. Никаких \`\`\`json\`\`\` обёрток.

Формат ответа:
{
  "text": "Краткое объяснение что сделано (на русском, 1-3 предложения)",
  "new_sheet": {
    "file_index": 0,
    "sheet_name": "Название_листа",
    "data": [["Заголовок1","Заголовок2"],["значение1","значение2"]]
  },
  "cell_styles": [
    {
      "file_index": 0,
      "sheet_name": "Лист1",
      "changes": [
        {"row": 2, "col": 0, "bgColor": "FFFFFF00"}
      ]
    }
  ]
}

ПРАВИЛА:
1. "text" — ОБЯЗАТЕЛЬНО всегда
2. "new_sheet" — только когда нужно создать новый лист с данными/расчётами
3. "cell_styles" — когда нужно покрасить/форматировать ячейки в СУЩЕСТВУЮЩЕМ листе
4. row/col в cell_styles — 0-based. Колонка ROW в данных файла — это row-индекс
5. Чтобы покрасить ВСЕЮ СТРОКУ — укажи изменения для каждой колонки строки (col: 0, 1, 2, ... до последней колонки с данными)
6. bgColor — AARRGGBB: жёлтый=FFFFFF00, оранжевый=FFFFA500, зелёный=FF92D050, красный=FFFF0000, голубой=FF00B0F0
7. fontColor — AARRGGBB цвет текста
8. bold: true/false
9. Числа в data — числа, текст — строки
10. Если просят покрасить строки — cell_styles, НЕ новый лист
11. Изображение — используй как образец структуры/оформления

КАК ЧИТАТЬ ОБЪЕДИНЁННЫЕ ЯЧЕЙКИ (MERGES):
В данных файла ячейки помечены специально:
- [merged:«Заголовок»] — эта ячейка входит в объединение, хозяин которого содержит текст «Заголовок»
- ↑merged — ячейка входит в объединение с пустым хозяином

В секции "ОБЪЕДИНЁННЫЕ ЯЧЕЙКИ И ИЕРАРХИЯ ЗАГОЛОВКОВ" показано:
- Какой текст стоит в хозяйской ячейке объединения
- Какие строки и столбцы (col N..M) это объединение охватывает
- Иерархия: для каждого столбца перечислены все заголовки сверху вниз через →

КРИТИЧЕСКИ ВАЖНО для работы с объединёнными заголовками:
- Если столбец P содержит в иерархии "Пласт А → с начала разработки → добыча нефти" — это значит, что данные в col P относятся к пласту А, показатель "с начала разработки", тип "добыча нефти"
- Объединённая ячейка-заголовок распространяется на ВСЕ столбцы внутри своего диапазона (col N..M)
- При подсчёте сумм/итогов нужно брать данные из ВСЕХ столбцов, которые входят в нужный заголовочный диапазон
- Если просят "данные с начала разработки" — ищи в иерархии заголовков столбец с таким текстом или синонимом ("накопленная", "с н.р.", "НР", "с начала")
- Если просят данные "по пластам" — каждый пласт это отдельный диапазон объединённого заголовка; перебери все такие диапазоны`;

type CellStyleChange = { row: number; col: number; bgColor?: string; fontColor?: string; bold?: boolean };
type CellStyleMutation = { fileId: string; sheetName: string; changes: CellStyleChange[] };

function extractJson(raw: string): Record<string, unknown> {
  raw = raw.trim().replace(/^```(?:json)?\s*/m, "").replace(/\s*```$/m, "").trim();
  try { return JSON.parse(raw); } catch { /* fall through */ }
  const m = raw.match(/\{[\s\S]*\}/);
  if (m) { try { return JSON.parse(m[0]); } catch { /* fall through */ } }
  return { text: raw || "ИИ вернул пустой ответ" };
}

async function callAi(
  prompt: string,
  files: ExcelFile[],
  settings: AiSettings,
  images: ChatImage[]
): Promise<{
  text: string;
  mutations?: { fileId: string; sheetName: string; data: CellValue[][] }[];
  styleMutations?: CellStyleMutation[];
}> {
  const selectedModel = settings.model === "__custom__" ? settings.customModel : settings.model;
  // Модели без поддержки vision — при наличии картинок автоматически переключаем на gpt-4o-mini
  const VISION_UNSUPPORTED = ["deepseek/deepseek-chat", "deepseek/deepseek-r1"];
  const effectiveModel = (images.length > 0 && VISION_UNSUPPORTED.includes(selectedModel))
    ? "openai/gpt-4o-mini"
    : selectedModel;
  const baseUrl = settings.baseUrl.replace(/\/$/, "");

  // Строим контекст файлов
  const contextParts: string[] = [];
  files.forEach((f, fi) => {
    const roleLabel = f.role === "main" ? " [ОСНОВНОЙ]" : f.role === "reference" ? " [ОБРАЗЕЦ]" : "";
    contextParts.push(`=== Файл ${fi} «${f.name}»${roleLabel} ===`);
    f.sheets.forEach((sh, si) => {
      const isActive = si === f.activeSheet;
      const totalRows = sh.cells.filter(r => r.some(c => c.v !== null)).length;
      const numCols = Math.min(sh.cells[0]?.length ?? 0, 26);
      const marker = isActive ? " [АКТИВНЫЙ — полные данные]" : ` [краткий просмотр, всего ${totalRows} строк]`;
      contextParts.push(`--- Лист «${sh.name}»${marker} ---`);
      // Для активного листа сначала выводим структуру объединений
      if (isActive && sh.merges.length > 0) {
        contextParts.push(buildMergeContext(sh, numCols));
        contextParts.push("\nДАННЫЕ ЛИСТА (в ячейках [merged:«...»] — принадлежность к объединению):");
      }
      contextParts.push(sheetToText(sh, isActive));
    });
  });

  const textBlock = `ДАННЫЕ ФАЙЛОВ:\n${contextParts.join("\n")}\n\nЗАДАНИЕ: ${prompt || "(см. изображения)"}\n\nОтветь ТОЛЬКО JSON.`;

  // Строим сообщение (с картинками или без)
  type ContentPart = { type: "text"; text: string } | { type: "image_url"; image_url: { url: string } };
  let userContent: string | ContentPart[];
  if (images.length > 0) {
    userContent = [{ type: "text", text: textBlock } as ContentPart];
    for (const img of images) {
      (userContent as ContentPart[]).push({
        type: "image_url",
        image_url: { url: img.dataUrl },
      });
    }
  } else {
    userContent = textBlock;
  }

  const resp = await fetch(`${baseUrl}/chat/completions`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${settings.apiKey}`,
    },
    body: JSON.stringify({
      model: effectiveModel,
      messages: [
        { role: "system", content: SYSTEM_PROMPT },
        { role: "user", content: userContent },
      ],
      max_tokens: 4000,
      temperature: 0.1,
      ...(settings.reasoningEffort !== "none"
        ? { reasoning_effort: settings.reasoningEffort }
        : {}),
    }),
  });

  if (!resp.ok) {
    const err = await resp.json().catch(() => ({})) as { error?: { message?: string } };
    throw new Error(err.error?.message || `Ошибка ${resp.status}`);
  }

  const json = await resp.json() as { choices: { message: { content: string } }[] };
  const raw = json.choices?.[0]?.message?.content ?? "{}";
  const result = extractJson(raw);

  const mutations: { fileId: string; sheetName: string; data: CellValue[][] }[] = [];
  if (result.new_sheet) {
    const ns = result.new_sheet as { file_index?: number; sheet_name: string; data: CellValue[][] };
    const targetFile = files[ns.file_index ?? 0] ?? files[0];
    if (targetFile) mutations.push({ fileId: targetFile.id, sheetName: ns.sheet_name, data: ns.data });
  }

  const styleMutations: CellStyleMutation[] = [];
  if (Array.isArray(result.cell_styles)) {
    for (const cs of result.cell_styles as { file_index?: number; sheet_name: string; changes: CellStyleChange[] }[]) {
      const targetFile = files[cs.file_index ?? 0] ?? files[0];
      if (targetFile && cs.sheet_name && Array.isArray(cs.changes)) {
        styleMutations.push({ fileId: targetFile.id, sheetName: cs.sheet_name, changes: cs.changes });
      }
    }
  }

  return {
    text: (result.text as string) || "Готово!",
    mutations: mutations.length ? mutations : undefined,
    styleMutations: styleMutations.length ? styleMutations : undefined,
  };
}

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
    text: "Загрузи один или несколько Excel-файлов. Я вижу все их содержимое и могу выполнять задания — считать, трансформировать, создавать новые листы.\n\nМожно прикрепить **скриншот графика** — я пойму как его воспроизвести на твоих данных.",
    ts: getTime(),
  }]);
  const [aiInput, setAiInput] = useState("");
  const [pendingImages, setPendingImages] = useState<ChatImage[]>([]);
  const [aiThinking, setAiThinking] = useState(false);
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const navigate = useNavigate();
  const [aiSettings, setAiSettings] = useState<AiSettings>(loadSettings);
  const [, setAiError] = useState<string | null>(null);
  const [isListening, setIsListening] = useState(false);

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
    });
  }, [loadFile]);

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
      const result = await callAi(text, files, aiSettings, imgs);

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
              if (!cells[row]) return;
              if (!cells[row][col]) cells[row][col] = { v: null };
              const prev = cells[row][col];
              const newStyle: CellStyle = { ...(prev.s ?? {}) };
              if (bgColor) newStyle.bgColor = `#${bgColor.slice(2)}`;
              if (fontColor) newStyle.color = `#${fontColor.slice(2)}`;
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

      setMessages((prev) => [...prev, { role: "ai", text: result.text, ts: getTime(), chartData, chartTitle }]);
    } catch (err) {
      const msg = err instanceof Error ? err.message : "Неизвестная ошибка";
      setAiError(msg);
      setMessages((prev) => [...prev, { role: "ai", text: `⚠️ Ошибка: ${msg}`, ts: getTime() }]);
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
              onClick={() => { setActiveFileId(f.id); setEditing(null); setSelected(null); }}
              onKeyDown={(e) => { if (e.key === "Enter") { setActiveFileId(f.id); setEditing(null); setSelected(null); } }}
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
          <button onClick={() => fileInputRef.current?.click()}
            className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs text-muted-foreground hover:text-foreground hover:bg-secondary transition-all flex-shrink-0">
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
              {/* Formula bar */}
              <div className="flex items-center gap-2 px-3 py-1.5 border-b border-border/40 flex-shrink-0"
                style={{ background: "rgba(255,255,255,0.01)" }}>
                <span className="text-xs font-mono text-muted-foreground w-12 text-center flex-shrink-0">
                  {selected ? `${colLetter(selected.col)}${selected.row + 1}` : "—"}
                </span>
                <div className="h-4 w-px bg-border/60" />
                <span className="text-xs font-mono text-foreground flex-1 truncate">
                  {selected ? String(activeSheet.cells[selected.row]?.[selected.col]?.v ?? "") : ""}
                </span>
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
          {files.length === 0 && (
            <div className="flex-1 flex flex-col items-center justify-center gap-6 p-8 cursor-pointer"
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => { e.preventDefault(); handleFiles(e.dataTransfer.files); }}
              onClick={() => fileInputRef.current?.click()}>
              <div className="w-20 h-20 rounded-3xl bg-secondary flex items-center justify-center"
                style={{ boxShadow: "0 0 60px rgba(52,211,153,0.08)" }}>
                <Icon name="FileSpreadsheet" size={40} className="text-primary" />
              </div>
              <div className="text-center">
                <p className="text-foreground font-semibold text-lg mb-1">Перетащи Excel-файл</p>
                <p className="text-muted-foreground text-sm">или нажми чтобы выбрать · .xlsx, .xls, .csv</p>
              </div>
            </div>
          )}

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

            {files.length > 0 && (
              <div className="px-3 py-2 border-b border-border/30 flex flex-wrap gap-1.5">
                {files.map((f) => (
                  <div key={f.id} className="flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] tag-emerald">
                    <Icon name="FileSpreadsheet" size={10} />
                    <span className="max-w-[80px] truncate">{f.name}</span>
                  </div>
                ))}
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
                  placeholder={isListening ? "🎤 Говорите..." : (files.length === 0 ? "Сначала загрузи файл..." : "Задание для ИИ...")}
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
      <input ref={imageInputRef} type="file" accept="image/*" multiple className="hidden"
        onChange={(e) => { if (e.target.files) Array.from(e.target.files).forEach(loadImage); e.target.value = ""; }} />
    </div>
  );
}