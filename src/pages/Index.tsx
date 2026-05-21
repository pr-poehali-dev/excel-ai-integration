import { useState, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";
import Icon from "@/components/ui/icon";

// ─── Types ────────────────────────────────────────────────────────────────────

type CellValue = string | number | boolean | null;

interface SheetData {
  name: string;
  data: CellValue[][];
  colWidths: number[];
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

interface ChatMessage {
  role: "user" | "ai";
  text: string;
  ts: string;
  refs?: string[];
}

interface Selection {
  fileId: string;
  sheet: number;
  row: number;
  col: number;
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

function parseWorkbook(wb: XLSX.WorkBook, name: string, id: string): ExcelFile {
  const sheets: SheetData[] = wb.SheetNames.map((sn) => {
    const ws = wb.Sheets[sn];
    const raw: CellValue[][] = XLSX.utils.sheet_to_json(ws, {
      header: 1,
      defval: null,
      blankrows: true,
    }) as CellValue[][];
    // ensure at least 50 rows × 20 cols
    const ROWS = Math.max((raw.length || 0) + 10, 50);
    const COLS = Math.max(...raw.map((r) => r.length), 20);
    const data: CellValue[][] = Array.from({ length: ROWS }, (_, r) =>
      Array.from({ length: COLS }, (_, c) => raw[r]?.[c] ?? null)
    );
    const colWidths = Array.from({ length: COLS }, () => 100);
    return { name: sn, data, colWidths };
  });
  return { id, name, role: null, sheets, activeSheet: 0, workbook: wb, isDirty: false };
}

function buildWorkbook(file: ExcelFile): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();
  file.sheets.forEach((sh) => {
    const ws = XLSX.utils.aoa_to_sheet(sh.data);
    XLSX.utils.book_append_sheet(wb, ws, sh.name);
  });
  return wb;
}

function fileToContext(file: ExcelFile): string {
  return file.sheets
    .map((sh) => {
      const rows = sh.data.filter((r) => r.some((c) => c !== null));
      const preview = rows
        .slice(0, 40)
        .map((r) => r.join("\t"))
        .join("\n");
      return `### Файл «${file.name}» / Лист «${sh.name}»\n${preview}`;
    })
    .join("\n\n");
}

// ─── Mock AI ─────────────────────────────────────────────────────────────────

async function mockAiProcess(
  prompt: string,
  files: ExcelFile[]
): Promise<{ text: string; mutations?: { fileId: string; sheetName: string; data: CellValue[][] }[] }> {
  const lower = prompt.toLowerCase();

  // Simulate a delay
  await new Promise((r) => setTimeout(r, 1600));

  if (files.length === 0) {
    return { text: "Загрузи хотя бы один файл, чтобы я мог с ним работать." };
  }

  const mainFile = files.find((f) => f.role === "main") ?? files[0];
  const refFile = files.find((f) => f.role === "reference");

  // Aggregate / sum command
  if (lower.includes("сумм") || lower.includes("итог") || lower.includes("итого")) {
    const sh = mainFile.sheets[mainFile.activeSheet];
    const header = sh.data[0];
    const numericCols: number[] = [];
    header?.forEach((h, i) => {
      if (typeof h === "string" && (h.toLowerCase().includes("сумм") || h.toLowerCase().includes("цен") || h.toLowerCase().includes("кол"))) {
        numericCols.push(i);
      }
    });
    if (numericCols.length === 0) numericCols.push(...[...Array(header?.length ?? 3).keys()].slice(1, 4));

    const newSheet: CellValue[][] = [["Столбец", "Сумма", "Среднее", "Кол-во строк"]];
    numericCols.forEach((ci) => {
      const vals = sh.data
        .slice(1)
        .map((r) => Number(r[ci]))
        .filter((v) => !isNaN(v) && v !== 0);
      const sum = vals.reduce((a, b) => a + b, 0);
      newSheet.push([header?.[ci] ?? `Столбец ${ci + 1}`, Math.round(sum * 100) / 100, vals.length ? Math.round((sum / vals.length) * 100) / 100 : 0, vals.length]);
    });

    return {
      text: `Выполнено! Создан новый лист «Итоги_ИИ» в файле «${mainFile.name}» с агрегированными суммами по числовым столбцам (${numericCols.length} шт.). Данные из ${sh.data.length - 1} строк.`,
      mutations: [{ fileId: mainFile.id, sheetName: "Итоги_ИИ", data: newSheet }],
    };
  }

  // Copy + transform from ref to main
  if ((lower.includes("образец") || lower.includes("формат") || lower.includes("шаблон")) && refFile) {
    const refSh = refFile.sheets[0];
    const mainSh = mainFile.sheets[mainFile.activeSheet];
    const newRows: CellValue[][] = [refSh.data[0] ?? mainSh.data[0]];
    mainSh.data.slice(1).forEach((row) => newRows.push([...row]));
    return {
      text: `Готово! Создал лист «Результат_ИИ» в «${mainFile.name}» — применил заголовки из образца «${refFile.name}» и заполнил данными из основного файла.`,
      mutations: [{ fileId: mainFile.id, sheetName: "Результат_ИИ", data: newRows }],
    };
  }

  // Deduplicate
  if (lower.includes("дубл") || lower.includes("уникальн")) {
    const sh = mainFile.sheets[mainFile.activeSheet];
    const seen = new Set<string>();
    const unique = sh.data.filter((row) => {
      const key = JSON.stringify(row);
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
    return {
      text: `Удалено дублей: ${sh.data.length - unique.length}. Уникальных строк: ${unique.length}. Результат записан на лист «Без_дублей».`,
      mutations: [{ fileId: mainFile.id, sheetName: "Без_дублей", data: unique }],
    };
  }

  // Sort
  if (lower.includes("сортир") || lower.includes("упоряд")) {
    const sh = mainFile.sheets[mainFile.activeSheet];
    const header = sh.data[0] ?? [];
    const rows = [...sh.data.slice(1)].sort((a, b) => {
      const va = a[1] ?? "";
      const vb = b[1] ?? "";
      return String(va).localeCompare(String(vb), "ru");
    });
    return {
      text: `Данные отсортированы по второму столбцу. Результат — лист «Сортировка».`,
      mutations: [{ fileId: mainFile.id, sheetName: "Сортировка", data: [header, ...rows] }],
    };
  }

  // Generic with context
  const ctx = files.map((f) => `«${f.name}» (${f.sheets.length} лист(ов), ${f.sheets[0]?.data?.length ?? 0} строк)`).join(", ");
  return {
    text: `Вижу ${files.length} загруженных файл(а): ${ctx}.\n\nЯ могу:\n• **Посчитать итоги/суммы** → «посчитай суммы»\n• **Убрать дубли** → «удали дубликаты»\n• **Сортировать** → «отсортируй данные»\n• **Применить формат из образца** → «возьми формат из файла 2»\n\nНапиши конкретное задание!`,
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
    <input
      ref={ref}
      value={v}
      onChange={(e) => setV(e.target.value)}
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

// ─── Main Component ───────────────────────────────────────────────────────────

const VISIBLE_ROWS = 60;
const VISIBLE_COLS = 26;

export default function Index() {
  const [files, setFiles] = useState<ExcelFile[]>([]);
  const [activeFileId, setActiveFileId] = useState<string | null>(null);
  const [editing, setEditing] = useState<{ row: number; col: number } | null>(null);
  const [selected, setSelected] = useState<{ row: number; col: number } | null>(null);
  const [messages, setMessages] = useState<ChatMessage[]>([
    {
      role: "ai",
      text: "Загрузи один или несколько Excel-файлов. Я вижу все их содержимое и могу выполнять задания — считать, трансформировать, создавать новые листы. Пример: «возьми данные из Файл1, образец из Файл2, сделай итоговую таблицу на новом листе».",
      ts: getTime(),
    },
  ]);
  const [aiInput, setAiInput] = useState("");
  const [aiThinking, setAiThinking] = useState(false);
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const chatEndRef = useRef<HTMLDivElement>(null);
  const tableRef = useRef<HTMLDivElement>(null);

  const activeFile = files.find((f) => f.id === activeFileId) ?? null;
  const activeSheet = activeFile ? activeFile.sheets[activeFile.activeSheet] : null;

  // ── Load file ──
  const loadFile = useCallback((file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target!.result as ArrayBuffer);
      const wb = XLSX.read(data, { type: "array" });
      const id = uid();
      const ef = parseWorkbook(wb, file.name, id);
      setFiles((prev) => {
        const next = [...prev, ef];
        return next;
      });
      setActiveFileId(id);
      setMessages((prev) => [
        ...prev,
        {
          role: "ai",
          text: `Файл **«${file.name}»** загружен: ${ef.sheets.length} лист(а/ов), первый лист — ${ef.sheets[0]?.data?.filter((r) => r.some((c) => c !== null)).length ?? 0} строк данных.`,
          ts: getTime(),
          refs: [file.name],
        },
      ]);
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const handleFiles = useCallback(
    (fileList: FileList) => {
      Array.from(fileList).forEach((f) => {
        if (f.name.match(/\.(xlsx|xls|csv)$/i)) loadFile(f);
      });
    },
    [loadFile]
  );

  // ── Cell update ──
  const updateCell = useCallback(
    (row: number, col: number, value: CellValue) => {
      if (!activeFile) return;
      setFiles((prev) =>
        prev.map((f) => {
          if (f.id !== activeFile.id) return f;
          const sheets = f.sheets.map((sh, si) => {
            if (si !== f.activeSheet) return sh;
            const data = sh.data.map((r, ri) =>
              ri === row ? r.map((c, ci) => (ci === col ? value : c)) : r
            );
            return { ...sh, data };
          });
          return { ...f, sheets, isDirty: true };
        })
      );
    },
    [activeFile]
  );

  // ── Add new sheet ──
  const addSheet = useCallback(
    (fileId: string, name: string, data: CellValue[][]) => {
      setFiles((prev) =>
        prev.map((f) => {
          if (f.id !== fileId) return f;
          const ROWS = Math.max(data.length + 10, 50);
          const COLS = Math.max(...data.map((r) => r.length), 20);
          const padded = Array.from({ length: ROWS }, (_, r) =>
            Array.from({ length: COLS }, (_, c) => data[r]?.[c] ?? null)
          );
          const existing = f.sheets.findIndex((s) => s.name === name);
          let sheets = [...f.sheets];
          if (existing >= 0) {
            sheets[existing] = { name, data: padded, colWidths: Array(COLS).fill(100) };
          } else {
            sheets = [...sheets, { name, data: padded, colWidths: Array(COLS).fill(100) }];
          }
          return { ...f, sheets, activeSheet: sheets.length - 1, isDirty: true };
        })
      );
    },
    []
  );

  // ── Save file ──
  const saveFile = useCallback(
    (fileId: string) => {
      const f = files.find((ff) => ff.id === fileId);
      if (!f) return;
      const wb = buildWorkbook(f);
      XLSX.writeFile(wb, f.name);
      setFiles((prev) => prev.map((ff) => (ff.id === fileId ? { ...ff, isDirty: false } : ff)));
    },
    [files]
  );

  // ── AI send ──
  const handleAiSend = async () => {
    const text = aiInput.trim();
    if (!text || aiThinking) return;
    setMessages((prev) => [...prev, { role: "user", text, ts: getTime() }]);
    setAiInput("");
    setAiThinking(true);
    try {
      const result = await mockAiProcess(text, files);
      if (result.mutations) {
        result.mutations.forEach((m) => addSheet(m.fileId, m.sheetName, m.data));
        setActiveFileId(result.mutations[0].fileId);
      }
      setMessages((prev) => [...prev, { role: "ai", text: result.text, ts: getTime() }]);
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
      if (editing || !selected || !activeSheet) return;
      const { row, col } = selected;
      const maxRow = activeSheet.data.length - 1;
      const maxCol = (activeSheet.data[0]?.length ?? 1) - 1;
      if (e.key === "ArrowDown") { e.preventDefault(); setSelected({ row: Math.min(row + 1, maxRow), col }); }
      if (e.key === "ArrowUp") { e.preventDefault(); setSelected({ row: Math.max(row - 1, 0), col }); }
      if (e.key === "ArrowRight") { e.preventDefault(); setSelected({ row, col: Math.min(col + 1, maxCol) }); }
      if (e.key === "ArrowLeft") { e.preventDefault(); setSelected({ row, col: Math.max(col - 1, 0) }); }
      if (e.key === "Enter" || e.key === "F2") { e.preventDefault(); setEditing(selected); }
      if (e.key === "Delete" || e.key === "Backspace") { updateCell(row, col, null); }
      if (e.key.length === 1 && !e.ctrlKey && !e.metaKey) {
        setEditing(selected);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [editing, selected, activeSheet, updateCell]);

  // ── Save shortcut ──
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

  const visibleRows = activeSheet ? Math.min(activeSheet.data.length, VISIBLE_ROWS) : 0;
  const visibleCols = activeSheet ? Math.min(activeSheet.data[0]?.length ?? 0, VISIBLE_COLS) : 0;

  return (
    <div className="h-screen flex flex-col overflow-hidden">
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
            <button
              key={f.id}
              onClick={() => { setActiveFileId(f.id); setEditing(null); setSelected(null); }}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium flex-shrink-0 transition-all ${
                f.id === activeFileId
                  ? "bg-primary/10 text-primary border border-primary/30"
                  : "text-muted-foreground hover:text-foreground hover:bg-secondary"
              }`}
            >
              <Icon name="FileSpreadsheet" size={13} />
              <span className="max-w-[120px] truncate">{f.name}</span>
              {f.isDirty && <span className="w-1.5 h-1.5 rounded-full bg-amber-400 flex-shrink-0" />}
              {f.role && (
                <span className={`text-[10px] px-1 rounded ${f.role === "main" ? "tag-emerald" : "tag-blue"}`}>
                  {f.role === "main" ? "осн." : "обр."}
                </span>
              )}
              <button
                onClick={(e) => { e.stopPropagation(); setFiles((prev) => prev.filter((ff) => ff.id !== f.id)); if (activeFileId === f.id) setActiveFileId(files.find((ff) => ff.id !== f.id)?.id ?? null); }}
                className="ml-0.5 opacity-40 hover:opacity-100 transition-opacity"
              >
                <Icon name="X" size={11} />
              </button>
            </button>
          ))}
          <button
            onClick={() => fileInputRef.current?.click()}
            className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs text-muted-foreground hover:text-foreground hover:bg-secondary transition-all flex-shrink-0"
          >
            <Icon name="Plus" size={13} />
            <span className="hidden sm:inline">Открыть</span>
          </button>
        </div>

        {/* Right actions */}
        <div className="flex items-center gap-2 flex-shrink-0">
          {activeFile && (
            <>
              {/* Role assign */}
              <select
                value={activeFile.role ?? ""}
                onChange={(e) => setFiles((prev) => prev.map((f) => f.id === activeFileId ? { ...f, role: (e.target.value as "main" | "reference") || null } : f))}
                className="text-xs bg-secondary border border-border/60 text-foreground rounded-lg px-2 py-1.5 outline-none hidden sm:block"
              >
                <option value="">— роль —</option>
                <option value="main">Основной</option>
                <option value="reference">Образец</option>
              </select>
              <button
                onClick={() => saveFile(activeFile.id)}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold transition-all ${activeFile.isDirty ? "btn-primary" : "bg-secondary text-muted-foreground"}`}
              >
                <Icon name="Save" size={13} />
                <span className="hidden sm:inline">Сохранить</span>
              </button>
            </>
          )}
          <button
            onClick={() => setSidebarOpen((p) => !p)}
            className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs text-muted-foreground hover:text-foreground hover:bg-secondary transition-all"
          >
            <Icon name="MessageSquare" size={14} />
          </button>
        </div>
      </header>

      {/* ── Main ── */}
      <div className="flex flex-1 overflow-hidden">

        {/* ── Table area ── */}
        <div className="flex flex-col flex-1 overflow-hidden min-w-0">

          {/* Sheet tabs + formula bar */}
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
                  {selected ? String(activeSheet.data[selected.row]?.[selected.col] ?? "") : ""}
                </span>
              </div>

              {/* Sheet name tabs */}
              <div className="flex items-center border-b border-border/40 px-2 gap-0.5 flex-shrink-0 overflow-x-auto scrollbar-thin"
                style={{ background: "rgba(255,255,255,0.01)" }}>
                {activeFile.sheets.map((sh, si) => (
                  <button
                    key={si}
                    onClick={() => setFiles((prev) => prev.map((f) => f.id === activeFileId ? { ...f, activeSheet: si } : f))}
                    className={`px-3 py-2 text-xs font-medium border-b-2 -mb-px transition-all whitespace-nowrap ${
                      si === activeFile.activeSheet
                        ? "border-primary text-primary"
                        : "border-transparent text-muted-foreground hover:text-foreground"
                    }`}
                  >
                    {sh.name}
                  </button>
                ))}
                <button
                  onClick={() => {
                    const name = `Лист${activeFile.sheets.length + 1}`;
                    setFiles((prev) => prev.map((f) => {
                      if (f.id !== activeFileId) return f;
                      const empty = Array.from({ length: 50 }, () => Array(20).fill(null));
                      const sheets = [...f.sheets, { name, data: empty, colWidths: Array(20).fill(100) }];
                      return { ...f, sheets, activeSheet: sheets.length - 1, isDirty: true };
                    }));
                  }}
                  className="px-2 py-2 text-xs text-muted-foreground hover:text-foreground transition-colors flex-shrink-0"
                >
                  <Icon name="Plus" size={13} />
                </button>
              </div>
            </>
          )}

          {/* ── Drop zone / empty state ── */}
          {files.length === 0 && (
            <div
              className="flex-1 flex flex-col items-center justify-center gap-6 p-8 cursor-pointer"
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => { e.preventDefault(); handleFiles(e.dataTransfer.files); }}
              onClick={() => fileInputRef.current?.click()}
            >
              <div className="w-20 h-20 rounded-3xl bg-secondary flex items-center justify-center"
                style={{ boxShadow: "0 0 60px rgba(52,211,153,0.08)" }}>
                <Icon name="FileSpreadsheet" size={40} className="text-primary" />
              </div>
              <div className="text-center">
                <p className="text-lg font-semibold text-foreground mb-2">Открой Excel-файл</p>
                <p className="text-sm text-muted-foreground mb-1">Перетащи один или несколько файлов сюда</p>
                <p className="text-xs text-muted-foreground">.xlsx · .xls · .csv</p>
              </div>
              <div className="grid grid-cols-3 gap-3 max-w-sm text-center">
                {[
                  { icon: "Edit3", t: "Редактируй ячейки", d: "Как в Excel" },
                  { icon: "Layers", t: "Несколько файлов", d: "Все доступны ИИ" },
                  { icon: "Brain", t: "ИИ-задания", d: "Новые листы автоматом" },
                ].map((f, i) => (
                  <div key={i} className="p-3 rounded-xl border border-border/30" style={{ background: "rgba(255,255,255,0.02)" }}>
                    <Icon name={f.icon} size={18} className="text-primary mx-auto mb-1.5" />
                    <p className="text-xs font-medium text-foreground">{f.t}</p>
                    <p className="text-[11px] text-muted-foreground">{f.d}</p>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* ── Spreadsheet ── */}
          {activeFile && activeSheet && (
            <div
              ref={tableRef}
              className="flex-1 overflow-auto scrollbar-thin"
              style={{ background: "hsl(220,16%,5.5%)" }}
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => { e.preventDefault(); handleFiles(e.dataTransfer.files); }}
            >
              <table className="border-collapse" style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: "12px" }}>
                <thead>
                  <tr>
                    {/* Row num header */}
                    <th className="sticky left-0 top-0 z-30 w-10 text-center text-[11px] text-muted-foreground font-normal border-b border-r border-border/50 select-none"
                      style={{ background: "hsl(220,14%,8%)", minWidth: 40 }} />
                    {Array.from({ length: visibleCols }, (_, ci) => (
                      <th
                        key={ci}
                        className="sticky top-0 z-20 text-center text-[11px] text-muted-foreground font-medium border-b border-r border-border/50 select-none px-1 py-1"
                        style={{ background: "hsl(220,14%,8%)", minWidth: activeSheet.colWidths[ci] ?? 100 }}
                      >
                        {colLetter(ci)}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {Array.from({ length: visibleRows }, (_, ri) => (
                    <tr key={ri} className="group">
                      {/* Row number */}
                      <td className="sticky left-0 z-10 text-center text-[11px] text-muted-foreground border-b border-r border-border/30 select-none px-1"
                        style={{ background: "hsl(220,14%,8%)", minWidth: 40 }}>
                        {ri + 1}
                      </td>
                      {Array.from({ length: visibleCols }, (_, ci) => {
                        const isSelected = selected?.row === ri && selected?.col === ci;
                        const isEditing = editing?.row === ri && editing?.col === ci;
                        const val = activeSheet.data[ri]?.[ci];
                        return (
                          <td
                            key={ci}
                            className={`relative border-b border-r border-border/20 px-1.5 py-0.5 cursor-cell select-none transition-colors ${
                              isSelected
                                ? "bg-primary/10 outline outline-1 outline-primary z-10"
                                : "hover:bg-white/[0.02]"
                            }`}
                            style={{ minWidth: activeSheet.colWidths[ci] ?? 100, maxWidth: 300, height: 24 }}
                            onClick={() => { setSelected({ row: ri, col: ci }); setEditing(null); }}
                            onDoubleClick={() => { setSelected({ row: ri, col: ci }); setEditing({ row: ri, col: ci }); }}
                          >
                            {isEditing ? (
                              <CellEditor
                                value={val}
                                onCommit={(v) => { updateCell(ri, ci, v); setEditing(null); setSelected({ row: ri, col: ci + 1 }); }}
                                onCancel={() => setEditing(null)}
                              />
                            ) : (
                              <span className={`block truncate text-[12px] ${typeof val === "number" ? "text-right text-emerald-300/90" : "text-foreground/90"}`}>
                                {val !== null && val !== undefined ? String(val) : ""}
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
                  {files.length === 0 ? "Нет файлов" : `${files.length} файл(а) загружено`}
                </p>
              </div>
              {aiThinking && <Icon name="Loader2" size={14} className="text-primary spinner flex-shrink-0" />}
            </div>

            {/* File context pills */}
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
                  <div
                    className={`max-w-[86%] rounded-xl px-3 py-2 text-xs leading-relaxed ${
                      msg.role === "ai"
                        ? "rounded-tl-sm border border-border/40"
                        : "btn-amber rounded-tr-sm"
                    }`}
                    style={msg.role === "ai" ? { background: "rgba(255,255,255,0.03)" } : {}}
                  >
                    <p dangerouslySetInnerHTML={{ __html: renderMarkdown(msg.text) }} />
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
                    <button
                      key={q}
                      onClick={() => setAiInput(q)}
                      className="w-full text-left text-[11px] px-2.5 py-1.5 rounded-lg border border-border/50 text-muted-foreground hover:text-foreground hover:border-primary/30 transition-all"
                      style={{ background: "rgba(255,255,255,0.02)" }}
                    >
                      {q}
                    </button>
                  ))}
                </div>
              </div>
            )}

            {/* Input */}
            <div className="p-3 border-t border-border/40 flex-shrink-0">
              <div className="flex gap-2">
                <textarea
                  value={aiInput}
                  onChange={(e) => setAiInput(e.target.value)}
                  onKeyDown={(e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleAiSend(); } }}
                  placeholder={files.length === 0 ? "Сначала загрузи файл..." : "Задание для ИИ..."}
                  disabled={files.length === 0 || aiThinking}
                  rows={2}
                  className="flex-1 px-3 py-2 rounded-lg border border-border/60 text-xs text-foreground placeholder:text-muted-foreground resize-none focus:outline-none focus:border-primary/50 transition-all disabled:opacity-40"
                  style={{ background: "rgba(255,255,255,0.03)", fontFamily: "'IBM Plex Sans', sans-serif" }}
                />
                <button
                  onClick={handleAiSend}
                  disabled={!aiInput.trim() || aiThinking || files.length === 0}
                  className="w-9 h-9 self-end rounded-lg btn-primary flex items-center justify-center flex-shrink-0 disabled:opacity-30 disabled:cursor-not-allowed"
                >
                  <Icon name="Send" size={13} />
                </button>
              </div>
              <p className="text-[10px] text-muted-foreground mt-1.5 text-center">Ctrl+S — сохранить · Enter — отправить</p>
            </div>
          </div>
        )}
      </div>

      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        multiple
        className="hidden"
        onChange={(e) => e.target.files && handleFiles(e.target.files)}
      />
    </div>
  );
}