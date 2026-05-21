import { useState, useRef, useCallback } from "react";
import Icon from "@/components/ui/icon";

type FileStatus = "idle" | "loading" | "ready" | "error";
type AiStatus = "idle" | "thinking" | "done";

interface Column {
  name: string;
  type: string;
  nulls: number;
  sample: string;
}

interface SheetInfo {
  name: string;
  rows: number;
  cols: number;
  columns: Column[];
}

interface ChatMessage {
  role: "user" | "ai";
  text: string;
  ts: string;
}

const MOCK_SHEET: SheetInfo = {
  name: "Продажи_Q1",
  rows: 4821,
  cols: 12,
  columns: [
    { name: "ID", type: "Целое", nulls: 0, sample: "1, 2, 3..." },
    { name: "Дата", type: "Дата", nulls: 3, sample: "2024-01-15" },
    { name: "Менеджер", type: "Текст", nulls: 0, sample: "Иванов А." },
    { name: "Продукт", type: "Текст", nulls: 0, sample: "Ноутбук Pro" },
    { name: "Количество", type: "Целое", nulls: 7, sample: "1–50" },
    { name: "Цена", type: "Число", nulls: 0, sample: "12 500.00" },
    { name: "Сумма", type: "Число", nulls: 7, sample: "625 000.00" },
    { name: "Регион", type: "Текст", nulls: 12, sample: "Москва" },
    { name: "Категория", type: "Текст", nulls: 0, sample: "Электроника" },
    { name: "Статус", type: "Текст", nulls: 4, sample: "Выполнен" },
    { name: "Канал", type: "Текст", nulls: 18, sample: "Онлайн" },
    { name: "Комментарий", type: "Текст", nulls: 389, sample: "—" },
  ],
};

const MOCK_AI_REPLIES: Record<string, string> = {
  default:
    "Данные загружены и проанализированы. В таблице обнаружено **12 столбцов** и **4 821 строка**. Есть пропуски в полях «Дата» (3), «Количество» (7) и «Комментарий» (389) — рекомендую заполнить или исключить перед финальным анализом.",
  пропуск:
    "Пропущенные значения найдены в 5 полях. Критичные: «Сумма» (7 строк) — скорее всего формульные ошибки. Некритичные: «Комментарий» (389) — опциональное поле.",
  сумм: "Суммарная выручка по колонке «Сумма» составляет приблизительно **310 млн руб.** (при среднем чеке ~64 200 руб.). Топ-регион — Москва: ~38% объёма.",
  менедж:
    "По менеджерам: в данных встречается ~24 уникальных имени. Рекомендую сводную таблицу по полю «Менеджер» → «Сумма» для ранжирования.",
};

function getTime() {
  return new Date().toLocaleTimeString("ru-RU", { hour: "2-digit", minute: "2-digit" });
}

function getAiReply(input: string): string {
  const lower = input.toLowerCase();
  for (const key of Object.keys(MOCK_AI_REPLIES)) {
    if (key !== "default" && lower.includes(key)) return MOCK_AI_REPLIES[key];
  }
  return MOCK_AI_REPLIES.default;
}

const TYPE_COLORS: Record<string, string> = {
  Целое: "tag-blue",
  Число: "tag-emerald",
  Дата: "tag-amber",
  Текст: "bg-purple-500/10 text-purple-300 border border-purple-500/20",
};

export default function Index() {
  const [fileStatus, setFileStatus] = useState<FileStatus>("idle");
  const [fileName, setFileName] = useState<string | null>(null);
  const [sheet, setSheet] = useState<SheetInfo | null>(null);
  const [dragOver, setDragOver] = useState(false);
  const [activeTab, setActiveTab] = useState<"structure" | "preview" | "stats">("structure");
  const [messages, setMessages] = useState<ChatMessage[]>([
    {
      role: "ai",
      text: "Привет! Загрузи Excel-файл слева, и я сразу проанализирую его структуру. Потом можешь задать любой вопрос по данным.",
      ts: getTime(),
    },
  ]);
  const [input, setInput] = useState("");
  const [aiStatus, setAiStatus] = useState<AiStatus>("idle");
  const fileInputRef = useRef<HTMLInputElement>(null);
  const chatEndRef = useRef<HTMLDivElement>(null);

  const handleFile = useCallback((file: File) => {
    if (!file.name.match(/\.(xlsx|xls|csv)$/i)) {
      setFileStatus("error");
      return;
    }
    setFileName(file.name);
    setFileStatus("loading");
    setTimeout(() => {
      setFileStatus("ready");
      setSheet(MOCK_SHEET);
      setMessages((prev) => [
        ...prev,
        {
          role: "ai",
          text: `Файл **«${file.name}»** успешно загружен. Обнаружен лист «${MOCK_SHEET.name}» — ${MOCK_SHEET.rows.toLocaleString("ru")} строк, ${MOCK_SHEET.cols} колонок. Структура отображена слева. Можешь спрашивать!`,
          ts: getTime(),
        },
      ]);
    }, 1800);
  }, []);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragOver(false);
      const file = e.dataTransfer.files[0];
      if (file) handleFile(file);
    },
    [handleFile]
  );

  const handleSend = () => {
    const text = input.trim();
    if (!text || aiStatus === "thinking") return;
    setMessages((prev) => [...prev, { role: "user", text, ts: getTime() }]);
    setInput("");
    setAiStatus("thinking");
    setTimeout(() => {
      setMessages((prev) => [
        ...prev,
        { role: "ai", text: getAiReply(text), ts: getTime() },
      ]);
      setAiStatus("done");
      setTimeout(() => setAiStatus("idle"), 500);
      setTimeout(() => chatEndRef.current?.scrollIntoView({ behavior: "smooth" }), 100);
    }, 1400);
  };

  const renderMarkdown = (text: string) =>
    text
      .replace(/\*\*(.*?)\*\*/g, '<strong class="text-emerald-400">$1</strong>')
      .replace(/«(.*?)»/g, '<span class="text-amber-300">«$1»</span>');

  return (
    <div className="min-h-screen flex flex-col">
      {/* Header */}
      <header
        className="border-b border-border/60 px-6 py-4 flex items-center justify-between backdrop-blur-sm sticky top-0 z-50"
        style={{ background: "hsla(220,16%,6%,0.92)" }}
      >
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 rounded-lg btn-primary flex items-center justify-center flex-shrink-0">
            <Icon name="Sparkles" size={16} />
          </div>
          <div>
            <span className="font-semibold text-foreground tracking-tight">DataMind</span>
            <span className="text-muted-foreground text-sm ml-2 hidden sm:inline">Excel AI Analyzer</span>
          </div>
        </div>
        <div className="flex items-center gap-2">
          {fileStatus === "ready" && (
            <div className="flex items-center gap-2 px-3 py-1.5 rounded-full tag-emerald text-xs font-medium animate-fade-in">
              <span className="w-1.5 h-1.5 rounded-full bg-emerald-400" />
              Файл загружен
            </div>
          )}
          <button
            onClick={() => fileInputRef.current?.click()}
            className="btn-amber px-4 py-2 rounded-lg text-sm font-semibold flex items-center gap-2"
          >
            <Icon name="Upload" size={15} />
            Загрузить файл
          </button>
        </div>
      </header>

      {/* Main layout */}
      <div className="flex-1 flex overflow-hidden" style={{ height: "calc(100vh - 65px)" }}>

        {/* Left panel */}
        <div className="w-full md:w-[54%] lg:w-[58%] flex flex-col border-r border-border/60 overflow-hidden">

          {/* Drop zone */}
          <div className="p-4 border-b border-border/40">
            <div
              className={`relative border-2 border-dashed rounded-xl p-6 text-center cursor-pointer transition-all duration-300 ${
                dragOver ? "drop-zone-active" : "border-border/40 hover:border-border/70"
              } ${fileStatus === "ready" ? "py-4" : ""}`}
              style={{ background: "rgba(52,211,153,0.015)" }}
              onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
              onDragLeave={() => setDragOver(false)}
              onDrop={handleDrop}
              onClick={() => fileInputRef.current?.click()}
            >
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                className="hidden"
                onChange={(e) => e.target.files?.[0] && handleFile(e.target.files[0])}
              />

              {fileStatus === "idle" && (
                <div className="animate-fade-in">
                  <div className="w-12 h-12 rounded-2xl bg-secondary mx-auto mb-3 flex items-center justify-center">
                    <Icon name="FileSpreadsheet" size={24} className="text-muted-foreground" />
                  </div>
                  <p className="text-sm font-medium text-foreground mb-1">Перетащи файл сюда или нажми</p>
                  <p className="text-xs text-muted-foreground">.xlsx · .xls · .csv</p>
                </div>
              )}

              {fileStatus === "loading" && (
                <div className="flex items-center justify-center gap-3 animate-fade-in">
                  <Icon name="Loader2" size={20} className="text-primary spinner" />
                  <span className="text-sm text-muted-foreground">Анализирую структуру...</span>
                </div>
              )}

              {fileStatus === "error" && (
                <div className="flex items-center justify-center gap-2 animate-fade-in">
                  <Icon name="AlertCircle" size={18} className="text-destructive" />
                  <span className="text-sm text-destructive">Неподдерживаемый формат файла</span>
                </div>
              )}

              {fileStatus === "ready" && fileName && (
                <div className="flex items-center justify-between animate-fade-in">
                  <div className="flex items-center gap-3">
                    <div className="w-9 h-9 rounded-lg tag-emerald flex items-center justify-center flex-shrink-0">
                      <Icon name="FileCheck" size={18} />
                    </div>
                    <div className="text-left">
                      <p className="text-sm font-semibold text-foreground">{fileName}</p>
                      <p className="text-xs text-muted-foreground">
                        {sheet?.rows.toLocaleString("ru")} строк · {sheet?.cols} колонок
                      </p>
                    </div>
                  </div>
                  <button
                    className="text-xs text-muted-foreground hover:text-foreground transition-colors px-2 py-1"
                    onClick={(e) => {
                      e.stopPropagation();
                      setFileStatus("idle");
                      setSheet(null);
                      setFileName(null);
                    }}
                  >
                    Заменить
                  </button>
                </div>
              )}
            </div>
          </div>

          {/* Sheet tabs */}
          {fileStatus === "ready" && sheet && (
            <div className="flex-1 flex flex-col overflow-hidden animate-fade-in">
              <div className="flex border-b border-border/40 px-4 gap-1">
                {(["structure", "preview", "stats"] as const).map((tab) => (
                  <button
                    key={tab}
                    onClick={() => setActiveTab(tab)}
                    className={`px-4 py-3 text-xs font-semibold tracking-wide transition-all border-b-2 -mb-px ${
                      activeTab === tab
                        ? "border-primary text-primary"
                        : "border-transparent text-muted-foreground hover:text-foreground"
                    }`}
                  >
                    {tab === "structure" && "Структура"}
                    {tab === "preview" && "Превью данных"}
                    {tab === "stats" && "Статистика"}
                  </button>
                ))}
                <div className="ml-auto flex items-center">
                  <span className="text-xs text-muted-foreground font-mono px-2">{sheet.name}</span>
                </div>
              </div>

              <div className="flex-1 overflow-y-auto scrollbar-thin p-4">
                {activeTab === "structure" && (
                  <div className="space-y-2">
                    {sheet.columns.map((col, i) => (
                      <div
                        key={i}
                        className="flex items-center gap-3 p-3 rounded-lg border border-border/40 hover:border-border/70 transition-colors"
                        style={{ background: "rgba(255,255,255,0.02)" }}
                      >
                        <span className="text-xs font-mono text-muted-foreground w-6 text-right flex-shrink-0">{i + 1}</span>
                        <span className="text-sm font-medium text-foreground flex-1 min-w-0 truncate">{col.name}</span>
                        <span className={`text-xs px-2 py-0.5 rounded-md font-medium flex-shrink-0 ${TYPE_COLORS[col.type] || "tag-blue"}`}>
                          {col.type}
                        </span>
                        <span className="text-xs font-mono text-muted-foreground flex-shrink-0 hidden sm:block w-24 truncate">
                          {col.sample}
                        </span>
                        {col.nulls > 0 && (
                          <span className="text-xs tag-amber px-2 py-0.5 rounded-md flex-shrink-0">
                            {col.nulls} пуст.
                          </span>
                        )}
                      </div>
                    ))}
                  </div>
                )}

                {activeTab === "preview" && (
                  <div className="overflow-x-auto">
                    <table className="w-full text-xs font-mono">
                      <thead>
                        <tr>
                          {sheet.columns.slice(0, 6).map((col) => (
                            <th key={col.name} className="text-left text-muted-foreground pb-2 pr-4 font-medium whitespace-nowrap">
                              {col.name}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {[0, 1, 2, 3, 4].map((row) => (
                          <tr key={row} className="border-t border-border/30">
                            {sheet.columns.slice(0, 6).map((col) => (
                              <td key={col.name} className="py-2 pr-4 text-muted-foreground whitespace-nowrap">
                                {col.sample.split(",")[0]?.trim() || "—"}
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    <p className="text-xs text-muted-foreground mt-4 text-center">
                      Показаны первые 5 строк · 6 из {sheet.cols} колонок
                    </p>
                  </div>
                )}

                {activeTab === "stats" && (
                  <div className="grid grid-cols-2 gap-3">
                    {[
                      { label: "Всего строк", value: sheet.rows.toLocaleString("ru"), icon: "Rows3", color: "text-primary" },
                      { label: "Колонок", value: String(sheet.cols), icon: "Columns3", color: "text-primary" },
                      { label: "Листов", value: "1", icon: "Layers", color: "text-amber-400" },
                      { label: "Заполнены", value: "96.2%", icon: "CheckCircle2", color: "text-primary" },
                      { label: "Пустых ячеек", value: "433", icon: "AlertCircle", color: "text-amber-400" },
                      { label: "Тип данных", value: "4 типа", icon: "Tag", color: "text-purple-400" },
                    ].map((stat, i) => (
                      <div key={i} className="p-4 rounded-xl border border-border/40 card-glow" style={{ background: "rgba(255,255,255,0.02)" }}>
                        <div className="flex items-center gap-2 mb-2">
                          <Icon name={stat.icon} size={15} className={stat.color} />
                          <span className="text-xs text-muted-foreground">{stat.label}</span>
                        </div>
                        <p className="text-2xl font-semibold text-foreground">{stat.value}</p>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          )}

          {/* Empty state */}
          {fileStatus === "idle" && (
            <div className="flex-1 flex flex-col items-center justify-center gap-4 p-8 text-center">
              <div className="grid grid-cols-3 gap-3 max-w-xs">
                {[
                  { icon: "FileSearch", label: "Анализ структуры", color: "text-primary" },
                  { icon: "Brain", label: "ИИ-вопросы", color: "text-amber-400" },
                  { icon: "BarChart3", label: "Статистика", color: "text-purple-400" },
                ].map((f, i) => (
                  <div
                    key={i}
                    className="p-3 rounded-xl border border-border/30 text-center"
                    style={{ background: "rgba(255,255,255,0.02)" }}
                  >
                    <Icon name={f.icon} size={20} className={`${f.color} mx-auto mb-2`} />
                    <p className="text-xs text-muted-foreground">{f.label}</p>
                  </div>
                ))}
              </div>
              <p className="text-sm text-muted-foreground max-w-xs">
                Загрузи Excel или CSV-файл, чтобы начать анализ
              </p>
            </div>
          )}
        </div>

        {/* Right panel — AI Chat */}
        <div className="hidden md:flex flex-col flex-1 overflow-hidden">
          <div className="px-5 py-4 border-b border-border/40 flex items-center gap-3">
            <div className="w-8 h-8 rounded-full btn-primary flex items-center justify-center flex-shrink-0">
              <Icon name="Bot" size={16} />
            </div>
            <div className="flex-1 min-w-0">
              <p className="text-sm font-semibold text-foreground">ИИ-аналитик</p>
              <p className="text-xs text-muted-foreground">Задай любой вопрос по данным</p>
            </div>
            <div
              className={`w-2 h-2 rounded-full flex-shrink-0 transition-colors ${
                aiStatus === "thinking" ? "bg-amber-400" : "bg-primary"
              }`}
            />
          </div>

          <div className="flex-1 overflow-y-auto scrollbar-thin p-5 space-y-4">
            {messages.map((msg, i) => (
              <div key={i} className={`flex gap-3 animate-fade-in ${msg.role === "user" ? "flex-row-reverse" : ""}`}>
                {msg.role === "ai" && (
                  <div className="w-7 h-7 rounded-full btn-primary flex items-center justify-center flex-shrink-0 mt-0.5">
                    <Icon name="Sparkles" size={13} />
                  </div>
                )}
                <div
                  className={`max-w-[85%] rounded-2xl px-4 py-3 text-sm leading-relaxed ${
                    msg.role === "ai"
                      ? "rounded-tl-sm border border-border/40"
                      : "btn-amber rounded-tr-sm"
                  }`}
                  style={msg.role === "ai" ? { background: "rgba(255,255,255,0.03)" } : {}}
                >
                  {msg.role === "ai" ? (
                    <p dangerouslySetInnerHTML={{ __html: renderMarkdown(msg.text) }} />
                  ) : (
                    <p>{msg.text}</p>
                  )}
                  <p className={`text-xs mt-1.5 ${msg.role === "ai" ? "text-muted-foreground" : "opacity-60"}`}>
                    {msg.ts}
                  </p>
                </div>
              </div>
            ))}

            {aiStatus === "thinking" && (
              <div className="flex gap-3 animate-fade-in">
                <div className="w-7 h-7 rounded-full btn-primary flex items-center justify-center flex-shrink-0">
                  <Icon name="Sparkles" size={13} />
                </div>
                <div
                  className="px-4 py-3 rounded-2xl rounded-tl-sm border border-border/40"
                  style={{ background: "rgba(255,255,255,0.03)" }}
                >
                  <div className="flex gap-1.5 items-center h-4">
                    <span className="w-1.5 h-1.5 rounded-full bg-primary animate-bounce" style={{ animationDelay: "0ms" }} />
                    <span className="w-1.5 h-1.5 rounded-full bg-primary animate-bounce" style={{ animationDelay: "150ms" }} />
                    <span className="w-1.5 h-1.5 rounded-full bg-primary animate-bounce" style={{ animationDelay: "300ms" }} />
                  </div>
                </div>
              </div>
            )}
            <div ref={chatEndRef} />
          </div>

          {messages.length <= 2 && (
            <div className="px-5 pb-2">
              <p className="text-xs text-muted-foreground mb-2">Быстрые вопросы:</p>
              <div className="flex flex-wrap gap-2">
                {["Какие данные пропущены?", "Посчитай сумму выручки", "Топ менеджеры по продажам"].map((q) => (
                  <button
                    key={q}
                    onClick={() => setInput(q)}
                    className="text-xs px-3 py-1.5 rounded-full border border-border/60 text-muted-foreground hover:text-foreground hover:border-primary/50 transition-all"
                    style={{ background: "rgba(255,255,255,0.02)" }}
                  >
                    {q}
                  </button>
                ))}
              </div>
            </div>
          )}

          <div className="p-4 border-t border-border/40">
            <div className="flex gap-2 items-end">
              <div className="flex-1 relative">
                <textarea
                  value={input}
                  onChange={(e) => setInput(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter" && !e.shiftKey) {
                      e.preventDefault();
                      handleSend();
                    }
                  }}
                  placeholder={fileStatus === "ready" ? "Спроси что-нибудь о данных..." : "Сначала загрузи файл..."}
                  disabled={fileStatus !== "ready"}
                  rows={1}
                  className="w-full px-4 py-3 rounded-xl border border-border/60 text-sm text-foreground placeholder:text-muted-foreground resize-none focus:outline-none focus:border-primary/50 focus:ring-1 focus:ring-primary/20 transition-all disabled:opacity-40"
                  style={{ background: "rgba(255,255,255,0.03)", fontFamily: "'IBM Plex Sans', sans-serif" }}
                />
              </div>
              <button
                onClick={handleSend}
                disabled={!input.trim() || aiStatus === "thinking" || fileStatus !== "ready"}
                className="w-11 h-11 rounded-xl btn-primary flex items-center justify-center flex-shrink-0 disabled:opacity-30 disabled:cursor-not-allowed"
                style={{ transform: "none" }}
              >
                <Icon name="Send" size={16} />
              </button>
            </div>
            <p className="text-xs text-muted-foreground mt-2 text-center">
              Enter — отправить · Shift+Enter — новая строка
            </p>
          </div>
        </div>
      </div>
    </div>
  );
}