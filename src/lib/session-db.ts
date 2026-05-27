// IndexedDB helper для сохранения сессий проекта
// Хранит: Excel-файлы (ArrayBuffer), PDF-страницы (URL), DOCX (текст), базу знаний

const DB_NAME = "datamind_session";
const DB_VERSION = 2;

export interface SavedSession {
  id: string;
  savedAt: number;
  name: string;
  excelFiles: SavedExcelFile[];
  docFiles: SavedDocFile[];
  knowledgeBase: KnowledgeEntry[];
  messages: SavedMessage[];
}

export interface SavedExcelFile {
  id: string;
  name: string;
  role: "main" | "reference" | null;
  activeSheet: number;
  buffer: ArrayBuffer;
}

export interface SavedDocFile {
  id: string;
  name: string;
  type: "pdf" | "docx";
  role: "report" | "protocol" | "database" | null;
  text: string;
  html?: string;
  pageCount?: number;
  pageImageUrls?: string[];
  buffer?: ArrayBuffer;
}

// Категории базы знаний
export type KnowledgeCategory =
  | "objects"      // Объекты разработки, пласты, залежи
  | "reserves"     // Запасы нефти и газа
  | "rules"        // Правила интерпретации, методики
  | "terms"        // Термины и сленг
  | "tables"       // Таблицы и справочники (Excel)
  | "docs"         // Нормативные документы (PDF/DOCX)
  | "custom";      // Пользовательские

export type KnowledgeSourceType = "text" | "excel" | "pdf" | "docx";

export interface KnowledgeEntry {
  id: string;
  title: string;
  content: string;            // основной текст (всегда есть, для файлов — извлечённый)
  category: KnowledgeCategory;
  sourceType: KnowledgeSourceType;
  enabled: boolean;
  updatedAt: number;
  // Для файловых источников
  fileName?: string;
  pageImageUrls?: string[];   // PDF → страницы как картинки
  pageFrom?: number;
  pageTo?: number;
}

export interface SavedMessage {
  role: "user" | "ai";
  text: string;
  ts: string;
}

// Категории: метаданные для UI
export const KNOWLEDGE_CATEGORIES: Record<KnowledgeCategory, { label: string; icon: string; color: string }> = {
  objects:  { label: "Объекты и пласты",     icon: "Layers",       color: "text-emerald-400" },
  reserves: { label: "Запасы",               icon: "BarChart3",    color: "text-blue-400" },
  rules:    { label: "Правила и методики",   icon: "BookMarked",   color: "text-amber-400" },
  terms:    { label: "Термины и сленг",      icon: "FileText",     color: "text-purple-400" },
  tables:   { label: "Таблицы (Excel)",      icon: "Table2",       color: "text-green-400" },
  docs:     { label: "Документы",            icon: "FileStack",    color: "text-sky-400" },
  custom:   { label: "Прочее",               icon: "MoreHorizontal", color: "text-muted-foreground" },
};

// ── Шаблоны для нефтяной отрасли ─────────────────────────────────────────────
export const OIL_KNOWLEDGE_TEMPLATES: Omit<KnowledgeEntry, "id" | "updatedAt">[] = [
  {
    title: "Объекты разработки — состав пластов",
    category: "objects",
    sourceType: "text",
    enabled: true,
    content: `Объект разработки — это сумма всех пластов, которые входят в него.
При упоминании объекта всегда подразумеваются все его пласты суммарно.

Пример (ШАБЛОН — замени на данные своего месторождения):
Объект АС10-11 включает пласты:
- K1br пл.АС10/1 (залежь 1)
- K1g пл.АС11/1 (залежь 1, р-он скв.5R-1ПО)
- K1g пл.АС11/1 (залежь 4)

Объект БС8 включает пласты:
- K2g пл.БС8/1 (залежь 1)
- K2g пл.БС8/2 (залежь 1)

Источник для уточнения: Таблица «База данных» или раздел VII ПТД «Принципиальные положения».`,
  },
  {
    title: "Запасы нефти и газа — источник данных",
    category: "reserves",
    sourceType: "text",
    enabled: true,
    content: `Когда речь идёт о запасах нефти и газа:
- Использовать ТОЛЬКО данные на 01.01. текущего года из Таблицы 2 Протокола ЦКР (который мы готовим в работе)
- НЕ использовать балансовые запасы из других источников
- При расхождении с балансом — сообщать, выделяя красным

Текущий год: ${new Date().getFullYear()}
Эталон запасов: Таблица 2 Протокола ЦКР на 01.01.${new Date().getFullYear()}

Запасы по категориям: АВ1+В2 — основная категория для расчётов КИН.
Категории запасов: АВ1, В2, С1, С2.`,
  },
  {
    title: "Правила работы с отчётом (ПТД)",
    category: "rules",
    sourceType: "text",
    enabled: true,
    content: `Правила актуализации текста отчёта:
1. Сохранять ВСЕ стили оригинального Word-документа
2. Изменять ТОЛЬКО числовые значения, не трогая структуру и стиль текста
3. Все изменения закрашивать зелёным цветом в Word
4. При любых расхождениях таблиц с текстом — выделять красным и сообщать
5. Эталон данных: таблицы протокола ЦКР → таблицы отчёта → текст отчёта
6. Сначала утверждаем данные в таблицах ЦКР, потом правим текст

Главы для актуализации: 3, 5, 6.
По запросу: проверка всех глав с цифрами на соответствие таблицам.`,
  },
  {
    title: "Формулы Excel 2010",
    category: "rules",
    sourceType: "text",
    enabled: false,
    content: `Все формулы писать для Excel 2010 на русском языке.

Примеры часто используемых формул:
=СУММ(B5:B10)
=ВПР(A2;Таблица!$A:$D;3;0)
=ПРОМЕЖУТОЧНЫЕ.ИТОГИ(9;B5:B100)
=ЕСЛИ(A2>0;B2/A2;"")
=СУММЕСЛИ(A:A;"АС10-11";B:B)
=СЧЁТЕСЛИ(C:C;"добывающая")
=СРЗНАЧ(D2:D100)
=ОКРУГЛ(B2/1000;3)

Синтаксис: русские названия функций, разделитель аргументов — точка с запятой (;)`,
  },
  {
    title: "Термины и сокращения",
    category: "terms",
    sourceType: "text",
    enabled: true,
    content: `Сокращения и их расшифровки:
ПТД — Проектный технологический документ
ЦКР — Центральная комиссия по разработке
КИН — Коэффициент извлечения нефти
Квыт — Коэффициент вытеснения
Кохв — Коэффициент охвата
ОРД — Одновременно-раздельная добыча
ОРЗ — Одновременно-раздельная закачка
ГС — Горизонтальная скважина
ГТМ — Геолого-технические мероприятия
ПРС — Подземный ремонт скважин
КРС — Капитальный ремонт скважин
ФОНДдоб — Фонд добывающих скважин
ФОНДнаг — Фонд нагнетательных скважин
ВНФ — Водонефтяной фактор
ГФ — Газовый фактор
Дебит — объём продукции скважины в единицу времени (т/сут, м3/сут)
Приёмистость — объём закачки нагнетательной скважины (м3/сут)`,
  },
  {
    title: "Система разработки — правила описания",
    category: "objects",
    sourceType: "text",
    enabled: false,
    content: `Правила описания системы разработки:

Типы систем размещения скважин:
- Равномерная сетка — скважины расположены по регулярной сетке
- Неравномерная (очаговая) — очаговое заводнение
- Однорядная — добывающие в один ряд
- Многорядная — несколько рядов добывающих между нагнетательными

Расстояние между скважинами указывать в метрах.
При описании объекта всегда указывать:
1. Тип системы размещения
2. Расстояние между скважинами / рядами
3. Тип заводнения (площадное, рядное, очаговое, естественный режим)
4. Фонд скважин (добывающие, нагнетательные, контрольные, водозаборные)`,
  },
  {
    title: "Накопленная добыча — источник и единицы",
    category: "reserves",
    sourceType: "text",
    enabled: false,
    content: `Единицы измерения:
- Нефть: тыс. т (тысячи тонн)
- Газ растворённый: млн м³
- Газ свободный: млн м³
- Жидкость: тыс. м³
- Закачка воды: тыс. м³

Накопленные показатели брать из базы данных на 01.01. текущего года.
Текущие годовые показатели — из базы данных за последний завершённый год.
При сравнении с проектом указывать % выполнения: факт/проект × 100%.`,
  },
];

// ── IndexedDB ─────────────────────────────────────────────────────────────────

function openDB(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = (e) => {
      const db = (e.target as IDBOpenDBRequest).result;
      if (!db.objectStoreNames.contains("sessions")) {
        db.createObjectStore("sessions", { keyPath: "id" });
      }
      if (!db.objectStoreNames.contains("knowledge")) {
        db.createObjectStore("knowledge", { keyPath: "id" });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

export async function saveSession(session: SavedSession): Promise<void> {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction("sessions", "readwrite");
    tx.objectStore("sessions").put(session);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

export async function loadSession(id = "current"): Promise<SavedSession | null> {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction("sessions", "readonly");
    const req = tx.objectStore("sessions").get(id);
    req.onsuccess = () => resolve(req.result ?? null);
    req.onerror = () => reject(req.error);
  });
}

export async function listSessions(): Promise<Pick<SavedSession, "id" | "name" | "savedAt">[]> {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction("sessions", "readonly");
    const req = tx.objectStore("sessions").getAll();
    req.onsuccess = () => resolve(
      (req.result as SavedSession[])
        .map(s => ({ id: s.id, name: s.name, savedAt: s.savedAt }))
        .sort((a, b) => b.savedAt - a.savedAt)
    );
    req.onerror = () => reject(req.error);
  });
}

export async function deleteSession(id: string): Promise<void> {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction("sessions", "readwrite");
    tx.objectStore("sessions").delete(id);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

export async function saveKnowledge(entries: KnowledgeEntry[]): Promise<void> {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction("knowledge", "readwrite");
    const store = tx.objectStore("knowledge");
    store.clear();
    entries.forEach(e => store.put(e));
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

export async function loadKnowledge(): Promise<KnowledgeEntry[]> {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction("knowledge", "readonly");
    const req = tx.objectStore("knowledge").getAll();
    req.onsuccess = () => resolve(req.result ?? []);
    req.onerror = () => reject(req.error);
  });
}

// Форматирует активные записи базы знаний для системного промпта ИИ
export function formatKnowledgeForAI(entries: KnowledgeEntry[]): string {
  const active = entries.filter(e => e.enabled);
  if (!active.length) return "";

  const byCategory: Partial<Record<KnowledgeCategory, KnowledgeEntry[]>> = {};
  for (const e of active) {
    if (!byCategory[e.category]) byCategory[e.category] = [];
    byCategory[e.category]!.push(e);
  }

  const sections: string[] = [];
  for (const [cat, items] of Object.entries(byCategory)) {
    const catLabel = KNOWLEDGE_CATEGORIES[cat as KnowledgeCategory]?.label ?? cat;
    sections.push(`## ${catLabel}\n` + items!.map(e => {
      // Для PDF/DOCX с картинками — только заголовок (картинки идут отдельно как изображения)
      if ((e.sourceType === "pdf") && e.pageImageUrls?.length) {
        return `### ${e.title}\n[Документ передан как изображения страниц]`;
      }
      const MAX = 8000;
      const text = e.content.length > MAX ? e.content.slice(0, MAX) + "\n[...обрезано]" : e.content;
      return `### ${e.title}\n${text}`;
    }).join("\n\n"));
  }

  return `═══ БАЗА ЗНАНИЙ ПРОЕКТА ═══\n${sections.join("\n\n")}\n═══════════════════════════\n`;
}

// Возвращает URL картинок из PDF-записей базы знаний (для vision-запросов)
export function getKnowledgePdfImages(entries: KnowledgeEntry[]): { title: string; urls: string[] }[] {
  return entries
    .filter(e => e.enabled && e.sourceType === "pdf" && e.pageImageUrls?.length)
    .map(e => {
      const all = e.pageImageUrls!;
      const from = Math.max(1, e.pageFrom ?? 1) - 1;
      const to = Math.min(all.length, e.pageTo ?? all.length);
      return { title: e.title, urls: all.slice(from, to) };
    });
}
