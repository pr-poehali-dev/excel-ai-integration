// IndexedDB helper для сохранения сессий проекта
// Хранит: Excel-файлы (ArrayBuffer), PDF-страницы (URL), DOCX (текст), базу знаний

const DB_NAME = "datamind_session";
const DB_VERSION = 1;

export interface SavedSession {
  id: string;           // "current" для последней сессии
  savedAt: number;      // timestamp
  name: string;         // отображаемое имя
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
  buffer: ArrayBuffer;  // оригинальный файл
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
  buffer?: ArrayBuffer; // оригинальный файл (для DOCX)
}

export interface KnowledgeEntry {
  id: string;
  title: string;         // например "Объекты разработки АС10-11"
  content: string;       // текст правила/словаря
  enabled: boolean;
  updatedAt: number;
}

export interface SavedMessage {
  role: "user" | "ai";
  text: string;
  ts: string;
}

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

// ── База знаний ──────────────────────────────────────────────────────────────

export async function saveKnowledge(entries: KnowledgeEntry[]): Promise<void> {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction("knowledge", "readwrite");
    const store = tx.objectStore("knowledge");
    // Очищаем и записываем заново
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

// Форматирует базу знаний для передачи в системный промпт ИИ
export function formatKnowledgeForAI(entries: KnowledgeEntry[]): string {
  const active = entries.filter(e => e.enabled);
  if (!active.length) return "";
  return `═══ БАЗА ЗНАНИЙ ПРОЕКТА ═══\n` +
    active.map(e => `### ${e.title}\n${e.content}`).join("\n\n") +
    `\n═══════════════════════════\n`;
}
