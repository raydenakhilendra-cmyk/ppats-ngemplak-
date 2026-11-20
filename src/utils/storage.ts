import {
  BerkasRecord,
  CustomTemplateEntry,
  OfficeSettings,
  TemplateDocKey,
} from '../types/ppats';
import { DEFAULT_OFFICE_SETTINGS, INITIAL_SEED_RECORDS } from './seedData';

const DB_NAME = 'sim_ppats_db';
const DB_VERSION = 1;
const STORE_RECORDS = 'records';
const STORE_SETTINGS = 'settings';
const STORE_TEMPLATES = 'templates';

const LS_FALLBACK_RECORDS = 'sim_ppats_records_v1';
const LS_FALLBACK_SETTINGS = 'sim_ppats_settings_v1';
const LS_FALLBACK_TEMPLATES = 'sim_ppats_templates_v1';
const LS_SEED_INITIALIZED = 'sim_ppats_seeded_v1';

function openDatabase(): Promise<IDBDatabase | null> {
  return new Promise((resolve) => {
    if (typeof window === 'undefined' || !window.indexedDB) {
      resolve(null);
      return;
    }
    try {
      const req = window.indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(STORE_RECORDS)) {
          db.createObjectStore(STORE_RECORDS, { keyPath: 'id' });
        }
        if (!db.objectStoreNames.contains(STORE_SETTINGS)) {
          db.createObjectStore(STORE_SETTINGS, { keyPath: 'key' });
        }
        if (!db.objectStoreNames.contains(STORE_TEMPLATES)) {
          db.createObjectStore(STORE_TEMPLATES, { keyPath: 'key' });
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
}

async function getAllFromStore<T>(storeName: string): Promise<T[] | null> {
  const db = await openDatabase();
  if (!db) return null;
  return new Promise((resolve) => {
    try {
      const tx = db.transaction(storeName, 'readonly');
      const store = tx.objectStore(storeName);
      const req = store.getAll();
      req.onsuccess = () => resolve((req.result as T[]) ?? []);
      req.onerror = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
}

async function putIntoStore<T>(storeName: string, value: T): Promise<boolean> {
  const db = await openDatabase();
  if (!db) return false;
  return new Promise((resolve) => {
    try {
      const tx = db.transaction(storeName, 'readwrite');
      const store = tx.objectStore(storeName);
      store.put(value);
      tx.oncomplete = () => resolve(true);
      tx.onerror = () => resolve(false);
    } catch {
      resolve(false);
    }
  });
}

async function deleteFromStore(storeName: string, key: string): Promise<boolean> {
  const db = await openDatabase();
  if (!db) return false;
  return new Promise((resolve) => {
    try {
      const tx = db.transaction(storeName, 'readwrite');
      const store = tx.objectStore(storeName);
      store.delete(key);
      tx.oncomplete = () => resolve(true);
      tx.onerror = () => resolve(false);
    } catch {
      resolve(false);
    }
  });
}

async function clearStore(storeName: string): Promise<boolean> {
  const db = await openDatabase();
  if (!db) return false;
  return new Promise((resolve) => {
    try {
      const tx = db.transaction(storeName, 'readwrite');
      const store = tx.objectStore(storeName);
      store.clear();
      tx.oncomplete = () => resolve(true);
      tx.onerror = () => resolve(false);
    } catch {
      resolve(false);
    }
  });
}

export async function loadRecords(): Promise<BerkasRecord[]> {
  const idbRows = await getAllFromStore<BerkasRecord>(STORE_RECORDS);
  let records: BerkasRecord[] | null = idbRows;

  if (records === null) {
    try {
      const raw = localStorage.getItem(LS_FALLBACK_RECORDS);
      records = raw ? (JSON.parse(raw) as BerkasRecord[]) : [];
    } catch {
      records = [];
    }
  }

  const alreadySeeded = localStorage.getItem(LS_SEED_INITIALIZED) === '1';
  if (records.length === 0 && !alreadySeeded) {
    for (const seed of INITIAL_SEED_RECORDS) {
      await saveRecord(seed);
    }
    localStorage.setItem(LS_SEED_INITIALIZED, '1');
    return [...INITIAL_SEED_RECORDS].sort((a, b) => a.no_urut - b.no_urut);
  }

  if (records.length > 0 && !alreadySeeded) {
    localStorage.setItem(LS_SEED_INITIALIZED, '1');
  }

  return [...records].sort((a, b) => a.no_urut - b.no_urut);
}

export async function saveRecord(record: BerkasRecord): Promise<void> {
  const ok = await putIntoStore(STORE_RECORDS, record);
  if (!ok) {
    const raw = localStorage.getItem(LS_FALLBACK_RECORDS);
    const list: BerkasRecord[] = raw ? JSON.parse(raw) : [];
    const idx = list.findIndex((r) => r.id === record.id);
    if (idx >= 0) {
      list[idx] = record;
    } else {
      list.push(record);
    }
    localStorage.setItem(LS_FALLBACK_RECORDS, JSON.stringify(list));
  }
  localStorage.setItem(LS_SEED_INITIALIZED, '1');
}

export async function deleteRecord(id: string): Promise<void> {
  const ok = await deleteFromStore(STORE_RECORDS, id);
  if (!ok) {
    const raw = localStorage.getItem(LS_FALLBACK_RECORDS);
    const list: BerkasRecord[] = raw ? JSON.parse(raw) : [];
    localStorage.setItem(
      LS_FALLBACK_RECORDS,
      JSON.stringify(list.filter((r) => r.id !== id))
    );
  }
}

export async function loadSettings(): Promise<OfficeSettings> {
  const rows = await getAllFromStore<{ key: string; value: OfficeSettings }>(STORE_SETTINGS);
  if (rows) {
    const found = rows.find((r) => r.key === 'office_settings');
    if (found?.value) {
      return { ...DEFAULT_OFFICE_SETTINGS, ...found.value };
    }
    return { ...DEFAULT_OFFICE_SETTINGS };
  }

  try {
    const raw = localStorage.getItem(LS_FALLBACK_SETTINGS);
    if (raw) {
      return { ...DEFAULT_OFFICE_SETTINGS, ...JSON.parse(raw) };
    }
  } catch {
    // ignore
  }
  return { ...DEFAULT_OFFICE_SETTINGS };
}

export async function saveSettings(settings: OfficeSettings): Promise<void> {
  const ok = await putIntoStore(STORE_SETTINGS, { key: 'office_settings', value: settings });
  if (!ok) {
    localStorage.setItem(LS_FALLBACK_SETTINGS, JSON.stringify(settings));
  }
}

export async function loadCustomTemplates(): Promise<Record<TemplateDocKey, CustomTemplateEntry>> {
  const rows = await getAllFromStore<CustomTemplateEntry>(STORE_TEMPLATES);
  const result = {} as Record<TemplateDocKey, CustomTemplateEntry>;
  if (rows) {
    for (const entry of rows) {
      result[entry.key] = entry;
    }
    return result;
  }

  try {
    const raw = localStorage.getItem(LS_FALLBACK_TEMPLATES);
    if (raw) {
      return JSON.parse(raw);
    }
  } catch {
    // ignore
  }
  return result;
}

export async function saveCustomTemplate(entry: CustomTemplateEntry): Promise<void> {
  const ok = await putIntoStore(STORE_TEMPLATES, entry);
  if (!ok) {
    const existing = await loadCustomTemplates();
    existing[entry.key] = entry;
    localStorage.setItem(LS_FALLBACK_TEMPLATES, JSON.stringify(existing));
  }
}

export async function deleteCustomTemplate(key: TemplateDocKey): Promise<void> {
  const ok = await deleteFromStore(STORE_TEMPLATES, key);
  if (!ok) {
    const existing = await loadCustomTemplates();
    delete existing[key];
    localStorage.setItem(LS_FALLBACK_TEMPLATES, JSON.stringify(existing));
  }
}

export interface DatabaseBackupPayload {
  app: 'SIM-PPATS';
  version: 1;
  exportedAt: string;
  records: BerkasRecord[];
  settings: OfficeSettings;
  templates: CustomTemplateEntry[];
}

export async function exportDatabaseBackup(): Promise<DatabaseBackupPayload> {
  const records = await loadRecords();
  const settings = await loadSettings();
  const templatesMap = await loadCustomTemplates();
  return {
    app: 'SIM-PPATS',
    version: 1,
    exportedAt: new Date().toISOString(),
    records,
    settings,
    templates: Object.values(templatesMap),
  };
}

export async function importDatabaseBackup(jsonText: string): Promise<{
  recordCount: number;
  templateCount: number;
}> {
  const parsed = JSON.parse(jsonText) as Partial<DatabaseBackupPayload>;
  if (parsed.app !== 'SIM-PPATS' || !Array.isArray(parsed.records)) {
    throw new Error('Format file backup tidak valid atau bukan file backup SIM-PPATS.');
  }

  await clearStore(STORE_RECORDS);
  await clearStore(STORE_TEMPLATES);
  localStorage.removeItem(LS_FALLBACK_RECORDS);
  localStorage.removeItem(LS_FALLBACK_TEMPLATES);

  for (const rec of parsed.records) {
    await saveRecord(rec);
  }
  if (parsed.settings) {
    await saveSettings({ ...DEFAULT_OFFICE_SETTINGS, ...parsed.settings });
  }
  const templates = Array.isArray(parsed.templates) ? parsed.templates : [];
  for (const tpl of templates) {
    await saveCustomTemplate(tpl);
  }
  localStorage.setItem(LS_SEED_INITIALIZED, '1');

  return {
    recordCount: parsed.records.length,
    templateCount: templates.length,
  };
}
