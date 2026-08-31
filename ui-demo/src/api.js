/**
 * API-клиент для взаимодействия с бэкендом аудита.
 * HTTP-запросы + SSE-подписка на прогресс.
 */

const BASE = '/api';

/** POST /api/login — авторизация */
export async function login(username, password) {
  const res = await fetch(`${BASE}/login`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ login: username, password }),
    credentials: 'include',
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.detail || 'Ошибка авторизации');
  }
  return res.json();
}

/** GET /api/types — список типов документов */
export async function fetchDocTypes() {
  const res = await fetch(`${BASE}/types`, { credentials: 'include' });
  if (!res.ok) throw new Error('Не удалось загрузить типы документов');
  return res.json();
}

/** GET /api/types/{docType}/rules — {editable, sections, rules[]} (база + кастом) */
export async function fetchRules(docType) {
  const res = await fetch(`${BASE}/types/${docType}/rules`, { credentials: 'include' });
  if (!res.ok) throw new Error('Не удалось загрузить правила проверки');
  return res.json();
}

/** POST /api/types/{docType}/rules/draft — спец-модель формулирует правило из сырья */
export async function draftRule(docType, payload) {
  const res = await fetch(`${BASE}/types/${docType}/rules/draft`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
    credentials: 'include',
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.detail || 'Не удалось сформулировать правило');
  }
  return res.json();
}

/** POST /api/types/{docType}/rules — сохранить пользовательское правило */
export async function createRule(docType, payload) {
  const res = await fetch(`${BASE}/types/${docType}/rules`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
    credentials: 'include',
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.detail || 'Не удалось сохранить правило');
  }
  return res.json();
}

/** PUT /api/types/{docType}/rules/{index} — изменить пользовательское правило */
export async function updateRule(docType, index, payload) {
  const res = await fetch(`${BASE}/types/${docType}/rules/${index}`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
    credentials: 'include',
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.detail || 'Не удалось изменить правило');
  }
  return res.json();
}

/** DELETE /api/types/{docType}/rules/{index} — удалить пользовательское правило */
export async function deleteRule(docType, index) {
  const res = await fetch(`${BASE}/types/${docType}/rules/${index}`, {
    method: 'DELETE',
    credentials: 'include',
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.detail || 'Не удалось удалить правило');
  }
  return res.json();
}

/** POST /api/audit — запуск аудита (file + doc_type) */
export async function startAudit(file, docType) {
  const form = new FormData();
  form.append('file', file);
  form.append('doc_type', docType);

  const res = await fetch(`${BASE}/audit`, {
    method: 'POST',
    body: form,
    credentials: 'include',
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.detail || 'Ошибка запуска аудита');
  }
  return res.json();
}

/** POST /api/audit/cross — запуск сквозной сверки (3 файла под одним полем files) */
export async function startCrossAudit(files) {
  const form = new FormData();
  files.forEach((f) => form.append('files', f));

  const res = await fetch(`${BASE}/audit/cross`, {
    method: 'POST',
    body: form,
    credentials: 'include',
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.detail || 'Ошибка запуска сквозной проверки');
  }
  return res.json();
}

/**
 * GET /api/audit/{id}/events — SSE-подписка на прогресс.
 * Возвращает EventSource. Вызывает onEvent(type, data) для каждого события.
 */
export function subscribeToProgress(sessionId, onEvent) {
  const source = new EventSource(`${BASE}/audit/${sessionId}/events`);

  // flag: terminalnoe sobytie (complete/error) polucheno
  let finished = false;

  const eventTypes = [
    'audit_start', 'parsing_target', 'parsing_target_done',
    'checking_rules', 'checking_rules_done',
    'complete', 'error',
  ];

  eventTypes.forEach((type) => {
    source.addEventListener(type, (e) => {
      // terminalnoe sobytie - pomechaem potok kak zavershyonnyy
      if (type === 'complete' || type === 'error') {
        finished = true;
      }
      onEvent(type, JSON.parse(e.data));
    });
  });

  source.onerror = () => {
    // shtatnoe zakrytie SSE posle terminalnogo sobytiya - ne lozhnaya oshibka
    if (finished) {
      source.close();
      return;
    }
    onEvent('error', { message: 'Потеряно соединение с сервером' });
    source.close();
  };

  return source;
}

/** URL для скачивания Excel */
export function getDownloadUrl(sessionId) {
  return `${BASE}/audit/${sessionId}/download`;
}

/* ───────────── Генерация документов (бэкенд :8090 через прокси /gen) ───────────── */

// Префикс /gen срезается прокси: оба бэкенда отдают /api/types/*, без него была бы коллизия
const GEN_BASE = '/gen/api';

/** GET /gen/api/graph — узлы графа генерации: {doc_type, title, code, position} */
export async function fetchGenTypes() {
  const res = await fetch(`${GEN_BASE}/graph`, { credentials: 'include' });
  if (!res.ok) throw new Error('Не удалось загрузить типы документов для генерации');
  const data = await res.json();
  return data.nodes || [];
}

/** GET /gen/api/types/{docType}/schema — {title, fields:[{key,label,required,source,hint}]} */
export async function fetchGenSchema(docType) {
  const res = await fetch(`${GEN_BASE}/types/${docType}/schema`, { credentials: 'include' });
  if (!res.ok) throw new Error('Не удалось загрузить форму документа');
  return res.json();
}

/** GET /gen/api/orgs — список известных организаций (подсказки для поля ООО) */
export async function fetchOrgs() {
  const res = await fetch(`${GEN_BASE}/orgs`, { credentials: 'include' });
  if (!res.ok) throw new Error('Не удалось загрузить список организаций');
  return res.json();
}

/** GET /gen/api/org/suggest?org=… — {field_key: [values]} ранее введённых значений по этому ООО */
export async function fetchOrgSuggest(org) {
  const res = await fetch(`${GEN_BASE}/org/suggest?org=${encodeURIComponent(org)}`, { credentials: 'include' });
  if (!res.ok) throw new Error('Не удалось загрузить подсказки по организации');
  return res.json();
}

/** URL нативного скачивания .docx (грузится через скрытый iframe — браузер сохраняет сам) */
export function getGenDownloadUrl(docType, params) {
  const qs = new URLSearchParams(params).toString();
  return `${GEN_BASE}/types/${docType}/download?${qs}`;
}
