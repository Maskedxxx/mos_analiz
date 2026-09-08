/**
 * API-клиент для взаимодействия с бэкендом аудита.
 * HTTP-запросы + SSE-подписка на прогресс.
 */

const BASE = '/api';

/**
 * Единый разбор ответа бэкенда (F26.1). Сеть/прокси без JSON → «Сервер недоступен»;
 * 401 → событие `audit:unauthorized` (App показывает экран входа); detail-массив
 * (422 от pydantic) → читаемая строка; иначе — detail или запасной текст.
 */
function formatDetail(d) {
  if (!d) return '';
  if (typeof d === 'string') return d;
  if (Array.isArray(d)) {
    return d.map((x) => (x && x.msg ? `${(x.loc || []).slice(-1)[0] || ''}: ${x.msg}`.replace(/^: /, '') : String(x))).join('; ');
  }
  if (typeof d === 'object' && d.detail) return formatDetail(d.detail);
  return JSON.stringify(d);
}

async function request(url, options, fallback) {
  let res;
  try {
    res = await fetch(url, { credentials: 'include', ...options });
  } catch {
    throw new Error('Сервер недоступен, попробуйте позже');
  }
  if (res.ok) return res;
  if (res.status === 401) {
    window.dispatchEvent(new Event('audit:unauthorized'));
    throw new Error('Сессия истекла — войдите снова');
  }
  let err = null;
  try { err = await res.json(); } catch { err = null; }
  if (!err) throw new Error(res.status >= 500 ? 'Сервер недоступен, попробуйте позже' : fallback);
  throw new Error(formatDetail(err.detail) || fallback);
}

/** POST /api/login — авторизация */
export async function login(username, password) {
  // Прямой fetch: 401 здесь — неверный логин/пароль, а не истёкшая сессия (без события unauthorized).
  const res = await fetch(`${BASE}/login`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ login: username, password }),
    credentials: 'include',
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(formatDetail(err.detail) || 'Ошибка авторизации');
  }
  return res.json();
}

/** GET /api/types — список типов документов */
export async function fetchDocTypes() {
  const res = await request(`${BASE}/types`, {}, 'Не удалось загрузить типы документов');
  return res.json();
}

/** GET /api/types/{docType}/rules — {editable, sections, rules[]} (база + кастом) */
export async function fetchRules(docType) {
  const res = await request(`${BASE}/types/${docType}/rules`, {}, 'Не удалось загрузить правила проверки');
  return res.json();
}

/** POST /api/types/{docType}/rules/draft — спец-модель формулирует правило из сырья */
export async function draftRule(docType, payload) {
  const res = await request(`${BASE}/types/${docType}/rules/draft`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  }, 'Не удалось сформулировать правило');
  return res.json();
}

/** POST /api/types/{docType}/rules — сохранить пользовательское правило */
export async function createRule(docType, payload) {
  const res = await request(`${BASE}/types/${docType}/rules`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  }, 'Не удалось сохранить правило');
  return res.json();
}

/** PUT /api/types/{docType}/rules/{index} — изменить пользовательское правило */
export async function updateRule(docType, index, payload) {
  const res = await request(`${BASE}/types/${docType}/rules/${index}`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  }, 'Не удалось изменить правило');
  return res.json();
}

/** DELETE /api/types/{docType}/rules/{index} — удалить пользовательское правило */
export async function deleteRule(docType, index) {
  const res = await request(`${BASE}/types/${docType}/rules/${index}`, {
    method: 'DELETE',
  }, 'Не удалось удалить правило');
  return res.json();
}

/** POST /api/audit — запуск аудита (file + doc_type) */
export async function startAudit(file, docType) {
  const form = new FormData();
  form.append('file', file);
  form.append('doc_type', docType);

  const res = await request(`${BASE}/audit`, {
    method: 'POST',
    body: form,
  }, 'Ошибка запуска аудита');
  return res.json();
}

/** POST /api/audit/cross — запуск сквозной сверки (3 файла под одним полем files) */
export async function startCrossAudit(files) {
  const form = new FormData();
  files.forEach((f) => form.append('files', f));

  const res = await request(`${BASE}/audit/cross`, {
    method: 'POST',
    body: form,
  }, 'Ошибка запуска сквозной проверки');
  return res.json();
}

/**
 * GET /api/audit/{id}/events — SSE-подписка на прогресс.
 * Возвращает EventSource. Вызывает onEvent(type, data) для каждого события.
 */
export function subscribeToProgress(sessionId, onEvent) {
  const source = new EventSource(`${BASE}/audit/${sessionId}/events`);

  // Флаг: получено терминальное событие (complete / audit_error).
  let finished = false;

  // Серверное событие ошибки называется `audit_error`, НЕ `error`: имя `error` совпало бы
  // с DOM-событием обрыва соединения EventSource, и этот же слушатель поймал бы обрыв,
  // упав на JSON.parse(undefined) («"undefined" is not valid JSON») и пометив поток
  // завершённым — тогда onerror счёл бы закрытие штатным и молчал про реальный обрыв
  // (аудит устойчивости, находки 4.4b / 6.1). `ping` — keepalive сервера, держит поток «живым».
  const eventTypes = [
    'audit_start', 'parsing_target', 'parsing_target_done',
    'checking_rules', 'checking_rules_done',
    'ping', 'complete', 'audit_error',
  ];

  eventTypes.forEach((type) => {
    source.addEventListener(type, (e) => {
      if (type === 'complete' || type === 'audit_error') {
        finished = true;
      }
      // Внутренний контракт onEvent не меняем: серверный `audit_error` доходит как 'error'.
      onEvent(type === 'audit_error' ? 'error' : type, JSON.parse(e.data));
    });
  });

  source.onerror = () => {
    // Штатное закрытие SSE после терминального события — не ложная ошибка.
    if (finished) {
      source.close();
      return;
    }
    onEvent('error', { message: 'Потеряно соединение с сервером' });
    source.close();
  };

  return source;
}

/**
 * GET /api/audit/{id}/result — разовый опрос результата.
 * Возвращает {status, data}: 200 — готов (result), 202 — ещё выполняется,
 * 404 — сессия не найдена, 500 — аудит завершился с ошибкой.
 * Используется сторожевым таймером экрана прогресса и при восстановлении сессии.
 */
export async function fetchAuditResult(sessionId) {
  const res = await fetch(`${BASE}/audit/${sessionId}/result`, { credentials: 'include' });
  const data = await res.json().catch(() => ({}));
  return { status: res.status, data };
}

/** URL для скачивания Excel */
export function getDownloadUrl(sessionId) {
  return `${BASE}/audit/${sessionId}/download`;
}

/**
 * Скачать Excel-отчёт с проверкой ответа (F26.3): при 404/500 бросает Error с текстом
 * (у <a download> ошибка была бы «тихой» — браузер молча ничего не сохраняет, 4.5b).
 * Возвращает {blob, filename}.
 */
export async function downloadReport(sessionId) {
  const res = await request(getDownloadUrl(sessionId), {}, 'Отчёт недоступен');
  const cd = res.headers.get('Content-Disposition') || '';
  const m = /filename\*?=(?:UTF-8''|")?([^";]+)/i.exec(cd);
  const filename = m ? decodeURIComponent(m[1].replace(/"/g, '')) : `audit_${sessionId}.xlsx`;
  return { blob: await res.blob(), filename };
}

/* ───────────── Генерация документов (бэкенд :8090 через прокси /gen) ───────────── */

// Префикс /gen срезается прокси: оба бэкенда отдают /api/types/*, без него была бы коллизия
const GEN_BASE = '/gen/api';

/** GET /gen/api/graph — узлы графа генерации: {doc_type, title, code, position} */
export async function fetchGenTypes() {
  const res = await request(`${GEN_BASE}/graph`, {}, 'Не удалось загрузить типы документов для генерации');
  const data = await res.json();
  return data.nodes || [];
}

/** GET /gen/api/types/{docType}/schema — {title, fields:[{key,label,required,source,hint}]} */
export async function fetchGenSchema(docType) {
  const res = await request(`${GEN_BASE}/types/${docType}/schema`, {}, 'Не удалось загрузить форму документа');
  return res.json();
}

/** GET /gen/api/orgs — список известных организаций (подсказки для поля ООО) */
export async function fetchOrgs() {
  const res = await request(`${GEN_BASE}/orgs`, {}, 'Не удалось загрузить список организаций');
  return res.json();
}

/** GET /gen/api/org/suggest?org=… — {field_key: [values]} ранее введённых значений по этому ООО */
export async function fetchOrgSuggest(org) {
  const res = await request(`${GEN_BASE}/org/suggest?org=${encodeURIComponent(org)}`, {}, 'Не удалось загрузить подсказки по организации');
  return res.json();
}

/** URL нативного скачивания .docx (грузится через скрытый iframe — браузер сохраняет сам) */
export function getGenDownloadUrl(docType, params) {
  const qs = new URLSearchParams(params).toString();
  return `${GEN_BASE}/types/${docType}/download?${qs}`;
}
