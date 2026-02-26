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

/**
 * GET /api/audit/{id}/events — SSE-подписка на прогресс.
 * Возвращает EventSource. Вызывает onEvent(type, data) для каждого события.
 */
export function subscribeToProgress(sessionId, onEvent) {
  const source = new EventSource(`${BASE}/audit/${sessionId}/events`);

  const eventTypes = [
    'audit_start', 'parsing_target', 'parsing_target_done',
    'parsing_template', 'parsing_template_done',
    'rule_done', 'complete', 'error',
  ];

  eventTypes.forEach((type) => {
    source.addEventListener(type, (e) => {
      onEvent(type, JSON.parse(e.data));
    });
  });

  source.onerror = () => {
    onEvent('error', { message: 'Потеряно соединение с сервером' });
    source.close();
  };

  return source;
}

/** URL для скачивания Excel */
export function getDownloadUrl(sessionId) {
  return `${BASE}/audit/${sessionId}/download`;
}
