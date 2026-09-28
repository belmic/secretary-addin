// Контуры (workspaces): изолированные «места работы» — у каждого своя папка в Drive,
// свой vault, свой глоссарий на ядре. Контур выбирается в popup и фиксируется при старте записи.
//   Мой диск/Secretary/ws-<id>/inbox/recordings/{rec_id}/
export const DEFAULT_WORKSPACES = [
  { id: 'vesco', name: 'Vesco' },
  { id: 'genesis', name: 'Genesis' },
  { id: 'ft-st', name: 'FT-ST' },
  { id: 'personal', name: 'Personal' },
];
export const UNSORTED = 'unsorted'; // запись без контура (старые записи, dev-режим)

export const workspaceFolder = (id) => `ws-${id || UNSORTED}`;

export function recordingsPath(workspaceId) {
  return ['Secretary', workspaceFolder(workspaceId), 'inbox', 'recordings'];
}

// «FT-ST Holding» → «ft-st-holding»: латиница, цифры, дефис; ≤ 24 символа.
export function workspaceId(name) {
  return String(name ?? '')
    .toLowerCase()
    .normalize('NFKD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 24)
    .replace(/-+$/, '');
}

export function addWorkspace(list, name) {
  const id = workspaceId(name);
  if (!id) throw new Error('Название контура: латиница, цифры, дефис');
  if (id === UNSORTED) throw new Error('Имя «unsorted» зарезервировано');
  if (list.some((w) => w.id === id)) throw new Error(`Контур «${id}» уже есть`);
  return [...list, { id, name: String(name).trim() }];
}

export function workspaceName(list, id) {
  return list.find((w) => w.id === id)?.name ?? (id === UNSORTED || !id ? 'без контура' : id);
}
