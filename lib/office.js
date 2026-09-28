// Office.js: колбэк → Promise. Статус сравнивается строкой 'succeeded'
// (= Office.AsyncResultStatus.Succeeded) — модуль тестируется без office.js.
export function call(fn) {
  return new Promise((resolve, reject) => {
    try {
      fn((r) =>
        r.status === 'succeeded' ? resolve(r.value) : reject(new Error(r.error?.message || 'Office: ошибка вызова')),
      );
    } catch (e) {
      reject(e);
    }
  });
}

export const SENT_CATEGORY = 'Secretary ✓';

// Категория «Secretary ✓» в мастер-списке ящика (один раз) и на отправленном письме.
export async function ensureMasterCategory(mailbox) {
  const list = await call((cb) => mailbox.masterCategories.getAsync(cb));
  if (!list.some((c) => c.displayName === SENT_CATEGORY))
    await call((cb) => mailbox.masterCategories.addAsync([{ displayName: SENT_CATEGORY, color: 'Preset4' }], cb));
}

export async function markSent(mailbox, itemId) {
  const m = await call((cb) => mailbox.loadItemByIdAsync(itemId, cb));
  try {
    await call((cb) => m.categories.addAsync([SENT_CATEGORY], cb));
  } finally {
    await call((cb) => m.unloadAsync(cb)).catch(() => {});
  }
}
