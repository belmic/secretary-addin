// Панель «В Secretary»: выделенные письма → пакеты → Google Drive контура (docs/spec.md).
import { createAuth, NeedsAuthError } from './lib/auth.js';
import { createDrive } from './lib/shared/drive.js';
import { buildPackage } from './lib/package.js';
import { uploadPackage, loadWorkspaces } from './lib/uploader.js';
import { call, ensureMasterCategory, markSent } from './lib/office.js';

const VERSION = '1.2.0';
const PAGES = new URL('./', location.href).href;
const LS_WS = 'secretary-workspace';
const LS_WS_LIST = 'secretary-workspaces';
const DEFAULT_WS = [
  { id: 'vesco', name: 'Vesco' },
  { id: 'genesis', name: 'Genesis' },
  { id: 'ft-st', name: 'FT-ST' },
  { id: 'personal', name: 'Личное' },
];

const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
const lsGet = (k, d) => {
  try {
    return JSON.parse(localStorage.getItem(k)) ?? d;
  } catch {
    return d;
  }
};
const lsSet = (k, v) => {
  try {
    localStorage.setItem(k, JSON.stringify(v));
  } catch {
    /* нет localStorage — выбор просто не запомнится */
  }
};

function openDialog(url) {
  return new Promise((resolve) => {
    Office.context.ui.displayDialogAsync(url, { height: 70, width: 40 }, (r) => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) return resolve({ error: r.error.message });
      const dlg = r.value;
      dlg.addEventHandler(Office.EventType.DialogMessageReceived, (e) => {
        dlg.close();
        try {
          resolve(JSON.parse(e.message));
        } catch {
          resolve({ error: 'ответ окна входа не распознан' });
        }
      });
      dlg.addEventHandler(Office.EventType.DialogEventReceived, (e) =>
        resolve({ error: e.error === 12006 ? 'окно входа закрыто' : `ошибка окна ${e.error}` }),
      );
    });
  });
}

const auth = createAuth({ storage: localStorage, openDialog, pagesBase: PAGES });
const drive = createDrive({ getAccessToken: (o) => auth.getAccessToken(o) });
const folderCache = {};
let selected = []; // [{itemId, subject, hasAttachment}]
const status = new Map(); // itemId → {text, cls}
let sending = false;

// ---------- вход ----------
function renderAuth(error) {
  const st = auth.state();
  const box = $('auth');
  if (st === 'connected') {
    box.innerHTML = 'Google Drive: подключён ✓';
  } else if (st === 'unconfigured') {
    box.innerHTML = 'Заполните Client ID и Client Secret в «Настройках» ниже.';
    $('app').querySelector('details').open = true;
  } else {
    box.innerHTML = '<button id="login" class="main">Войти в Google</button> <span class="hint">личный аккаунт с папкой Secretary</span>';
    $('login').onclick = login;
  }
  if (error) box.innerHTML += `<div class="err">${esc(error)}</div>`;
  updateSend();
}

async function login() {
  try {
    await auth.connect();
    renderAuth();
    await refreshWorkspaces();
  } catch (e) {
    renderAuth(e.message);
  }
}

// ---------- контуры ----------
function renderWorkspaces(list) {
  const cur = lsGet(LS_WS, 'vesco');
  $('ws').innerHTML = list
    .map((w) => `<option value="${esc(w.id)}"${w.id === cur ? ' selected' : ''}>${esc(w.name || w.id)}</option>`)
    .join('');
}

async function refreshWorkspaces() {
  renderWorkspaces(lsGet(LS_WS_LIST, DEFAULT_WS));
  if (auth.state() !== 'connected') return;
  try {
    const list = await loadWorkspaces(drive);
    if (list?.length) {
      lsSet(LS_WS_LIST, list);
      renderWorkspaces(list);
    }
  } catch (e) {
    if (e instanceof NeedsAuthError) renderAuth(e.message);
  }
}

// ---------- выделение ----------
function renderList() {
  $('list').innerHTML = selected
    .map((s) => {
      const st = status.get(s.itemId);
      return `<li><span class="subj" title="${esc(s.subject)}">${s.hasAttachment ? '📎 ' : ''}${esc(s.subject || '(без темы)')}</span>
        <span class="st ${st?.cls ?? ''}">${esc(st?.text ?? '')}</span></li>`;
    })
    .join('');
  updateSend();
}

function updateSend() {
  const n = selected.length;
  $('send').textContent = n ? `Отправить (${n})` : 'Выделите письма';
  $('send').disabled = sending || !n || auth.state() !== 'connected';
}

async function refreshSelection() {
  if (sending) return; // список не меняем посреди отправки
  try {
    selected = await call((cb) => Office.context.mailbox.getSelectedItemsAsync(cb));
  } catch {
    const it = Office.context.mailbox.item;
    selected = it?.itemId ? [{ itemId: it.itemId, subject: it.subject, hasAttachment: (it.attachments ?? []).length > 0 }] : [];
  }
  status.clear();
  $('summary').textContent = '';
  renderList();
}

// ---------- отправка ----------
async function send() {
  const workspace = $('ws').value;
  lsSet(LS_WS, workspace);
  const intent = $('sign').checked ? 'sign' : 'store';
  const diag = Office.context.mailbox.diagnostics;
  const client = { host: diag?.hostName, version: diag?.hostVersion, platform: Office.context.platform };
  const count = { stored: 0, duplicate: 0, error: 0 };
  sending = true;
  updateSend();
  await ensureMasterCategory(Office.context.mailbox).catch(() => {});
  for (const sel of [...selected]) {
    const set = (text, cls = '') => {
      status.set(sel.itemId, { text, cls });
      renderList();
    };
    try {
      set('читаю…');
      const pkg = await buildPackage(Office.context.mailbox, sel, { workspace, intent, client, addinVersion: VERSION });
      set('загрузка 0%');
      const r = await uploadPackage(drive, pkg, {
        workspace,
        cache: folderCache,
        onProgress: (p) => set(`загрузка ${Math.round(p * 100)}%`),
      });
      await markSent(Office.context.mailbox, sel.itemId).catch(() => {});
      count[r.status]++;
      const skipped = pkg.meta.skipped.filter((s) => s.reason !== 'inline').length;
      set(r.status === 'duplicate' ? 'уже было' : `готово${skipped ? ` (пропущено ${skipped})` : ''}`, 'ok');
    } catch (e) {
      count.error++;
      set(`ошибка: ${e.message}`, 'err');
      if (e instanceof NeedsAuthError) {
        renderAuth(e.message);
        break;
      }
    }
  }
  sending = false;
  $('summary').textContent = `отправлено ${count.stored}, уже было ${count.duplicate}, ошибок ${count.error}`;
  updateSend();
}

// ---------- старт ----------
Office.onReady(() => {
  const ok =
    Office.context.requirements.isSetSupported('Mailbox', '1.15') &&
    typeof Office.context.mailbox?.getSelectedItemsAsync === 'function'; // Outlook 2016 ложно отвечает true на «1.15»
  if (!ok) {
    $('unsupported').hidden = false;
    return;
  }
  $('app').hidden = false;
  $('ver').textContent = VERSION;
  const c = auth.client();
  $('client-id').value = c.clientId;
  $('client-secret').value = c.clientSecret;
  $('save-client').onclick = () => {
    auth.setClient($('client-id').value, $('client-secret').value);
    renderAuth();
  };
  $('logout').onclick = () => {
    auth.disconnect();
    renderAuth();
  };
  $('ws').onchange = () => lsSet(LS_WS, $('ws').value);
  $('send').onclick = () => send();
  renderAuth();
  refreshWorkspaces();
  refreshSelection();
  Office.context.mailbox.addHandlerAsync(Office.EventType.SelectedItemsChanged, refreshSelection);
});
