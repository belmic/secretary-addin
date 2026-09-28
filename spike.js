// Спайки plan-outlook.md: 0.2 — извлечение выделенных писем, 0.3 — вход Google из диалога Office.
// Только новый Outlook / Outlook в браузере (WebView2 / браузер: современный JS).
// В отчёт не попадают ни содержимое писем, ни имена файлов, ни токены — только размеры, типы, время.

const report = { spike_version: '0.2.0', generated_at: new Date().toISOString() };
const $ = (id) => document.getElementById(id);
const render = () => ($('out').value = JSON.stringify(report, null, 2));
const status = (t) => ($('status').textContent = t || '');

// Office-колбэк → Promise с замером времени.
function call(fn) {
  const t0 = performance.now();
  return new Promise((resolve) => {
    try {
      fn((r) =>
        resolve({
          ok: r.status === Office.AsyncResultStatus.Succeeded,
          value: r.value,
          error: r.error?.message,
          ms: Math.round(performance.now() - t0),
        }),
      );
    } catch (e) {
      resolve({ ok: false, error: `throw: ${e.message}`, ms: Math.round(performance.now() - t0) });
    }
  });
}

// ---------- 0.2 извлечение ----------
async function extraction() {
  const mbx = Office.context.mailbox;
  status('Извлечение…');
  const sel = await call((cb) => mbx.getSelectedItemsAsync(cb));
  const items = [];
  for (const s of sel.value ?? []) {
    const it = { itemType: s.itemType, hasAttachment: s.hasAttachment };
    const ld = await call((cb) => mbx.loadItemByIdAsync(s.itemId, cb));
    it.load = { ok: ld.ok, ms: ld.ms, error: ld.error };
    if (!ld.ok) {
      items.push(it);
      continue;
    }
    const m = ld.value;
    it.has_internetMessageId = !!m.internetMessageId;
    it.api = {
      getAsFileAsync: typeof m.getAsFileAsync === 'function',
      getAttachmentContentAsync: typeof m.getAttachmentContentAsync === 'function',
      categories: !!m.categories,
    };
    if (it.api.getAsFileAsync) {
      const f = await call((cb) => m.getAsFileAsync(cb));
      it.eml = { ok: f.ok, ms: f.ms, error: f.error, bytes_approx: f.ok ? Math.round((f.value.length * 3) / 4) : null };
    }
    it.attachments = [];
    for (const a of m.attachments ?? []) {
      const ai = {
        ext: (a.name || '').includes('.') ? a.name.split('.').pop().slice(0, 8).toLowerCase() : '',
        size: a.size,
        attachmentType: a.attachmentType,
        isInline: a.isInline,
        contentType: a.contentType,
      };
      if (it.api.getAttachmentContentAsync) {
        const c = await call((cb) => m.getAttachmentContentAsync(a.id, cb));
        ai.content = { ok: c.ok, ms: c.ms, error: c.error, format: c.value?.format, length: c.value?.content?.length };
      }
      it.attachments.push(ai);
      status(`Извлечение… письмо ${items.length + 1}, вложение ${it.attachments.length}`);
    }
    const un = await call((cb) => m.unloadAsync(cb));
    it.unload_ok = un.ok;
    items.push(it);
  }
  report.extraction = { selected: sel.value?.length ?? 0, select_error: sel.error, items };
  status('');
  render();
}

// Категория «Secretary ✓»: мастер-список ящика + пометка первого выделенного письма (и снятие).
async function categoryTest() {
  const NAME = 'Secretary ✓';
  const mbx = Office.context.mailbox;
  const r = {};
  status('Категории…');
  const mc = await call((cb) => mbx.masterCategories.getAsync(cb));
  r.master_get = { ok: mc.ok, error: mc.error };
  if (mc.ok && !mc.value.some((c) => c.displayName === NAME)) {
    const add = await call((cb) =>
      mbx.masterCategories.addAsync([{ displayName: NAME, color: Office.MailboxEnums.CategoryColor.Preset4 }], cb),
    );
    r.master_add = { ok: add.ok, error: add.error };
  } else r.master_exists = mc.ok;
  const sel = await call((cb) => mbx.getSelectedItemsAsync(cb));
  const first = sel.value?.[0];
  if (first) {
    const ld = await call((cb) => mbx.loadItemByIdAsync(first.itemId, cb));
    if (ld.ok) {
      const m = ld.value;
      const a = await call((cb) => m.categories.addAsync([NAME], cb));
      r.item_add = { ok: a.ok, error: a.error, ms: a.ms };
      const rm = await call((cb) => m.categories.removeAsync([NAME], cb));
      r.item_remove = { ok: rm.ok, error: rm.error };
      await call((cb) => m.unloadAsync(cb));
    } else r.load_error = ld.error;
  } else r.note = 'нет выделенного письма';
  report.categories = r;
  status('');
  render();
}

// ---------- 0.3 вход Google ----------
const LS_KEY = 'secretary-spike-google';
const b64url = (bytes) =>
  btoa(String.fromCharCode(...new Uint8Array(bytes)))
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');

function openDialog(url) {
  return new Promise((resolve) => {
    Office.context.ui.displayDialogAsync(url, { height: 70, width: 40 }, (r) => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) return resolve({ error: `dialog: ${r.error.message}` });
      const dlg = r.value;
      dlg.addEventHandler(Office.EventType.DialogMessageReceived, (e) => {
        dlg.close();
        try {
          resolve(JSON.parse(e.message));
        } catch {
          resolve({ error: 'dialog: не JSON' });
        }
      });
      // закрыли окно / навигация не удалась (код ошибки диалога Office)
      dlg.addEventHandler(Office.EventType.DialogEventReceived, (e) => resolve({ error: `dialog event ${e.error}` }));
    });
  });
}

async function driveJson(token, url) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Drive ${res.status}: ${(await res.text()).slice(0, 200)}`);
  return res.json();
}

async function googleSpike() {
  const clientId = $('client-id').value.trim();
  const clientSecret = $('client-secret').value.trim();
  if (!clientId || !clientSecret) return status('Заполните Client ID и Client Secret');
  localStorage.setItem(LS_KEY, JSON.stringify({ clientId, clientSecret }));
  const g = (report.google = { redirect_uri: new URL('auth-callback.html', location.href).href });
  render();
  try {
    status('Открываю вход Google…');
    const verifier = b64url(crypto.getRandomValues(new Uint8Array(32)));
    const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
    const state = b64url(crypto.getRandomValues(new Uint8Array(16)));
    const auth = new URL('https://accounts.google.com/o/oauth2/v2/auth');
    // access_type=online, без prompt=consent: refresh-токен не выпускается — вход расширения не
    // затрагивается (отзыв или лимит refresh-токенов этого клиента его бы задел).
    auth.search = new URLSearchParams({
      client_id: clientId,
      redirect_uri: g.redirect_uri,
      response_type: 'code',
      scope: 'https://www.googleapis.com/auth/drive.file',
      access_type: 'online',
      code_challenge: challenge,
      code_challenge_method: 'S256',
      state,
    });
    const start = new URL('auth-start.html', location.href);
    start.searchParams.set('u', auth.href);
    const t0 = performance.now();
    const back = await openDialog(start.href);
    g.dialog_ms = Math.round(performance.now() - t0);
    if (back.error) throw new Error(`вход: ${back.error}`);
    if (back.state !== state) throw new Error('вход: state не совпал');
    g.code_received = !!back.code;

    status('Обмен кода на токен…');
    const tokRes = await fetch('https://oauth2.googleapis.com/token', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        code: back.code,
        client_id: clientId,
        client_secret: clientSecret,
        redirect_uri: g.redirect_uri,
        grant_type: 'authorization_code',
        code_verifier: verifier,
      }),
    });
    const tok = await tokRes.json();
    if (!tokRes.ok) throw new Error(`token ${tokRes.status}: ${tok.error} ${tok.error_description ?? ''}`);
    g.token_ok = true;

    status('Читаю Drive…');
    const q = (s) => `https://www.googleapis.com/drive/v3/files?fields=files(id,name)&q=${encodeURIComponent(s)}`;
    const root = await driveJson(
      tok.access_token,
      q("name='Secretary' and mimeType='application/vnd.google-apps.folder' and trashed=false"),
    );
    g.secretary_folders = root.files.length;
    if (!root.files.length) throw new Error('папка Secretary не видна (drive.file: другой OAuth-клиент?)');
    const reg = await driveJson(
      tok.access_token,
      q(`'${root.files[0].id}' in parents and name='workspaces.json' and trashed=false`),
    );
    if (!reg.files.length) throw new Error('workspaces.json не найден');
    const ws = await driveJson(tok.access_token, `https://www.googleapis.com/drive/v3/files/${reg.files[0].id}?alt=media`);
    g.workspaces = (ws.workspaces ?? ws).map?.((w) => w.id ?? w) ?? Object.keys(ws);
    g.result = 'OK → транспорт A';
  } catch (e) {
    g.error = e.message;
    g.result = /disallowed_useragent|403/.test(e.message) ? 'вход заблокирован → транспорт B' : 'ошибка';
  }
  status('');
  render();
}

function copy() {
  const ta = $('out');
  ta.focus();
  ta.select();
  let ok;
  try {
    ok = document.execCommand('copy');
  } catch {
    ok = false;
  }
  status(ok ? 'Скопировано — вставьте в чат' : 'Выделено — нажмите Ctrl+C');
}

Office.onReady((info) => {
  report.host = { platform: info.platform, hostName: Office.context.mailbox?.diagnostics?.hostName };
  try {
    const saved = JSON.parse(localStorage.getItem(LS_KEY) || '{}');
    $('client-id').value = saved.clientId ?? '';
    $('client-secret').value = saved.clientSecret ?? '';
  } catch {
    /* localStorage недоступен — поля пустые */
  }
  $('btn-extract').onclick = () => extraction().catch((e) => ((report.extraction = { error: e.message }), render()));
  $('btn-cat').onclick = () => categoryTest().catch((e) => ((report.categories = { error: e.message }), render()));
  $('btn-google').onclick = () => googleSpike();
  $('btn-copy').onclick = copy;
  render();
});
