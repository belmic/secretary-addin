// Письмо → пакет для Drive (docs/spec.md «Почта из Outlook»). Строго по одному письму:
// loadItemByIdAsync → EML → вложения → unloadAsync (всегда, в finally).
import { call } from './office.js';
import { attachmentNames, emlName, folderName, ymd, CODE_LEN } from './names.js';

export const MAX_ATTACHMENT_BYTES = 50 * 1024 * 1024; // больше — в skipped (файл всё равно есть в EML)

export async function sha256hex(text) {
  const d = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
  return [...new Uint8Array(d)].map((b) => b.toString(16).padStart(2, '0')).join('');
}

export function b64ToBytes(b64) {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

const addr = (a) => (a ? { name: a.displayName || '', email: a.emailAddress || '' } : null);
const iso = (d) => (d instanceof Date && !isNaN(d) ? d.toISOString() : null);

export async function buildPackage(mailbox, sel, opts) {
  const m = await call((cb) => mailbox.loadItemByIdAsync(sel.itemId, cb));
  try {
    return await fromLoaded(m, sel, opts);
  } finally {
    await call((cb) => m.unloadAsync(cb)).catch(() => {});
  }
}

async function fromLoaded(
  m,
  sel,
  { workspace, intent = 'store', client = {}, addinVersion = '', userEmail = '', maxAttachmentBytes = MAX_ATTACHMENT_BYTES },
) {
  const imid = m.internetMessageId || '';
  const mailId = (await sha256hex(imid || sel.itemId)).slice(0, 12);
  const subject = m.subject ?? sel.subject ?? '';
  const date = ymd(m.dateTimeCreated);
  // counterpart: отправитель входящего; для своего (исходящего) — первый получатель
  const fromEmail = m.from?.emailAddress || '';
  const outgoing = userEmail && fromEmail.toLowerCase() === userEmail.toLowerCase();
  const counterpart = outgoing ? m.to?.[0]?.emailAddress || fromEmail : fromEmail;
  const folder = folderName({ date, counterpart, subject, code: mailId.slice(0, CODE_LEN) });
  const eml = emlName({ date, subject });
  const files = [];
  const skipped = [];
  let emlBytes = null;

  if (typeof m.getAsFileAsync === 'function') {
    try {
      const bytes = b64ToBytes(await call((cb) => m.getAsFileAsync(cb)));
      files.push({ path: eml, bytes, mime: 'message/rfc822' });
      emlBytes = bytes.length;
    } catch (e) {
      skipped.push({ name: eml, reason: `error: ${e.message}` });
    }
  }

  const list = m.attachments ?? [];
  const names = attachmentNames(
    date,
    list.map((a) => (a.attachmentType === 'item' ? `${a.name || 'message'}.eml` : a.name)),
  );
  const attachments = [];
  for (const [i, a] of list.entries()) {
    const rec = {
      name: a.name,
      size: a.size,
      content_type: a.contentType ?? null,
      is_inline: !!a.isInline,
      kind: a.attachmentType,
      saved_as: null,
    };
    attachments.push(rec);
    if (a.isInline) {
      skipped.push({ name: a.name, reason: 'inline' });
      continue;
    }
    if (a.attachmentType !== 'cloud' && a.size > maxAttachmentBytes) {
      skipped.push({ name: a.name, reason: 'too_large' });
      continue;
    }
    try {
      const c = await call((cb) => m.getAttachmentContentAsync(a.id, cb));
      if (c.format === 'url') {
        rec.url = c.content; // облачное вложение (OneDrive/SharePoint): только ссылка
        continue;
      }
      let name = names[i];
      if (c.format === 'iCalendar' && !/\.ics$/i.test(name)) name = `${name.replace(/\.[^.]*$/, '')}.ics`;
      const bytes = c.format === 'base64' ? b64ToBytes(c.content) : new TextEncoder().encode(c.content);
      const mime =
        c.format === 'base64' ? a.contentType || 'application/octet-stream' : c.format === 'eml' ? 'message/rfc822' : 'text/calendar';
      rec.saved_as = `attachments/${name}`;
      files.push({ path: rec.saved_as, bytes, mime });
    } catch (e) {
      skipped.push({ name: a.name, reason: `error: ${e.message}` });
    }
  }

  const meta = {
    schema: 1,
    source: 'outlook-addin',
    addin_version: addinVersion,
    workspace,
    intent,
    mail_id: mailId,
    folder,
    eml_name: eml,
    internet_message_id: imid || null,
    conversation_id: m.conversationId ?? sel.conversationId ?? null,
    subject,
    from: addr(m.from),
    to: (m.to ?? []).map(addr),
    cc: (m.cc ?? []).map(addr),
    received_at: iso(m.dateTimeCreated),
    eml_saved: emlBytes !== null,
    eml_bytes: emlBytes,
    attachments,
    skipped,
    client,
  };
  return { mailId, folder, itemId: sel.itemId, subject, meta, files };
}

// Ошибка одного письма не останавливает остальные.
export async function* buildPackages(mailbox, selected, opts) {
  for (const sel of selected) {
    try {
      yield { sel, pkg: await buildPackage(mailbox, sel, opts) };
    } catch (e) {
      yield { sel, error: e.message };
    }
  }
}
