// Пакет письма → Google Drive: Secretary/ws-<контур>/inbox/email/{mail_id}/
// message.eml → attachments/* → meta.json ПОСЛЕДНИМ. Есть meta.json → «уже отправлено».
// Файл, созданный прошлой (оборванной) попыткой, не перезаливается: и multipart, и resumable
// создают файл в Drive только целиком.
import { workspaceFolder } from './shared/workspaces.js';

const FOLDER = 'application/vnd.google-apps.folder';
export const SMALL_BYTES = 5 * 1024 * 1024; // до 5 МБ — одним запросом (multipart), больше — resumable

export const emailPath = (workspace) => ['Secretary', workspaceFolder(workspace), 'inbox', 'email'];

export async function uploadPackage(drive, pkg, { workspace, cache = {}, now = () => new Date(), onProgress } = {}) {
  const base = emailPath(workspace);
  const emailDir = await drive.ensureFolderPath(base, cache);
  const existing = await drive.findChild(emailDir, pkg.mailId, FOLDER);
  if (existing && (await drive.findChild(existing.id, 'meta.json'))) return { status: 'duplicate', folderId: existing.id };
  const dir = existing?.id ?? (await drive.ensureFolderPath([...base, pkg.mailId], cache));

  const total = pkg.files.reduce((s, f) => s + f.bytes.length, 0) || 1;
  let done = 0;
  let attDir = null;
  for (const f of pkg.files) {
    const slash = f.path.indexOf('/');
    const name = slash >= 0 ? f.path.slice(slash + 1) : f.path;
    const parent = slash >= 0 ? (attDir ??= await drive.ensureFolderPath([...base, pkg.mailId, 'attachments'], cache)) : dir;
    if (!(await drive.findChild(parent, name))) {
      const blob = new Blob([f.bytes], { type: f.mime });
      if (blob.size <= SMALL_BYTES) await drive.upsertFile({ parentId: parent, name, content: blob, mimeType: f.mime });
      else
        await drive.uploadResumable({
          blob,
          name,
          parentId: parent,
          mimeType: f.mime,
          onProgress: (sent) => onProgress?.((done + sent) / total),
        });
    }
    done += f.bytes.length;
    onProgress?.(done / total);
  }
  const meta = { ...pkg.meta, sent_at: now().toISOString() };
  await drive.upsertFile({ parentId: dir, name: 'meta.json', content: JSON.stringify(meta, null, 2), mimeType: 'application/json' });
  return { status: 'stored', folderId: dir };
}

// Реестр контуров Secretary/workspaces.json (пишет расширение) → [{id, name}] или null.
export async function loadWorkspaces(drive, base = 'https://www.googleapis.com') {
  const sec = await drive.findChild('root', 'Secretary', FOLDER);
  if (!sec) return null;
  const f = await drive.findChild(sec.id, 'workspaces.json');
  if (!f) return null;
  const res = await drive.request(`${base}/drive/v3/files/${f.id}?alt=media`);
  return (await res.json()).workspaces ?? null;
}
