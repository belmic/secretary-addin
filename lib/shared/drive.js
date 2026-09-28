// Google Drive API v3 (scope drive.file: приложение видит только свои файлы и папки).
//   ensureFolderPath — найти или создать цепочку папок;
//   uploadResumable  — resumable upload чанками по 8 MiB с продолжением по URI сессии;
//   upsertFile       — создать или перезаписать небольшой файл (manifest.json, events.jsonl).
// Ошибки: 401 → обновить токен и повторить; 429/5xx/сеть → экспоненциальный backoff.

export const CHUNK_BYTES = 8 * 1024 * 1024; // кратно 256 KiB
const FOLDER_MIME = 'application/vnd.google-apps.folder';

export class DriveError extends Error {
  constructor(message, { status, retryable = false, cause } = {}) {
    super(message, { cause });
    this.name = 'DriveError';
    this.status = status;
    this.retryable = retryable;
  }
}

const isRetryableStatus = (s) => s === 429 || s === 408 || s >= 500;

export function createDrive({
  getAccessToken, // ({force}) => token
  fetchImpl = (...a) => fetch(...a),
  base = 'https://www.googleapis.com',
  sleep = (ms) => new Promise((r) => setTimeout(r, ms)),
  maxRetries = 5, // 1, 2, 4, 8, 16 с внутри одного вызова; дальше — решает очередь
  chunkBytes = CHUNK_BYTES,
  log,
}) {
  // Один HTTP-запрос с авторизацией, повтором на 401 и backoff на временных ошибках.
  async function request(url, init = {}, { okStatuses = [] } = {}) {
    let refreshed = false;
    for (let attempt = 0; ; attempt++) {
      const token = await getAccessToken({ force: false });
      let res;
      try {
        res = await fetchImpl(url, { ...init, headers: { ...init.headers, Authorization: `Bearer ${token}` } });
      } catch (e) {
        if (attempt >= maxRetries) throw new DriveError(`network: ${e.message}`, { retryable: true, cause: e });
        log?.warn('network error, retry', attempt + 1, e.message);
        await sleep(1000 * 2 ** attempt);
        continue;
      }
      if (res.ok || okStatuses.includes(res.status)) return res;
      if (res.status === 401 && !refreshed) {
        refreshed = true;
        await getAccessToken({ force: true });
        attempt--;
        continue;
      }
      if (isRetryableStatus(res.status) && attempt < maxRetries) {
        log?.warn(`HTTP ${res.status}, retry`, attempt + 1);
        await sleep(1000 * 2 ** attempt);
        continue;
      }
      const text = await res.text().catch(() => '');
      throw new DriveError(`Drive ${res.status}: ${text.slice(0, 300)}`, {
        status: res.status,
        retryable: isRetryableStatus(res.status),
      });
    }
  }

  const json = async (url, init, opts) => (await request(url, init, opts)).json();
  const q = (s) => s.replace(/\\/g, '\\\\').replace(/'/g, "\\'");

  async function findChild(parentId, name, mimeType) {
    const cond = [`name = '${q(name)}'`, `'${parentId}' in parents`, 'trashed = false'];
    if (mimeType) cond.push(`mimeType = '${mimeType}'`);
    const url = `${base}/drive/v3/files?${new URLSearchParams({
      q: cond.join(' and '),
      fields: 'files(id,name,modifiedTime)',
      spaces: 'drive',
      orderBy: 'createdTime',
    })}`;
    return (await json(url)).files?.[0] ?? null;
  }

  async function createFolder(parentId, name) {
    return json(`${base}/drive/v3/files?fields=id,name`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name, mimeType: FOLDER_MIME, parents: [parentId] }),
    });
  }

  const drive = {
    request,

    about: () => json(`${base}/drive/v3/about?fields=user(emailAddress,displayName)`),

    // Цепочка папок от корня «Мой диск». cache: Map/obj path → id (передаёт вызывающий).
    async ensureFolderPath(names, cache = {}) {
      let parent = 'root';
      let path = '';
      for (const name of names) {
        path += `/${name}`;
        if (cache[path]) {
          parent = cache[path];
          continue;
        }
        const found = await findChild(parent, name, FOLDER_MIME);
        parent = (found ?? (await createFolder(parent, name))).id;
        cache[path] = parent;
      }
      return parent;
    },

    findChild,

    // Resumable upload. state = {uploadUri?} — сохраняется вызывающим через onSession,
    // чтобы после падения продолжить с места (без дублей: файл создаётся только в конце).
    async uploadResumable({ blob, name, parentId, mimeType, uploadUri, onSession, onProgress }) {
      const total = blob.size;
      let uri = uploadUri;
      let offset = 0;

      if (uri) {
        const st = await queryStatus(uri, total);
        if (st.done) return st.file;
        if (st.expired) uri = null;
        else offset = st.offset;
      }
      if (!uri) {
        const res = await request(`${base}/upload/drive/v3/files?uploadType=resumable&fields=id,name,size,md5Checksum`, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json; charset=UTF-8',
            'X-Upload-Content-Type': mimeType,
            'X-Upload-Content-Length': String(total),
          },
          body: JSON.stringify({ name, parents: [parentId], mimeType }),
        });
        uri = res.headers.get('Location');
        if (!uri) throw new DriveError('resumable: нет Location');
        await onSession?.(uri);
        offset = 0;
      }

      while (true) {
        const end = Math.min(offset + chunkBytes, total);
        let res;
        try {
          res = await request(
            uri,
            {
              method: 'PUT',
              headers: { 'Content-Range': total ? `bytes ${offset}-${end - 1}/${total}` : 'bytes */0' },
              body: blob.slice(offset, end),
            },
            { okStatuses: [308] },
          );
        } catch (e) {
          if (e.status === 404 || e.status === 410) await onSession?.(null); // сессия умерла
          throw e;
        }
        if (res.status === 308) {
          offset = rangeEnd(res.headers.get('Range'));
          onProgress?.(offset, total);
          continue;
        }
        onProgress?.(total, total);
        return res.json();
      }
    },

    // Небольшой файл целиком: create (multipart) или update содержимого (media PATCH).
    async upsertFile({ parentId, name, content, mimeType, fileId }) {
      let id = fileId;
      if (!id) id = (await findChild(parentId, name))?.id;
      if (id) {
        const res = await request(
          `${base}/upload/drive/v3/files/${id}?uploadType=media&fields=id`,
          { method: 'PATCH', headers: { 'Content-Type': mimeType }, body: content },
          { okStatuses: [404] },
        );
        if (res.status !== 404) return res.json();
        // файл удалён пользователем — создадим заново
      }
      const boundary = `mr${crypto.getRandomValues(new Uint32Array(2)).join('')}`;
      const body = new Blob([
        `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n`,
        JSON.stringify({ name, parents: [parentId], mimeType }),
        `\r\n--${boundary}\r\nContent-Type: ${mimeType}\r\n\r\n`,
        content,
        `\r\n--${boundary}--`,
      ]);
      return json(`${base}/upload/drive/v3/files?uploadType=multipart&fields=id`, {
        method: 'POST',
        headers: { 'Content-Type': `multipart/related; boundary=${boundary}` },
        body,
      });
    },

    folderUrl: (id) => `https://drive.google.com/drive/folders/${id}`,
  };

  // PUT с пустым телом и Content-Range: bytes */total → сколько уже принято.
  async function queryStatus(uri, total) {
    let res;
    try {
      res = await request(
        uri,
        { method: 'PUT', headers: { 'Content-Range': `bytes */${total}` } },
        { okStatuses: [308, 404, 410] },
      );
    } catch (e) {
      if (e.status === 400) return { expired: true };
      throw e;
    }
    if (res.status === 404 || res.status === 410) return { expired: true };
    if (res.status === 308) return { offset: rangeEnd(res.headers.get('Range')) };
    return { done: true, file: await res.json() };
  }

  return drive;
}

// "bytes=0-1048575" → 1048576; нет заголовка → 0.
export function rangeEnd(range) {
  const m = /bytes=\d+-(\d+)/.exec(range ?? '');
  return m ? Number(m[1]) + 1 : 0;
}
