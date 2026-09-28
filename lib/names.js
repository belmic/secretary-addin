// Имена по naming.config.yaml (единая точка правды; маски v1 заморожены):
//   папка письма — маска inbox.email без .md: {date}_{counterpart}_{subject}_{code}
//   письмо       — маска drive.eml без __{uid}: {date}__{subject}.eml
//   вложение     — маска drive.attachment без __{uid}: {date}__{slug}.{ext}
// {uid} (f-NNNN из files.json) присваивает ядро при приёме — надстройка его не знает.
// slugify: strip_chars, управляющие, эмодзи; пробелы → «-»; повторы «-» схлопываются;
// края без пробелов/точек/дефисов; ≤ 40 символов на поле; CON/PRN/… → «_»; пусто → fallback.
const STRIP = /[<>:"/\\|?*#%&{}[\]$!`;@=\u0000-\u001f]/g; // eslint-disable-line no-control-regex
const EMOJI = /\p{Extended_Pictographic}|\u200d|\ufe0f/gu;
const RESERVED = /^(con|prn|aux|nul|com[1-9]|lpt[1-9])$/i;
export const MAX_FIELD_LEN = 40;
export const CODE_LEN = 4;

export function slugify(value, { max = MAX_FIELD_LEN, fallback = 'untitled' } = {}) {
  const trim = (s) => s.replace(/^[\s.-]+|[\s.-]+$/g, '');
  let s = trim(
    String(value ?? '')
      .replace(STRIP, '')
      .replace(EMOJI, '')
      .replace(/\s+/g, '-')
      .replace(/-{2,}/g, '-'),
  );
  s = trim(s.slice(0, max));
  if (!s) return fallback;
  return RESERVED.test(s) ? `_${s}` : s;
}

// Адрес целиком (keep_full): juan@vesco.es → juan-vesco-es
export const slugAddress = (email) => slugify(String(email ?? '').toLowerCase().replace(/[@.]/g, '-'), { fallback: 'unknown' });

// Дата письма — YYYY-MM-DD по часам компьютера (как видит пользователь в Outlook).
export function ymd(date) {
  const d = date instanceof Date && !isNaN(date) ? date : new Date();
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

export const folderName = ({ date, counterpart, subject, code }) =>
  `${date}_${slugAddress(counterpart)}_${slugify(subject)}_${code}`;

export const emlName = ({ date, subject }) => `${date}__${slugify(subject)}.eml`;

// Имена вложений письма: {date}__{slug}.{ext}; совпадения (без учёта регистра) → slug-2, slug-3.
export function attachmentNames(date, names) {
  const used = new Set();
  return names.map((n) => {
    const s = String(n ?? '');
    const dot = s.lastIndexOf('.');
    const rawExt = dot > 0 ? s.slice(dot + 1) : '';
    const ext = /^[\p{L}\p{N}]{1,10}$/u.test(rawExt) ? rawExt.toLowerCase() : '';
    const slug = slugify(ext ? s.slice(0, dot) : s, { fallback: 'attachment' });
    let out = `${date}__${slug}${ext ? `.${ext}` : ''}`;
    for (let i = 2; used.has(out.toLowerCase()); i++) out = `${date}__${slug}-${i}${ext ? `.${ext}` : ''}`;
    used.add(out.toLowerCase());
    return out;
  });
}
