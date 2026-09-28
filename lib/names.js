// Безопасные имена файлов вложений (docs/spec.md «Почта из Outlook»).
const BAD = /[<>:"/\\|?*#%&{}[\]$!`;@=\u0000-\u001f]/g; // eslint-disable-line no-control-regex
const RESERVED = /^(con|prn|aux|nul|com[1-9]|lpt[1-9])$/i;

export function safeName(name, max = 80, fallback = 'attachment') {
  let s = String(name ?? '')
    .replace(BAD, '')
    .replace(/\s+/g, ' ')
    .replace(/^[.\s]+|[.\s]+$/g, '');
  if (!s) s = fallback;
  const dot = s.lastIndexOf('.');
  let base = dot > 0 ? s.slice(0, dot) : s;
  let ext = dot > 0 ? s.slice(dot) : '';
  if (ext.length > 16) {
    base = s;
    ext = '';
  }
  if (RESERVED.test(base)) base = `_${base}`;
  if (base.length + ext.length > max) base = base.slice(0, max - ext.length).replace(/[.\s]+$/, '') || fallback;
  return base + ext;
}

// Совпадения без учёта регистра (Windows/macOS) → «-2», «-3» перед расширением.
export function uniqueNames(names) {
  const used = new Set();
  return names.map((n) => {
    const s = safeName(n);
    const dot = s.lastIndexOf('.');
    const base = dot > 0 ? s.slice(0, dot) : s;
    const ext = dot > 0 ? s.slice(dot) : '';
    let out = s;
    for (let i = 2; used.has(out.toLowerCase()); i++) out = `${base}-${i}${ext}`;
    used.add(out.toLowerCase());
    return out;
  });
}
