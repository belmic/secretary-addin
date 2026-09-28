// Вход Google из надстройки: диалог Office (displayDialogAsync) → Authorization Code + PKCE
// → refresh-токен. Тот же OAuth-клиент, что у расширения (drive.file видит папки Secretary).
// Хранение — только localStorage панели на этой машине (не roamingSettings: они в корп-ящике).
// Выход — только забыть токены: отзыв (revoke) снял бы доступ и у расширения.
const KEY = 'secretary-auth';
const AUTH_URL = 'https://accounts.google.com/o/oauth2/v2/auth';
const TOKEN_URL = 'https://oauth2.googleapis.com/token';
export const SCOPE = 'https://www.googleapis.com/auth/drive.file';
const MARGIN_MS = 5 * 60_000;

export class NeedsAuthError extends Error {
  constructor(msg = 'Войдите в Google (кнопка «Войти»)') {
    super(msg);
    this.name = 'NeedsAuthError';
  }
}

const b64url = (bytes) =>
  btoa(String.fromCharCode(...new Uint8Array(bytes)))
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');

export function createAuth({ storage, fetchImpl = (...a) => fetch(...a), openDialog, pagesBase, now = () => Date.now() }) {
  const load = () => {
    try {
      return JSON.parse(storage.getItem(KEY) || '{}');
    } catch {
      return {};
    }
  };
  const save = (patch) => storage.setItem(KEY, JSON.stringify({ ...load(), ...patch }));

  async function token(params) {
    const res = await fetchImpl(TOKEN_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams(params),
    });
    const j = await res.json().catch(() => ({}));
    if (!res.ok) {
      const e = new Error(`Google: ${j.error || res.status} ${j.error_description || ''}`.trim());
      e.code = j.error;
      throw e;
    }
    return j;
  }

  return {
    client() {
      const a = load();
      return { clientId: a.clientId || '', clientSecret: a.clientSecret || '' };
    },
    setClient(clientId, clientSecret) {
      save({ clientId: clientId.trim(), clientSecret: clientSecret.trim() });
    },
    state() {
      const a = load();
      if (!a.clientId || !a.clientSecret) return 'unconfigured';
      return a.refreshToken ? 'connected' : 'disconnected';
    },

    async connect() {
      const a = load();
      if (!a.clientId || !a.clientSecret) throw new Error('Заполните Client ID и Client Secret (настройки)');
      const redirectUri = `${pagesBase}auth-callback.html`;
      const verifier = b64url(crypto.getRandomValues(new Uint8Array(32)));
      const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
      const state = b64url(crypto.getRandomValues(new Uint8Array(16)));
      const url = new URL(AUTH_URL);
      url.search = new URLSearchParams({
        client_id: a.clientId,
        redirect_uri: redirectUri,
        response_type: 'code',
        scope: SCOPE,
        access_type: 'offline',
        prompt: 'consent', // иначе Google не выдаст refresh-токен повторно
        code_challenge: challenge,
        code_challenge_method: 'S256',
        state,
      });
      const back = await openDialog(`${pagesBase}auth-start.html?u=${encodeURIComponent(url.href)}`);
      if (back.error) throw new Error(`Вход Google: ${back.error}`);
      if (back.state !== state) throw new Error('Вход Google: state не совпал');
      const tok = await token({
        code: back.code,
        client_id: a.clientId,
        client_secret: a.clientSecret,
        redirect_uri: redirectUri,
        grant_type: 'authorization_code',
        code_verifier: verifier,
      });
      if (!tok.refresh_token) throw new Error('Google не выдал refresh-токен — повторите вход');
      save({ refreshToken: tok.refresh_token, accessToken: tok.access_token, expiresAt: now() + tok.expires_in * 1000 });
    },

    async getAccessToken({ force = false } = {}) {
      const a = load();
      if (!a.refreshToken) throw new NeedsAuthError();
      if (!force && a.accessToken && a.expiresAt - now() > MARGIN_MS) return a.accessToken;
      try {
        const tok = await token({
          grant_type: 'refresh_token',
          refresh_token: a.refreshToken,
          client_id: a.clientId,
          client_secret: a.clientSecret,
        });
        save({ accessToken: tok.access_token, expiresAt: now() + tok.expires_in * 1000 });
        return tok.access_token;
      } catch (e) {
        if (e.code === 'invalid_grant') {
          save({ refreshToken: null, accessToken: null });
          throw new NeedsAuthError('Доступ к Google истёк или отозван — войдите заново');
        }
        throw e;
      }
    },

    disconnect() {
      save({ refreshToken: null, accessToken: null, expiresAt: 0 });
    },
  };
}
