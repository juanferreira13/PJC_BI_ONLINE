import crypto from 'node:crypto';

const COOKIE = 'pjc_session';
const MAX_AGE = 60 * 60 * 12;

function b64url(input) {
  return Buffer.from(input).toString('base64url');
}

function safeEqual(a, b) {
  const ab = Buffer.from(String(a ?? ''));
  const bb = Buffer.from(String(b ?? ''));
  if (ab.length !== bb.length) return false;
  return crypto.timingSafeEqual(ab, bb);
}

function secret() {
  const s = process.env.PJC_SESSION_SECRET;
  if (!s || s.length < 24) throw new Error('PJC_SESSION_SECRET ausente ou curta.');
  return s;
}

export function validCredentials(email, password) {
  const expectedEmail = process.env.PJC_LOGIN_EMAIL;
  const expectedPassword = process.env.PJC_LOGIN_PASSWORD;
  if (!expectedEmail || !expectedPassword) return false;
  return safeEqual(String(email).trim().toLowerCase(), expectedEmail.trim().toLowerCase()) &&
         safeEqual(password, expectedPassword);
}

export function createSession(email) {
  const payload = {
    email: String(email).trim().toLowerCase(),
    exp: Math.floor(Date.now() / 1000) + MAX_AGE
  };
  const body = b64url(JSON.stringify(payload));
  const sig = crypto.createHmac('sha256', secret()).update(body).digest('base64url');
  return `${body}.${sig}`;
}

export function verifySession(request) {
  try {
    const cookies = request.headers.get('cookie') || '';
    const raw = cookies.split(';').map(v => v.trim()).find(v => v.startsWith(`${COOKIE}=`));
    if (!raw) return null;
    const token = decodeURIComponent(raw.slice(COOKIE.length + 1));
    const [body, sig] = token.split('.');
    if (!body || !sig) return null;
    const expected = crypto.createHmac('sha256', secret()).update(body).digest('base64url');
    if (!safeEqual(sig, expected)) return null;
    const payload = JSON.parse(Buffer.from(body, 'base64url').toString('utf8'));
    if (!payload.exp || payload.exp < Math.floor(Date.now() / 1000)) return null;
    return payload;
  } catch {
    return null;
  }
}

export function sessionCookie(token) {
  return `${COOKIE}=${encodeURIComponent(token)}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=${MAX_AGE}`;
}

export function clearSessionCookie() {
  return `${COOKIE}=; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=0`;
}

export function redirectToLogin(request) {
  const url = new URL('/', request.url);
  return new Response(null, { status: 302, headers: { Location: url.toString(), 'Cache-Control': 'no-store' } });
}
