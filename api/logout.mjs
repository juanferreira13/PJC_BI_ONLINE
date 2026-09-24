import { clearSessionCookie } from './_auth.mjs';

export function POST(request) {
  return new Response(null, {
    status: 302,
    headers: {
      Location: new URL('/', request.url).toString(),
      'Set-Cookie': clearSessionCookie(),
      'Cache-Control': 'no-store'
    }
  });
}

export function GET(request) { return POST(request); }
