import fs from 'node:fs';
import path from 'node:path';
import { redirectToLogin, verifySession } from './_auth.mjs';

export function GET(request) {
  if (!verifySession(request)) return redirectToLogin(request);
  const file = path.join(process.cwd(), 'private', 'BI_OPERACOES_PJC.html');
  const html = fs.readFileSync(file, 'utf8');
  return new Response(html, {
    headers: {
      'Content-Type': 'text/html; charset=utf-8',
      'Cache-Control': 'private, no-store, no-cache, must-revalidate'
    }
  });
}
