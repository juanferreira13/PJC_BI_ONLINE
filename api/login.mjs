import { createSession, sessionCookie, validCredentials } from './_auth.mjs';

export async function POST(request) {
  let data = {};
  const type = request.headers.get('content-type') || '';
  try {
    if (type.includes('application/json')) data = await request.json();
    else {
      const form = await request.formData();
      data = Object.fromEntries(form.entries());
    }
  } catch {
    return Response.json({ ok: false, message: 'Dados inválidos.' }, { status: 400 });
  }

  if (!validCredentials(data.email, data.password)) {
    return Response.json({ ok: false, message: 'E-mail ou senha incorretos.' }, {
      status: 401,
      headers: { 'Cache-Control': 'no-store' }
    });
  }

  const token = createSession(data.email);
  return Response.json({ ok: true }, {
    headers: {
      'Set-Cookie': sessionCookie(token),
      'Cache-Control': 'no-store'
    }
  });
}
