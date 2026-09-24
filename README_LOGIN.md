# Login protegido no Vercel

Este pacote protege o BI no servidor: `dados_bi.js` e o HTML do painel ficam fora da pasta pública e só são entregues quando há uma sessão válida.

## 1. Variáveis no Vercel

Abra **Vercel > Project > Settings > Environment Variables** e crie:

- `PJC_LOGIN_EMAIL` = e-mail autorizado
- `PJC_LOGIN_PASSWORD` = senha autorizada
- `PJC_SESSION_SECRET` = uma chave aleatória longa (recomendado: 40+ caracteres)

Marque as variáveis para **Production** e, se quiser testar em previews, também **Preview**.

Depois clique em **Redeploy**.

## 2. Uso

- `/` = tela de login
- `/app` = BI protegido
- `/api/logout` = encerra a sessão

A sessão dura 12 horas e usa cookie `HttpOnly`, `Secure` e `SameSite=Lax`.

## 3. Segurança

Não coloque a senha diretamente no HTML, JavaScript público ou no GitHub. Mantenha o repositório privado porque os scripts de automação contêm a origem dos dados.
