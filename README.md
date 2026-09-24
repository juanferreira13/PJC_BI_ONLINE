# BI de Operações — PJC

Pacote preparado para GitHub + Vercel com atualização automática diária.

## Estrutura

- `public/` — únicos arquivos publicados pelo Vercel.
- `automation/` — planilha e scripts PowerShell usados para atualizar os dados.
- `.github/workflows/atualizar-bi.yml` — baixa a planilha do OneDrive, gera `public/dados_bi.js` e faz commit automático.
- `vercel.json` — publica somente `public/` e desativa cache do painel.

## Subir no GitHub

Crie um repositório **privado**, extraia este ZIP dentro dele e rode:

```bash
git init
git add .
git commit -m "BI PJC online"
git branch -M main
git remote add origin URL_DO_REPOSITORIO
git push -u origin main
```

## Publicar no Vercel

1. Importe o repositório no Vercel.
2. Framework Preset: `Other`.
3. O `vercel.json` define `public` como diretório publicado.
4. Faça o primeiro deploy.

O endereço raiz abre diretamente o BI de Operações.

## Atualização automática

O workflow roda todos os dias às **06:00 (horário de Brasília)** e também pode ser executado manualmente em:

`GitHub > Actions > Atualizar BI PJC > Run workflow`

Fluxo:

1. baixa a planilha atual do OneDrive;
2. executa `Gerar-BI.ps1`;
3. atualiza `public/dados_bi.js`;
4. faz commit e push;
5. o Vercel detecta o commit e publica automaticamente.

## Segurança

Este pacote evita publicar a planilha `.xlsx` e os scripts PowerShell porque somente `public/` é usado como saída do Vercel.

**Importante:** os dados contidos em `public/dados_bi.js` continuam sendo dados do BI e ficam acessíveis a quem tiver acesso ao site. Antes de usar em produção com informações confidenciais, proteja o site com autenticação/controle de acesso apropriado.
