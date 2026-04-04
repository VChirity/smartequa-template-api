# Deploy no Render (API da cantina + templates)

O app Flutter chama `https://smartequa-template-api.onrender.com`. Depois de qualquer alteração em `api_server.py` ou `pix_cantina.py`, o serviço no Render precisa ser atualizado.

## Opção A — Blueprint (recomendado)

1. No [Render Dashboard](https://dashboard.render.com), **New** → **Blueprint**.
2. Conecte o repositório GitHub/GitLab do **SmartEqua**.
3. O arquivo `render.yaml` na raiz do repo define o serviço com **Root Directory** = `Templates app`.
4. Ao criar o blueprint, defina os **secrets** (não ficam no Git):
   - **MERCADO_PAGO_ACCESS_TOKEN** — já deve existir (PIX).
   - **FIREBASE_SERVICE_ACCOUNT_JSON** — JSON completo da conta de serviço Firebase (uma linha). Obtenha em: Firebase Console → Configurações do projeto → Contas de serviço → Gerar nova chave privada. A conta precisa de acesso ao **Realtime Database** do projeto `equa-sec-apk`.
   - **GEMINI_API_KEY** — se usar transcrever/corrigir.

## Opção B — Serviço Web já existente

1. **Settings** do serviço `smartequa-template-api` (ou o nome que você usa).
2. **Root Directory** = `Templates app`.
3. **Build Command** = `pip install -r requirements.txt`
4. **Start Command** = `gunicorn api_server:app --bind 0.0.0.0:$PORT --workers 2`
5. Em **Environment**, adicione **FIREBASE_SERVICE_ACCOUNT_JSON** (secret) como no passo 4 da Opção A.
6. **Manual Deploy** → **Deploy latest commit**.

## Conferir se a quitação automática funciona

Depois do deploy, com o app logado como aluno com dívida:

- `POST /api/pix/settle-debt-balance` com header `Authorization: Bearer <idToken>` deve retornar `{"ok": true}` se o saldo cobrir a dívida.
- Após pagar um PIX de dívida, o app chama `POST /api/pix/confirm-debt` com `{"txid":"..."}` — deve retornar `{"ok": true}`.

Se aparecer `Servidor sem Firebase Admin`, a variável **FIREBASE_SERVICE_ACCOUNT_JSON** não está definida ou o JSON é inválido.
