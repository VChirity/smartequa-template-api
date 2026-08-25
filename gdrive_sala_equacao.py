# -*- coding: utf-8 -*-
"""
Arquivos da Sala Equação no Google Drive pessoal (Gmail).

A service account NÃO tem cota no Drive comum. O dono da pasta autoriza
UMA VEZ (OAuth). Professor/aluno continuam só no Smart Equação.

A pasta raiz deve ficar RESTRITA (não pública). Só cada arquivo publicado
ganha link de leitura, para abrir no app. Aluno não lista a pasta.
Subpastas: 2026 / turma / disciplina / Nº bimestre / Aulas|Entregas.
"""
import io
import json
import os
import re
import secrets
import time
import uuid
from pathlib import Path
from urllib.parse import urlencode

import requests as http_requests
from flask import jsonify, redirect, render_template_string, request, send_file

DEFAULT_FOLDER_ID = '1xYGoBFkHosbgFN5hKLjn2GB_DZN0gluq'
MAX_FILE_BYTES = 50 * 1024 * 1024
DRIVE_SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/userinfo.email',
    'openid',
]
RTDB_OAUTH_PATH = 'sala_equacao_secrets/drive_oauth'
RTDB_OAUTH_STATE = 'sala_equacao_secrets/drive_oauth_state'

CONNECT_PAGE = '''<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Conectar Drive — Sala Equação</title>
<style>
  body { font-family: Arial, sans-serif; max-width: 720px; margin: 24px auto; padding: 0 16px; color: #222; }
  h1 { font-size: 1.4rem; }
  .ok { background: #e8f5e9; border: 1px solid #81c784; padding: 12px 16px; border-radius: 8px; }
  .warn { background: #fff8e1; border: 1px solid #ffcc80; padding: 12px 16px; border-radius: 8px; }
  .err { background: #ffebee; border: 1px solid #ef9a9a; padding: 12px 16px; border-radius: 8px; }
  a.btn, button { display: inline-block; background: #f57c00; color: #fff; border: 0; padding: 12px 18px;
    border-radius: 8px; font-size: 1rem; text-decoration: none; cursor: pointer; }
  label { display: block; margin-top: 12px; font-weight: bold; }
  input { width: 100%; padding: 8px; box-sizing: border-box; }
  ol { line-height: 1.5; }
  code { background: #f5f5f5; padding: 1px 4px; }
  details { margin-top: 16px; }
</style>
</head>
<body>
<h1>Sala Equação — Google Drive</h1>
{% if connected %}
  <div class="ok">
    <b>Pronto.</b> Drive autorizado como <b>{{ email }}</b>.<br>
    Pasta raiz (restrita, só você vê a lista): a API cria sozinha
    <code>2026 / turma / disciplina / bimestre / Aulas ou Entregas</code>.<br>
    Pode voltar no app e publicar o PDF.
  </div>
{% else %}
  <div class="warn">
    A pasta já está escolhida. <b>Deixe ela Restrita</b> (não “qualquer pessoa com o link”).
    Aluno não deve abrir essa pasta. Ele só vê o arquivo da atividade no app.
  </div>
  {% if error %}<p class="err">{{ error }}</p>{% endif %}
  {% if not has_client %}
  <p>Um passo único no Google Cloud (projeto <code>equa-sec-apk</code>):</p>
  <ol>
    <li>Abra <a href="https://console.cloud.google.com/apis/credentials/consent?project=equa-sec-apk" target="_blank">Tela de consentimento</a>
      → Externo → e-mail de suporte = seu Gmail dos 5 TB → em Test users, adicione esse mesmo Gmail.</li>
    <li>Abra <a href="https://console.cloud.google.com/apis/credentials?project=equa-sec-apk" target="_blank">Credenciais</a>
      → Criar credenciais → ID do cliente OAuth → <b>Aplicativo da Web</b>.</li>
    <li>Em “URIs de redirecionamento”, cole exatamente:<br>
      <code>{{ redirect_uri }}</code></li>
    <li>Copie o ID e o Secret e cole abaixo.</li>
  </ol>
  <form method="post" action="/api/sala-equacao/drive-oauth/credentials">
    <label>ID do cliente</label>
    <input name="client_id" required autocomplete="off">
    <label>Segredo do cliente</label>
    <input name="client_secret" required autocomplete="off">
    <p><button type="submit">Salvar e continuar</button></p>
  </form>
  {% else %}
  <p>Credenciais do Google já estão no servidor. Agora autorize <b>o Gmail dono da pasta</b> (o dos 5 TB):</p>
  <p><a class="btn" href="/api/sala-equacao/drive-oauth/start">Autorizar pasta do Drive</a></p>
  {% endif %}
{% endif %}
<p style="margin-top:24px;color:#666;font-size:.9rem">
  Pasta: <code>{{ folder_id }}</code><br>
  Status: {{ storage }}
</p>
</body>
</html>
'''


def _upload_root():
    raw = os.environ.get('SALA_EQUACAO_UPLOAD_DIR') or 'data/sala_equacao'
    path = Path(raw)
    path.mkdir(parents=True, exist_ok=True)
    return path


def _index_path():
    return _upload_root() / 'index.json'


def _load_index():
    p = _index_path()
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding='utf-8'))
    except Exception:
        return {}


def _folder_id():
    return (os.environ.get('GOOGLE_DRIVE_SALA_FOLDER_ID') or DEFAULT_FOLDER_ID).strip()


def _sa_info():
    raw = os.environ.get('FIREBASE_SERVICE_ACCOUNT_JSON')
    if raw:
        return json.loads(raw)
    if os.path.exists('firebase-service-account.json'):
        with open('firebase-service-account.json', 'r', encoding='utf-8') as f:
            return json.load(f)
    return None


def _ensure_firebase():
    import firebase_admin
    from firebase_admin import credentials
    try:
        firebase_admin.get_app()
        return True, None
    except ValueError:
        info = _sa_info()
        if not info:
            return False, 'Servidor sem Firebase Admin'
        cred = credentials.Certificate(info)
        firebase_admin.initialize_app(cred, {
            'databaseURL': 'https://equa-sec-apk-default-rtdb.firebaseio.com',
        })
        return True, None


def _rtdb_get(path):
    ok, err = _ensure_firebase()
    if not ok:
        return None
    from firebase_admin import db
    return db.reference(path).get()


def _rtdb_set(path, value):
    ok, err = _ensure_firebase()
    if not ok:
        return False, err
    from firebase_admin import db
    db.reference(path).set(value)
    return True, None


def _oauth_saved():
    data = _rtdb_get(RTDB_OAUTH_PATH)
    return data if isinstance(data, dict) else {}


def _client_id():
    return (
        (os.environ.get('GOOGLE_DRIVE_CLIENT_ID') or '').strip()
        or (_oauth_saved().get('clientId') or '').strip()
    )


def _client_secret():
    return (
        (os.environ.get('GOOGLE_DRIVE_CLIENT_SECRET') or '').strip()
        or (_oauth_saved().get('clientSecret') or '').strip()
    )


def _refresh_token():
    return (
        (os.environ.get('GOOGLE_DRIVE_REFRESH_TOKEN') or '').strip()
        or (_oauth_saved().get('refreshToken') or '').strip()
    )


def _public_base():
    env = (os.environ.get('SALA_EQUACAO_PUBLIC_BASE') or os.environ.get('PUBLIC_API_URL') or '').strip()
    if env:
        return env.rstrip('/')
    default = 'https://smartequa-template-api.onrender.com'
    if not request:
        return default
    root = (request.url_root or '').rstrip('/')
    proto = (request.headers.get('X-Forwarded-Proto') or '').split(',')[0].strip()
    if proto == 'https' and root.startswith('http://'):
        root = 'https://' + root[len('http://'):]
    if 'onrender.com' in root and root.startswith('http://'):
        root = 'https://' + root[len('http://'):]
    return root or default


def _redirect_uri():
    return f'{_public_base()}/api/sala-equacao/drive-oauth/callback'


def _connect_url():
    return f'{_public_base()}/api/sala-equacao/drive-conectar'


def _drive_service():
    refresh = _refresh_token()
    client_id = _client_id()
    client_secret = _client_secret()
    if not (refresh and client_id and client_secret):
        return None, 'drive_oauth_ausente'
    try:
        from google.oauth2.credentials import Credentials
        from googleapiclient.discovery import build
    except ImportError as e:
        return None, str(e)
    creds = Credentials(
        token=None,
        refresh_token=refresh,
        token_uri='https://oauth2.googleapis.com/token',
        client_id=client_id,
        client_secret=client_secret,
        scopes=DRIVE_SCOPES,
    )
    service = build('drive', 'v3', credentials=creds, cache_discovery=False)
    return service, None


def _verify_uid():
    try:
        from firebase_admin import auth
        ok, err = _ensure_firebase()
        if not ok:
            return None, (jsonify({'ok': False, 'error': err or 'Servidor sem Firebase Admin'}), 503)
        header = request.headers.get('Authorization', '')
        if not header.startswith('Bearer '):
            return None, (jsonify({'ok': False, 'error': 'Token ausente'}), 401)
        decoded = auth.verify_id_token(header.replace('Bearer ', '', 1))
        return decoded.get('uid'), None
    except Exception as e:
        return None, (jsonify({'ok': False, 'error': str(e)}), 401)


def _safe_segment(raw, fallback='Pasta'):
    s = (raw or '').strip()
    s = s.replace('/', ' ').replace('\\', ' ').replace('\n', ' ')
    s = re.sub(r'\s+', ' ', s).strip(' .')
    if not s:
        s = fallback
    return s[:120]


def _safe_filename(name):
    s = _safe_segment(name, 'documento')
    s = re.sub(r'[<>:"|?*]', '_', s)
    return s or 'documento'


def _escape_q(name):
    return name.replace('\\', '\\\\').replace("'", "\\'")


def _find_or_create_folder(service, name, parent_id):
    q = (
        f"name = '{_escape_q(name)}' and '{parent_id}' in parents "
        f"and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    )
    res = service.files().list(
        q=q,
        fields='files(id,name)',
        pageSize=5,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = res.get('files') or []
    if files:
        return files[0]['id']
    created = service.files().create(
        body={
            'name': name,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parent_id],
        },
        fields='id',
        supportsAllDrives=True,
    ).execute()
    return created['id']


def _ensure_path(service, segments, root_id):
    parent = root_id
    for seg in segments:
        parent = _find_or_create_folder(service, seg, parent)
    return parent


def _make_file_readable_by_link(service, file_id):
    """Só o arquivo, nunca a pasta. Aluno não lista o Drive."""
    try:
        service.permissions().create(
            fileId=file_id,
            body={'type': 'anyone', 'role': 'reader'},
            fields='id',
            supportsAllDrives=True,
        ).execute()
    except Exception:
        pass


def _file_url(file_id):
    return f'{_public_base()}/api/sala-equacao/files/{file_id}'


def _parse_meta(form):
    ano = _safe_segment(form.get('ano'), str(__import__('datetime').datetime.now().year))
    turma = _safe_segment(form.get('turma'), 'Turma')
    disciplina = _safe_segment(form.get('disciplina'), 'Disciplina')
    bimestre_raw = form.get('bimestre') or '1'
    try:
        bimestre_n = max(1, min(4, int(re.sub(r'\D', '', str(bimestre_raw)) or '1')))
    except (TypeError, ValueError):
        bimestre_n = 1
    categoria = _safe_segment(form.get('categoria'), 'Aulas')
    categoria = 'Entregas' if categoria.lower() == 'entregas' else 'Aulas'
    aluno = (form.get('alunoNome') or '').strip()
    segments = [ano, turma, disciplina, f'{bimestre_n}º bimestre', categoria]
    if categoria == 'Entregas' and aluno:
        segments.append(_safe_segment(aluno, 'Aluno'))
    return segments


def _verify_can_write_folder(service):
    folder_id = _folder_id()
    meta = service.files().get(
        fileId=folder_id,
        fields='id,name,mimeType,capabilities',
        supportsAllDrives=True,
    ).execute()
    caps = meta.get('capabilities') or {}
    if caps.get('canAddChildren') is False:
        raise RuntimeError('Esta conta não pode criar arquivos nessa pasta. Entre com o Gmail dono dela.')
    about = service.about().get(fields='user').execute()
    email = ((about.get('user') or {}).get('emailAddress')) or ''
    return meta.get('name') or folder_id, email


def _upload_drive_oauth(data, filename, mime, segments):
    service, err = _drive_service()
    if service is None:
        return None, err
    from googleapiclient.http import MediaIoBaseUpload
    parent_id = _ensure_path(service, segments, _folder_id())
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=False)
    created = service.files().create(
        body={'name': filename, 'parents': [parent_id]},
        media_body=media,
        fields='id,name,mimeType',
        supportsAllDrives=True,
    ).execute()
    file_id = created['id']
    _make_file_readable_by_link(service, file_id)
    return {
        'fileId': file_id,
        'fileName': created.get('name') or filename,
        'mimeType': created.get('mimeType') or mime,
        'url': f'https://drive.google.com/file/d/{file_id}/preview',
        'viewUrl': f'https://drive.google.com/file/d/{file_id}/view',
        'downloadUrl': f'https://drive.google.com/uc?export=download&id={file_id}',
        'path': ' / '.join(segments),
        'storage': 'gdrive',
    }, None


def register_gdrive_sala_routes(app):
    app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_BYTES

    @app.route('/api/sala-equacao/drive-conectar', methods=['GET'])
    def sala_drive_conectar():
        drive, _ = _drive_service()
        email = (_oauth_saved().get('email') or '') if drive else ''
        return render_template_string(
            CONNECT_PAGE,
            connected=bool(drive),
            email=email,
            has_client=bool(_client_id() and _client_secret()),
            redirect_uri=_redirect_uri(),
            folder_id=_folder_id(),
            storage='gdrive-oauth' if drive else 'precisa autorizar',
            error=request.args.get('erro') or '',
        )

    @app.route('/api/sala-equacao/drive-oauth/credentials', methods=['POST'])
    def sala_drive_save_credentials():
        client_id = (request.form.get('client_id') or '').strip()
        client_secret = (request.form.get('client_secret') or '').strip()
        if not client_id or not client_secret:
            return redirect('/api/sala-equacao/drive-conectar?erro=Cole+o+ID+e+o+segredo')
        saved = _oauth_saved()
        saved['clientId'] = client_id
        saved['clientSecret'] = client_secret
        ok, err = _rtdb_set(RTDB_OAUTH_PATH, saved)
        if not ok:
            return redirect('/api/sala-equacao/drive-conectar?erro=' + (err or 'falha'))
        return redirect('/api/sala-equacao/drive-oauth/start')

    @app.route('/api/sala-equacao/drive-oauth/start', methods=['GET'])
    def sala_drive_oauth_start():
        if not (_client_id() and _client_secret()):
            return redirect('/api/sala-equacao/drive-conectar?erro=Falta+ID+e+segredo+do+cliente')
        state = secrets.token_urlsafe(24)
        _rtdb_set(RTDB_OAUTH_STATE, {'state': state, 'at': int(time.time())})
        params = {
            'client_id': _client_id(),
            'redirect_uri': _redirect_uri(),
            'response_type': 'code',
            'scope': ' '.join(DRIVE_SCOPES),
            'access_type': 'offline',
            'prompt': 'consent',
            'include_granted_scopes': 'true',
            'state': state,
        }
        return redirect('https://accounts.google.com/o/oauth2/v2/auth?' + urlencode(params))

    @app.route('/api/sala-equacao/drive-oauth/callback', methods=['GET'])
    def sala_drive_oauth_callback():
        err = request.args.get('error')
        if err:
            return redirect('/api/sala-equacao/drive-conectar?erro=' + err)
        code = request.args.get('code') or ''
        state = request.args.get('state') or ''
        saved_state = _rtdb_get(RTDB_OAUTH_STATE) or {}
        if not code or state != (saved_state.get('state') or ''):
            return redirect('/api/sala-equacao/drive-conectar?erro=Estado+OAuth+invalido.+Tente+de+novo.')
        token_res = http_requests.post(
            'https://oauth2.googleapis.com/token',
            data={
                'code': code,
                'client_id': _client_id(),
                'client_secret': _client_secret(),
                'redirect_uri': _redirect_uri(),
                'grant_type': 'authorization_code',
            },
            timeout=30,
        )
        payload = token_res.json() if token_res.content else {}
        refresh = (payload.get('refresh_token') or '').strip()
        if not refresh:
            return redirect(
                '/api/sala-equacao/drive-conectar?erro='
                'Google+nao+devolveu+refresh_token.+Revogue+o+acesso+em+'
                'myaccount.google.com/permissions+e+autorize+de+novo.'
            )
        saved = _oauth_saved()
        saved['clientId'] = _client_id()
        saved['clientSecret'] = _client_secret()
        saved['refreshToken'] = refresh
        _rtdb_set(RTDB_OAUTH_PATH, saved)
        try:
            service, derr = _drive_service()
            if service is None:
                raise RuntimeError(derr or 'Drive indisponível')
            _name, email = _verify_can_write_folder(service)
            saved['email'] = email
            saved['connectedAtMillis'] = int(time.time() * 1000)
            _rtdb_set(RTDB_OAUTH_PATH, saved)
        except Exception as e:
            saved.pop('refreshToken', None)
            _rtdb_set(RTDB_OAUTH_PATH, saved)
            return redirect(
                '/api/sala-equacao/drive-conectar?erro='
                'Entre+com+o+Gmail+dono+da+pasta.+Detalhe:+' + str(e)[:180]
            )
        return redirect('/api/sala-equacao/drive-conectar')

    @app.route('/api/sala-equacao/drive-status', methods=['GET'])
    def sala_drive_status():
        drive, _ = _drive_service()
        email = (_oauth_saved().get('email') or '') if drive else ''
        return jsonify({
            'ok': bool(drive),
            'storage': 'gdrive-oauth' if drive else 'nao-autorizado',
            'folderId': _folder_id(),
            'email': email,
            'connectUrl': _connect_url(),
            'hint': (
                'Pasta restrita; subpastas 2026/turma/... criadas na publicação. '
                'Aluno não vê a pasta, só o arquivo da atividade.'
            ),
        })

    @app.route('/api/sala-equacao/files/<file_id>', methods=['GET', 'HEAD'])
    def sala_serve_file(file_id):
        rec = _load_index().get(file_id)
        if not rec:
            return jsonify({'ok': False, 'error': 'Arquivo não encontrado'}), 404
        path = _upload_root() / rec['relPath']
        if not path.exists():
            return jsonify({'ok': False, 'error': 'Arquivo ausente no disco'}), 404
        return send_file(
            path,
            mimetype=rec.get('mimeType') or 'application/octet-stream',
            as_attachment=request.args.get('download') == '1',
            download_name=rec.get('fileName') or path.name,
            max_age=86400,
        )

    @app.route('/api/sala-equacao/upload', methods=['POST'])
    def sala_drive_upload():
        uid, err = _verify_uid()
        if err is not None:
            return err

        uploaded = request.files.get('file')
        if uploaded is None or not uploaded.filename:
            return jsonify({'ok': False, 'error': 'Arquivo ausente'}), 400

        data = uploaded.read()
        if not data:
            return jsonify({'ok': False, 'error': 'Arquivo vazio'}), 400
        if len(data) > MAX_FILE_BYTES:
            return jsonify({'ok': False, 'error': 'Arquivo acima de 50 MB'}), 413

        segments = _parse_meta(request.form)
        filename = _safe_filename(uploaded.filename)
        mime = uploaded.mimetype or 'application/octet-stream'

        drive_payload, drive_err = (None, None)
        try:
            drive_payload, drive_err = _upload_drive_oauth(data, filename, mime, segments)
        except Exception as e:
            drive_err = str(e)
            drive_payload = None

        if drive_payload:
            drive_payload['ok'] = True
            drive_payload['uploadedByUid'] = uid
            return jsonify(drive_payload)

        msg = (
            'Falta autorizar o Google Drive uma vez (Gmail dono da pasta).'
            if drive_err == 'drive_oauth_ausente'
            else (drive_err or 'Falha no Google Drive')
        )
        return jsonify({
            'ok': False,
            'error': msg,
            'code': 'drive_not_connected' if drive_err == 'drive_oauth_ausente' else 'drive_upload_failed',
            'connectUrl': _connect_url(),
        }), 503
