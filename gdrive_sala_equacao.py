# -*- coding: utf-8 -*-
"""
Arquivos da Sala Equação.

O Google Drive com conta de serviço NÃO funciona em pasta comum do Gmail:
a service account não tem cota ("Service Accounts do not have storage quota").

Solução: gravar no disco do servidor (NAS persiste; Render é backup).
Opcional: Drive via OAuth do dono da pasta (GOOGLE_DRIVE_REFRESH_TOKEN).
"""
import io
import json
import os
import re
import time
import uuid
from pathlib import Path

from flask import jsonify, request, send_file

DEFAULT_FOLDER_ID = '1RCciGsZ3yMmveqCeT7ih7Pi1cjVNTuPi'
MAX_FILE_BYTES = 50 * 1024 * 1024
DRIVE_SCOPES = ['https://www.googleapis.com/auth/drive']


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


def _save_index(idx):
    _index_path().write_text(json.dumps(idx, ensure_ascii=False, indent=2), encoding='utf-8')


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


def _drive_service():
    """Só usa OAuth do usuário. Service account não tem cota no Drive pessoal."""
    refresh = (os.environ.get('GOOGLE_DRIVE_REFRESH_TOKEN') or '').strip()
    client_id = (os.environ.get('GOOGLE_DRIVE_CLIENT_ID') or '').strip()
    client_secret = (os.environ.get('GOOGLE_DRIVE_CLIENT_SECRET') or '').strip()
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
        import firebase_admin
        from firebase_admin import auth, credentials
        try:
            firebase_admin.get_app()
        except ValueError:
            info = _sa_info()
            if not info:
                return None, (jsonify({'ok': False, 'error': 'Servidor sem Firebase Admin'}), 503)
            cred = credentials.Certificate(info)
            firebase_admin.initialize_app(cred, {
                'databaseURL': 'https://equa-sec-apk-default-rtdb.firebaseio.com',
            })
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


def _disk_folder_name(name):
    s = _safe_segment(name, 'pasta')
    s = re.sub(r'[^\w\-À-ÿ. ]+', '_', s, flags=re.UNICODE)
    s = re.sub(r'\s+', '_', s).strip('_')
    return (s or 'pasta')[:80]


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


def _make_public(service, file_id):
    try:
        service.permissions().create(
            fileId=file_id,
            body={'type': 'anyone', 'role': 'reader'},
            fields='id',
            supportsAllDrives=True,
        ).execute()
    except Exception:
        pass


def _public_base():
    env = (os.environ.get('SALA_EQUACAO_PUBLIC_BASE') or os.environ.get('PUBLIC_API_URL') or '').strip()
    if env:
        return env.rstrip('/')
    return request.url_root.rstrip('/')


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


def _save_to_disk(data, filename, mime, segments, uid):
    file_id = uuid.uuid4().hex
    rel_parts = [_disk_folder_name(s) for s in segments]
    dest_dir = _upload_root().joinpath(*rel_parts)
    dest_dir.mkdir(parents=True, exist_ok=True)
    stored_name = f'{file_id}_{_safe_filename(filename)}'
    dest = dest_dir / stored_name
    dest.write_bytes(data)

    rec = {
        'id': file_id,
        'fileName': filename,
        'mimeType': mime,
        'relPath': str(Path(*rel_parts) / stored_name).replace('\\', '/'),
        'pathLabel': ' / '.join(segments),
        'uploadedByUid': uid,
        'uploadedAtMillis': int(time.time() * 1000),
    }
    idx = _load_index()
    idx[file_id] = rec
    _save_index(idx)
    return rec


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
    _make_public(service, file_id)
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

    @app.route('/api/sala-equacao/drive-status', methods=['GET'])
    def sala_drive_status():
        drive, _ = _drive_service()
        return jsonify({
            'ok': True,
            'storage': 'gdrive-oauth' if drive else 'disk',
            'folderId': _folder_id(),
            'folderUrl': f'https://drive.google.com/drive/folders/{_folder_id()}',
            'hint': (
                'Arquivos gravados no servidor (sem cota do Google). '
                'Drive opcional: defina GOOGLE_DRIVE_REFRESH_TOKEN do dono da pasta.'
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

        rec = _save_to_disk(data, filename, mime, segments, uid)
        url = _file_url(rec['id'])
        return jsonify({
            'ok': True,
            'fileId': rec['id'],
            'fileName': rec['fileName'],
            'mimeType': rec['mimeType'],
            'url': url,
            'viewUrl': url,
            'downloadUrl': f'{url}?download=1',
            'path': rec['pathLabel'],
            'storage': 'disk',
            'uploadedByUid': uid,
        })
