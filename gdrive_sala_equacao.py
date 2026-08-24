# -*- coding: utf-8 -*-
"""
Upload de documentos da Sala Equação para uma pasta do Google Drive.

Pasta raiz (pública / compartilhada):
  https://drive.google.com/drive/folders/1RCciGsZ3yMmveqCeT7ih7Pi1cjVNTuPi

Subpastas: {ano} / {turma} / {disciplina} / {Nº bimestre} / Aulas|Entregas [/ aluno]

Usa FIREBASE_SERVICE_ACCOUNT_JSON (mesma conta do Admin).
A pasta do Drive precisa ser compartilhada como EDITOR com o e-mail dessa conta.
"""
import io
import json
import os
import re

from flask import jsonify, request

DEFAULT_FOLDER_ID = '1RCciGsZ3yMmveqCeT7ih7Pi1cjVNTuPi'
MAX_FILE_BYTES = 50 * 1024 * 1024
DRIVE_SCOPES = ['https://www.googleapis.com/auth/drive']


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
    info = _sa_info()
    if not info:
        return None, 'Servidor sem FIREBASE_SERVICE_ACCOUNT_JSON'
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
    except ImportError as e:
        return None, f'Dependência Google Drive ausente: {e}'
    creds = service_account.Credentials.from_service_account_info(info, scopes=DRIVE_SCOPES)
    service = build('drive', 'v3', credentials=creds, cache_discovery=False)
    return service, None


def _verify_uid():
    try:
        import firebase_admin
        from firebase_admin import auth, credentials, db as fb_db
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


def register_gdrive_sala_routes(app):
    app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_BYTES

    @app.route('/api/sala-equacao/drive-status', methods=['GET'])
    def sala_drive_status():
        info = _sa_info() or {}
        email = info.get('client_email')
        folder_id = _folder_id()
        service, err = _drive_service()
        reachable = False
        detail = err
        if service:
            try:
                service.files().get(
                    fileId=folder_id,
                    fields='id,name',
                    supportsAllDrives=True,
                ).execute()
                reachable = True
                detail = None
            except Exception as e:
                detail = str(e)
        return jsonify({
            'ok': reachable,
            'folderId': folder_id,
            'folderUrl': f'https://drive.google.com/drive/folders/{folder_id}',
            'serviceAccountEmail': email,
            'error': detail,
            'hint': None if reachable else (
                f'Compartilhe a pasta do Drive como Editor com {email}'
                if email else 'Configure FIREBASE_SERVICE_ACCOUNT_JSON no servidor'
            ),
        })

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

        ano = _safe_segment(request.form.get('ano'), str(__import__('datetime').datetime.now().year))
        turma = _safe_segment(request.form.get('turma'), 'Turma')
        disciplina = _safe_segment(request.form.get('disciplina'), 'Disciplina')
        bimestre_raw = request.form.get('bimestre') or '1'
        try:
            bimestre_n = max(1, min(4, int(str(bimestre_raw).strip()[0])))
        except (TypeError, ValueError):
            bimestre_n = 1
        categoria = _safe_segment(request.form.get('categoria'), 'Aulas')
        if categoria.lower() not in ('aulas', 'entregas'):
            categoria = 'Aulas'
        categoria = 'Entregas' if categoria.lower() == 'entregas' else 'Aulas'
        aluno = (request.form.get('alunoNome') or '').strip()

        segments = [ano, turma, disciplina, f'{bimestre_n}º bimestre', categoria]
        if categoria == 'Entregas' and aluno:
            segments.append(_safe_segment(aluno, 'Aluno'))

        filename = _safe_segment(uploaded.filename, 'documento')
        mime = uploaded.mimetype or 'application/octet-stream'

        service, derr = _drive_service()
        if service is None:
            return jsonify({'ok': False, 'error': derr}), 503

        folder_id = _folder_id()
        try:
            parent_id = _ensure_path(service, segments, folder_id)
            from googleapiclient.http import MediaIoBaseUpload
            media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=False)
            created = service.files().create(
                body={'name': filename, 'parents': [parent_id]},
                media_body=media,
                fields='id,name,mimeType,webViewLink,webContentLink',
                supportsAllDrives=True,
            ).execute()
            file_id = created['id']
            _make_public(service, file_id)
        except Exception as e:
            msg = str(e)
            info = _sa_info() or {}
            email = info.get('client_email') or ''
            if 'File not found' in msg or 'notFound' in msg or '404' in msg:
                msg = (
                    'A pasta do Drive não está acessível para o servidor. '
                    f'Compartilhe como Editor com {email}'
                )
            return jsonify({'ok': False, 'error': msg, 'serviceAccountEmail': email}), 502

        preview = f'https://drive.google.com/file/d/{file_id}/preview'
        view = f'https://drive.google.com/file/d/{file_id}/view'
        download = f'https://drive.google.com/uc?export=download&id={file_id}'
        return jsonify({
            'ok': True,
            'fileId': file_id,
            'fileName': created.get('name') or filename,
            'mimeType': created.get('mimeType') or mime,
            'url': preview,
            'viewUrl': view,
            'downloadUrl': download,
            'path': ' / '.join(segments),
            'uploadedByUid': uid,
        })
