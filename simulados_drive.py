"""Simulados (OMR) – imagem corrigida (overlay + quadro de notas) no Google Drive.

Usa a mesma autorização OAuth do Drive da Sala Equação (gdrive_sala_equacao), mas grava em
<SIMULADOS_DRIVE_PARENT_ID>/Simulados correção/<ano>/<Nº Bimestre>/<turma>/<aluno>.jpg.
Arquivos NÃO são compartilhados por link: só o dono da pasta os abre no Drive; o app lê pela rota proxy (admin).
"""
import io
import os
import re

from flask import jsonify, request, Response

import gdrive_sala_equacao as G

DEFAULT_PARENT_ID = '1iQbHalbZIjKSgr42nNP_CvS23wn1oIw7'
ROOT_NAME = 'Simulados correção'
MAX_BYTES = 3 * 1024 * 1024


def _parent_id():
    return (os.environ.get('SIMULADOS_DRIVE_PARENT_ID') or DEFAULT_PARENT_ID).strip()


def _admin():
    uid, err = G._verify_admin()
    return uid, err


def _find_file(service, name, parent_id):
    q = (f"name = '{G._escape_q(name)}' and '{parent_id}' in parents and trashed = false "
         f"and mimeType != 'application/vnd.google-apps.folder'")
    res = service.files().list(q=q, fields='files(id,name,appProperties)', pageSize=5,
                               supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    return res.get('files') or []


def _trash_simulados_file(service, file_id):
    """Manda para a lixeira do Drive só arquivos criados por Simulados (appProperties.simulados=1)."""
    try:
        meta = service.files().get(fileId=file_id, fields='id,appProperties,trashed', supportsAllDrives=True).execute()
    except Exception as e:
        if 'notFound' in str(e) or '404' in str(e):
            return 'missing'
        raise
    if (meta.get('appProperties') or {}).get('simulados') != '1':
        return 'forbidden'
    if not meta.get('trashed'):
        service.files().update(fileId=file_id, body={'trashed': True}, supportsAllDrives=True).execute()
    return 'trashed'


def register_simulados_drive_routes(app):
    @app.route('/api/simulados/drive-check', methods=['GET'])
    def simulados_drive_check():
        # diagnóstico sem dados sensíveis: só diz se o servidor alcança a pasta-mãe e pode escrever nela
        service, derr = G._drive_service()
        if service is None:
            return jsonify({'ok': False, 'stage': 'oauth', 'error': derr or 'Drive não autorizado'}), 503
        try:
            meta = service.files().get(fileId=_parent_id(), fields='id,trashed,capabilities(canAddChildren)',
                                       supportsAllDrives=True).execute()
            return jsonify({'ok': True, 'parentReachable': True, 'trashed': bool(meta.get('trashed')),
                            'canAddChildren': bool((meta.get('capabilities') or {}).get('canAddChildren'))})
        except Exception as e:
            return jsonify({'ok': False, 'stage': 'parent', 'error': str(e)[:200]}), 500

    @app.route('/api/simulados/correcao/<file_id>', methods=['DELETE'])
    def simulados_correcao_delete(file_id):
        uid, err = _admin()
        if err is not None:
            return err
        if not re.fullmatch(r'[0-9A-Za-z_-]{10,100}', file_id or ''):
            return jsonify({'ok': False, 'error': 'id inválido'}), 400
        service, derr = G._drive_service()
        if service is None:
            return jsonify({'ok': False, 'error': derr or 'Drive não autorizado'}), 503
        try:
            st = _trash_simulados_file(service, file_id)
            if st == 'forbidden':
                return jsonify({'ok': False, 'error': 'Arquivo não é de Simulados'}), 403
            return jsonify({'ok': True, 'status': st})
        except Exception as e:
            return jsonify({'ok': False, 'error': str(e)[:300]}), 500

    @app.route('/api/simulados/correcao-upload', methods=['POST'])
    def simulados_correcao_upload():
        uid, err = _admin()
        if err is not None:
            return err
        up = request.files.get('file')
        if up is None:
            return jsonify({'ok': False, 'error': 'Arquivo ausente'}), 400
        data = up.read()
        if not data or len(data) > MAX_BYTES:
            return jsonify({'ok': False, 'error': 'Imagem vazia ou acima de 3 MB'}), 413
        if not data.startswith(b'\xff\xd8'):
            return jsonify({'ok': False, 'error': 'Só JPEG'}), 415
        f = request.form
        ano = G._safe_segment(f.get('ano'), 'Ano')
        try:
            bim = max(1, min(4, int(re.sub(r'\D', '', str(f.get('bimestre') or '')) or '0')))
        except ValueError:
            bim = 0
        if not bim:
            return jsonify({'ok': False, 'error': 'Bimestre inválido'}), 400
        turma = G._safe_segment(f.get('turma'), 'Turma')
        aluno = G._safe_filename(f.get('aluno') or 'Aluno')
        aluno_id = re.sub(r'[^0-9A-Za-z_-]', '', str(f.get('alunoId') or ''))[:40]
        exam_id = re.sub(r'[^0-9A-Za-z_-]', '', str(f.get('examId') or ''))[:40]
        replace_id = re.sub(r'[^0-9A-Za-z_-]', '', str(f.get('replaceFileId') or ''))[:100]
        service, derr = G._drive_service()
        if service is None:
            return jsonify({'ok': False, 'code': 'drive_not_connected', 'error': derr or 'Drive não autorizado',
                            'connectUrl': G._connect_url()}), 503
        try:
            from googleapiclient.http import MediaIoBaseUpload
            folder = G._ensure_path(service, [ROOT_NAME, ano, f'{bim}º Bimestre', turma], _parent_id())
            name = f'{aluno}.jpg'
            props = {'simulados': '1', 'alunoId': aluno_id, 'examId': exam_id}
            media = MediaIoBaseUpload(io.BytesIO(data), mimetype='image/jpeg', resumable=False)
            same = [x for x in _find_file(service, name, folder)]
            mine = [x for x in same if (x.get('appProperties') or {}).get('alunoId') in (aluno_id, None, '')]
            if mine:
                fid = mine[0]['id']
                service.files().update(fileId=fid, media_body=media, body={'appProperties': props},
                                       fields='id', supportsAllDrives=True).execute()
            else:
                if same:  # homônimo de outro aluno na mesma turma
                    name = f'{aluno} ({aluno_id}).jpg'
                fid = service.files().create(body={'name': name, 'parents': [folder], 'appProperties': props},
                                             media_body=media, fields='id', supportsAllDrives=True).execute()['id']
            if replace_id and replace_id != fid:
                _trash_simulados_file(service, replace_id)
            return jsonify({'ok': True, 'fileId': fid, 'name': name,
                            'viewUrl': f'https://drive.google.com/file/d/{fid}/view',
                            'path': f'{ROOT_NAME} / {ano} / {bim}º Bimestre / {turma}'})
        except Exception as e:
            return jsonify({'ok': False, 'error': str(e)[:300]}), 500

    @app.route('/api/simulados/correcao/<file_id>', methods=['GET'])
    def simulados_correcao_get(file_id):
        uid, err = _admin()
        if err is not None:
            return err
        if not re.fullmatch(r'[0-9A-Za-z_-]{10,100}', file_id or ''):
            return jsonify({'ok': False, 'error': 'id inválido'}), 400
        service, derr = G._drive_service()
        if service is None:
            return jsonify({'ok': False, 'error': derr or 'Drive não autorizado'}), 503
        try:
            meta = service.files().get(fileId=file_id, fields='id,mimeType,appProperties', supportsAllDrives=True).execute()
            if (meta.get('appProperties') or {}).get('simulados') != '1':
                return jsonify({'ok': False, 'error': 'Arquivo não é de Simulados'}), 403
            data = service.files().get_media(fileId=file_id, supportsAllDrives=True).execute()
            return Response(data, mimetype='image/jpeg', headers={'Cache-Control': 'private, max-age=3600'})
        except Exception as e:
            return jsonify({'ok': False, 'error': str(e)[:200]}), 404
