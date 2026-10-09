"""Proteção das rotas do Assistente de Redação (/api/transcrever e /api/corrigir).

- Exige token do Firebase (Authorization: Bearer <idToken>) de um usuário cadastrado
  (aluno, funcionário, admin; professor só se aprovado).
- Limite simples por usuário (ou por IP, sem login): ESSAY_RATE_MAX pedidos em
  ESSAY_RATE_WINDOW_S segundos (padrão 20 em 10 min) por processo.
- ESSAY_AUTH_ENFORCE: '1' bloqueia (401/429); '0' só registra no log e deixa passar.
  A variável de ambiente, se existir, tem prioridade sobre ENFORCE_DEFAULT.
"""
import os
import threading
import time
from collections import deque

from flask import jsonify, request

ENFORCE_DEFAULT = '0'

RATE_MAX = int(os.environ.get('ESSAY_RATE_MAX', '20'))
RATE_WINDOW = int(os.environ.get('ESSAY_RATE_WINDOW_S', '600'))

_hits = {}
_lock = threading.Lock()


def _enforced():
    return (os.environ.get('ESSAY_AUTH_ENFORCE') or ENFORCE_DEFAULT).strip() == '1'


def _client_ip():
    xff = request.headers.get('X-Forwarded-For', '')
    return (xff.split(',')[0].strip() if xff else '') or request.remote_addr or '?'


def _rate_ok(key):
    now = time.time()
    with _lock:
        dq = _hits.setdefault(key, deque())
        while dq and now - dq[0] > RATE_WINDOW:
            dq.popleft()
        if len(dq) >= RATE_MAX:
            return False
        dq.append(now)
        if len(_hits) > 5000:
            for k in [k for k, v in _hits.items() if not v]:
                _hits.pop(k, None)
        return True


def _check_user(token):
    """Devolve (uid, None) se o token é válido e o usuário pode usar; senão (None, motivo)."""
    try:
        from firebase_admin_routes import _init_firebase
        if not _init_firebase():
            return None, 'Servidor sem Firebase Admin'
        from firebase_admin import auth, db
        try:
            uid = auth.verify_id_token(token).get('uid')
        except Exception:
            return None, 'Token inválido ou expirado'
        if not uid:
            return None, 'Token sem usuário'
        # Lê só os campos necessários (não o usuário inteiro).
        base = db.reference(f'usuarios/{uid}')
        role = str(base.child('role').get() or '')
        if not role:
            return None, 'Usuário não cadastrado'
        if role == 'professor':
            if base.child('isApproved').get() is not True and base.child('isProfAdmin').get() is not True:
                return None, 'Professor ainda não aprovado'
        return uid, None
    except Exception as e:
        return None, f'Falha ao verificar login: {str(e)[:120]}'


def essay_guard(route):
    """Chamar no início da rota. Devolve uma resposta de erro ou None (pode seguir)."""
    enforce = _enforced()
    hdr = request.headers.get('Authorization', '')
    token = hdr[7:].strip() if hdr.startswith('Bearer ') else ''
    if token:
        uid, err = _check_user(token)
    else:
        uid, err = None, 'Token ausente'
    ip = _client_ip()
    key = ('u:' + uid) if uid else ('ip:' + ip)
    if uid is None:
        if enforce:
            return jsonify({'error': 'Faça login para usar o Assistente de Redação.', 'motivo': err}), 401
        print(f'[essay] {route}: sem login válido ({err}) ip={ip} - permitido (ESSAY_AUTH_ENFORCE=0)')
    if not _rate_ok(key):
        if enforce:
            return jsonify({'error': 'Muitas solicitações seguidas. Aguarde alguns minutos e tente de novo.'}), 429
        print(f'[essay] {route}: limite excedido para {key} - permitido (ESSAY_AUTH_ENFORCE=0)')
    return None
