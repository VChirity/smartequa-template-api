# -*- coding: utf-8 -*-
"""
Rotas PIX Cantina (Mercado Pago) para o app Smart Equação.
NÃO altera nenhuma rota de templates - apenas adiciona endpoints /api/pix/*
Quitação de dívida: confirm-debt e settle-debt-balance (Firebase Admin no servidor).
"""
import json
import os
import uuid
from datetime import datetime, timedelta, timezone

import requests
from flask import request, jsonify


def _get_token():
    return os.environ.get('MERCADO_PAGO_ACCESS_TOKEN')


def _mp_headers():
    token = _get_token()
    if not token:
        return None
    return {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json',
    }


def _get_fb_db():
    try:
        import firebase_admin
        from firebase_admin import credentials, db as fb_db
        try:
            firebase_admin.get_app()
        except ValueError:
            service_account_json = os.environ.get('FIREBASE_SERVICE_ACCOUNT_JSON')
            if service_account_json:
                cred_dict = json.loads(service_account_json)
                cred = credentials.Certificate(cred_dict)
            elif os.path.exists('firebase-service-account.json'):
                cred = credentials.Certificate('firebase-service-account.json')
            else:
                return None
            firebase_admin.initialize_app(cred, {
                'databaseURL': 'https://equa-sec-apk-default-rtdb.firebaseio.com',
            })
        return fb_db
    except Exception:
        return None


def _verify_bearer_uid():
    fb_db = _get_fb_db()
    if fb_db is None:
        return None, (jsonify({'ok': False, 'error': 'Servidor sem Firebase Admin'}), 503)
    try:
        from firebase_admin import auth
        auth_header = request.headers.get('Authorization', '')
        if not auth_header.startswith('Bearer '):
            return None, (jsonify({'ok': False, 'error': 'Token ausente'}), 401)
        id_token = auth_header.replace('Bearer ', '')
        decoded = auth.verify_id_token(id_token)
        return decoded['uid'], None
    except Exception as e:
        return None, (jsonify({'ok': False, 'error': str(e)}), 401)


def register_pix_routes(app):
    """Registra as rotas PIX no app Flask. Chamar passando o app."""

    @app.route('/api/pix/check', methods=['GET'])
    def pix_check():
        """Teste de conexão com Mercado Pago (token válido)."""
        headers = _mp_headers()
        if not headers:
            return jsonify({'status': 'ERRO', 'detalhe': 'Token ausente'}), 503
        try:
            r = requests.get(
                'https://api.mercadopago.com/users/me',
                headers=headers,
                timeout=10,
            )
            r.raise_for_status()
            data = r.json()
            return jsonify({
                'status': 'ONLINE',
                'integracao': 'Mercado Pago',
                'loja': data.get('site_id', ''),
                'mensagem': 'Token válido e comunicando com Mercado Pago!',
            })
        except Exception as e:
            return jsonify({'status': 'ERRO', 'detalhe': str(e)}), 500

    @app.route('/api/pix/create', methods=['POST'])
    def pix_create():
        """Cria cobrança PIX. Body: { "valor": number, "descricao": string }."""
        headers = _mp_headers()
        if not headers:
            return jsonify({'error': 'Servidor não configurado', 'detail': 'Token MP ausente'}), 503

        try:
            body = request.get_json() or {}
            valor = body.get('valor')
            descricao = body.get('descricao') or 'Pedido Cantina'

            try:
                transaction_amount = float(valor)
            except (TypeError, ValueError):
                return jsonify({'error': 'Valor inválido'}), 400

            if transaction_amount <= 0:
                return jsonify({'error': 'Valor inválido'}), 400

            payload = {
                'transaction_amount': transaction_amount,
                'description': descricao,
                'payment_method_id': 'pix',
                'payer': {'email': 'cliente@cantina.com'},
            }
            # Evita PIX com validade curta por padrão do Mercado Pago.
            payload['date_of_expiration'] = (
                datetime.now(timezone.utc) + timedelta(days=7)
            ).strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3] + 'Z'

            headers = dict(headers)
            headers['X-Idempotency-Key'] = str(uuid.uuid4())

            r = requests.post(
                'https://api.mercadopago.com/v1/payments',
                json=payload,
                headers=headers,
                timeout=15,
            )
            r.raise_for_status()
            payment = r.json()

            poi = payment.get('point_of_interaction') or {}
            tx_data = poi.get('transaction_data') or {}
            qr_code = tx_data.get('qr_code') or ''
            qr_code_base64 = tx_data.get('qr_code_base64')
            payment_id = payment.get('id')

            return jsonify({
                'txid': str(payment_id),
                'copiaECola': qr_code,
                'qrCodeBase64': qr_code_base64,
            })
        except requests.exceptions.HTTPError as e:
            err_detail = e.response.text if e.response else str(e)
            try:
                err_detail = e.response.json().get('message', err_detail)
            except Exception:
                pass
            return jsonify({'error': 'Erro ao criar cobrança', 'detail': err_detail}), 500
        except Exception as e:
            return jsonify({'error': 'Erro ao criar cobrança', 'detail': str(e)}), 500

    @app.route('/api/pix/status/<txid>', methods=['GET'])
    def pix_status(txid):
        """Consulta status do pagamento PIX. Retorna { "pago": bool, "status_mp": string, "pagante": opcional }."""
        headers = _mp_headers()
        if not headers:
            return jsonify({'pago': False}), 503

        try:
            r = requests.get(
                f'https://api.mercadopago.com/v1/payments/{txid}',
                headers=headers,
                timeout=10,
            )
            if r.status_code == 404:
                return jsonify({'pago': False})
            r.raise_for_status()
            data = r.json()
            status = data.get('status', '')
            pagante = ''
            payer = data.get('payer') or {}
            first = (payer.get('first_name') or '').strip()
            last = (payer.get('last_name') or '').strip()
            if first or last:
                pagante = f'{first} {last}'.strip()
            result = {'pago': status == 'approved', 'status_mp': status}
            if pagante:
                result['pagante'] = pagante
            return jsonify(result)
        except Exception:
            return jsonify({'pago': False}), 500

    @app.route('/api/pix/confirm-debt', methods=['POST'])
    def pix_confirm_debt():
        """
        Quita dívida da cantina após PIX aprovado (Mercado Pago).
        Body JSON: { "txid": string }
        Header: Authorization: Bearer <Firebase ID token do aluno>
        """
        uid, err = _verify_bearer_uid()
        if err:
            return err
        fb_db = _get_fb_db()
        if fb_db is None:
            return jsonify({'ok': False, 'error': 'Firebase Admin não configurado'}), 503
        body = request.get_json() or {}
        txid = (body.get('txid') or '').strip()
        if not txid:
            return jsonify({'ok': False, 'error': 'txid obrigatório'}), 400
        headers = _mp_headers()
        if not headers:
            return jsonify({'ok': False, 'error': 'Mercado Pago não configurado'}), 503
        try:
            r = requests.get(
                f'https://api.mercadopago.com/v1/payments/{txid}',
                headers=headers,
                timeout=15,
            )
            if r.status_code != 200:
                return jsonify({'ok': False, 'error': 'Pagamento não encontrado'}), 400
            payment = r.json()
            if payment.get('status') != 'approved':
                return jsonify({'ok': False, 'error': 'Pagamento ainda não aprovado'}), 400
        except Exception as e:
            return jsonify({'ok': False, 'error': str(e)}), 500

        ref_pix = fb_db.reference(f'cantina_pending_pix/{txid}')
        entry = ref_pix.get()
        if not entry or not isinstance(entry, dict):
            return jsonify({'ok': False, 'error': 'PIX pendente não registrado na nuvem'}), 400
        if entry.get('type') != 'debt':
            return jsonify({'ok': False, 'error': 'Tipo de PIX inválido'}), 400
        entry_uid = (entry.get('userId') or '').strip()
        if entry_uid != uid:
            return jsonify({'ok': False, 'error': 'Usuário não corresponde ao PIX'}), 403

        ref_debt = fb_db.reference(f'cantina_pending_debts/{uid}')
        debt = ref_debt.get()
        if not debt:
            try:
                ref_pix.delete()
            except Exception:
                pass
            return jsonify({'ok': True, 'alreadyCleared': True})

        debt_balance_used = float(entry.get('debtBalanceUsed') or 0)
        if debt_balance_used > 0:
            uref = fb_db.reference(f'usuarios/{uid}')
            snap = uref.get()
            if isinstance(snap, dict):
                current = float(snap.get('cantinaSaldo') or 0)
                nome = (snap.get('cantinaNome') or snap.get('nome') or '').strip()
                new_bal = max(0.0, current - debt_balance_used)
                uref.update({'cantinaSaldo': new_bal, 'cantinaNome': nome or None})

        try:
            ref_debt.delete()
        except Exception:
            pass
        try:
            ref_pix.delete()
        except Exception:
            pass
        return jsonify({'ok': True})

    @app.route('/api/pix/settle-debt-balance', methods=['POST'])
    def pix_settle_debt_balance():
        """
        Quita dívida usando apenas o saldo da carteira (servidor aplica débito e remove bloqueio).
        Header: Authorization: Bearer <Firebase ID token do aluno>
        """
        uid, err = _verify_bearer_uid()
        if err:
            return err
        fb_db = _get_fb_db()
        if fb_db is None:
            return jsonify({'ok': False, 'error': 'Firebase Admin não configurado'}), 503
        ref_debt = fb_db.reference(f'cantina_pending_debts/{uid}')
        debt = ref_debt.get()
        if not debt or not isinstance(debt, dict):
            return jsonify({'ok': False, 'error': 'Sem dívida pendente'}), 400
        total = float(debt.get('totalAmount') or 0)
        paid = float(debt.get('paidAmount') or 0)
        pending = max(0.0, total - paid)
        if pending <= 0.001:
            try:
                ref_debt.delete()
            except Exception:
                pass
            return jsonify({'ok': True})

        uref = fb_db.reference(f'usuarios/{uid}')
        snap = uref.get()
        if not isinstance(snap, dict):
            return jsonify({'ok': False, 'error': 'Usuário não encontrado'}), 400
        balance = float(snap.get('cantinaSaldo') or 0)
        if balance + 1e-6 < pending:
            return jsonify({'ok': False, 'error': 'Saldo insuficiente'}), 400
        nome = (snap.get('cantinaNome') or snap.get('nome') or '').strip()
        new_bal = balance - pending
        uref.update({'cantinaSaldo': new_bal, 'cantinaNome': nome or None})
        try:
            ref_debt.delete()
        except Exception:
            pass
        return jsonify({'ok': True})
