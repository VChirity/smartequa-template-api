# -*- coding: utf-8 -*-
"""
Rotas PIX Cantina (Mercado Pago) para o app Smart Equa├º├úo.
N├âO altera nenhuma rota de templates - apenas adiciona endpoints /api/pix/*
Quita├º├úo de d├¡vida: confirm-debt e settle-debt-balance (Firebase Admin no servidor).
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


def _webhook_public_url():
    return (
        os.environ.get('PIX_WEBHOOK_URL')
        or 'https://smartequa-template-api.onrender.com/api/pix/webhook'
    ).rstrip('/')


def _mp_get_payment(txid):
    headers = _mp_headers()
    if not headers:
        return None
    r = requests.get(
        f'https://api.mercadopago.com/v1/payments/{txid}',
        headers=headers,
        timeout=15,
    )
    if r.status_code != 200:
        return None
    return r.json()


def _payer_name(payment):
    payer = (payment or {}).get('payer') or {}
    first = (payer.get('first_name') or '').strip()
    last = (payer.get('last_name') or '').strip()
    return f'{first} {last}'.strip()


def _load_pending_entry(fb_db, txid, uid=None):
    entry = fb_db.reference(f'cantina_pending_pix/{txid}').get()
    if isinstance(entry, dict):
        return entry
    if uid:
        for path in (
            f'cantina_pending_pix_by_user/{uid}/{txid}',
            f'usuarios/{uid}/cantinaPendingPix/{txid}',
        ):
            alt = fb_db.reference(path).get()
            if isinstance(alt, dict):
                return alt
        return None
    by_user = fb_db.reference('cantina_pending_pix_by_user').get() or {}
    if isinstance(by_user, dict):
        for _uid, txs in by_user.items():
            if isinstance(txs, dict) and txid in txs and isinstance(txs[txid], dict):
                return txs[txid]
    return None


def _cleanup_pending(fb_db, txid, uid):
    paths = [f'cantina_pending_pix/{txid}']
    if uid:
        paths.append(f'cantina_pending_pix_by_user/{uid}/{txid}')
        paths.append(f'usuarios/{uid}/cantinaPendingPix/{txid}')
    for p in paths:
        try:
            fb_db.reference(p).delete()
        except Exception:
            pass


def _purchase_has_txid(fb_db, uid, txid):
    data = fb_db.reference(f'usuarios/{uid}/cantinaPurchases').get()
    if not isinstance(data, dict):
        return False
    expected = {f'pix_balance_{txid}', f'pix_cart_{txid}'}
    if any(k in data for k in expected):
        return True
    for rec in data.values():
        if isinstance(rec, dict) and str(rec.get('pixTransactionId') or '') == str(txid):
            return True
    return False


def _write_purchase(fb_db, uid, purchase_id, data):
    data = dict(data)
    data['id'] = purchase_id
    data['userId'] = uid
    fb_db.reference(f'usuarios/{uid}/cantinaPurchases/{purchase_id}').set(data)
    try:
        fb_db.reference(f'cantina_vendas_index/{purchase_id}').set(data)
    except Exception:
        pass


def _dec_stock(fb_db, items):
    if not items:
        return
    for item in items:
        if not isinstance(item, dict):
            continue
        pid = str(item.get('productId') or '').strip()
        if not pid or pid == 'custom':
            continue
        try:
            qty = int(item.get('quantity') or 1)
        except (TypeError, ValueError):
            qty = 1
        if qty <= 0:
            continue
        ref = fb_db.reference(f'cantina_stock/{pid}')

        def txn(cur):
            try:
                n = int(cur or 0)
            except (TypeError, ValueError):
                n = 0
            return max(0, n - qty)

        try:
            ref.transaction(txn)
        except Exception:
            pass


def _begin_settle_lock(fb_db, txid):
    ref = fb_db.reference(f'cantina_pix_settled/{txid}')
    started = {'ok': False}
    now = datetime.now(timezone.utc)

    def txn(cur):
        if isinstance(cur, dict):
            st = cur.get('status')
            if st == 'settled':
                return cur
            if st == 'processing':
                raw = str(cur.get('at') or '')
                try:
                    at = datetime.fromisoformat(raw.replace('Z', '+00:00'))
                    if at.tzinfo is None:
                        at = at.replace(tzinfo=timezone.utc)
                    if now - at < timedelta(seconds=120):
                        return cur
                except Exception:
                    return cur
        started['ok'] = True
        return {'status': 'processing', 'at': now.isoformat()}

    try:
        ref.transaction(txn)
    except Exception:
        return False
    return started['ok']


def _finish_settle_lock(fb_db, txid, kind):
    try:
        fb_db.reference(f'cantina_pix_settled/{txid}').set({
            'status': 'settled',
            'type': kind,
            'at': datetime.now(timezone.utc).isoformat(),
        })
    except Exception:
        pass


def _abort_settle_lock(fb_db, txid):
    try:
        fb_db.reference(f'cantina_pix_settled/{txid}').delete()
    except Exception:
        pass


def _settle_approved_pix(txid, payment=None):
    """Confirma PIX pago no Mercado Pago e aplica no Firebase (app pode estar fechado)."""
    fb_db = _get_fb_db()
    if fb_db is None:
        return False, 'no_firebase'
    txid = str(txid or '').strip()
    if not txid:
        return False, 'no_txid'
    if payment is None:
        payment = _mp_get_payment(txid)
    if not payment or payment.get('status') != 'approved':
        return False, 'not_approved'
    entry = _load_pending_entry(fb_db, txid)
    if not isinstance(entry, dict):
        return True, 'no_pending'
    uid = str(entry.get('userId') or '').strip()
    kind = str(entry.get('type') or '').strip() or 'balance'
    if not uid:
        return False, 'no_user'
    if _purchase_has_txid(fb_db, uid, txid) and kind != 'debt':
        _cleanup_pending(fb_db, txid, uid)
        _finish_settle_lock(fb_db, txid, kind)
        return True, 'already'
    if not _begin_settle_lock(fb_db, txid):
        return True, 'locked'
    try:
        pagante = _payer_name(payment)
        created = str(entry.get('createdAt') or datetime.now(timezone.utc).isoformat())
        amount = float(entry.get('amount') or payment.get('transaction_amount') or 0)
        buyer = str(entry.get('buyerName') or entry.get('sourceProductName') or '').strip()
        if kind == 'debt':
            ref_debt = fb_db.reference(f'cantina_pending_debts/{uid}')
            debt = ref_debt.get()
            if debt:
                used = float(entry.get('debtBalanceUsed') or 0)
                if used > 0:
                    uref = fb_db.reference(f'usuarios/{uid}')
                    snap = uref.get()
                    if isinstance(snap, dict):
                        current = float(snap.get('cantinaSaldo') or 0)
                        nome = (snap.get('cantinaNome') or snap.get('nome') or '').strip()
                        uref.update({
                            'cantinaSaldo': max(0.0, current - used),
                            'cantinaNome': nome or None,
                        })
                try:
                    ref_debt.delete()
                except Exception:
                    pass
            _cleanup_pending(fb_db, txid, uid)
            _finish_settle_lock(fb_db, txid, 'debt')
            return True, 'debt'
        if kind == 'balance':
            pid = f'pix_balance_{txid}'
            existing = fb_db.reference(f'usuarios/{uid}/cantinaPurchases/{pid}').get()
            if not existing:
                uref = fb_db.reference(f'usuarios/{uid}')
                snap = uref.get()
                snap = snap if isinstance(snap, dict) else {}
                current = float(snap.get('cantinaSaldo') or 0)
                nome = (snap.get('cantinaNome') or snap.get('nome') or buyer).strip()
                uref.update({'cantinaSaldo': current + amount, 'cantinaNome': nome or None})
                _write_purchase(fb_db, uid, pid, {
                    'dateTime': created,
                    'items': [],
                    'totalPaid': amount,
                    'paymentMethod': 'pix',
                    'studentName': buyer or nome,
                    'pixTransactionId': txid,
                    'pixPayerName': pagante,
                    'purchaserRoleLabel': str(entry.get('purchaserRoleLabel') or ''),
                    'isCustomPayment': False,
                    'isBalanceAddition': True,
                    'retrieved': False,
                })
            _cleanup_pending(fb_db, txid, uid)
            _finish_settle_lock(fb_db, txid, 'balance')
            return True, 'balance'
        if kind == 'cart':
            pid = f'pix_cart_{txid}'
            existing = fb_db.reference(f'usuarios/{uid}/cantinaPurchases/{pid}').get()
            if not existing and not _purchase_has_txid(fb_db, uid, txid):
                items = entry.get('items') or []
                stock_items = entry.get('stockItems') or items
                recipient = str(entry.get('recipientName') or '').strip()
                _write_purchase(fb_db, uid, pid, {
                    'dateTime': created,
                    'items': items,
                    'totalPaid': amount,
                    'paymentMethod': 'pix',
                    'studentName': recipient or buyer,
                    'pixTransactionId': txid,
                    'pixPayerName': pagante,
                    'purchaserRoleLabel': str(entry.get('purchaserRoleLabel') or ''),
                    'isCustomPayment': False,
                    'isBalanceAddition': False,
                    'retrieved': False,
                    'scheduledRetrievalDate': entry.get('scheduledRetrievalDate'),
                    'retrievalTimeSlot': entry.get('retrievalTimeSlot'),
                })
                _dec_stock(fb_db, stock_items)
            _cleanup_pending(fb_db, txid, uid)
            _finish_settle_lock(fb_db, txid, 'cart')
            return True, 'cart'
        _cleanup_pending(fb_db, txid, uid)
        _finish_settle_lock(fb_db, txid, kind)
        return True, 'ok'
    except Exception as e:
        _abort_settle_lock(fb_db, txid)
        return False, str(e)


def register_pix_routes(app):
    """Registra as rotas PIX no app Flask. Chamar passando o app."""

    @app.route('/api/pix/check', methods=['GET'])
    def pix_check():
        """Teste de conex├úo com Mercado Pago (token v├ílido)."""
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
                'mensagem': 'Token v├ílido e comunicando com Mercado Pago!',
            })
        except Exception as e:
            return jsonify({'status': 'ERRO', 'detalhe': str(e)}), 500

    @app.route('/api/pix/create', methods=['POST'])
    def pix_create():
        """Cria cobran├ºa PIX. Body: { "valor": number, "descricao": string }."""
        headers = _mp_headers()
        if not headers:
            return jsonify({'error': 'Servidor n├úo configurado', 'detail': 'Token MP ausente'}), 503

        try:
            body = request.get_json() or {}
            valor = body.get('valor')
            descricao = body.get('descricao') or 'Pedido Cantina'

            try:
                transaction_amount = float(valor)
            except (TypeError, ValueError):
                return jsonify({'error': 'Valor inv├ílido'}), 400

            if transaction_amount <= 0:
                return jsonify({'error': 'Valor inv├ílido'}), 400

            payload = {
                'transaction_amount': transaction_amount,
                'description': descricao,
                'payment_method_id': 'pix',
                'payer': {'email': 'cliente@cantina.com'},
                'notification_url': _webhook_public_url(),
            }
            # Evita PIX com validade curta por padr├úo do Mercado Pago.
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
            return jsonify({'error': 'Erro ao criar cobran├ºa', 'detail': err_detail}), 500
        except Exception as e:
            return jsonify({'error': 'Erro ao criar cobran├ºa', 'detail': str(e)}), 500

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

    @app.route('/api/pix/webhook', methods=['POST', 'GET'])
    def pix_webhook():
        """Mercado Pago avisa quando o PIX e aprovado. App pode estar fechado."""
        txid = (
            request.args.get('id')
            or request.args.get('data.id')
            or ''
        ).strip()
        topic = (request.args.get('topic') or request.args.get('type') or '').strip().lower()
        body = request.get_json(silent=True) or {}
        if not txid:
            data = body.get('data') if isinstance(body, dict) else None
            if isinstance(data, dict):
                txid = str(data.get('id') or '').strip()
            if not topic:
                topic = str(body.get('type') or body.get('action') or '').lower()
        if topic and 'payment' not in topic and topic not in ('payment', 'ipn'):
            return jsonify({'ok': True, 'ignored': True}), 200
        if not txid:
            return jsonify({'ok': True, 'ignored': True}), 200
        _settle_approved_pix(txid)
        return jsonify({'ok': True}), 200

    @app.route('/api/pix/settle-mine', methods=['POST'])
    def pix_settle_mine():
        """Aluno reabre o app: servidor confere os PIX dele no Mercado Pago e aplica."""
        uid, err = _verify_bearer_uid()
        if err:
            return err
        fb_db = _get_fb_db()
        if fb_db is None:
            return jsonify({'ok': False}), 503
        settled = 0
        seen = set()
        mine = fb_db.reference(f'cantina_pending_pix_by_user/{uid}').get() or {}
        extra = fb_db.reference(f'usuarios/{uid}/cantinaPendingPix').get() or {}
        all_map = {}
        if isinstance(mine, dict):
            all_map.update(mine)
        if isinstance(extra, dict):
            all_map.update(extra)
        for txid, entry in all_map.items():
            txid = str(txid or '').strip()
            if not txid or txid in seen:
                continue
            seen.add(txid)
            ok, _reason = _settle_approved_pix(txid)
            if ok and _reason not in ('not_approved', 'no_pending', 'locked'):
                if _reason in ('debt', 'balance', 'cart', 'already', 'ok'):
                    settled += 1
        return jsonify({'ok': True, 'settled': settled})

    @app.route('/api/pix/confirm-debt', methods=['POST'])
    def pix_confirm_debt():
        uid, err = _verify_bearer_uid()
        if err:
            return err
        fb_db = _get_fb_db()
        if fb_db is None:
            return jsonify({'ok': False, 'error': 'Firebase Admin nao configurado'}), 503
        body = request.get_json() or {}
        txid = (body.get('txid') or '').strip()
        if not txid:
            return jsonify({'ok': False, 'error': 'txid obrigatorio'}), 400
        entry = _load_pending_entry(fb_db, txid, uid)
        if isinstance(entry, dict):
            entry_uid = (entry.get('userId') or '').strip()
            if entry_uid and entry_uid != uid:
                return jsonify({'ok': False, 'error': 'Usuario nao corresponde ao PIX'}), 403
        ok, reason = _settle_approved_pix(txid)
        if ok:
            return jsonify({'ok': True, 'alreadyCleared': reason in ('already', 'no_pending')})
        return jsonify({'ok': False, 'error': reason}), 400

    @app.route('/api/pix/settle-debt-balance', methods=['POST'])
    def pix_settle_debt_balance():
        """
        Quita d├¡vida usando apenas o saldo da carteira (servidor aplica d├®bito e remove bloqueio).
        Header: Authorization: Bearer <Firebase ID token do aluno>
        """
        uid, err = _verify_bearer_uid()
        if err:
            return err
        fb_db = _get_fb_db()
        if fb_db is None:
            return jsonify({'ok': False, 'error': 'Firebase Admin n├úo configurado'}), 503
        ref_debt = fb_db.reference(f'cantina_pending_debts/{uid}')
        debt = ref_debt.get()
        if not debt or not isinstance(debt, dict):
            return jsonify({'ok': False, 'error': 'Sem d├¡vida pendente'}), 400
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
            return jsonify({'ok': False, 'error': 'Usu├írio n├úo encontrado'}), 400
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
