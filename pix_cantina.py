# -*- coding: utf-8 -*-
"""
Rotas PIX Cantina (Mercado Pago) para o app Smart Equação.
NÃO altera nenhuma rota de templates - apenas adiciona endpoints /api/pix/*
"""
import os
import uuid
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
        """Consulta status do pagamento PIX. Retorna { "pago": bool, "status_mp": string }."""
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
            return jsonify({'pago': status == 'approved', 'status_mp': status})
        except Exception:
            return jsonify({'pago': False}), 500
