"""
Firebase Admin Routes - Rotas para operações administrativas no Firebase
Requer credenciais de Service Account do Firebase

Para configurar:
1. No Firebase Console: Configurações do projeto → Contas de serviço → Gerar nova chave privada
2. Copie o JSON inteiro
3. No Render: adicione variável de ambiente FIREBASE_SERVICE_ACCOUNT_JSON com o JSON

Estas rotas permitem:
- Deletar usuários do Firebase Auth E do Realtime Database
- Alterar senha de usuários
- Atualizar email de usuários no Auth
"""

from flask import Blueprint, request, jsonify
import hashlib
import os
import json

firebase_admin_bp = Blueprint('firebase_admin', __name__, url_prefix='/api/admin')

_firebase_initialized = False
_init_error = None

def _init_firebase():
    """Inicializa o Firebase Admin SDK se ainda não estiver inicializado"""
    global _firebase_initialized, _init_error
    
    if _firebase_initialized:
        return True
    
    if _init_error:
        return False
    
    try:
        import firebase_admin
        from firebase_admin import credentials
        
        if firebase_admin._apps:
            _firebase_initialized = True
            return True
        
        cred_json = os.getenv('FIREBASE_SERVICE_ACCOUNT_JSON')
        
        if cred_json:
            cred_dict = json.loads(cred_json)
            cred = credentials.Certificate(cred_dict)
            firebase_admin.initialize_app(cred, {
                'databaseURL': 'https://equa-sec-apk-default-rtdb.firebaseio.com'
            })
            _firebase_initialized = True
            print('Firebase Admin SDK inicializado com sucesso!')
            return True
        else:
            cred_file = 'firebase-service-account.json'
            if os.path.exists(cred_file):
                cred = credentials.Certificate(cred_file)
                firebase_admin.initialize_app(cred, {
                    'databaseURL': 'https://equa-sec-apk-default-rtdb.firebaseio.com'
                })
                _firebase_initialized = True
                print('Firebase Admin SDK inicializado com arquivo local!')
                return True
            else:
                _init_error = 'Credenciais do Firebase Admin não encontradas'
                print(f'AVISO: {_init_error}')
                return False
                
    except Exception as e:
        _init_error = str(e)
        print(f'Erro ao inicializar Firebase Admin SDK: {e}')
        return False


def _verify_request_token(id_token):
    """
    Verifica se o token do usuário é válido e se tem permissão de admin.
    Para simplificar, vamos confiar no token enviado e verificar no banco de dados.
    """
    if not id_token:
        return None, 'Token não fornecido'
    
    try:
        from firebase_admin import auth, db
        
        decoded_token = auth.verify_id_token(id_token)
        uid = decoded_token.get('uid')
        
        if not uid:
            return None, 'UID não encontrado no token'
        
        ref = db.reference(f'usuarios/{uid}')
        user_data = ref.get()
        
        if not user_data:
            return None, 'Usuário não encontrado no banco de dados'
        
        # 'isAdmin' não vale mais: era um campo que o próprio usuário podia gravar.
        is_admin = user_data.get('role') == 'admin'
        is_prof_admin = user_data.get('isProfAdmin') == True
        
        if not (is_admin or is_prof_admin):
            return None, 'Usuário não tem permissão de administrador'
        
        return uid, None
        
    except Exception as e:
        return None, f'Erro ao verificar token: {str(e)}'


@firebase_admin_bp.route('/check', methods=['GET'])
def check_admin_api():
    """Verifica se a API de admin está configurada e funcionando"""
    initialized = _init_firebase()
    
    return jsonify({
        'configured': initialized,
        'message': 'Firebase Admin API configurada' if initialized else (_init_error or 'API não configurada')
    })


@firebase_admin_bp.route('/delete-user', methods=['POST'])
def delete_user():
    """
    Deleta um usuário do Firebase Auth E do Realtime Database.
    Requer token de admin no header Authorization.
    Body JSON: { "userId": "uid_do_usuario" }
    """
    if not _init_firebase():
        return jsonify({'error': 'Firebase Admin não configurado', 'configured': False}), 503
    
    auth_header = request.headers.get('Authorization', '')
    id_token = auth_header.replace('Bearer ', '') if auth_header.startswith('Bearer ') else auth_header
    
    admin_uid, error = _verify_request_token(id_token)
    if error:
        return jsonify({'error': error}), 401
    
    try:
        from firebase_admin import auth, db
        
        body = request.get_json() or {}
        user_id = body.get('userId')
        
        if not user_id:
            return jsonify({'error': 'userId é obrigatório'}), 400
        
        if user_id == admin_uid:
            return jsonify({'error': 'Não é possível deletar a própria conta'}), 400
        
        auth_deleted = False
        db_deleted = False
        
        try:
            auth.delete_user(user_id)
            auth_deleted = True
            print(f'Usuário {user_id} deletado do Firebase Auth')
        except auth.UserNotFoundError:
            print(f'Usuário {user_id} não encontrado no Auth (pode já ter sido deletado)')
        except Exception as e:
            print(f'Aviso: Não foi possível deletar do Auth: {e}')
        
        try:
            ref = db.reference(f'usuarios/{user_id}')
            ref.delete()
            db_deleted = True
            print(f'Usuário {user_id} deletado do Realtime Database')
        except Exception as e:
            print(f'Aviso: Não foi possível deletar do Database: {e}')
        
        try:
            dup_ref = db.reference('duplicate_check')
            user_ref = db.reference(f'usuarios/{user_id}')
            user_data = user_ref.get()
            
            if user_data:
                import hashlib
                
                email = (user_data.get('email') or '').strip().lower().replace(' ', '')
                if email:
                    email_hash = hashlib.sha256(email.encode()).hexdigest()
                    db.reference(f'duplicate_check/emails/{email_hash}').delete()
                
                telefone = (user_data.get('telefone') or '').replace(' ', '').replace('-', '').replace('(', '').replace(')', '')
                if len(telefone) >= 8:
                    phone_hash = hashlib.sha256(telefone.encode()).hexdigest()
                    db.reference(f'duplicate_check/phones/{phone_hash}').delete()
                
                nome = (user_data.get('nome') or '').strip().lower().replace('  ', ' ')
                if nome:
                    nome_hash = hashlib.sha256(nome.encode()).hexdigest()
                    db.reference(f'duplicate_check/nomes/{nome_hash}').delete()
        except Exception as e:
            print(f'Aviso: Erro ao limpar duplicate_check: {e}')
        
        if auth_deleted or db_deleted:
            return jsonify({
                'success': True,
                'authDeleted': auth_deleted,
                'dbDeleted': db_deleted,
                'message': 'Usuário excluído com sucesso'
            })
        else:
            return jsonify({'error': 'Não foi possível excluir o usuário'}), 500
            
    except Exception as e:
        print(f'Erro ao excluir usuário: {e}')
        return jsonify({'error': f'Erro ao excluir usuário: {str(e)}'}), 500


@firebase_admin_bp.route('/update-password', methods=['POST'])
def update_password():
    """
    Atualiza a senha de um usuário no Firebase Auth.
    Requer token de admin no header Authorization.
    Body JSON: { "userId": "uid", "newPassword": "nova_senha" }
    """
    if not _init_firebase():
        return jsonify({'error': 'Firebase Admin não configurado', 'configured': False}), 503
    
    auth_header = request.headers.get('Authorization', '')
    id_token = auth_header.replace('Bearer ', '') if auth_header.startswith('Bearer ') else auth_header
    
    admin_uid, error = _verify_request_token(id_token)
    if error:
        return jsonify({'error': error}), 401
    
    try:
        from firebase_admin import auth
        
        body = request.get_json() or {}
        user_id = body.get('userId')
        new_password = body.get('newPassword')
        
        if not user_id:
            return jsonify({'error': 'userId é obrigatório'}), 400
        
        if not new_password:
            return jsonify({'error': 'newPassword é obrigatório'}), 400
        
        if len(new_password) < 6:
            return jsonify({'error': 'A senha deve ter pelo menos 6 caracteres'}), 400
        
        auth.update_user(user_id, password=new_password)
        
        print(f'Senha do usuário {user_id} atualizada com sucesso')
        
        return jsonify({
            'success': True,
            'message': 'Senha atualizada com sucesso'
        })
        
    except auth.UserNotFoundError:
        return jsonify({'error': 'Usuário não encontrado no Firebase Auth'}), 404
    except Exception as e:
        print(f'Erro ao atualizar senha: {e}')
        return jsonify({'error': f'Erro ao atualizar senha: {str(e)}'}), 500


def _normalize_login_email(raw):
    email = (raw or '').strip().lower().replace(' ', '')
    if not email:
        return ''
    if '@' not in email:
        email = email + '@colegioequacao.com'
    return email


def _dup_hash(value):
    return hashlib.sha256(value.encode('utf-8')).hexdigest()


def _sync_duplicate_email(db, old_email, new_email):
    old_n = _normalize_login_email(old_email)
    new_n = _normalize_login_email(new_email)
    if old_n and old_n != new_n:
        db.reference(f'duplicate_check/emails/{_dup_hash(old_n)}').delete()
    if new_n:
        db.reference(f'duplicate_check/emails/{_dup_hash(new_n)}').set(True)


@firebase_admin_bp.route('/update-user', methods=['POST'])
def update_user():
    """
    Atualiza nome, e-mail e senha no Firebase Auth e no Realtime Database.
    O e-mail do banco só muda depois que o login (Auth) aceitar.
    Body JSON: { "userId", "email", "nome", "newPassword", "telefone", "dadosAdicionais" }
    """
    if not _init_firebase():
        return jsonify({'error': 'Firebase Admin não configurado', 'configured': False}), 503

    auth_header = request.headers.get('Authorization', '')
    id_token = auth_header.replace('Bearer ', '') if auth_header.startswith('Bearer ') else auth_header

    admin_uid, error = _verify_request_token(id_token)
    if error:
        return jsonify({'error': error}), 401

    try:
        from firebase_admin import auth, db

        body = request.get_json() or {}
        user_id = body.get('userId')
        if not user_id:
            return jsonify({'error': 'userId é obrigatório'}), 400

        new_email = _normalize_login_email(body.get('email'))
        nome = (body.get('nome') or '').strip()
        telefone = body.get('telefone')
        dados_adicionais = body.get('dadosAdicionais')
        new_password = body.get('newPassword') or body.get('password') or ''
        if isinstance(new_password, str):
            new_password = new_password.strip()
        else:
            new_password = ''

        if new_email and '@' not in new_email:
            return jsonify({'error': 'E-mail inválido.'}), 400
        if new_password and len(new_password) < 6:
            return jsonify({'error': 'A senha deve ter pelo menos 6 caracteres.'}), 400

        ref = db.reference(f'usuarios/{user_id}')
        atual = ref.get() or {}
        email_atual = _normalize_login_email(atual.get('email'))

        auth_kwargs = {}
        if new_email and new_email != email_atual:
            auth_kwargs['email'] = new_email
        if new_password:
            auth_kwargs['password'] = new_password

        auth_updated = False
        if auth_kwargs:
            try:
                auth.update_user(user_id, **auth_kwargs)
                auth_updated = True
                print(f'Auth do usuário {user_id} atualizado: {list(auth_kwargs.keys())}')
            except auth.UserNotFoundError:
                return jsonify({
                    'error': 'Esta conta não existe no login. O e-mail e a senha do banco não foram alterados.',
                }), 404
            except Exception as e:
                msg = str(e)
                baixo = msg.lower()
                if 'email_exists' in baixo or 'already exists' in baixo or 'email already' in baixo:
                    return jsonify({'error': 'Esse e-mail já está em uso por outra conta. Nada foi alterado.'}), 409
                print(f'Erro ao atualizar Auth de {user_id}: {e}')
                return jsonify({
                    'error': f'Não foi possível atualizar o login (e-mail/senha). Nada foi alterado no banco. {msg}',
                }), 400

        db_data = {}
        if new_email:
            db_data['email'] = new_email
        if nome:
            db_data['nome'] = nome
        if telefone is not None:
            db_data['telefone'] = telefone
        if dados_adicionais:
            db_data['dadosAdicionais'] = dados_adicionais

        db_updated = False
        if db_data:
            ref.update(db_data)
            db_updated = True
            if new_email:
                try:
                    _sync_duplicate_email(db, atual.get('email'), new_email)
                except Exception as e:
                    print(f'Aviso: duplicate_check de e-mail: {e}')

        return jsonify({
            'success': True,
            'authUpdated': auth_updated,
            'dbUpdated': db_updated,
            'email': new_email or email_atual,
            'message': 'Dados atualizados com sucesso',
        })

    except Exception as e:
        print(f'Erro ao atualizar usuário: {e}')
        return jsonify({'error': f'Erro ao atualizar usuário: {str(e)}'}), 500


def register_firebase_admin_routes(app):
    """Registra as rotas de admin no app Flask"""
    app.register_blueprint(firebase_admin_bp)
    print('Rotas de Firebase Admin registradas em /api/admin/*')
