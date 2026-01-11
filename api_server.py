from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from generators.quadro_notas_generator import gerar_quadro_notas_word

app = Flask(__name__)
CORS(app)  # Permitir requisições do Flutter

@app.route('/')
def home():
    return jsonify({
        'status': 'online',
        'message': 'Template App API - Servidor rodando!',
        'endpoints': [
            '/api/gerar-quadro-notas'
        ]
    })

@app.route('/api/gerar-quadro-notas', methods=['POST'])
def gerar_quadro_notas():
    """
    Endpoint para gerar Quadro de Notas em Word
    Recebe JSON com dados do quadro e retorna arquivo .docx
    """
    try:
        dados = request.json
        
        if not dados:
            return jsonify({'error': 'Nenhum dado recebido'}), 400
        
        # Gerar documento Word
        arquivo = gerar_quadro_notas_word(dados)
        
        # Nome do arquivo para download
        turma = dados.get('turma', 'Turma').replace(' ', '_')
        bimestre = dados.get('bimestre', 'Bimestre').replace(' ', '_')
        ano = dados.get('ano', '2026')
        nome_arquivo = f'Quadro_Notas_{turma}_{bimestre}_{ano}.docx'
        
        # Retornar arquivo para download
        return send_file(
            arquivo,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=nome_arquivo
        )
    except FileNotFoundError as e:
        return jsonify({'error': f'Template não encontrado: {str(e)}'}), 404
    except Exception as e:
        print(f'Erro ao gerar quadro de notas: {str(e)}')
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5000))
    
    print('=' * 60)
    print('🚀 Template App API Server')
    print('=' * 60)
    print(f'Servidor rodando na porta: {port}')
    print('Endpoint disponível: /api/gerar-quadro-notas')
    print('=' * 60)
    app.run(debug=False, host='0.0.0.0', port=port)
