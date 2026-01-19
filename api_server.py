from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from generators.quadro_notas_generator import gerar_quadro_notas_word
from generators.contrato_generator import gerar_contrato_word
from generators.termo_imagem_generator import gerar_termo_imagem_word
import os
import base64
import json
from dotenv import load_dotenv
import google.generativeai as genai
from regras_correcao import PROMPT_REGRAS
from PIL import Image
from io import BytesIO
from num2words import num2words

# Carregar variáveis de ambiente
load_dotenv()

app = Flask(__name__)
CORS(app)  # Permitir requisições do Flutter

def converter_desconto_extenso(desconto_str):
    """Converte desconto numérico para extenso com decimais corretos"""
    try:
        valor = float(desconto_str.replace(',', '.'))
        
        # Separar parte inteira e decimal
        partes = str(valor).split('.')
        parte_inteira = int(partes[0])
        
        # Converter parte inteira
        extenso_inteira = num2words(parte_inteira, lang='pt_BR')
        
        # Se tem parte decimal
        if len(partes) > 1 and partes[1] != '0':
            parte_decimal = partes[1].ljust(2, '0')[:2]  # Pegar 2 dígitos
            decimal_num = int(parte_decimal)
            extenso_decimal = num2words(decimal_num, lang='pt_BR')
            return f"{extenso_inteira} vírgula {extenso_decimal} por cento"
        else:
            return f"{extenso_inteira} por cento"
    except:
        return desconto_str

@app.route('/')
def home():
    return jsonify({
        'status': 'online',
        'message': 'Template App API - Servidor rodando!',
        'endpoints': [
            '/api/gerar-quadro-notas',
            '/api/gerar-contrato',
            '/api/gerar-termo-imagem',
            '/api/transcrever',
            '/api/corrigir'
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

@app.route('/api/gerar-contrato', methods=['POST'])
def gerar_contrato():
    """
    Endpoint para gerar Contrato em Word
    Recebe JSON com dados do contrato e retorna arquivo .docx
    """
    try:
        dados = request.json
        
        if not dados:
            return jsonify({'error': 'Nenhum dado recebido'}), 400
        
        # Converter desconto para extenso se houver
        if 'desconto' in dados and dados['desconto']:
            dados['desconto_extenso'] = converter_desconto_extenso(dados['desconto'])
        
        # Gerar documento Word
        arquivo = gerar_contrato_word(dados)
        
        # Nome do arquivo para download
        nome_aluno = dados.get('nome_aluno', 'Aluno').replace(' ', '_')
        ano = dados.get('ano_letivo', '2025')
        tem_desconto = dados.get('tem_desconto', False)
        tipo = 'Desconto' if tem_desconto else 'Normal'
        nome_arquivo = f'Contrato_{tipo}_{nome_aluno}_{ano}.docx'
        
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
        print(f'Erro ao gerar contrato: {str(e)}')
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/gerar-termo-imagem', methods=['POST'])
def gerar_termo_imagem():
    """
    Endpoint para gerar Termo de Imagem e Voz em Word
    Recebe JSON com dados do termo e retorna arquivo .docx
    """
    try:
        dados = request.json
        
        if not dados:
            return jsonify({'error': 'Nenhum dado recebido'}), 400
        
        # Gerar documento Word
        arquivo = gerar_termo_imagem_word(dados)
        
        # Nome do arquivo para download
        nome_aluno = dados.get('nome_aluno', 'Aluno').replace(' ', '_')
        tipo_termo = dados.get('tipo_termo', 'institucional')
        tipo_nome = 'Publicidade' if tipo_termo == 'publicidade' else 'Institucional'
        nome_arquivo = f'Termo_Imagem_Voz_{tipo_nome}_{nome_aluno}.docx'
        
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
        print(f'Erro ao gerar termo de imagem: {str(e)}')
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/transcrever', methods=['POST'])
def transcrever_redacao():
    """
    Endpoint para transcrever redação manuscrita via OCR usando Google Gemini
    Recebe imagem em Base64 e retorna texto transcrito
    """
    try:
        dados = request.json
        
        if not dados or 'imagem' not in dados:
            return jsonify({'error': 'Imagem não fornecida'}), 400
        
        # Configurar API do Gemini
        api_key = os.getenv('GEMINI_API_KEY')
        if not api_key:
            raise Exception('GEMINI_API_KEY não configurada nas variáveis de ambiente')
        print(f'🔑 Transcrever - Usando API Key: {api_key[:20]}...')
        genai.configure(api_key=api_key)
        
        # Decodificar imagem Base64
        imagem_base64 = dados['imagem']
        if ',' in imagem_base64:
            imagem_base64 = imagem_base64.split(',')[1]
        
        imagem_bytes = base64.b64decode(imagem_base64)
        imagem = Image.open(BytesIO(imagem_bytes))
        
        # Prompt para OCR
        prompt_ocr = """Aja como um OCR especializado em transcrever redações manuscritas.
        
Transcreva EXATAMENTE o que você vê na imagem, palavra por palavra.
NÃO corrija erros ortográficos ou gramaticais.
NÃO adicione ou remova palavras.
NÃO interprete ou melhore o texto.
Apenas transcreva fielmente o que está escrito.
        
Se houver palavras ilegíveis, marque com [ilegível].
Mantenha a estrutura de parágrafos."""
        
        # Sistema de fallback: tenta múltiplos modelos
        modelos = [
            'models/gemini-2.5-flash',      # Melhor: mais recente e rápido
            'models/gemini-flash-latest',   # Fallback 1: sempre atualizado
            'models/gemini-2.5-pro'         # Fallback 2: mais poderoso
        ]
        
        texto_transcrito = None
        ultimo_erro = None
        
        for modelo_nome in modelos:
            try:
                print(f'🔄 Tentando modelo: {modelo_nome}')
                model = genai.GenerativeModel(modelo_nome)
                response = model.generate_content([prompt_ocr, imagem])
                texto_transcrito = response.text
                print(f'✅ Sucesso com modelo: {modelo_nome}')
                break  # Sucesso! Sair do loop
            except Exception as e:
                ultimo_erro = str(e)
                print(f'❌ Falha com {modelo_nome}: {ultimo_erro}')
                continue  # Tentar próximo modelo
        
        # Se nenhum modelo funcionou, retornar erro
        if texto_transcrito is None:
            raise Exception(f'Nenhum modelo funcionou. Último erro: {ultimo_erro}')
        
        return jsonify({
            'sucesso': True,
            'texto': texto_transcrito
        })
        
    except Exception as e:
        print(f'Erro ao transcrever: {str(e)}')
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/corrigir', methods=['POST'])
def corrigir_redacao():
    """
    Endpoint para corrigir redação usando Google Gemini
    Recebe JSON com tema e texto, retorna correção estruturada
    """
    try:
        dados = request.json
        
        if not dados or 'tema' not in dados or 'texto' not in dados:
            return jsonify({'error': 'Tema e texto são obrigatórios'}), 400
        
        tema = dados['tema']
        texto = dados['texto']
        
        # Configurar API do Gemini
        api_key = os.getenv('GEMINI_API_KEY')
        if not api_key:
            raise Exception('GEMINI_API_KEY não configurada nas variáveis de ambiente')
        print(f'🔑 Corrigir - Usando API Key: {api_key[:20]}...')
        genai.configure(api_key=api_key)
        
        # Usar modelo Gemini 1.5 Flash (mais estável)
        model = genai.GenerativeModel('gemini-1.5-flash')
        print('✅ Modelo Gemini 1.5 Flash configurado para correção')
        
        # Montar prompt completo
        prompt_completo = f"""{PROMPT_REGRAS}

=== TEMA PROPOSTO ===
{tema}

=== REDAÇÃO DO ALUNO ===
{texto}

=== INSTRUÇÕES FINAIS ===
Analise a redação acima considerando o tema proposto e os critérios estabelecidos.
Retorne APENAS um JSON válido (sem markdown, sem ```json) com a seguinte estrutura:
{{
  "nota_final": <número de 0 a 1000>,
  "competencia_1": {{"nota": <0-200>, "comentario": "..."}},
  "competencia_2": {{"nota": <0-200>, "comentario": "..."}},
  "competencia_3": {{"nota": <0-200>, "comentario": "..."}},
  "competencia_4": {{"nota": <0-200>, "comentario": "..."}},
  "competencia_5": {{"nota": <0-200>, "comentario": "..."}},
  "pontos_fortes": ["...", "..."],
  "pontos_fracos": ["...", "..."],
  "sugestoes": ["...", "..."]
}}"""
        
        # Gerar correção
        response = model.generate_content(prompt_completo)
        resposta_texto = response.text.strip()
        
        # Limpar markdown se houver
        if resposta_texto.startswith('```'):
            resposta_texto = resposta_texto.split('\n', 1)[1]
            resposta_texto = resposta_texto.rsplit('```', 1)[0]
        
        # Parsear JSON
        correcao = json.loads(resposta_texto)
        
        return jsonify({
            'sucesso': True,
            'correcao': correcao
        })
        
    except json.JSONDecodeError as e:
        print(f'Erro ao parsear JSON da IA: {str(e)}')
        print(f'Resposta recebida: {resposta_texto}')
        return jsonify({
            'error': 'Erro ao processar resposta da IA',
            'detalhes': str(e),
            'resposta_bruta': resposta_texto
        }), 500
    except Exception as e:
        print(f'Erro ao corrigir: {str(e)}')
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
