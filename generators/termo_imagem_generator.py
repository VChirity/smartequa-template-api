import os
from docxtpl import DocxTemplate
from io import BytesIO

def gerar_termo_imagem_word(dados):
    """
    Gera termo de imagem e voz Word a partir dos dados fornecidos
    
    Args:
        dados: Dicionário com os dados do termo
        
    Returns:
        BytesIO: Arquivo Word em memória
    """
    # Determinar qual template usar
    tipo_termo = dados.get('tipo_termo', 'institucional')
    
    if tipo_termo == 'publicidade':
        template_path = os.path.join('templates_contratos', 'IMAGEM-E-VOZ-ALUNO-PUBLICIDADE_1.docx')
    else:
        template_path = os.path.join('templates_contratos', 'IMAGEM-E-VOZ-ALUNO-INSTITUCIONAL_1.docx')
    
    if not os.path.exists(template_path):
        raise FileNotFoundError(f'Template não encontrado: {template_path}')
    
    # Carregar template
    doc = DocxTemplate(template_path)
    
    # Preparar contexto com os dados (mesmas tags do Streamlit)
    contexto = {
        # Responsável
        'responsavel1': dados.get('nome_responsavel', ''),
        'cpf_responsavel': dados.get('cpf_responsavel', ''),
        'endereco_completo': dados.get('endereco', ''),
        
        # Aluno
        'nome_aluno': dados.get('nome_aluno', ''),
        'naturalidade_aluno': dados.get('naturalidade_aluno', ''),
        'nasc_aluno': dados.get('nasc_aluno', ''),
        'cpf_aluno': dados.get('cpf_aluno', ''),
        
        # Data
        'data_extenso': dados.get('data_extenso', ''),
    }
    
    # Renderizar template
    doc.render(contexto)
    
    # Salvar em memória
    arquivo_bytes = BytesIO()
    doc.save(arquivo_bytes)
    arquivo_bytes.seek(0)
    
    return arquivo_bytes
