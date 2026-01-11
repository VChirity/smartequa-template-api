import os
from docxtpl import DocxTemplate
from io import BytesIO

def gerar_contrato_word(dados):
    """
    Gera contrato Word a partir dos dados fornecidos
    
    Args:
        dados: Dicionário com os dados do contrato
        
    Returns:
        BytesIO: Arquivo Word em memória
    """
    # Determinar qual template usar
    tem_desconto = dados.get('tem_desconto', False)
    
    if tem_desconto:
        template_path = os.path.join('templates_contratos', 'template_contratoDESCONTO2025_2.docx')
    else:
        template_path = os.path.join('templates_contratos', 'template_contrato2025_2.docx')
    
    if not os.path.exists(template_path):
        raise FileNotFoundError(f'Template não encontrado: {template_path}')
    
    # Carregar template
    doc = DocxTemplate(template_path)
    
    # Preparar contexto com os dados (mesmas tags do Streamlit)
    contexto = {
        # Responsável 1
        'responsavel1': dados.get('nome_responsavel1', ''),
        'cpf_responsavel': dados.get('cpf_responsavel1', ''),
        'endereco_completo': dados.get('endereco_responsavel1', ''),
        'bairro': dados.get('bairro_responsavel1', ''),
        'cep': dados.get('cep_responsavel1', ''),
        'naturalidade_resp1': dados.get('naturalidade_resp1', ''),
        'nasc_resp1': dados.get('nasc_resp1', ''),
        
        # Responsável 2
        'responsavel2': dados.get('nome_responsavel2', ''),
        'cpf2': dados.get('cpf_responsavel2', ''),
        'endereco2': dados.get('endereco_responsavel2', ''),
        'bairro2': dados.get('bairro_responsavel2', ''),
        'cep2': dados.get('cep_responsavel2', ''),
        
        # Aluno
        'nome_aluno': dados.get('nome_aluno', ''),
        'ano': dados.get('ano_escolar', ''),
        'ano_letivo': dados.get('ano_letivo', '2025'),
        'data_extenso': dados.get('data_extenso', ''),
        'naturalidade_aluno': dados.get('naturalidade_aluno', ''),
        'nasc_aluno': dados.get('nasc_aluno', ''),
        'cpf_aluno': dados.get('cpf_aluno', ''),
    }
    
    # Se tiver desconto, adicionar campos específicos
    if tem_desconto:
        contexto['desconto'] = dados.get('desconto', '')
        contexto['desconto_extenso'] = dados.get('desconto_extenso', '')
    
    # Renderizar template
    doc.render(contexto)
    
    # Salvar em memória
    arquivo_bytes = BytesIO()
    doc.save(arquivo_bytes)
    arquivo_bytes.seek(0)
    
    return arquivo_bytes
