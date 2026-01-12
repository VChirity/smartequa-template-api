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
    # Usar o novo template único CONTRATO_EQUAÇÃO_2026.docx
    template_path = os.path.join('templates_contratos', 'CONTRATO_EQUAÇÃO_2026.docx')
    
    if not os.path.exists(template_path):
        raise FileNotFoundError(f'Template não encontrado: {template_path}')
    
    # Carregar template
    doc = DocxTemplate(template_path)
    
    # Preparar contexto com os dados
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
        'ano_letivo': dados.get('ano_letivo', '2026'),
        'data_extenso': dados.get('data_extenso', ''),
        'naturalidade_aluno': dados.get('naturalidade_aluno', ''),
        'nasc_aluno': dados.get('nasc_aluno', ''),
        'cpf_aluno': dados.get('cpf_aluno', ''),
        
        # Desconto (sempre incluir, calculado automaticamente)
        'desconto': dados.get('desconto', ''),
        'desconto_extenso': dados.get('desconto_extenso', ''),
        
        # Mensalidades mensais (mens_jan a mens_dez)
        'mens_jan': dados.get('mens_jan', ''),
        'mens_fev': dados.get('mens_fev', ''),
        'mens_mar': dados.get('mens_mar', ''),
        'mens_abr': dados.get('mens_abr', ''),
        'mens_mai': dados.get('mens_mai', ''),
        'mens_jun': dados.get('mens_jun', ''),
        'mens_jul': dados.get('mens_jul', ''),
        'mens_ago': dados.get('mens_ago', ''),
        'mens_set': dados.get('mens_set', ''),
        'mens_out': dados.get('mens_out', ''),
        'mens_nov': dados.get('mens_nov', ''),
        'mens_dez': dados.get('mens_dez', ''),
        
        # Valores por extenso (extenso_jan a extenso_dez)
        'extenso_jan': dados.get('extenso_jan', ''),
        'extenso_fev': dados.get('extenso_fev', ''),
        'extenso_mar': dados.get('extenso_mar', ''),
        'extenso_abr': dados.get('extenso_abr', ''),
        'extenso_mai': dados.get('extenso_mai', ''),
        'extenso_jun': dados.get('extenso_jun', ''),
        'extenso_jul': dados.get('extenso_jul', ''),
        'extenso_ago': dados.get('extenso_ago', ''),
        'extenso_set': dados.get('extenso_set', ''),
        'extenso_out': dados.get('extenso_out', ''),
        'extenso_nov': dados.get('extenso_nov', ''),
        'extenso_dez': dados.get('extenso_dez', ''),
    }
    
    # Renderizar template
    doc.render(contexto)
    
    # Salvar em memória
    arquivo_bytes = BytesIO()
    doc.save(arquivo_bytes)
    arquivo_bytes.seek(0)
    
    return arquivo_bytes
