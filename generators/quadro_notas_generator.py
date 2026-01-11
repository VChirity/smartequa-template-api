from docxtpl import DocxTemplate
import io
import os

def gerar_quadro_notas_word(dados):
    """
    Gera documento Word usando template com tags Jinja2
    ⚠️ IMPORTANTE: Usa templates_quadros/notas/ (NÃO templates/)
    """
    print('=' * 60)
    print('DEBUG: Iniciando geração de Word')
    print(f'DEBUG: Dados recebidos: {dados}')
    print('=' * 60)
    
    # Carregar template da pasta SEPARADA
    template_path = os.path.join('templates_quadros', 'notas', 'quadro_notas_template.docx')
    
    if not os.path.exists(template_path):
        raise FileNotFoundError(f'Template não encontrado: {template_path}')
    
    print(f'DEBUG: Template encontrado em: {template_path}')
    
    doc = DocxTemplate(template_path)
    
    # Preparar contexto para o template
    context = {
        'bimestre': dados.get('bimestre', ''),
        'ano': dados.get('ano', ''),
        'data_entrega': dados.get('data_entrega', ''),
        'turma': dados.get('turma', ''),
        'tipo_av_1': dados.get('tipo_av_1', ''),
        'tipo_av_2': dados.get('tipo_av_2', ''),
        'tipo_av_3': dados.get('tipo_av_3', ''),
        'tipo_av_4': dados.get('tipo_av_4', ''),
        'tipo_av_5': dados.get('tipo_av_5', ''),
        'pont_1': dados.get('pont_1', ''),
        'pont_2': dados.get('pont_2', ''),
        'pont_3': dados.get('pont_3', ''),
        'pont_4': dados.get('pont_4', ''),
        'pont_5': dados.get('pont_5', ''),
        'tipo_calculo': dados.get('tipo_calculo', ''),
        'explicacao_calculo': dados.get('explicacao_calculo', ''),
        'professor': dados.get('professor', ''),
        'disciplina': dados.get('disciplina', ''),
        'alunos': dados.get('alunos', []),
    }
    
    print(f'DEBUG: Contexto preparado: {context}')
    print(f'DEBUG: Número de alunos: {len(context["alunos"])}')
    
    # Renderizar template
    print('DEBUG: Renderizando template...')
    doc.render(context)
    print('DEBUG: Template renderizado com sucesso!')
    
    # Salvar em memória
    arquivo = io.BytesIO()
    doc.save(arquivo)
    arquivo.seek(0)
    
    print('DEBUG: Arquivo Word gerado com sucesso!')
    print('=' * 60)
    
    return arquivo
