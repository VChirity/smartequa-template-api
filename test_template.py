from docxtpl import DocxTemplate
import os

# Testar template
template_path = os.path.join('templates_quadros', 'notas', 'quadro_notas_template_FINAL.docx')

print(f"Carregando template: {template_path}")
doc = DocxTemplate(template_path)

# Dados de teste
dados_teste = {
    'bimestre': '4º Bimestre',
    'ano': '2025',
    'data_entrega': '18/01/2026',
    'turma': '3° Ano',
    'tipo_av_1': 'Teste Bimestral',
    'tipo_av_2': '',
    'tipo_av_3': '',
    'tipo_av_4': '',
    'tipo_av_5': '',
    'pont_1': '10',
    'pont_2': '',
    'pont_3': '',
    'pont_4': '',
    'pont_5': '',
    'tipo_calculo': 'Nota única',
    'explicacao_calculo': '',
    'professor': 'Victor Chirity',
    'disciplina': 'Matemática',
    'alunos': [
        {'numero': '01', 'nome': 'Aluno Teste 1', 'av1': '7,00', 'av2': '', 'av3': '', 'av4': '', 'av5': '', 'media': '7,00'},
        {'numero': '02', 'nome': 'Aluno Teste 2', 'av1': '8,00', 'av2': '', 'av3': '', 'av4': '', 'av5': '', 'media': '8,00'},
    ]
}

print(f"\nNúmero de alunos: {len(dados_teste['alunos'])}")
print(f"Alunos: {dados_teste['alunos']}")

try:
    print("\nRenderizando template...")
    doc.render(dados_teste)
    
    output_path = 'teste_output.docx'
    doc.save(output_path)
    print(f"\n✅ Arquivo gerado com sucesso: {output_path}")
    print("Abra o arquivo para verificar se a tabela está preenchida corretamente.")
    
except Exception as e:
    print(f"\n❌ Erro ao renderizar: {e}")
    import traceback
    traceback.print_exc()
