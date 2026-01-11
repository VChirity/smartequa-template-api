# Solução Final para Templates de Imagem

## Problema
Os arquivos em `templates backup/` têm as tags corretas mas estão corrompidos quando salvos pelo Word.

## Solução Manual (Recomendada)

1. **Abra o arquivo limpo de `assets/`** no Word:
   - `assets/IMAGEM-E-VOZ-ALUNO-PUBLICIDADE_1.docx`
   - `assets/IMAGEM-E-VOZ-ALUNO-INSTITUCIONAL_1.docx`

2. **Edite as tags manualmente** (substitua os textos pelas tags):
   - (NOME COMPLETO DO RESPONSÁVEL) → {{responsavel1}}
   - (CPF DO RESPONSÁVEL) → {{cpf_responsavel}}
   - (ENDEREÇO COMPLETO DO RESPONSÁVEL) → {{endereco_completo}}
   - (NATURALIDADE DO RESPONSÁVEL) → {{naturalidade_resp1}}
   - (DATA DE NASCIMENTO DO RESPONSÁVEL) → {{nasc_resp1}}
   - (NOME COMPLETO DO ALUNO) → {{nome_aluno}}
   - (NATURALIDADE DO ALUNO) → {{naturalidade_aluno}}
   - (DATA DE NASCIMENTO DO ALUNO) → {{nasc_aluno}}
   - (CPF DO ALUNO) → {{cpf_aluno}}
   - (DATA DO DIA) → {{data_extenso}}

3. **SALVE o arquivo**

4. **IMEDIATAMENTE rode o script de consolidação:**
   ```bash
   python fix_image_final.py
   ```

5. **Teste abrindo no Word** - deve funcionar!

## Por que funciona assim?

- Arquivos de `assets/` têm estrutura XML limpa
- Você edita manualmente (Word fragmenta as tags ao salvar)
- Script consolida as tags fragmentadas ANTES de tentar abrir de novo
- Resultado: arquivo com tags corretas E estrutura XML limpa

## Alternativa: Usar os arquivos atuais

Se os arquivos em `templates/` (que acabei de gerar) abrirem no Word mas sem logo/formatação:
1. Adicione a logo manualmente
2. Ajuste a formatação
3. Salve
4. Rode `python fix_image_final.py`
