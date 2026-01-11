# 🔧 Correção do Template - Tabela Vazia

## Problema:
A tabela está sendo gerada vazia. Os dados estão sendo enviados corretamente pelo Flutter e recebidos pelo Python, mas o loop Jinja2 não está funcionando no template Word.

## Causa Provável:
As tags do loop têm **espaços ou caracteres invisíveis** que impedem o Jinja2 de processar corretamente.

## Solução:

### 1. Abrir o template Word
`G:\Projetos\SmartEqua\Templates app\templates_quadros\notas\quadro_notas_template.docx`

### 2. Localizar a tabela de alunos

Você vai ver 3 linhas dentro da tabela:
- Linha 1: `{% for aluno in alunos -%}`
- Linha 2: Com as tags dos dados
- Linha 3: `{% endfor -%}`

### 3. DELETAR as 3 linhas completamente

Selecione e delete as 3 linhas da tabela (não só o conteúdo, mas as linhas inteiras).

### 4. Inserir 3 novas linhas NA TABELA

**Linha 1 (início do loop):**
- Inserir nova linha na tabela
- Mesclar todas as células
- Copiar e colar EXATAMENTE isso (sem espaços extras):
```
{% for aluno in alunos -%}
```

**Linha 2 (dados):**
- Inserir nova linha na tabela
- **NÃO mesclar** - deixar 8 células separadas
- Em cada célula, copiar e colar EXATAMENTE (uma tag por célula):

Célula 1:
```
{{loop.index}}
```

Célula 2:
```
{{aluno.nome}}
```

Célula 3:
```
{{aluno.av1}}
```

Célula 4:
```
{{aluno.av2}}
```

Célula 5:
```
{{aluno.av3}}
```

Célula 6:
```
{{aluno.av4}}
```

Célula 7:
```
{{aluno.av5}}
```

Célula 8:
```
{{aluno.media}}
```

**Linha 3 (fim do loop):**
- Inserir nova linha na tabela
- Mesclar todas as células
- Copiar e colar EXATAMENTE isso (sem espaços extras):
```
{% endfor -%}
```

### 5. IMPORTANTE:
- Cada tag deve estar **sozinha** na célula
- **SEM espaços** antes ou depois
- **SEM Enter** dentro da célula
- Copiar e colar as tags deste documento para garantir que não tem caracteres invisíveis

### 6. Salvar o arquivo

### 7. Testar novamente

Recarregue a página do SmartEqua (F5) e gere um novo Word.

---

## Se ainda não funcionar:

Me avise e vou criar um template completamente novo do zero para você.
