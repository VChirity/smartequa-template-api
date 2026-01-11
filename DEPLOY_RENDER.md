# 🚀 Deploy do Template App no Render

## Pré-requisitos
- Conta no GitHub
- Conta no Render (https://render.com) - grátis

## Passo 1: Criar Repositório no GitHub

1. Criar novo repositório no GitHub (ex: `smartequa-template-api`)
2. **NÃO** incluir o `venv/` (já está no .gitignore)

## Passo 2: Fazer Push do Código

No terminal, dentro da pasta `Templates app`:

```bash
git init
git add .
git commit -m "Initial commit - Template App API"
git remote add origin https://github.com/SEU_USUARIO/smartequa-template-api.git
git push -u origin main
```

## Passo 3: Deploy no Render

1. Acesse https://render.com e faça login
2. Clique em **"New +"** → **"Web Service"**
3. Conecte seu repositório GitHub
4. Configure:
   - **Name:** `smartequa-template-api` (ou outro nome)
   - **Region:** Escolha a mais próxima (ex: Oregon)
   - **Branch:** `main`
   - **Root Directory:** deixe vazio (ou `.` se pedir)
   - **Runtime:** `Python 3`
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `python api_server.py`
   - **Instance Type:** `Free`

5. Clique em **"Create Web Service"**

## Passo 4: Aguardar Deploy

O Render vai:
- Instalar dependências
- Iniciar o servidor
- Fornecer uma URL (ex: `https://smartequa-template-api.onrender.com`)

⏱️ Primeiro deploy pode demorar 5-10 minutos.

## Passo 5: Testar a API

Acesse no navegador:
```
https://SEU-APP.onrender.com/
```

Deve retornar:
```json
{
  "status": "online",
  "message": "Template App API - Servidor rodando!",
  "endpoints": ["/api/gerar-quadro-notas"]
}
```

## Passo 6: Atualizar SmartEqua

Editar `lib/services/document_generator_service.dart`:

```dart
class DocumentGeneratorService {
  // Trocar localhost pela URL do Render
  static const String baseUrl = 'https://SEU-APP.onrender.com';
  
  // ... resto do código
}
```

Rebuild do Flutter:
```bash
flutter build web
```

## ⚠️ Importante

### Limitações do Plano Grátis:
- **Sleep após inatividade:** Servidor "dorme" após 15 min sem uso
- **Primeira requisição após sleep:** Demora ~30s para acordar
- **Solução:** Aceitar a demora ou fazer upgrade para plano pago ($7/mês)

### Manter Template Atualizado:
Sempre que modificar o template Word:
1. Fazer commit das mudanças
2. Push para GitHub
3. Render faz redeploy automático

## 🔧 Troubleshooting

### Erro "Template não encontrado"
- Verificar se a pasta `templates_quadros/notas/` está no repositório
- Verificar se o arquivo `quadro_notas_template.docx` está lá

### Erro de CORS
- Já está configurado no `api_server.py` com `CORS(app)`

### Logs
- No painel do Render, aba "Logs" mostra erros em tempo real

## 📝 Estrutura de Arquivos Necessária

```
Templates app/
├── api_server.py
├── requirements.txt
├── Procfile
├── runtime.txt
├── .gitignore
├── generators/
│   ├── __init__.py
│   └── quadro_notas_generator.py
└── templates_quadros/
    └── notas/
        └── quadro_notas_template.docx
```

## ✅ Checklist Final

- [ ] Código no GitHub
- [ ] Deploy no Render concluído
- [ ] URL da API funcionando
- [ ] SmartEqua atualizado com nova URL
- [ ] Rebuild do Flutter web
- [ ] Teste completo de geração de Word
