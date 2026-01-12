# ⚠️ IMPORTANTE: REINICIAR SERVIDOR PYTHON

## 🔴 PROBLEMA IDENTIFICADO:

O servidor Python está usando um **template antigo em cache**. As mudanças no código foram aplicadas, mas o servidor precisa ser reiniciado para carregar o novo template.

---

## ✅ SOLUÇÃO:

### **1. PARAR o servidor Python atual:**

No terminal onde o servidor está rodando, pressione:
```
Ctrl + C
```

Você vai ver algo como:
```
^C
Keyboard interrupt received, exiting.
```

### **2. INICIAR o servidor novamente:**

No mesmo terminal, execute:
```powershell
python api_server.py
```

Você vai ver:
```
============================================================
🚀 Template App API Server
============================================================
Servidor rodando na porta: 5000
Endpoint disponível: /api/gerar-quadro-notas
============================================================
 * Running on http://127.0.0.1:5000
 * Running on http://192.168.15.66:5000
```

### **3. Atualizar a página do Flutter:**

No Chrome, pressione **F5** ou **Ctrl+R**

### **4. Testar novamente:**

Gere o contrato e verifique se agora está usando o template correto com:
- ✅ Tabela 2026 (não mais 2025)
- ✅ Tags `{{mens_jan}}` a `{{mens_dez}}`
- ✅ Tags `{{extenso_jan}}` a `{{extenso_dez}}`

---

## 📋 VERIFICAÇÃO:

O arquivo correto está em:
```
G:\Projetos\SmartEqua\Templates app\templates_contratos\CONTRATO_EQUAÇÃO_2026.docx
```

O código Python está configurado para usar este arquivo:
```python
template_path = os.path.join('templates_contratos', 'CONTRATO_EQUAÇÃO_2026.docx')
```

**Tudo está correto no código, só precisa reiniciar o servidor!**

---

## 🎯 APÓS REINICIAR:

O servidor vai carregar o template atualizado e você vai ver no Word gerado:
- ✅ Tabela com "Anuidade 2026"
- ✅ Valores corretos (R$ 25.634,16 para 1º ao 4º Ano, etc.)
- ✅ Parágrafo com as tags mensais (mens_jan, mens_fev, etc.)
- ✅ Valores por extenso corretos

---

**REINICIE O SERVIDOR AGORA!** 🔄
