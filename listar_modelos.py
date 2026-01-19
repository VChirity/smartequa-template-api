import google.generativeai as genai

# Configurar API Key
api_key = 'AIzaSyBomqSUgc3m7HCZu_dQS0nUy2cPSDT1q7I'
genai.configure(api_key=api_key)

print("=" * 60)
print("🔍 LISTANDO MODELOS DISPONÍVEIS NA API KEY")
print("=" * 60)
print()

try:
    # Listar todos os modelos disponíveis
    models = genai.list_models()
    
    print(f"✅ Total de modelos encontrados: {len(list(models))}")
    print()
    
    # Listar novamente para iterar (list_models retorna um generator)
    models = genai.list_models()
    
    for model in models:
        print(f"📦 Modelo: {model.name}")
        print(f"   Display Name: {model.display_name}")
        print(f"   Suporta: {', '.join(model.supported_generation_methods)}")
        print()
        
except Exception as e:
    print(f"❌ Erro ao listar modelos: {str(e)}")
    import traceback
    traceback.print_exc()

print("=" * 60)
print("🔍 PROCURANDO MODELOS QUE SUPORTAM IMAGENS (generateContent)")
print("=" * 60)
print()

try:
    # Filtrar modelos que suportam generateContent (necessário para imagens)
    models = genai.list_models()
    modelos_com_imagem = [m for m in models if 'generateContent' in m.supported_generation_methods]
    
    print(f"✅ Modelos que suportam generateContent: {len(modelos_com_imagem)}")
    print()
    
    for model in modelos_com_imagem:
        print(f"✅ {model.name}")
        print(f"   Display: {model.display_name}")
        print()
        
except Exception as e:
    print(f"❌ Erro: {str(e)}")
    import traceback
    traceback.print_exc()
