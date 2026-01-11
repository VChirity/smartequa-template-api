import win32com.client
import os
import sys

def repair_with_word(filepath):
    """
    Usa o Word COM para abrir e reparar o arquivo corrompido
    """
    print(f"\n🔧 Reparando com Word: {filepath}")
    
    if not os.path.exists(filepath):
        print(f"❌ Arquivo não encontrado!")
        return False
    
    # Caminho absoluto
    abs_path = os.path.abspath(filepath)
    
    try:
        # Inicia o Word
        print("📂 Iniciando Microsoft Word...")
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        
        # Abre com modo de reparo
        print("🔄 Abrindo arquivo com modo de reparo...")
        # OpenAndRepair = True
        doc = word.Documents.Open(abs_path, False, False, False, "", "", False, "", "", 1, 0, True)
        
        print("💾 Salvando arquivo reparado...")
        doc.Save()
        
        print("🗑️ Fechando documento...")
        doc.Close()
        
        print("✅ Arquivo reparado com sucesso!")
        
        word.Quit()
        return True
        
    except Exception as e:
        print(f"❌ Erro: {e}")
        try:
            word.Quit()
        except:
            pass
        return False

def main():
    print("=" * 70)
    print("WORD REPAIR - Reparo usando Microsoft Word")
    print("=" * 70)
    
    templates = [
        "templates/IMAGEM-E-VOZ-ALUNO-PUBLICIDADE_1.docx",
        "templates/IMAGEM-E-VOZ-ALUNO-INSTITUCIONAL_1.docx"
    ]
    
    success_count = 0
    for template_path in templates:
        if repair_with_word(template_path):
            success_count += 1
    
    print("\n" + "=" * 70)
    print(f"✅ Concluído! {success_count}/{len(templates)} arquivos reparados")
    print("=" * 70)
    
    if success_count == len(templates):
        print("\n🎉 Arquivos reparados pelo Word!")
        print("💡 Tente abrir no Word agora.")
    else:
        print("\n⚠️ Alguns arquivos falharam.")
        print("\n📝 SOLUÇÃO MANUAL:")
        print("1. Abra cada arquivo no Word")
        print("2. Word vai perguntar se quer reparar - clique SIM")
        print("3. Salve o arquivo")
        print("4. Feche o Word")
        print("5. Rode: python fix_image_final.py")

if __name__ == '__main__':
    main()
