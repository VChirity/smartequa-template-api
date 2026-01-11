import streamlit as st
import streamlit.components.v1 as components
from docxtpl import DocxTemplate, RichText
from faker import Faker
import requests
import io
import os
import re
from datetime import datetime
from num2words import num2words

st.set_page_config(page_title="Smart Equação", page_icon="📝", layout="wide")
fake = Faker('pt_BR')

st.markdown("""
    <style>
    .stButton>button {
        width: 100%;
        border-radius: 5px;
        height: 3em;
    }
    </style>
""", unsafe_allow_html=True)

# JavaScript para formatação em tempo real
components.html("""
<script>
(function() {
    function formatarData() {
        const inputs = parent.document.querySelectorAll('input[aria-label*="Data de Nascimento"]');
        inputs.forEach(input => {
            if (!input.hasAttribute('data-date-formatted')) {
                input.setAttribute('data-date-formatted', 'true');
                input.addEventListener('input', function(e) {
                    let valor = e.target.value.replace(/[^0-9]/g, '');
                    let formatado = '';
                    
                    if (valor.length >= 1) {
                        formatado = valor.substring(0, 2);
                    }
                    if (valor.length >= 3) {
                        formatado += '/' + valor.substring(2, 4);
                    }
                    if (valor.length >= 5) {
                        formatado += '/' + valor.substring(4, 8);
                    }
                    
                    const cursorPos = e.target.selectionStart;
                    const oldLength = e.target.value.length;
                    
                    if (formatado !== e.target.value) {
                        e.target.value = formatado;
                        // Ajusta cursor
                        const newLength = formatado.length;
                        const diff = newLength - oldLength;
                        e.target.setSelectionRange(cursorPos + diff, cursorPos + diff);
                    }
                });
            }
        });
    }
    
    // Executa repetidamente para pegar inputs renderizados dinamicamente
    setInterval(formatarData, 300);
})();
</script>
""", height=0)

def limpar_cep(cep):
    return re.sub(r'\D', '', str(cep))

def formatar_nome(nome):
    if not nome: return ""
    preposicoes = ['da', 'de', 'do', 'das', 'dos', 'e']
    palavras = nome.lower().split()
    return " ".join([p if p in preposicoes else p.capitalize() for p in palavras])

def buscar_cep(tipo):
    cep_key = f"cep_{tipo}"
    cep_digitado = st.session_state.get(cep_key, "")
    cep_limpo = limpar_cep(cep_digitado)
    
    if len(cep_limpo) == 8:
        try:
            with st.spinner("Buscando CEP..."):
                r = requests.get(f"https://viacep.com.br/ws/{cep_limpo}/json/", timeout=5)
                data = r.json()
                if "erro" not in data:
                    st.session_state[f"endereco_{tipo}"] = f"{data['logradouro']}, "
                    st.session_state[f"bairro_{tipo}"] = data['bairro']
                else:
                    st.toast(f"CEP {tipo} não encontrado.", icon="❌")
        except:
            st.toast("Erro de conexão com ViaCEP.", icon="⚠️")

def converter_desconto():
    """Converte o desconto numérico para extenso automaticamente"""
    desconto = st.session_state.get('desconto', '')
    if desconto and desconto.strip():
        try:
            valor = float(desconto)
            extenso = num2words(valor, lang='pt_BR')
            st.session_state['desconto_extenso'] = extenso
        except:
            pass

def formatar_data(key):
    """Formata data automaticamente com barras (DD/MM/AAAA)"""
    data = st.session_state.get(key, '')
    # Remove tudo que não é número
    apenas_numeros = re.sub(r'\D', '', data)
    
    # Formata com barras
    if len(apenas_numeros) <= 2:
        formatado = apenas_numeros
    elif len(apenas_numeros) <= 4:
        formatado = f"{apenas_numeros[:2]}/{apenas_numeros[2:]}"
    elif len(apenas_numeros) <= 8:
        formatado = f"{apenas_numeros[:2]}/{apenas_numeros[2:4]}/{apenas_numeros[4:8]}"
    else:
        formatado = f"{apenas_numeros[:2]}/{apenas_numeros[2:4]}/{apenas_numeros[4:8]}"
    
    # Atualiza apenas se mudou
    if formatado != data:
        st.session_state[key] = formatado

TEMPLATES_DISPONIVEIS = {
    "📄 Contrato Padrão 2025": "template_contrato2025_2.docx",
    "💰 Contrato com Desconto": "template_contratoDESCONTO2025_2.docx",
    "📸 Termo de Imagem (Publicidade)": "IMAGEM-E-VOZ-ALUNO-PUBLICIDADE_1.docx",
    "🏫 Termo de Imagem (Institucional)": "IMAGEM-E-VOZ-ALUNO-INSTITUCIONAL_1.docx"
}

def preencher_faker():
    st.session_state.nome_resp1 = fake.name()
    st.session_state.cpf_resp1 = fake.cpf()
    st.session_state.cep_resp1 = fake.postcode()
    st.session_state.endereco_resp1 = f"{fake.street_name()}, {fake.building_number()}"
    st.session_state.bairro_resp1 = fake.bairro()
    st.session_state.naturalidade_resp1 = "Rio de Janeiro"
    st.session_state.nasc_resp1 = fake.date_of_birth(minimum_age=25, maximum_age=60).strftime('%d/%m/%Y')
    
    st.session_state.nome_resp2 = fake.name_female()
    st.session_state.cpf_resp2 = fake.cpf()
    st.session_state.cep_resp2 = fake.postcode()
    st.session_state.endereco_resp2 = f"{fake.street_name()}, {fake.building_number()}"
    st.session_state.bairro_resp2 = fake.bairro()
    
    st.session_state.nome_aluno = fake.first_name() + " " + fake.last_name()
    st.session_state.naturalidade_aluno = "Rio de Janeiro"
    st.session_state.nasc_aluno = fake.date_of_birth(minimum_age=6, maximum_age=15).strftime('%d/%m/%Y')
    st.session_state.cpf_aluno = fake.cpf()
    st.session_state.ano_letivo = "2026"
    
    meses = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro']
    hoje = datetime.now()
    st.session_state.data_extenso = f"{hoje.day} de {meses[hoje.month-1]} de {hoje.year}"

def gerar_documento(template_filename):
    template_path = f"templates/{template_filename}"
    if not os.path.exists(template_path):
        template_path = f"assets/{template_filename}"
    
    if not os.path.exists(template_path):
        return None
    
    try:
        doc = DocxTemplate(template_path)
        
        def rt(s):
            return RichText(str(s) if s else '', bold=True)
        
        context = {
            'responsavel1': rt(formatar_nome(st.session_state.get('nome_resp1', ''))),
            'cpf_responsavel': rt(st.session_state.get('cpf_resp1', '')),
            'endereco_completo': rt(st.session_state.get('endereco_resp1', '')),
            'bairro': rt(st.session_state.get('bairro_resp1', '')),
            'cep': rt(st.session_state.get('cep_resp1', '')),
            'naturalidade_resp1': rt(st.session_state.get('naturalidade_resp1', '')),
            'nasc_resp1': rt(st.session_state.get('nasc_resp1', '')),
            'responsavel2': rt(formatar_nome(st.session_state.get('nome_resp2', ''))),
            'cpf2': rt(st.session_state.get('cpf_resp2', '')),
            'endereco2': rt(st.session_state.get('endereco_resp2', '')),
            'bairro2': rt(st.session_state.get('bairro_resp2', '')),
            'cep2': rt(st.session_state.get('cep_resp2', '')),
            'nome_aluno': rt(formatar_nome(st.session_state.get('nome_aluno', ''))),
            'ano': rt(st.session_state.get('ano_escolar', '')),
            'ano_letivo': rt(st.session_state.get('ano_letivo', '')),
            'data_extenso': rt(st.session_state.get('data_extenso', '')),
            'naturalidade_aluno': rt(st.session_state.get('naturalidade_aluno', '')),
            'nasc_aluno': rt(st.session_state.get('nasc_aluno', '')),
            'cpf_aluno': rt(st.session_state.get('cpf_aluno', '')),
            'desconto': rt(st.session_state.get('desconto', '')),
            'desconto_extenso': rt(st.session_state.get('desconto_extenso', ''))
        }
        doc.render(context)
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio
    except Exception as e:
        st.error(f"Erro ao gerar documento: {e}")
        return None

st.title("🍊 Smart Equação - Gerador")

col_top, col_btn = st.columns([3, 1])
with col_btn:
    st.button("🧪 Preencher Teste", on_click=preencher_faker, type="secondary")

st.markdown("### 👤 Responsável 1 (Financeiro)")
c1, c2, c3 = st.columns([1, 2, 1])
c1.text_input("CEP", key="cep_resp1", on_change=buscar_cep, args=('resp1',), max_chars=9)
c2.text_input("Endereço", key="endereco_resp1")
c3.text_input("Bairro", key="bairro_resp1")
c4, c5 = st.columns([2, 1])
c4.text_input("Nome Completo", key="nome_resp1")
c5.text_input("CPF", key="cpf_resp1")
c6, c7 = st.columns([2, 1])
c6.text_input("Naturalidade", key="naturalidade_resp1", value="Rio de Janeiro", placeholder="Ex: São Paulo/SP")
c7.text_input("Data de Nascimento", key="nasc_resp1", placeholder="DD/MM/AAAA", help="Formato: 01/01/1980", max_chars=10, on_change=formatar_data, args=('nasc_resp1',))

st.markdown("---")
st.markdown("### 👥 Responsável 2 (Opcional)")
d1, d2, d3 = st.columns([1, 2, 1])
d1.text_input("CEP", key="cep_resp2", on_change=buscar_cep, args=('resp2',), max_chars=9)
d2.text_input("Endereço", key="endereco_resp2")
d3.text_input("Bairro", key="bairro_resp2")
d4, d5 = st.columns([2, 1])
d4.text_input("Nome Completo", key="nome_resp2")
d5.text_input("CPF", key="cpf_resp2")

st.markdown("---")
st.markdown("### 🎓 Aluno")
e1, e2, e3 = st.columns([2, 1, 1])
e1.text_input("Nome Aluno", key="nome_aluno")
e2.selectbox("Série", ["1º Ano", "2º Ano", "3º Ano", "4º Ano", "5º Ano", "6º Ano", "7º Ano", "8º Ano", "9º Ano"], key="ano_escolar")
e3.text_input("Ano Letivo", key="ano_letivo", value="2025")
st.text_input("Data Extenso", key="data_extenso")
e4, e5, e6 = st.columns([2, 1, 1])
e4.text_input("Naturalidade", key="naturalidade_aluno", value="Rio de Janeiro", placeholder="Ex: Rio de Janeiro/RJ")
e5.text_input("Data de Nascimento", key="nasc_aluno", placeholder="DD/MM/AAAA", help="Formato: 15/03/2010", max_chars=10, on_change=formatar_data, args=('nasc_aluno',))
e6.text_input("CPF do Aluno", key="cpf_aluno", placeholder="000.000.000-00")

st.markdown("---")
st.markdown("### 💰 Informações Financeiras (Apenas para contratos com desconto)")
f1, f2 = st.columns([1, 2])
f1.text_input("Desconto (%)", key="desconto", placeholder="Ex: 10", on_change=converter_desconto)
f2.text_input("Desconto por Extenso", key="desconto_extenso", placeholder="Ex: dez", help="Preenchido automaticamente ao digitar o desconto")

st.markdown("---")
st.markdown("## 📂 Documentos Disponíveis")

if not st.session_state.get('nome_resp1') or not st.session_state.get('nome_aluno'):
    st.warning("⚠️ Preencha pelo menos o Responsável 1 e o Aluno para habilitar os downloads.")
else:
    st.info(f"📋 Dados salvos: **{st.session_state.get('nome_aluno')}** - {st.session_state.get('ano_escolar', 'N/A')}")
    
    documentos_encontrados = 0
    
    for nome_exibido, arquivo_template in TEMPLATES_DISPONIVEIS.items():
        template_path = f"templates/{arquivo_template}"
        if not os.path.exists(template_path):
            template_path = f"assets/{arquivo_template}"
        
        if os.path.exists(template_path):
            documentos_encontrados += 1
            
            col1, col2 = st.columns([3, 1])
            with col1:
                st.markdown(f"**{nome_exibido}**")
            with col2:
                bio = gerar_documento(arquivo_template)
                if bio:
                    nome_aluno_limpo = st.session_state.get('nome_aluno', 'Aluno').replace(' ', '_')
                    
                    # Remove emojis e formata nome do documento (usa nome completo para diferenciar)
                    tipo_doc = re.sub(r'[^\w\s()-]', '', nome_exibido)  # Remove apenas emojis
                    tipo_doc = tipo_doc.strip().replace(' ', '_')
                    
                    st.download_button(
                        label=f"⬇️ Baixar",
                        data=bio.getvalue(),
                        file_name=f"{tipo_doc}_{nome_aluno_limpo}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"download_{arquivo_template}"
                    )
    
    if documentos_encontrados == 0:
        st.warning("⚠️ Nenhum template encontrado nas pastas 'templates/' ou 'assets/'. Adicione arquivos .docx para habilitar os downloads.")
