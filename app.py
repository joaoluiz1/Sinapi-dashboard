import streamlit as st
import pandas as pd
from text_unidecode import unidecode
import re
import io

# 1. Configuração da página para ocupar a largura total
st.set_page_config(page_title="SINAPI Dashboard Pro", layout="wide", page_icon="🏗️")

# --- Constantes ---
ESTADOS = ["AC", "AL", "AM", "AP", "BA", "CE", "DF", "ES", "GO", "MA",
           "MG", "MS", "MT", "PA", "PB", "PE", "PI", "PR", "RJ", "RN",
           "RO", "RR", "RS", "SC", "SE", "SP", "TO"]

TIPOS_ANALISE = {
    'Composições': ['CSD', 'CCD', 'CSE'],
    'Insumos': ['ISD', 'ICD', 'ISE']
}

# --- Funções de Processamento ---

def formatar_valor_br(valor):
    if pd.isna(valor) or valor == 0: return "---"
    try:
        return f"R$ {float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return valor

@st.cache_data(show_spinner=False)
def carregar_dados_upload(uploaded_file):
    """Lê o ficheiro carregado pelo usuário via botão"""
    dados = {}
    try:
        # Lendo o ficheiro em memória
        content = uploaded_file.read()
        with pd.ExcelFile(io.BytesIO(content)) as xls:
            for tipo, abas in TIPOS_ANALISE.items():
                for aba in abas:
                    if aba in xls.sheet_names:
                        df = pd.read_excel(xls, sheet_name=aba, skiprows=9)
                        df.columns = [str(c) for c in df.columns]
                        df = df.rename(columns={
                            df.columns[1]: "Código",
                            df.columns[2]: "Descrição",
                            df.columns[3]: "Unidade"
                        })
                        df["Código"] = df["Código"].astype(str).str.strip()
                        dados[aba] = df
        return dados
    except Exception as e:
        st.error(f"Erro ao processar o ficheiro: {e}")
        return None

# --- Interface ---

def main():
    st.sidebar.title("🏗️ SINAPI Manager")
    
    # BOTÃO PARA ESCOLHER O FICHEIRO (O QUE VOCÊ PEDIU)
    st.sidebar.subheader("1. Base de Dados")
    uploaded_file = st.sidebar.file_uploader(
        "Clique para escolher a planilha SINAPI", 
        type=["xlsx"],
        help="Selecione o arquivo Excel do SINAPI do seu diretório."
    )

    if uploaded_file:
        # Carrega os dados do ficheiro que o utilizador escolheu
        dados = carregar_dados_upload(uploaded_file)
        
        if dados:
            st.title(f"Consulta SINAPI: {uploaded_file.name}")
            
            # Filtros na Barra Lateral
            with st.sidebar.expander("Filtros de Busca", expanded=True):
                tipo_sel = st.selectbox("Tipo de Análise:", list(TIPOS_ANALISE.keys()))
                estado_sel = st.selectbox("Estado (UF):", ESTADOS, index=ESTADOS.index('MT'))
                
                st.divider()
                filtro_codigo = st.text_input("Filtrar por Código:")
                termos_contem = st.text_input("Contém (Parcial):")
                termos_nao_contem = st.text_input("Não contém (Palavra Exata):")

            # Lógica de cruzamento
            abas_alvo = TIPOS_ANALISE[tipo_sel]
            lista_dfs = []

            for aba in abas_alvo:
                if aba in dados:
                    temp = dados[aba].copy()
                    is_comp = (tipo_sel == 'Composições')
                    idx_uf = ESTADOS.index(estado_sel)
                    col_idx = (idx_uf * 2 + 4) if is_comp else (idx_uf + 5)
                    nome_col_preco = f"Preço_{aba}"
                    temp[nome_col_preco] = temp[temp.columns[col_idx]]
                    lista_dfs.append(temp[["Código", "Descrição", "Unidade", nome_col_preco]])

            if lista_dfs:
                res = lista_dfs[0]
                for d in lista_dfs[1:]:
                    res = res.merge(d, on=["Código", "Descrição", "Unidade"], how="left")

                # --- REGRAS DE FILTRO ---
                
                # 1. Filtro de Código
                if filtro_codigo:
                    res = res[res['Código'].str.contains(filtro_codigo.strip())]

                # 2. Filtros de Texto
                if termos_contem or termos_nao_contem:
                    res["_desc_limpa"] = res["Descrição"].str.lower().apply(unidecode)
                    
                    if termos_contem:
                        for t in [unidecode(x.strip().lower()) for x in termos_contem.split(',') if x.strip()]:
                            # PARCIAL: Acha parte da palavra
                            res = res[res["_desc_limpa"].str.contains(t, na=False)]
                    
                    if termos_nao_contem:
                        for t in [unidecode(x.strip().lower()) for x in termos_nao_contem.split(',') if x.strip()]:
                            # EXATO: Exclui apenas a palavra inteira (\b)
                            res = res[~res["_desc_limpa"].str.contains(rf'\b{t}\b', na=False)]
                    
                    res = res.drop(columns=["_desc_limpa"])

                # Formatação de Moeda
                for col in res.columns:
                    if "Preço_" in col:
                        res[col] = res[col].apply(formatar_valor_br)

                # --- EXIBIÇÃO ---
                st.write(f"**Itens encontrados:** {len(res)}")
                
                # Para evitar o erro do hide_index, limpamos o index antes de exibir
                # Removi o 'height' para que a planilha role na página inteira
                st.dataframe(
                    res.reset_index(drop=True), 
                    use_container_width=True
                )
                
                # Download
                csv = res.to_csv(index=False).encode('utf-8-sig')
                st.download_button("📂 Baixar CSV", csv, "extração_sinapi.csv", "text/csv")
    else:
        # Tela inicial quando nenhum ficheiro foi escolhido
        st.info("👋 Olá! Clique no botão à esquerda para carregar a sua planilha SINAPI (.xlsx) e começar.")

if __name__ == "__main__":
    main()