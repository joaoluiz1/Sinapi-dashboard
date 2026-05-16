"""
=========================================================
SINAPI Dashboard Pro
Criado por: João Luiz (joaoprefeitura)
Ano: 2026
=========================================================
"""
import streamlit as st
import pandas as pd
import openpyxl
from text_unidecode import unidecode
import re
import io
import os

# 1. Configuração da página
st.set_page_config(page_title="SINAPI Dashboard Pro", layout="wide", page_icon="🏗️")

# --- Constantes ---
ESTADOS = ["AC", "AL", "AM", "AP", "BA", "CE", "DF", "ES", "GO", "MA",
           "MG", "MS", "MT", "PA", "PB", "PE", "PI", "PR", "RJ", "RN",
           "RO", "RR", "RS", "SC", "SE", "SP", "TO"]

TIPOS_ANALISE = {
    'Composições': ['CSD', 'CCD', 'CSE'],
    'Insumos': ['ISD', 'ICD', 'ISE']
}

# --- Configuração do Caminho Relativo para os PDFs ---
# Esta linha descobre automaticamente onde o ficheiro app.py está alojado
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Procura a pasta PDFs_Separados exatamente no mesmo local onde o app.py está
PASTA_PDFS = os.path.join(BASE_DIR, "PDFs_Separados")

# --- Funções de Processamento ---

def formatar_valor_br(valor):
    if pd.isna(valor) or valor == 0: return "---"
    try:
        return f"R$ {float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return valor

@st.cache_data(show_spinner=True)
def carregar_dados_completos(uploaded_file):
    dados = {}
    try:
        content = uploaded_file.read()
        wb_formulas = openpyxl.load_workbook(io.BytesIO(content), data_only=False)
        
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
                        
                        ws = wb_formulas[aba]
                        codigos_reais = []
                        for i in range(len(df)):
                            row_excel = i + 11
                            val_celula = str(ws.cell(row=row_excel, column=2).value).strip()
                            if val_celula.upper().startswith('=HIPERLINK') or val_celula.upper().startswith('=HYPERLINK'):
                                match = re.search(r'[;,]\s*"?([^";,)]+)"?\s*\)$', val_celula)
                                if match: val_celula = match.group(1).strip()
                            if val_celula in ['None', '']: val_celula = str(df["Código"].iloc[i]).strip()
                            codigos_reais.append(val_celula)
                            
                        df["Código"] = codigos_reais
                        df["_desc_limpa"] = df["Descrição"].astype(str).str.lower().apply(unidecode)
                        dados[aba] = df
            
            if "Analítico" in xls.sheet_names:
                dados["Analítico"] = pd.read_excel(xls, sheet_name="Analítico")
                
        return dados
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
        return None

def obter_pdf_bytes(codigo):
    """
    Função adaptada para ambiente web.
    Em vez de abrir o ficheiro no servidor, lê os bytes para os enviar ao utilizador.
    """
    nome_arquivo = f"{str(codigo).zfill(8)}.pdf"
    caminho_pdf = os.path.join(PASTA_PDFS, nome_arquivo)
    
    if os.path.exists(caminho_pdf):
        with open(caminho_pdf, "rb") as f:
            return f.read(), nome_arquivo
    return None, None

def main():
    st.sidebar.title("🏗️ SINAPI Manager")
    uploaded_file = st.sidebar.file_uploader("Escolha a planilha SINAPI", type=["xlsx"])

    if uploaded_file:
        dados = carregar_dados_completos(uploaded_file)
        
        if dados:
            st.title(f"Consulta SINAPI: {uploaded_file.name}")
            
            with st.sidebar.expander("Filtros de Busca", expanded=True):
                tipo_sel = st.selectbox("Tipo de Análise:", list(TIPOS_ANALISE.keys()))
                estado_sel = st.selectbox("Estado (UF):", ESTADOS, index=ESTADOS.index('MT'))
                st.divider()
                filtro_codigo = st.text_input("Filtrar por Código:")
                termos_contem = st.text_area("Contém (Parcial):", height=150, help="Separe as palavras por vírgula ou quebra de linha.")
                termos_nao_contem = st.text_area("Não contém:", height=150, help="Separe as palavras por vírgula ou quebra de linha.")

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
                    lista_dfs.append(temp[["Código", "Descrição", "_desc_limpa", "Unidade", nome_col_preco]])

            if lista_dfs:
                res = lista_dfs[0]
                for d in lista_dfs[1:]:
                    res = res.merge(d, on=["Código", "Descrição", "_desc_limpa", "Unidade"], how="left")

                if (filtro_codigo or termos_contem or termos_nao_contem):
                    if filtro_codigo:
                        res = res[res['Código'].astype(str).str.contains(filtro_codigo.strip())]
                    if termos_contem:
                        for t in [unidecode(x.strip().lower()) for x in termos_contem.split(',') if x.strip()]:
                            res = res[res["_desc_limpa"].str.contains(t, na=False)]
                    if termos_nao_contem:
                        for t in [unidecode(x.strip().lower()) for x in termos_nao_contem.split(',') if x.strip()]:
                            res = res[~res["_desc_limpa"].str.contains(rf'\b{t}\b', na=False)]
                    
                res_exibicao = res.drop(columns=["_desc_limpa"])
                
                # --- PAINEL DE AÇÕES ---
                st.markdown("### ⚙️ Detalhar Item")
                
                opcoes_select = res_exibicao.apply(lambda row: f"{row['Código']} - {row['Descrição']}", axis=1).tolist()
                item_sel = st.selectbox("Selecione para ver detalhes:", ["-- Escolha o Item --"] + opcoes_select)
                
                if item_sel != "-- Escolha o Item --":
                    cod_puro = item_sel.split(" - ")[0].strip()
                    
                    # Lógica Condicional: Insumos (PDF) vs Composições (Analítico)
                    if tipo_sel == 'Insumos':
                        # Utiliza a nova função para buscar o PDF
                        pdf_data, nome_pdf = obter_pdf_bytes(cod_puro)
                        
                        if pdf_data:
                            # Botão de download nativo do Streamlit
                            st.download_button(
                                label=f"📄 Baixar/Ver Ficha Técnica ({cod_puro})",
                                data=pdf_data,
                                file_name=nome_pdf,
                                mime="application/pdf",
                                use_container_width=True
                            )
                        else:
                            st.error(f"Ficha técnica {cod_puro} não encontrada na pasta 'PDFs_Separados'.")
                            
                    elif tipo_sel == 'Composições':
                        if "Analítico" in dados:
                            df_ana = dados["Analítico"]
                            df_filtrado = df_ana[df_ana.iloc[:, 1].astype(str) == str(cod_puro)]
                            
                            if not df_filtrado.empty:
                                st.write(f"**Composição Analítica: {cod_puro}**")
                                st.dataframe(df_filtrado.iloc[:, [3, 4, 5, 6]], hide_index=True, use_container_width=True)
                            else:
                                st.warning("Composição detalhada não encontrada na aba 'Analítico'.")
                
                st.divider()

                # --- TABELA DE RESULTADOS ---
                for col in res_exibicao.columns:
                    if "Preço_" in col:
                        res_exibicao[col] = res_exibicao[col].apply(formatar_valor_br)

                st.write(f"**Itens encontrados:** {len(res_exibicao)}")
                
                if len(res_exibicao) > 0:
                    if len(res_exibicao) > 500:
                        st.warning("⚠️ Mostrando os primeiros 500 resultados.")
                        st.table(res_exibicao.head(500).reset_index(drop=True))
                    else:
                        st.table(res_exibicao.reset_index(drop=True))
                
                csv = res_exibicao.to_csv(index=False).encode('utf-8-sig')
                st.download_button("📂 Baixar CSV", csv, "extração_sinapi.csv", "text/csv")
    else:
        st.info("Carregue a planilha SINAPI no painel lateral.")

    st.sidebar.divider()
    st.sidebar.caption("© 2026 | João Luiz")

if __name__ == "__main__":
    main()