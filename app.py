# import pandas as pd
# import os 
# from tkinter import Tk
# from tkinter.filedialog import askopenfilename

# # === CONFIGURAÇÃO INICIAL ===
# Tk().withdraw()  # esconde a janela principal do Tkinter

# # 1. SELECIONAR O ARQUIVO DE DADOS (PLAN1)
# arquivo_dados = askopenfilename(
#     title="Selecione o arquivo Excel com a aba 'Plan1'",
#     filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
# )

# if not arquivo_dados:
#     print("❌ Nenhum arquivo de dados selecionado.")
#     exit()

# # 2. DEFINIR O CAMINHO PARA O ARQUIVO 'OBRAS.XLSX'
# arquivo_mapa = "obras.xlsx"

# if not os.path.exists(arquivo_mapa):
#     print(f"❌ Arquivo de mapa '{arquivo_mapa}' não encontrado na pasta de execução.")
#     print("Por favor, verifique se 'obras.xlsx' está na mesma pasta do script.")
#     exit()

# # === LER AS ABAS ===
# try:
#     dtype_map = {'CONVENIO': str}
#     plan1 = pd.read_excel(arquivo_dados, sheet_name="Plan1", dtype=dtype_map)
#     obras = pd.read_excel(arquivo_mapa, sheet_name="Obras")
# except ValueError as e:
#     print(f"❌ Erro ao ler abas. Verifique a existência das abas 'Plan1' e 'Obras': {e}")
#     exit()
# except Exception as e:
#      print(f"❌ Erro de leitura. Certifique-se de que os arquivos não estão abertos: {e}")
#      exit()

# # === NORMALIZAR NOMES DAS COLUNAS ===
# plan1.columns = plan1.columns.str.strip().str.upper()
# obras.columns = obras.columns.str.strip().str.upper()

# # =========================================================
# # === TRATAMENTO DO FORMATO DA DATA (NOVO!) ===
# # =========================================================

# NOME_COLUNA_DATA = "DATA" # Altere esta string se o nome real da coluna for diferente

# if NOME_COLUNA_DATA in plan1.columns:
#     try:
#         # 1. Converte para datetime
#         plan1[NOME_COLUNA_DATA] = pd.to_datetime(plan1[NOME_COLUNA_DATA], errors='coerce')
        
#         # 2. Formata para string "DD/MM/YYYY" e ignora NaT (Not a Time)
#         plan1[NOME_COLUNA_DATA] = plan1[NOME_COLUNA_DATA].dt.strftime('%d/%m/%Y').fillna('')
#         print(f"✔️ Coluna '{NOME_COLUNA_DATA}' formatada para string 'DD/MM/YYYY'.")
#     except Exception as e:
#         print(f"❌ Não foi possível formatar a coluna '{NOME_COLUNA_DATA}'. Erro: {e}")
# else:
#     print(f"ℹ️ Coluna '{NOME_COLUNA_DATA}' não encontrada para formatação de data.")

# # =========================================================
# # O RESTANTE DO CÓDIGO PERMANECE O MESMO
# # =========================================================

# # === VERIFICAÇÃO DE COLUNAS ESSENCIAIS ===
# colunas_plan1 = ["CONTA_CREDITO", "CONTA_DEBITO"]
# colunas_obras = ["CONTAS", "OBRAS"]

# for col in colunas_plan1:
#     if col not in plan1.columns:
#         print(f"❌ Coluna essencial '{col}' não encontrada na aba 'Plan1'.")
#         exit()

# for col in colunas_obras:
#     if col not in obras.columns:
#         print(f"❌ Coluna essencial '{col}' não encontrada na aba 'Obras' (arquivo '{arquivo_mapa}').")
#         exit()

# # === CRIAR UM DICIONÁRIO DE CONTAS → OBRAS ===
# mapa_obras = obras.drop_duplicates(subset=['CONTAS']).set_index("CONTAS")["OBRAS"].to_dict()

# # === PREENCHER AS COLUNAS DE OBRA NA PLAN1 ===
# # 1. OBRA_CREDITO
# if "OBRA_CREDITO" not in plan1.columns:
#     plan1["OBRA_CREDITO"] = ""

# plan1["OBRA_CREDITO"] = plan1["CONTA_CREDITO"].map(mapa_obras)
# plan1["OBRA_CREDITO"] = plan1["OBRA_CREDITO"].fillna("")

# # 2. OBRA_DEBITO
# if "OBRA_DEBITO" not in plan1.columns:
#     plan1["OBRA_DEBITO"] = ""

# plan1["OBRA_DEBITO"] = plan1["CONTA_DEBITO"].map(mapa_obras)
# plan1["OBRA_DEBITO"] = plan1["OBRA_DEBITO"].fillna("")


# # =========================================================
# # === TRATAMENTO E SALVAMENTO (AJUSTE PARA .xls) ===
# # =========================================================

# nome_base, extensao = os.path.splitext(arquivo_dados)
# extensao = extensao.lower()

# if extensao == '.xls':
#     arquivo_saida = nome_base + "_ATUALIZADO.xlsx"
#     print(f"⚠️ O arquivo original é '.xls'. O resultado será salvo em um novo arquivo '.xlsx'.")
    
#     try:
#         with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as writer:
#             plan1.to_excel(writer, sheet_name="Plan1", index=False)
        
#         print(f"   Arquivo de saída criado: {arquivo_saida}")
#     except Exception as e:
#         print(f"❌ Erro ao salvar o novo arquivo '{arquivo_saida}': {e}")
#         exit()

# else: 
#     arquivo_saida = arquivo_dados
    
#     try:
#         with pd.ExcelWriter(arquivo_saida, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
#             plan1.to_excel(writer, sheet_name="Plan1", index=False)
        
#         print(f"   Arquivo atualizado: {arquivo_saida}")
#     except Exception as e:
#         print(f"❌ Erro ao salvar o arquivo '{arquivo_saida}'. Verifique se está fechado: {e}")
#         exit()


# print("✅ Colunas 'OBRA_CREDITO' e 'OBRA_DEBITO' atualizadas com sucesso na aba 'Plan1'!")
# print("ℹ️ Contas não encontradas foram marcadas como VAZIO.")


import streamlit as st
import pandas as pd
import io
import os

# --- Funções de Processamento ---

# Usar st.cache_data garante que o processamento só é refeito se os dados mudarem.
@st.cache_data
def processar_planilha(plan1_df, obras_df):
    """
    Executa a lógica de normalização, mapeamento de obras e formatação de data.
    Retorna o DataFrame processado.
    """
    
    st.info("Iniciando processamento dos dados. Aguarde...")

    # Usar .copy() para garantir que a função seja pura e não altere o cache
    plan1_df = plan1_df.copy()
    obras_df = obras_df.copy()

    # === NORMALIZAR NOMES DAS COLUNAS ===
    plan1_df.columns = plan1_df.columns.str.strip().str.upper()
    obras_df.columns = obras_df.columns.str.strip().str.upper()

    # === VERIFICAÇÃO DE COLUNAS ESSENCIAIS ===
    colunas_plan1 = ["CONTA_CREDITO", "CONTA_DEBITO"]
    colunas_obras = ["CONTAS", "OBRAS"]

    for col in colunas_plan1:
        if col not in plan1_df.columns:
            st.error(f"❌ Coluna essencial '{col}' não encontrada na Planilha de Dados.")
            return None

    for col in colunas_obras:
        if col not in obras_df.columns:
            st.error(f"❌ Coluna essencial '{col}' não encontrada na Planilha de Obras.")
            return None
    
    # =========================================================
    # === TRATAMENTO DO FORMATO DA DATA ===
    # =========================================================
    NOME_COLUNA_DATA = "DATA"
    
    if NOME_COLUNA_DATA in plan1_df.columns:
        try:
            plan1_df[NOME_COLUNA_DATA] = pd.to_datetime(plan1_df[NOME_COLUNA_DATA], errors='coerce')
            plan1_df[NOME_COLUNA_DATA] = plan1_df[NOME_COLUNA_DATA].dt.strftime('%d/%m/%Y').fillna('')
            # st.success(f"✔️ Coluna '{NOME_COLUNA_DATA}' formatada para string 'DD/MM/YYYY'.") # Sucesso será dado no final
        except Exception:
            st.warning(f"❌ Não foi possível formatar a coluna '{NOME_COLUNA_DATA}'.")

    # =========================================================
    # === TRATAMENTO DA COLUNA CONVENIO (GARANTIR STRING) ===
    if 'CONVENIO' in plan1_df.columns:
        plan1_df['CONVENIO'] = plan1_df['CONVENIO'].astype(str).str.strip()

    # === CRIAR UM DICIONÁRIO DE CONTAS → OBRAS ===
    mapa_obras = obras_df.drop_duplicates(subset=['CONTAS']).set_index("CONTAS")["OBRAS"].to_dict()

    # === PREENCHER AS COLUNAS DE OBRA NA PLAN1 ===
    
    # 1. OBRA_CREDITO
    plan1_df["OBRA_CREDITO"] = plan1_df.get("OBRA_CREDITO", "")
    plan1_df["OBRA_CREDITO"] = plan1_df["CONTA_CREDITO"].map(mapa_obras)
    plan1_df["OBRA_CREDITO"] = plan1_df["OBRA_CREDITO"].fillna("")

    # 2. OBRA_DEBITO
    plan1_df["OBRA_DEBITO"] = plan1_df.get("OBRA_DEBITO", "")
    plan1_df["OBRA_DEBITO"] = plan1_df["CONTA_DEBITO"].map(mapa_obras)
    plan1_df["OBRA_DEBITO"] = plan1_df["OBRA_DEBITO"].fillna("")
    
    st.success("✅ Processamento de mapeamento concluído!")

    return plan1_df

@st.cache_data
def to_excel(df_processado):
    """Converte o DataFrame processado para um buffer de Excel."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_processado.to_excel(writer, sheet_name='Plan1_Atualizada', index=False)
    output.seek(0)
    return output

# --- Interface Streamlit ---

st.set_page_config(
    page_title="Mapeamento de Obras em Planilha Excel",
    layout="centered"
)

st.title("⚙️ Ferramenta de Mapeamento de Obras")
st.markdown("Faça o upload da Planilha de Dados (`Plan1`). O processamento iniciará automaticamente.")

# --- Configuração do Arquivo de Mapeamento Fixo ---
ARQUIVO_MAPA_NOME = "obras.xlsx"
obras_df = None

# Tenta carregar o arquivo de mapeamento do disco
if os.path.exists(ARQUIVO_MAPA_NOME):
    try:
        obras_df = pd.read_excel(ARQUIVO_MAPA_NOME, sheet_name="Obras")
        st.info(f"✔️ Planilha de Mapeamento ('Obras') carregada de `{ARQUIVO_MAPA_NOME}`.")
    except Exception as e:
        st.error(f"❌ Erro ao ler o arquivo '{ARQUIVO_MAPA_NOME}'. Verifique se a aba 'Obras' existe e se o arquivo não está corrompido.")
        obras_df = None
else:
    st.error(f"❌ Arquivo de mapa '{ARQUIVO_MAPA_NOME}' não encontrado na pasta de execução.")
    st.markdown("⚠️ **Atenção:** Coloque o arquivo `obras.xlsx` na mesma pasta onde você executa o Streamlit.")


# --- Upload do Arquivo de Dados ---
st.subheader("1. Upload da Planilha de Dados (Contém a aba 'Plan1')")
uploaded_file_data = st.file_uploader(
    "Selecione o arquivo com os dados (Transferencia.xls/xlsx, etc.)",
    type=['xlsx', 'xls']
)


plan1_df = None

# --- Leitura do Arquivo de Dados ---
if uploaded_file_data:
    try:
        # Força a coluna 'CONVENIO' a ser lida como string
        dtype_map = {'CONVENIO': str, 'FORMA_PGTO': str} 
        plan1_df = pd.read_excel(uploaded_file_data, sheet_name="Plan1", dtype=dtype_map)
        st.success("✔️ Planilha de Dados ('Plan1') lida com sucesso. Iniciando mapeamento...")

    except ValueError:
        st.error(f"❌ Erro de Leitura: Verifique se a aba 'Plan1' está nomeada corretamente no arquivo de dados.")
        plan1_df = None
    except Exception as e:
        st.error(f"❌ Erro inesperado durante a leitura do arquivo de dados: {e}")
        plan1_df = None

# --- Processamento e Download AUTOMÁTICO ---
if plan1_df is not None and obras_df is not None:
    st.subheader("2. Download do Resultado")

    # O processamento é chamado aqui, acionado pelo upload
    df_processado = processar_planilha(plan1_df, obras_df)
    
    if df_processado is not None:
        
        # Converte para Excel no buffer
        excel_data = to_excel(df_processado)

        # Mostra o botão de download
        st.download_button(
            label="📥 Baixar Planilha Atualizada (.xlsx)",
            data=excel_data,
            file_name="Planilha_Mapeada_Atualizada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="final_download_button",
            type="primary" # Destaca o botão
        )
        
        st.markdown("---")
        st.markdown("**Prévia dos Dados (5 Primeiras Linhas):**")
        st.dataframe(df_processado.head(5))
        st.balloons()