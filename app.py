# app.py
import pandas as pd
import streamlit as st
from io import BytesIO
import zipfile   # <-- usado pra verificar se o .xls é um .xlsx "disfarçado"

# ──────────────────────────────
# Configuração da página
# ──────────────────────────────
st.set_page_config(page_title="Análise de Atendimentos por Especialidade", layout="wide")
st.title("Análise de Atendimentos por Especialidade, Convênio e Data")
st.markdown("_Envie sua planilha de atendimentos para gerar uma tabela agrupada e visualização personalizada._")

# ──────────────────────────────
# Upload do arquivo
# ──────────────────────────────
uploaded_file = st.file_uploader("📎 Envie a planilha (.xls ou .xlsx)", type=["xls", "xlsx"])

if uploaded_file:
    file_name = uploaded_file.name.lower()
    file_ext  = file_name.split(".")[-1]

    try:
        # ────────────────────────────────────────────────────────────────
        # 1) Verifica se o .xls é, na verdade, um .xlsx disfarçado
        #    (um .xls não é ZIP, já o .xlsx é)
        # ────────────────────────────────────────────────────────────────
        file_bytes = BytesIO(uploaded_file.read())        # lê todo o arquivo na memória
        file_bytes.seek(0)

        if file_ext == "xls" and zipfile.is_zipfile(file_bytes):
            st.error(
                "⚠️ O arquivo enviado tem extensão **.xls**, mas internamente é um **.xlsx**. "
                "Renomeie para **.xlsx** ou exporte novamente e tente de novo."
            )
            st.stop()

        # ────────────────────────────────────────────────────────────────
        # 2) Leitura conforme a extensão
        # ────────────────────────────────────────────────────────────────
        if file_ext == "xls":
            df_raw = pd.read_excel(
                file_bytes,
                sheet_name="Report",
                header=None,
                engine="xlrd"            # usa xlrd 1.2.0
            )
        else:  # .xlsx
            df_raw = pd.read_excel(
                file_bytes,
                sheet_name="Report",
                header=None,
                engine="openpyxl"
            )

        # ────────────────────────────────────────────────────────────────
        # 3) Seu processamento original
        # ────────────────────────────────────────────────────────────────
        # Seleciona colunas: Especialidade (9), Convênio (6), Data (8)
        df = df_raw.iloc[:, [9, 6, 8]].copy()
        df.columns = ["Especialidade", "Convenio", "Data"]

        # Remove linhas vazias
        df.dropna(subset=["Especialidade", "Convenio", "Data"], inplace=True)

        # Classifica tipo de convênio
        df["TipoConvenio"] = df["Convenio"].apply(
            lambda x: "GRUPO" if "AMIL" in str(x).upper() else "EXTRA GRUPO"
        )

        # Converte data e remove inválidas
        df["Data"] = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce").dt.date
        df.dropna(subset=["Data"], inplace=True)

        # Agrupa dados
        resumo = (
            df.groupby(["Especialidade", "TipoConvenio", "Data"])
              .size()
              .reset_index(name="Total")
        )

        # Tabela pivô formatada
        tabela_formatada = resumo.pivot_table(
            index=["Especialidade", "TipoConvenio"],
            columns="Data",
            values="Total",
            fill_value=0
        )

        # ────────────────────────────────────────────────────────────────
        # 4) Exibição e download
        # ────────────────────────────────────────────────────────────────
        st.subheader("📊 Tabela de Atendimentos")
        st.dataframe(tabela_formatada, use_container_width=True)

        # Download em Excel
        buffer = BytesIO()
        tabela_formatada.to_excel(buffer)
        buffer.seek(0)

        st.download_button(
            label="⬇️ Baixar Tabela em Excel",
            data=buffer,
            file_name="atendimentos_formatados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # ────────────────────────────────────────────────────────────────
        # 5) Análise de volume de atendimentos por dia
        # ────────────────────────────────────────────────────────────────
        total_por_dia = df.groupby("Data").size().reset_index(name="TotalPacientes")

        if not total_por_dia.empty:
            dia_mais  = total_por_dia.sort_values("TotalPacientes", ascending=False).iloc[0]
            dia_menos = total_por_dia.sort_values("TotalPacientes", ascending=True ).iloc[0]

            st.markdown("### 🔍 Análise de Volume de Atendimentos")
            st.markdown(
                f"📈 **Maior movimento:** {dia_mais['Data'].strftime('%d/%m/%Y')} "
                f"com **{dia_mais['TotalPacientes']} pacientes**"
            )
            st.markdown(
                f"📉 **Menor movimento:** {dia_menos['Data'].strftime('%d/%m/%Y')} "
                f"com **{dia_menos['TotalPacientes']} pacientes**"
            )

    except Exception as e:
        st.error(f"❌ Erro ao processar o arquivo: {e}")
        st.stop()

# Rodapé
st.markdown("---")
st.markdown(
    "**Desenvolvido por Vitor Cavalcante Gomes para Luana – "
    "vitor.cavalcante@amil.com.br – www.vitorgomes.tech**"
)
