import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="Análise de Atendimentos por Especialidade", layout="wide")
st.title("Análise de Atendimentos por Especialidade, Convênio e Data")

st.markdown("_Envie sua planilha de atendimentos para gerar uma tabela agrupada e visualização personalizada._")

uploaded_file = st.file_uploader("📎 Envie a planilha (.xls ou .xlsx)", type=["xls", "xlsx"])

if uploaded_file:
    file_name = uploaded_file.name

    # Leitura do arquivo (tratamento para .xls e .xlsx)
    if file_name.endswith(".xls"):
        df_temp = pd.read_excel(uploaded_file, sheet_name="Report", engine="xlrd", header=None)
        buffer_xlsx = BytesIO()
        with pd.ExcelWriter(buffer_xlsx, engine="openpyxl") as writer:
            df_temp.to_excel(writer, sheet_name="Report", index=False, header=False)
        buffer_xlsx.seek(0)
        df = pd.read_excel(buffer_xlsx, sheet_name="Report", header=None)
    else:
        df = pd.read_excel(uploaded_file, sheet_name="Report", header=None)

    # Seleciona colunas: Especialidade (9), Convênio (6), Data (8)
    df = df.iloc[:, [9, 6, 8]]
    df.columns = ["Especialidade", "Convenio", "Data"]

    # Remove linhas vazias
    df = df.dropna(subset=["Especialidade", "Convenio", "Data"])

    # Classifica tipo de convênio
    df["TipoConvenio"] = df["Convenio"].apply(lambda x: "GRUPO" if "AMIL" in str(x).upper() else "EXTRA GRUPO")

    # Converte data e remove inválidas
    df["Data"] = pd.to_datetime(df["Data"], dayfirst=True, errors='coerce').dt.date
    df = df.dropna(subset=["Data"])

    # Agrupa dados
    resumo = df.groupby(["Especialidade", "TipoConvenio", "Data"]).size().reset_index(name="Total")

    # Tabela pivô formatada
    tabela_formatada = resumo.pivot_table(
        index=["Especialidade", "TipoConvenio"],
        columns="Data",
        values="Total",
        fill_value=0
    )

    st.subheader("📊 Tabela de Atendimentos")
    st.dataframe(tabela_formatada)

    # Download do Excel
    buffer = BytesIO()
    tabela_formatada.to_excel(buffer)
    buffer.seek(0)

    st.download_button(
        label="⬇️ Baixar Tabela em Excel",
        data=buffer,
        file_name="atendimentos_formatados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Análise de volume de atendimentos por dia
    total_por_dia = df.groupby("Data").size().reset_index(name="TotalPacientes")

    if not total_por_dia.empty:
        dia_mais = total_por_dia.sort_values("TotalPacientes", ascending=False).iloc[0]
        dia_menos = total_por_dia.sort_values("TotalPacientes", ascending=True).iloc[0]

        st.markdown("### 🔍 Análise de Volume de Atendimentos")
        st.markdown(f"📈 **Maior movimento:** {dia_mais['Data'].strftime('%d/%m/%Y')} com **{dia_mais['TotalPacientes']} pacientes**")
        st.markdown(f"📉 **Menor movimento:** {dia_menos['Data'].strftime('%d/%m/%Y')} com **{dia_menos['TotalPacientes']} pacientes**")

st.markdown("---")
st.markdown("**Desenvolvido por Vitor Cavalcante Gomes - vitor.cavalcante@amil.com.br - www.vitorgomes.tech**")
