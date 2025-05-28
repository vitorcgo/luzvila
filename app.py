import pandas as pd
import streamlit as st
from io import BytesIO
import zipfile
import unicodedata
import re

# Configuração da página
st.set_page_config(page_title="Análise de Atendimentos por Especialidade", layout="wide")
st.title("Análise de Atendimentos por Especialidade, Convênio e Data")
st.markdown("_Envie sua planilha de atendimentos para gerar uma tabela agrupada e visualização personalizada._")

# Função para limpar texto: remove acentos, espaços extras e deixa maiúsculo
def limpar_texto(texto):
    if not isinstance(texto, str):
        return ""
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    texto = texto.upper()
    texto = re.sub(r'\s+', ' ', texto).strip()
    return texto

# Função que classifica TipoConvenio conforme presença de "AMIL"
def detectar_tipo_convenio(convenio):
    convenio_limpo = limpar_texto(convenio)
    if "AMIL" in convenio_limpo:
        return "GRUPO"
    else:
        return "EXTRA GRUPO"

# Upload do arquivo
uploaded_file = st.file_uploader("📎 Envie a planilha (.xls ou .xlsx)", type=["xls", "xlsx"])

if uploaded_file:
    file_name = uploaded_file.name.lower()
    file_ext  = file_name.split(".")[-1]

    try:
        file_bytes = BytesIO(uploaded_file.read())
        file_bytes.seek(0)

        if file_ext == "xls" and zipfile.is_zipfile(file_bytes):
            st.error(
                "⚠️ O arquivo enviado tem extensão **.xls**, mas internamente é um **.xlsx**. "
                "Renomeie para **.xlsx** ou exporte novamente e tente de novo."
            )
            st.stop()

        if file_ext == "xls":
            df_raw = pd.read_excel(
                file_bytes,
                sheet_name="Report",
                header=None,
                engine="xlrd"
            )
        else:
            df_raw = pd.read_excel(
                file_bytes,
                sheet_name="Report",
                header=None,
                engine="openpyxl"
            )

        # Mostrar amostras para inspeção
        st.write("### Amostras das colunas originais")
        st.write("Especialidade:")
        st.write(df_raw.iloc[:, 9].head(10))
        st.write("Convenio:")
        st.write(df_raw.iloc[:, 6].head(10))
        st.write("Data:")
        st.write(df_raw.iloc[:, 8].head(10))

        # Seleciona colunas: Especialidade (9), Convenio (6), Data (8)
        df = df_raw.iloc[:, [9, 6, 8]].copy()
        df.columns = ["Especialidade", "Convenio", "Data"]

        # Limpa e normaliza coluna Convenio
        df["Convenio"] = df["Convenio"].astype(str).apply(limpar_texto)

        # Remove linhas com dados faltando ou vazios
        df = df.dropna(subset=["Especialidade", "Convenio", "Data"])
        df = df[(df["Especialidade"].str.strip() != "") & (df["Convenio"].str.strip() != "")]

        # Classifica TipoConvenio
        df["TipoConvenio"] = df["Convenio"].apply(detectar_tipo_convenio)

        # Converte coluna Data para datetime e remove datas inválidas
        df["Data"] = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce").dt.date
        df = df.dropna(subset=["Data"])

        # Inspeção dos dados limpos
        st.write("### Resumo após limpeza")
        st.write(f"Total linhas no arquivo original: {len(df_raw)}")
        st.write(f"Total linhas após limpeza: {len(df)}")
        st.write("Valores nulos por coluna:")
        st.write(df.isnull().sum())
        st.write("Valores únicos em Convenio:")
        st.write(df["Convenio"].unique())

        # Teste agrupamento
        grouped = df.groupby(["Especialidade", "TipoConvenio", "Data"]).size()
        st.write("Total agrupado (soma dos grupos):", grouped.sum())

        # Cria tabela pivô
        resumo = grouped.reset_index(name="Total")
        tabela_formatada = resumo.pivot_table(
            index=["Especialidade", "TipoConvenio"],
            columns="Data",
            values="Total",
            fill_value=0
        )

        # Exibe tabela formatada
        st.subheader("📊 Tabela de Atendimentos")
        st.dataframe(tabela_formatada, use_container_width=True)

        # Botão para download da tabela em Excel
        buffer = BytesIO()
        tabela_formatada.to_excel(buffer)
        buffer.seek(0)

        st.download_button(
            label="⬇️ Baixar Tabela em Excel",
            data=buffer,
            file_name="atendimentos_formatados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Análise de volume total por dia
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
