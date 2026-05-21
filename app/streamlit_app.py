import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
sys.path.append(str(ROOT_DIR))

import streamlit as st
import pandas as pd

from src.core.extrair_alunos_pdf import extrair_alunos_pdf
from src.core.processar_moodle import processar_moodle
from src.core.ranking_engajamento import finalizar_ranking
from src.core.exportar_relatorio_word import exportar_relatorio_word


st.set_page_config(
    page_title="AVA Intelligence | Athena Scientific",
    page_icon="📊",
    layout="wide"
)


st.title("📊 AVA Intelligence")
st.subheader("Athena Scientific — Monitoramento Acadêmico Moodle")


UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("outputs")

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


st.markdown("""
Sistema institucional para monitoramento de engajamento discente no Ambiente Virtual de Aprendizagem,
com extração automática de alunos, cruzamento com dados do Moodle, ranking de participação,
classificação de risco acadêmico, parecer pedagógico automatizado e exportação de relatório técnico.
""")


st.divider()


st.sidebar.title("Configurações da análise")

periodo_analisado = st.sidebar.text_input(
    "Período analisado",
    value="Últimos 7 dias"
)

st.sidebar.markdown("### Critérios de risco")
st.sidebar.info(
    "Até 3 dias sem acesso: sem risco\n\n"
    "4 a 7 dias: risco médio\n\n"
    "Mais de 7 dias: risco elevado"
)


st.header("1. Upload dos arquivos")

arquivo_pdf = st.file_uploader(
    "Envie o PDF com a lista de alunos",
    type=["pdf"]
)

arquivo_moodle = st.file_uploader(
    "Envie o relatório Moodle",
    type=["xlsx", "xls", "csv"]
)


df_final = None


if arquivo_pdf is not None:
    caminho_pdf = UPLOAD_DIR / arquivo_pdf.name

    with open(caminho_pdf, "wb") as f:
        f.write(arquivo_pdf.getbuffer())

    st.success(f"PDF salvo em: {caminho_pdf}")

    try:
        df_alunos = extrair_alunos_pdf(str(caminho_pdf))
        st.subheader("Alunos extraídos do PDF")
        st.dataframe(df_alunos, use_container_width=True)

    except Exception as e:
        st.error("Erro ao extrair alunos do PDF.")
        st.exception(e)
        df_alunos = None

else:
    df_alunos = None


if arquivo_moodle is not None:
    caminho_moodle = UPLOAD_DIR / arquivo_moodle.name

    with open(caminho_moodle, "wb") as f:
        f.write(arquivo_moodle.getbuffer())

    st.success(f"Relatório Moodle salvo em: {caminho_moodle}")

    try:
        df_moodle = processar_moodle(str(caminho_moodle))
        st.subheader("Dados processados do Moodle")
        st.dataframe(df_moodle, use_container_width=True)

    except Exception as e:
        st.error("Erro ao processar relatório Moodle.")
        st.exception(e)
        df_moodle = None

else:
    df_moodle = None


st.divider()


st.header("2. Ranking, risco e parecer pedagógico")


if df_alunos is not None and df_moodle is not None:

    if st.button("Gerar ranking de engajamento"):

        try:
            df_final = finalizar_ranking(df_alunos, df_moodle)

            st.session_state["df_final"] = df_final

            st.success("Ranking gerado com sucesso.")

        except Exception as e:
            st.error("Erro ao gerar ranking de engajamento.")
            st.exception(e)


if "df_final" in st.session_state:

    df_final = st.session_state["df_final"]

    st.subheader("Resultado final do monitoramento")
    st.dataframe(df_final, use_container_width=True)

    st.divider()

    st.header("3. Indicadores gerais")

    total_alunos = len(df_final)

    media_acessos = (
        df_final["acessos"].mean()
        if "acessos" in df_final.columns
        else 0
    )

    alunos_sem_acesso = (
        len(df_final[df_final["acessos"] == 0])
        if "acessos" in df_final.columns
        else 0
    )

    risco_medio = (
        len(
            df_final[
                df_final["risco"]
                .astype(str)
                .str.lower()
                .str.contains("médio|medio", na=False)
            ]
        )
        if "risco" in df_final.columns
        else 0
    )

    risco_elevado = (
        len(
            df_final[
                df_final["risco"]
                .astype(str)
                .str.lower()
                .str.contains("elevado", na=False)
            ]
        )
        if "risco" in df_final.columns
        else 0
    )

    col1, col2, col3, col4, col5 = st.columns(5)

    col1.metric("Total de alunos", total_alunos)
    col2.metric("Média de acessos", round(media_acessos, 2))
    col3.metric("Sem acesso", alunos_sem_acesso)
    col4.metric("Risco médio", risco_medio)
    col5.metric("Risco elevado", risco_elevado)

    st.divider()

    st.header("4. Exportações")

    caminho_excel = OUTPUT_DIR / "ranking_engajamento_ava.xlsx"

    df_final.to_excel(caminho_excel, index=False)

    with open(caminho_excel, "rb") as arquivo_excel:
        st.download_button(
            label="📥 Baixar ranking em Excel",
            data=arquivo_excel,
            file_name="ranking_engajamento_ava.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.subheader("Relatório técnico institucional em Word")

    if st.button("Gerar relatório técnico Word"):

        try:
            caminho_docx = exportar_relatorio_word(
                df_final,
                periodo_analisado=periodo_analisado
            )

            st.session_state["caminho_docx"] = caminho_docx

            st.success("Relatório técnico Word gerado com sucesso.")

        except Exception as e:
            st.error("Erro ao gerar relatório técnico Word.")
            st.exception(e)

    if "caminho_docx" in st.session_state:

        caminho_docx = st.session_state["caminho_docx"]

        with open(caminho_docx, "rb") as arquivo_docx:
            st.download_button(
                label="📄 Baixar relatório técnico Word",
                data=arquivo_docx,
                file_name="relatorio_tecnico_ava.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )


else:
    st.warning(
        "Envie o PDF dos alunos e o relatório Moodle para gerar o ranking, "
        "os indicadores e o relatório técnico."
    )


st.divider()

st.caption(
    "Athena Scientific — AVA Intelligence | Sistema de Learning Analytics para gestão acadêmica"
)
