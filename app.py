"""Consolidador de Notas Fiscais — interface Streamlit.

Execute com:  streamlit run app.py
"""

from __future__ import annotations

import hashlib
from datetime import datetime

import pandas as pd
import streamlit as st

from nf_extrator import (
    COLUNAS_PRINCIPAIS,
    agrupar_quantidades,
    consolidar,
    para_csv,
    para_excel,
    processar_arquivo,
    remover_duplicados,
)
from nf_extrator.modelos import ResultadoArquivo

st.set_page_config(page_title="Consolidador de NF-e", page_icon="🧾", layout="wide")


# --------------------------------------------------------------------------- #
# Processamento (com cache: reenviar o mesmo arquivo não reprocessa)
# --------------------------------------------------------------------------- #
@st.cache_data(show_spinner=False, max_entries=2000)
def _processar_em_cache(nome: str, digest: str, _conteudo: bytes) -> ResultadoArquivo:
    # O cache é indexado pelo nome + hash SHA-256; o conteúdo em si não é re-hasheado.
    return processar_arquivo(nome, _conteudo)


def processar_uploads(arquivos) -> list[ResultadoArquivo]:
    resultados = []
    barra = st.progress(0.0, text="Processando arquivos...")
    for i, arq in enumerate(arquivos, start=1):
        conteudo = arq.getvalue()
        digest = hashlib.sha256(conteudo).hexdigest()
        resultados.append(_processar_em_cache(arq.name, digest, conteudo))
        barra.progress(i / len(arquivos), text=f"Processando {i}/{len(arquivos)}: {arq.name}")
    barra.empty()
    return resultados


def tabela_resumo(resultados: list[ResultadoArquivo]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Arquivo": r.arquivo,
                "Tipo": r.tipo,
                "Status": r.status,
                "Notas": ", ".join(sorted(r.notas, key=lambda n: n.zfill(9))),
                "Itens": len(r.itens),
                "Mensagens": r.erro or " | ".join(r.avisos),
            }
            for r in resultados
        ]
    )


# --------------------------------------------------------------------------- #
# Interface
# --------------------------------------------------------------------------- #
st.title("🧾 Consolidador de Notas Fiscais")
st.caption(
    "Envie DANFEs (PDF) e/ou XMLs de NF-e. Os itens de todas as notas são reunidos "
    "em uma única tabela, pronta para download em Excel ou CSV."
)

with st.sidebar:
    st.header("Opções")
    mostrar_detalhes = st.toggle(
        "Mostrar colunas extras",
        value=False,
        help="Descrição, unidade, chave de acesso, arquivo de origem e tipo (PDF/XML).",
    )
    deduplicar = st.toggle(
        "Ignorar notas repetidas",
        value=True,
        help="Se a mesma nota for enviada mais de uma vez (ex.: PDF e XML), "
        "considera apenas um arquivo — o XML tem prioridade.",
    )
    agrupar = st.toggle(
        "Somar itens repetidos",
        value=False,
        help="Soma a quantidade de linhas com a mesma NF e o mesmo código de material.",
    )
    st.divider()
    st.markdown(
        "**Dica:** sempre que possível use o XML — ele é a fonte oficial da NF-e. "
        "PDFs escaneados (imagem) não são lidos."
    )

arquivos = st.file_uploader(
    "Selecione ou arraste os arquivos",
    type=["pdf", "xml"],
    accept_multiple_files=True,
    help="Você pode selecionar vários arquivos de uma vez.",
)

if not arquivos:
    st.info("Aguardando arquivos PDF ou XML.")
    st.stop()

resultados = processar_uploads(arquivos)
resumo = tabela_resumo(resultados)
df = consolidar(resultados)

removidos = 0
if deduplicar:
    df, removidos = remover_duplicados(df)
if agrupar:
    df = agrupar_quantidades(df)

# ---- Indicadores -----------------------------------------------------------
c1, c2, c3, c4 = st.columns(4)
c1.metric("Arquivos", len(resultados))
c2.metric("Notas fiscais", df["Nota Fiscal de Emissão"].nunique() if not df.empty else 0)
c3.metric("Itens", len(df))
c4.metric("Com erro/aviso", int((resumo["Status"] != "OK").sum()))

problemas = resumo[resumo["Status"] != "OK"]
if not problemas.empty:
    st.warning(f"{len(problemas)} arquivo(s) com erro ou aviso — veja os detalhes abaixo.")
if removidos:
    st.info(f"{removidos} linha(s) ignorada(s) por pertencerem a notas enviadas em duplicidade.")

with st.expander("Status por arquivo", expanded=not problemas.empty):
    st.dataframe(
        resumo,
        hide_index=True,
        width="stretch",
        column_config={"Mensagens": st.column_config.TextColumn(width="large")},
    )

# ---- Tabela consolidada ----------------------------------------------------
st.subheader("Itens consolidados")

if df.empty:
    st.error("Nenhum item foi extraído dos arquivos enviados.")
    st.stop()

busca = st.text_input("Filtrar", placeholder="Digite um número de NF, código de material...")
visao = df if mostrar_detalhes else df[COLUNAS_PRINCIPAIS]
if busca:
    termo = busca.strip().lower()
    mascara = visao.astype(str).apply(lambda col: col.str.lower().str.contains(termo, regex=False)).any(axis=1)
    visao = visao[mascara]

st.dataframe(
    visao,
    hide_index=True,
    width="stretch",
    column_config={
        "Quantidade": st.column_config.NumberColumn(format="%.4f"),
        "Chave de Acesso": st.column_config.TextColumn(width="medium"),
    },
)
st.caption(f"{len(visao)} de {len(df)} linha(s) exibida(s).")

# ---- Downloads -------------------------------------------------------------
carimbo = datetime.now().strftime("%Y%m%d_%H%M")
d1, d2, _ = st.columns([1, 1, 3])
d1.download_button(
    "⬇️ Baixar Excel (.xlsx)",
    data=para_excel(visao, resumo),
    file_name=f"notas_consolidadas_{carimbo}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    on_click="ignore",
    width="stretch",
)
d2.download_button(
    "⬇️ Baixar CSV",
    data=para_csv(visao),
    file_name=f"notas_consolidadas_{carimbo}.csv",
    mime="text/csv",
    on_click="ignore",
    width="stretch",
)
