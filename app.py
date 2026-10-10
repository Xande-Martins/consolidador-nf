"""Consolidador de Notas Fiscais (PDF/DANFE + XML de NF-e) — versão em arquivo único.

Execute com:  streamlit run app.py
"""

from __future__ import annotations

import hashlib
import io
import pandas as pd
import pdfplumber
import re
import streamlit as st
import unicodedata
import xml.etree.ElementTree as ET
from collections.abc import Iterable
from dataclasses import dataclass, field
from datetime import datetime
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ===========================================================================
# utils
# ===========================================================================
def remover_acentos(texto: str) -> str:
    """Remove acentos mantendo o comprimento do texto (útil para regex)."""
    nfkd = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def normalizar(texto: str) -> str:
    """Texto em maiúsculas, sem acentos e com 'º/°' padronizados para 'O'."""
    texto = texto.replace("º", "o").replace("°", "o").replace("ª", "a")
    return remover_acentos(texto).upper()


def numero_nf_limpo(valor: str) -> str:
    """'000.012.345' -> '12345'. Mantém '0' se o número for zero."""
    digitos = re.sub(r"\D", "", valor or "")
    return digitos.lstrip("0") or ("0" if digitos else "")


def numero_nf_da_chave(chave: str) -> str | None:
    """Extrai o nNF (posições 26-34) de uma chave de acesso de 44 dígitos.

    Layout: cUF(2) AAMM(4) CNPJ(14) mod(2) série(3) nNF(9) tpEmis(1) cNF(8) cDV(1)
    """
    digitos = re.sub(r"\D", "", chave or "")
    if len(digitos) != 44:
        return None
    return numero_nf_limpo(digitos[25:34])


def chave_valida(chave: str) -> bool:
    """Valida o dígito verificador (módulo 11) de uma chave de acesso."""
    digitos = re.sub(r"\D", "", chave or "")
    if len(digitos) != 44:
        return False
    corpo, dv = digitos[:43], int(digitos[43])
    pesos = [2, 3, 4, 5, 6, 7, 8, 9]
    soma = sum(int(d) * pesos[i % 8] for i, d in enumerate(reversed(corpo)))
    resto = soma % 11
    esperado = 0 if resto < 2 else 11 - resto
    return dv == esperado


def converter_decimal(valor: str | None) -> float | None:
    """Converte números no formato brasileiro ou americano para float.

    '1.234,5600' -> 1234.56 | '10,5' -> 10.5 | '10.5000' -> 10.5 | '1.000' -> 1000
    """
    if valor is None:
        return None
    s = str(valor).strip().replace(" ", "")
    if not s:
        return None
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif s.count(".") > 1 or re.fullmatch(r"\d{1,3}\.\d{3}", s):
        # '1.234.567' ou '1.000' (separador de milhar sem decimais)
        s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return None


# ===========================================================================
# modelos
# ===========================================================================
SEM_REFERENCIA = "N/D"


@dataclass
class ItemNF:
    """Um item (linha de produto) de uma Nota Fiscal."""

    nf_referencia: str
    nf_emissao: str
    codigo_material: str
    quantidade: float | None
    descricao: str = ""
    unidade: str = ""
    chave_acesso: str = ""
    arquivo: str = ""
    origem: str = ""  # "XML" ou "PDF"


@dataclass
class ResultadoArquivo:
    """Resultado do processamento de um arquivo enviado."""

    arquivo: str
    tipo: str
    itens: list[ItemNF] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)
    erro: str | None = None

    @property
    def notas(self) -> set[str]:
        return {i.nf_emissao for i in self.itens}

    @property
    def status(self) -> str:
        if self.erro:
            return "Erro"
        if self.avisos:
            return "Com avisos"
        return "OK"


# ===========================================================================
# referencias
# ===========================================================================
# Palavra-chave que indica Nota Fiscal
_KW_NF = r"(?:NOTAS?\s+FISCA(?:L|IS)(?:\s+ELETRONICAS?)?|DANFE|\bN\.?\s?F\.?(?:\s?-?\s?E)?S?\b)"

# Qualificadores que podem aparecer entre a palavra-chave e o número
_QUALIF = (
    r"(?:\s*(?:DE|DA|DO|DAS|DOS|REF\.?|REFERENCIA(?:DA|DAS)?|REFERENTE|ORIGEM|ORIGINAL|"
    r"VENDA|REMESSA|RETORNO|DEVOLUCAO|ENTRADA|SAIDA|NUMEROS?|NUM\.?|NO\.?|NROS?\.?|N\.|N(?=\s*\d)|"
    r"SERIE\s*\d{1,3}\s*[,-]?|[:#\-.])\s*)*"
)

# Um número de NF (aceita pontos de milhar: 000.012.345) ...
_NUM = r"\d{1,3}(?:\.\d{3}){1,2}|\d{1,9}"
# ... ou uma lista: "10, 11 e 12" / "10/11" não (barra costuma ser série)
_LISTA = rf"(?:{_NUM})(?:\s*(?:,|;|\bE\b)\s*(?:{_NUM}))*"

_RE_NF = re.compile(rf"{_KW_NF}{_QUALIF}\s*({_LISTA})(?!\d)")
_RE_CHAVE = re.compile(r"(?<!\d)(?:\d[\s.]?){43}\d(?!\d)")
_RE_PEDIDO = re.compile(
    r"(?:\bPEDIDO(?:\s+DE\s+COMPRAS?)?|\bPED\.|ORDEM\s+DE\s+COMPRA|\bXPED)"
    r"\s*(?:N\.|NO\.?|NUM\.?|NUMERO|[:#\-])*\s*([A-Z0-9][A-Z0-9./-]{2,})"
)


def _sem_duplicatas(valores: Iterable[str]) -> list[str]:
    vistos: list[str] = []
    for v in valores:
        if v and v not in vistos:
            vistos.append(v)
    return vistos


def numeros_nf_no_texto(texto: str, ignorar: str | None = None) -> list[str]:
    """Retorna os números de NF citados no texto (sem zeros à esquerda)."""
    if not texto:
        return []
    t = normalizar(texto)
    encontrados: list[str] = []

    # 1) Chaves de acesso de 44 dígitos (mais confiáveis)
    for m in _RE_CHAVE.finditer(t):
        chave = re.sub(r"\D", "", m.group(0))
        if chave_valida(chave):
            encontrados.append(numero_nf_da_chave(chave))
    t_sem_chaves = _RE_CHAVE.sub(" ", t)

    # 2) Menções textuais: "NF 123", "Nota Fiscal nº 456"...
    for m in _RE_NF.finditer(t_sem_chaves):
        for parte in re.split(r"\s*(?:,|;|\bE\b)\s*", m.group(1)):
            numero = numero_nf_limpo(parte)
            if numero and numero != "0":
                encontrados.append(numero)

    ignorar_limpo = numero_nf_limpo(ignorar) if ignorar else None
    return [n for n in _sem_duplicatas(encontrados) if n != ignorar_limpo]


def pedido_no_texto(texto: str) -> str | None:
    if not texto:
        return None
    m = _RE_PEDIDO.search(normalizar(texto))
    return m.group(1).rstrip(".-/") if m else None


def extrair_referencia(
    texto_complementar: str,
    nf_emissao: str | None = None,
    refs_estruturadas: Iterable[str] = (),
    usar_pedido: bool = True,
) -> str:
    """Monta o valor da coluna "Nota Fiscal de Referência".

    Prioridade: referências estruturadas (tags NFref do XML) + NFs citadas no texto;
    se nenhuma for encontrada, usa o número do pedido (quando houver); senão "N/D".
    """
    proprio = numero_nf_limpo(nf_emissao) if nf_emissao else None
    numeros = [n for n in refs_estruturadas if n and n != proprio]
    numeros += numeros_nf_no_texto(texto_complementar, ignorar=nf_emissao)
    numeros = _sem_duplicatas(numeros)
    if numeros:
        return ", ".join(numeros)
    if usar_pedido:
        pedido = pedido_no_texto(texto_complementar)
        if pedido:
            return f"Pedido {pedido}"
    return SEM_REFERENCIA


# ===========================================================================
# xml_parser
# ===========================================================================
def _local(tag: str) -> str:
    """'{http://www.portalfiscal.inf.br/nfe}det' -> 'det'."""
    return tag.rsplit("}", 1)[-1]


def _filho(elem: ET.Element | None, *caminho: str) -> ET.Element | None:
    """Navega pelos filhos diretos ignorando namespaces."""
    atual = elem
    for nome in caminho:
        if atual is None:
            return None
        atual = next((c for c in atual if _local(c.tag) == nome), None)
    return atual


def _texto(elem: ET.Element | None, *caminho: str) -> str:
    alvo = _filho(elem, *caminho)
    return (alvo.text or "").strip() if alvo is not None else ""


def _referencias_estruturadas(ide: ET.Element | None) -> list[str]:
    """Números das NFs informadas nos grupos NFref (refNFe, refNF, refNFP...)."""
    numeros: list[str] = []
    if ide is None:
        return numeros
    for nfref in (c for c in ide if _local(c.tag) == "NFref"):
        for ref in nfref:
            nome = _local(ref.tag)
            if nome in ("refNFe", "refNFeSig", "refCTe"):
                n = numero_nf_da_chave(ref.text or "")
            elif nome in ("refNF", "refNFP"):
                n = numero_nf_limpo(_texto(ref, "nNF"))
            elif nome == "refECF":
                n = numero_nf_limpo(_texto(ref, "nCOO"))
            else:
                n = None
            if n:
                numeros.append(n)
    return numeros


def _quantidade(valor: str) -> float | None:
    # No XML o separador decimal é sempre ponto ("10.0000")
    try:
        return float(valor) if valor else None
    except ValueError:
        return None


def processar_xml(conteudo: bytes, nome_arquivo: str) -> ResultadoArquivo:
    resultado = ResultadoArquivo(arquivo=nome_arquivo, tipo="XML")
    try:
        raiz = ET.fromstring(conteudo)
    except ET.ParseError as exc:
        resultado.erro = f"XML inválido: {exc}"
        return resultado

    blocos = [e for e in raiz.iter() if _local(e.tag) == "infNFe"]
    if not blocos:
        nome_raiz = _local(raiz.tag)
        resultado.erro = (
            f"O arquivo não contém uma NF-e (elemento raiz: <{nome_raiz}>). "
            "XMLs de evento, cancelamento ou CT-e não são suportados."
        )
        return resultado

    for inf in blocos:
        ide = _filho(inf, "ide")
        nf_emissao = numero_nf_limpo(_texto(ide, "nNF"))
        if not nf_emissao:
            # fallback: atributo Id="NFe<chave>"
            nf_emissao = numero_nf_da_chave(inf.get("Id", "")) or ""
        if not nf_emissao:
            resultado.avisos.append("Número da NF (nNF) não encontrado.")

        chave = re.sub(r"\D", "", inf.get("Id", ""))
        inf_cpl = _texto(inf, "infAdic", "infCpl")
        referencia = extrair_referencia(
            inf_cpl,
            nf_emissao=nf_emissao,
            refs_estruturadas=_referencias_estruturadas(ide),
        )

        dets = [c for c in inf if _local(c.tag) == "det"]
        if not dets:
            resultado.avisos.append(f"NF {nf_emissao or '?'}: nenhum item <det> encontrado.")
        for det in dets:
            prod = _filho(det, "prod")
            qtd_txt = _texto(prod, "qCom")
            quantidade = _quantidade(qtd_txt)
            if quantidade is None:
                resultado.avisos.append(
                    f"NF {nf_emissao}, item {det.get('nItem', '?')}: quantidade inválida ({qtd_txt!r})."
                )
            resultado.itens.append(
                ItemNF(
                    nf_referencia=referencia,
                    nf_emissao=nf_emissao,
                    codigo_material=_texto(prod, "cProd"),
                    quantidade=quantidade,
                    descricao=_texto(prod, "xProd"),
                    unidade=_texto(prod, "uCom"),
                    chave_acesso=chave if len(chave) == 44 else "",
                    arquivo=nome_arquivo,
                    origem="XML",
                )
            )
    return resultado


# ===========================================================================
# pdf_parser
# ===========================================================================
_RE_CHAVE = re.compile(r"(?<!\d)(?:\d[ .]?){43}\d(?!\d)")

_RE_NUMERO_DANFE = [
    # "Nº 000.012.345" / "N. 000.012.345" (formato padrão do DANFE)
    re.compile(r"\bN\s*O?\s*[.:]?\s*(\d{3}\.\d{3}\.\d{3})\b"),
    re.compile(r"(?:NF-?E|NOTA\s+FISCAL(?:\s+ELETRONICA)?)\s*N\s*O?\s*[.:]?\s*(\d{1,9})\b"),
]

# Linha de item no layout padrão:
# CÓDIGO  DESCRIÇÃO  NCM/SH  O/CST  CFOP  UN  QUANT  V.UNIT  V.TOTAL ...
_RE_LINHA_ITEM = re.compile(
    r"^(?P<cod>\S+)\s+(?P<desc>.+?)\s+"
    r"(?P<ncm>\d{8}|\d{4}\.\d{2}\.\d{2})\s+"
    r"(?P<cst>\d{2,4})\s+"
    r"(?P<cfop>\d{4}|\d\.\d{3})\s+"
    r"(?P<un>[A-Za-z][\w/.-]{0,7})\s+"
    r"(?P<qtd>\d{1,3}(?:\.\d{3})*,\d+|\d+(?:,\d+)?)\s+"
    r"(?P<vunit>\d{1,3}(?:\.\d{3})*,\d+|\d+,\d+)"
)

_RE_INFO_COMPL = re.compile(r"INFORMA[ÇC][ÕO]ES\s+COMPLEMENTARES|DADOS\s+ADICIONAIS", re.I)
_RE_RESERVADO = re.compile(r"RESERVADO\s+AO\s+FISCO", re.I)


@dataclass
class _Pagina:
    texto: str
    tabelas: list[list[list[str | None]]]
    info_compl: str
    chave: str | None
    numero: str  # número da NF lido do campo "Nº" (usado se não houver chave)

    @property
    def ident(self) -> str:
        return self.chave or (f"NF:{self.numero}" if self.numero else "")


@dataclass
class _Nota:
    paginas: list[_Pagina] = field(default_factory=list)

    @property
    def chave(self) -> str | None:
        return next((p.chave for p in self.paginas if p.chave), None)

    @property
    def ident(self) -> str:
        return next((p.ident for p in self.paginas if p.ident), "")


# --------------------------------------------------------------------------- #
# Leitura das páginas
# --------------------------------------------------------------------------- #
def _chave_da_pagina(texto: str) -> str | None:
    for m in _RE_CHAVE.finditer(texto):
        chave = re.sub(r"\D", "", m.group(0))
        if chave[20:22] in ("55", "65") and chave_valida(chave):
            return chave
    return None


def _texto_info_complementar(page: pdfplumber.page.Page, texto_pagina: str) -> str:
    """Recorta o quadro "Informações Complementares" pela posição na página."""
    try:
        achados = page.search(_RE_INFO_COMPL.pattern, regex=True, case=False)
    except Exception:  # noqa: BLE001 - versões antigas do pdfplumber
        achados = []

    if achados:
        topo = min(a["top"] for a in achados)
        x_dir = page.width
        for r in page.search(_RE_RESERVADO.pattern, regex=True, case=False) or []:
            if r["top"] >= topo - 5 and r["x0"] > page.width * 0.3:
                x_dir = min(x_dir, r["x0"] - 2)
        recorte = page.crop((0, topo, x_dir, page.height), strict=False)
        texto = recorte.extract_text() or ""
        return _RE_INFO_COMPL.sub(" ", texto).strip()

    # Fallback textual: tudo após o título do quadro
    m = _RE_INFO_COMPL.search(texto_pagina)
    if not m:
        return ""
    trecho = texto_pagina[m.end():]
    return _RE_RESERVADO.sub(" ", trecho).strip()


def _ler_paginas(pdf: pdfplumber.PDF) -> list[_Pagina]:
    paginas = []
    for page in pdf.pages:
        texto = page.extract_text() or ""
        try:
            tabelas = page.extract_tables() or []
        except Exception:  # noqa: BLE001
            tabelas = []
        paginas.append(
            _Pagina(
                texto=texto,
                tabelas=tabelas,
                info_compl=_texto_info_complementar(page, texto) if texto else "",
                chave=_chave_da_pagina(texto),
                numero=_numero_no_texto(texto),
            )
        )
    return paginas


def _agrupar_por_nota(paginas: list[_Pagina]) -> list[_Nota]:
    """Páginas consecutivas da mesma nota (mesma chave ou, sem chave, mesmo Nº)
    são agrupadas; páginas sem identificação seguem a nota anterior."""
    notas: list[_Nota] = []
    for pag in paginas:
        atual = notas[-1].ident if notas else ""
        if not notas or (pag.ident and atual and pag.ident != atual and not _mesma_nota(pag, notas[-1])):
            notas.append(_Nota())
        notas[-1].paginas.append(pag)
    return notas


def _mesma_nota(pag: _Pagina, nota: _Nota) -> bool:
    """Compara pelo número quando só uma das partes tem a chave legível."""
    numero_nota = numero_nf_da_chave(nota.chave) if nota.chave else nota.ident.removeprefix("NF:")
    numero_pag = numero_nf_da_chave(pag.chave) if pag.chave else pag.numero
    if pag.chave and nota.chave:
        return pag.chave == nota.chave
    return bool(numero_pag) and numero_pag == numero_nota


# --------------------------------------------------------------------------- #
# Extração dos campos
# --------------------------------------------------------------------------- #
def _numero_no_texto(texto_pagina: str) -> str:
    texto = normalizar(texto_pagina)
    for padrao in _RE_NUMERO_DANFE:
        m = padrao.search(texto)
        if m:
            return numero_nf_limpo(m.group(1))
    return ""


def _numero_nf(nota: _Nota) -> str:
    if nota.chave:
        return numero_nf_da_chave(nota.chave) or ""
    return next((p.numero for p in nota.paginas if p.numero), "")


def _itens_por_texto(nota: _Nota) -> list[dict]:
    itens = []
    for pag in nota.paginas:
        for linha in pag.texto.splitlines():
            m = _RE_LINHA_ITEM.match(linha.strip())
            if m:
                itens.append(
                    {
                        "codigo": m["cod"],
                        "descricao": m["desc"].strip(),
                        "unidade": m["un"],
                        "quantidade": converter_decimal(m["qtd"]),
                    }
                )
    return itens


def _indice_coluna(cabecalho: list[str], *termos: str, excluir: tuple[str, ...] = ()) -> int | None:
    for i, cel in enumerate(cabecalho):
        if any(t in cel for t in termos) and not any(e in cel for e in excluir):
            return i
    return None


def _itens_por_tabela(nota: _Nota) -> list[dict]:
    itens = []
    for pag in nota.paginas:
        for tabela in pag.tabelas:
            col_cod = col_qtd = col_desc = col_un = None
            for linha in tabela:
                celulas = [(c or "").strip() for c in linha]
                normal = [normalizar(c).replace("\n", " ") for c in celulas]

                # Procura (ou reencontra) o cabeçalho do quadro de produtos
                cod = _indice_coluna(normal, "COD", excluir=("CODIGO NCM",))
                qtd = _indice_coluna(normal, "QUANT", "QTD", "QTDE")
                if cod is not None and qtd is not None and cod != qtd:
                    col_cod, col_qtd = cod, qtd
                    col_desc = _indice_coluna(normal, "DESCRI")
                    col_un = next((i for i, c in enumerate(normal) if c in ("UN", "UNID", "UND", "UNID.")), None)
                    continue
                if col_cod is None or max(col_cod, col_qtd) >= len(celulas):
                    continue

                codigos = [c for c in celulas[col_cod].split("\n") if c.strip()]
                qtds = [q for q in celulas[col_qtd].split("\n") if q.strip()]
                if not codigos or len(codigos) != len(qtds):
                    continue
                descs = celulas[col_desc].split("\n") if col_desc is not None else []
                uns = celulas[col_un].split("\n") if col_un is not None and col_un < len(celulas) else []
                for i, (c, q) in enumerate(zip(codigos, qtds)):
                    quantidade = converter_decimal(q)
                    if quantidade is None:
                        continue
                    itens.append(
                        {
                            "codigo": c.strip(),
                            # descrições podem quebrar linha; só usamos se o número bater
                            "descricao": descs[i].strip() if len(descs) == len(codigos) else "",
                            "unidade": uns[i].strip() if len(uns) == len(codigos) else "",
                            "quantidade": quantidade,
                        }
                    )
    return itens


# --------------------------------------------------------------------------- #
# Ponto de entrada
# --------------------------------------------------------------------------- #
def processar_pdf(conteudo: bytes, nome_arquivo: str) -> ResultadoArquivo:
    resultado = ResultadoArquivo(arquivo=nome_arquivo, tipo="PDF")
    try:
        with pdfplumber.open(io.BytesIO(conteudo)) as pdf:
            paginas = _ler_paginas(pdf)
    except Exception as exc:  # noqa: BLE001
        resultado.erro = f"Não foi possível abrir o PDF: {exc}"
        return resultado

    if not any(p.texto.strip() for p in paginas):
        resultado.erro = (
            "O PDF não possui texto selecionável (provavelmente é uma imagem escaneada). "
            "Envie o XML da nota ou um PDF gerado digitalmente."
        )
        return resultado

    for nota in _agrupar_por_nota(paginas):
        nf_emissao = _numero_nf(nota)
        if not nf_emissao:
            resultado.avisos.append("Número da NF não identificado em uma das notas do PDF.")

        info_compl = "\n".join(p.info_compl for p in nota.paginas if p.info_compl)
        referencia = extrair_referencia(info_compl, nf_emissao=nf_emissao)

        por_texto = _itens_por_texto(nota)
        por_tabela = _itens_por_tabela(nota)
        itens = por_texto if len(por_texto) >= len(por_tabela) else por_tabela

        if not itens:
            resultado.avisos.append(
                f"NF {nf_emissao or '?'}: nenhum item reconhecido no PDF. "
                "O layout pode ser diferente do padrão — prefira o XML."
            )
        for it in itens:
            if it["quantidade"] is None:
                resultado.avisos.append(f"NF {nf_emissao}: quantidade ilegível no item {it['codigo']}.")
            resultado.itens.append(
                ItemNF(
                    nf_referencia=referencia,
                    nf_emissao=nf_emissao,
                    codigo_material=it["codigo"],
                    quantidade=it["quantidade"],
                    descricao=it["descricao"],
                    unidade=it["unidade"],
                    chave_acesso=nota.chave or "",
                    arquivo=nome_arquivo,
                    origem="PDF",
                )
            )
    return resultado


# ===========================================================================
# processador
# ===========================================================================
COLUNAS = {
    "nf_referencia": "Nota Fiscal de Referência",
    "nf_emissao": "Nota Fiscal de Emissão",
    "codigo_material": "Código do Material",
    "quantidade": "Quantidade",
    "descricao": "Descrição",
    "unidade": "Unidade",
    "chave_acesso": "Chave de Acesso",
    "arquivo": "Arquivo",
    "origem": "Origem",
}
COLUNAS_PRINCIPAIS = list(COLUNAS.values())[:4]


def detectar_tipo(nome: str, conteudo: bytes) -> str | None:
    """Identifica o tipo pelo conteúdo (assinatura) e, em último caso, pela extensão."""
    inicio = conteudo[:1024].lstrip(b"\xef\xbb\xbf \t\r\n")
    if inicio.startswith(b"%PDF"):
        return "pdf"
    if inicio.startswith(b"<"):
        return "xml"
    ext = nome.lower().rsplit(".", 1)[-1]
    return ext if ext in ("pdf", "xml") else None


def processar_arquivo(nome: str, conteudo: bytes) -> ResultadoArquivo:
    tipo = detectar_tipo(nome, conteudo)
    try:
        if tipo == "xml":
            return processar_xml(conteudo, nome)
        if tipo == "pdf":
            return processar_pdf(conteudo, nome)
    except Exception as exc:  # noqa: BLE001 - nunca derrubar o lote inteiro
        return ResultadoArquivo(arquivo=nome, tipo=(tipo or "?").upper(), erro=f"Erro inesperado: {exc}")
    return ResultadoArquivo(arquivo=nome, tipo="?", erro="Formato não suportado (envie PDF ou XML).")


def processar_arquivos(arquivos: Iterable[tuple[str, bytes]]) -> list[ResultadoArquivo]:
    return [processar_arquivo(nome, conteudo) for nome, conteudo in arquivos]


def consolidar(resultados: Iterable[ResultadoArquivo]) -> pd.DataFrame:
    """Une os itens de todos os arquivos em um único DataFrame."""
    linhas = [vars(item) for r in resultados for item in r.itens]
    df = pd.DataFrame(linhas, columns=list(COLUNAS.keys())).rename(columns=COLUNAS)
    df["Quantidade"] = pd.to_numeric(df["Quantidade"], errors="coerce")
    return df


def remover_duplicados(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Remove notas enviadas em duplicidade (ex.: PDF e XML da mesma NF).

    Quando a mesma NF aparece em mais de um arquivo, mantém apenas o primeiro
    arquivo (dando preferência ao XML, que é a fonte oficial).
    """
    if df.empty:
        return df, 0
    # Identificador da nota: chave de acesso (única no país); sem chave, o número da NF
    ident = df["Chave de Acesso"].where(df["Chave de Acesso"] != "", "NF:" + df["Nota Fiscal de Emissão"])
    ordenado = df.assign(_id=ident, _prio=(df["Origem"] != "XML").astype(int)).sort_values("_prio", kind="stable")
    arquivo_escolhido = ordenado.groupby("_id", sort=False)["Arquivo"].first()
    manter = df["Arquivo"] == ident.map(arquivo_escolhido)
    manter |= df["Nota Fiscal de Emissão"].eq("") & df["Chave de Acesso"].eq("")
    return df[manter].reset_index(drop=True), int((~manter).sum())


def agrupar_quantidades(df: pd.DataFrame) -> pd.DataFrame:
    """Soma as quantidades de itens repetidos (mesma NF + mesmo código)."""
    if df.empty:
        return df
    chaves = ["Nota Fiscal de Referência", "Nota Fiscal de Emissão", "Código do Material"]
    return (
        df.groupby(chaves, as_index=False, sort=False, dropna=False)
        .agg(
            {
                "Quantidade": "sum",
                "Descrição": "first",
                "Unidade": "first",
                "Chave de Acesso": "first",
                "Arquivo": lambda s: ", ".join(dict.fromkeys(s)),
                "Origem": lambda s: ", ".join(dict.fromkeys(s)),
            }
        )[df.columns]
    )


# ===========================================================================
# exportacao
# ===========================================================================
def para_excel(df: pd.DataFrame, resumo: pd.DataFrame | None = None) -> bytes:
    """Gera um .xlsx com a aba "Itens" (e "Arquivos", se o resumo for informado)."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Itens", index=False)
        if resumo is not None and not resumo.empty:
            resumo.to_excel(writer, sheet_name="Arquivos", index=False)

        for ws in writer.book.worksheets:
            for cel in ws[1]:
                cel.font = Font(bold=True, color="FFFFFF")
                cel.fill = PatternFill("solid", fgColor="1F4E78")
                cel.alignment = Alignment(vertical="center", wrap_text=True)
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for i, coluna in enumerate(ws.iter_cols(values_only=True), start=1):
                largura = max(len(str(v)) if v is not None else 0 for v in coluna)
                ws.column_dimensions[get_column_letter(i)].width = min(max(largura + 2, 10), 60)

        ws_itens = writer.book["Itens"]
        if "Quantidade" in df.columns:
            col = df.columns.get_loc("Quantidade") + 1
            for (cel,) in ws_itens.iter_rows(min_row=2, min_col=col, max_col=col):
                cel.number_format = "#,##0.####"
    return buffer.getvalue()


def para_csv(df: pd.DataFrame) -> bytes:
    """CSV no padrão do Excel em português: separador ';', decimal ',' e BOM UTF-8."""
    return df.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")


# ===========================================================================
# Interface Streamlit
# ===========================================================================
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
