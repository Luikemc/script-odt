import re
import sys
import copy
import shutil
import zipfile
import tempfile
import platform
import subprocess
from pathlib import Path
from xml.etree import ElementTree as ET

# ============================================================
# SUBLINHA CABECALHOS - V6
# ============================================================
#
# Objetivo:
#   Encontrar cabeçalhos de conversas no formato:
#
#   [10/08/2000 11:17:20] Lucas Ribeiro:
#   [10/08/2000, 11:17:20] Lucas Ribeiro:
#   [ 10/08/2000 , 11:17:20 ] Lucas Ribeiro:
#   10/08/2000 11:17:20 Lucas Ribeiro:          <- V6: sem colchetes
#   10/08/2000, 11:17:20 Lucas Ribeiro:         <- V6: sem colchetes
#
#   e sublinhar SOMENTE até o ":" do nome.
#
# O que mudou da V5 para a V6:
#
#   1) FORMATOS ACEITOS: agora o programa aceita .odt, .docx E .doc
#      (antes só .odt).
#
#        .odt  -> processado diretamente (é um zip com XML).
#        .docx -> processado diretamente (também é um zip com XML,
#                 mas com um esquema (schema) totalmente diferente
#                 do ODF: word/document.xml, elementos <w:p>, <w:r>,
#                 <w:t> etc.). Por isso foi escrito um processador
#                 dedicado (bloco "DOCX" abaixo). Como o DOCX permite
#                 formatação DIRETA no "run" (<w:rPr><w:u .../></w:rPr>),
#                 não precisamos criar um estilo nomeado como no ODT -
#                 isso simplifica bastante o processo.
#
#        .doc  -> este é o Word 97-2003, um formato BINÁRIO (OLE2),
#                 não é XML e não dá para editar diretamente com
#                 Python puro preservando a formatação. A única forma
#                 confiável é usar o próprio LibreOffice, em modo
#                 headless (sem abrir interface gráfica), para:
#
#                     .doc  --(LibreOffice)--> .docx
#                     [processa o .docx com o mesmo código acima]
#                     .docx --(LibreOffice)--> .doc
#
#                 Por isso, para processar arquivos .doc, é
#                 necessário ter o LibreOffice instalado na máquina
#                 (gratuito: https://www.libreoffice.org/download/).
#                 Se ele não for encontrado, o programa avisa
#                 claramente e não trava sem explicação.
#
#   2) CABEÇALHO SEM COLCHETES: o padrão de busca (PADRAO) agora
#      aceita tanto "[data hora] Nome:" quanto "data hora Nome:"
#      (sem os colchetes). As duas formas são tratadas como
#      alternativas completas (ou casa o bloco inteiro com colchete
#      de abertura E de fechamento, ou casa sem colchete nenhum) -
#      isso evita casar um colchete "solto" por engano.
#
# Características mantidas (das versões anteriores, agora válidas
# tanto para ODT quanto para DOCX):
#   - preserva o arquivo original;
#   - mantém sublinhados existentes;
#   - completa sublinhados parciais;
#   - entende texto dividido em vários trechos (spans/runs);
#   - pode ser executado várias vezes;
#   - gera relatório detalhado;
#   - aceita arquivo arrastado para o .exe;
#   - cria automaticamente *_formatado.<extensao original>.
#
# Limitação conhecida (cosmética, não gera perda de texto):
#   se o cabeçalho contiver uma tabulação, quebra de linha, ou (só
#   no DOCX) um trecho com formatação "mista" dentro do mesmo
#   pedaço de texto NO MEIO do trecho a sublinhar, esse pedaço
#   específico pode não receber o traço de sublinhado (o texto ao
#   redor é sublinhado normalmente). É raro acontecer dentro de
#   "[data hora] Nome:", mas fique ciente.
#
#   Para .doc: como o arquivo passa por duas conversões via
#   LibreOffice (doc -> docx -> doc), pequenos detalhes de
#   formatação do arquivo original (fontes muito específicas,
#   layout de página etc.) podem sofrer pequenos ajustes causados
#   pelo próprio LibreOffice durante a conversão. O texto e o
#   sublinhado do cabeçalho, porém, são preservados normalmente.
#
# ============================================================

# ------------------------------------------------------------
# PADRÃO DO CABEÇALHO (comum a ODT e DOCX)
# ------------------------------------------------------------
#
# Ou casa o bloco "[data hora]" completo (com os dois colchetes),
# ou casa "data hora" sem colchete nenhum. Isso evita reconhecer
# um colchete que abre sem fechar (ou vice-versa) como se fizesse
# parte do cabeçalho.

PADRAO = re.compile(
    r'(?:'
    r'\[\s*\d{2}/\d{2}/\d{4}\s*(?:,\s*)?\d{2}:\d{2}:\d{2}\s*\]'
    r'|'
    r'\d{2}/\d{2}/\d{4}\s*(?:,\s*)?\d{2}:\d{2}:\d{2}'
    r')'
    r'\s*[^:\n\r]+?:'
)


# ============================================================
# ============================================================
#   BLOCO ODT
# ============================================================
# ============================================================

NS_ODT = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
}

for prefixo, uri in NS_ODT.items():
    ET.register_namespace(prefixo, uri)


def odt_texto_especial(node):
    if node.tag == f"{{{NS_ODT['text']}}}s":
        quantidade = int(node.attrib.get(f"{{{NS_ODT['text']}}}c", "1"))
        return " " * quantidade
    if node.tag == f"{{{NS_ODT['text']}}}tab":
        return "\t"
    if node.tag == f"{{{NS_ODT['text']}}}line-break":
        return "\n"
    return ""


def odt_construir_mapa(paragrafo):
    partes = []
    posicao = 0

    def adicionar(node, texto, tipo, pai):
        nonlocal posicao
        if not texto:
            return
        partes.append({
            "node": node, "texto": texto,
            "inicio": posicao, "fim": posicao + len(texto),
            "tipo": tipo, "pai": pai,
        })
        posicao += len(texto)

    def visitar(node):
        if node.text:
            adicionar(node, node.text, "text", None)
        for filho in list(node):
            especial = odt_texto_especial(filho)
            if especial:
                adicionar(filho, especial, "special", node)
            else:
                visitar(filho)
            if filho.tail:
                adicionar(filho, filho.tail, "tail", node)

    visitar(paragrafo)
    return partes


def odt_criar_estilo_sublinhado(styles_root):
    """
    Cria (ou, se já existir - por exemplo, gerado por uma versão
    mais antiga deste script -, ATUALIZA) o estilo de sublinhado.

    Importante: mesmo quando o estilo já existe, sempre garantimos
    que ele tenha tanto o atributo legado (fo:text-decoration)
    quanto os atributos modernos (style:text-underline-*). Sem
    isso, um arquivo gerado por uma versão antiga do script (que só
    gravava o atributo legado) continuaria sem sublinhado de
    verdade no Word e ao converter para .doc, mesmo já estando
    marcado como "sublinhado" pelo estilo antigo.
    """

    nome = "SublinhadoAutomatico"

    estilo = None
    for candidato in styles_root.iter(f"{{{NS_ODT['style']}}}style"):
        if candidato.attrib.get(f"{{{NS_ODT['style']}}}name") == nome:
            estilo = candidato
            break

    if estilo is None:
        estilo = ET.SubElement(
            styles_root, f"{{{NS_ODT['style']}}}style",
            {f"{{{NS_ODT['style']}}}name": nome, f"{{{NS_ODT['style']}}}family": "text"}
        )

    propriedades = estilo.find(f"{{{NS_ODT['style']}}}text-properties")
    if propriedades is None:
        propriedades = ET.SubElement(estilo, f"{{{NS_ODT['style']}}}text-properties")

    # Atributo legado (LibreOffice / Google Docs).
    propriedades.set(f"{{{NS_ODT['fo']}}}text-decoration", "underline")
    # Atributos modernos (exigidos para o sublinhado aparecer no Word
    # e para sobreviver a uma conversão .odt -> .doc pelo LibreOffice).
    propriedades.set(f"{{{NS_ODT['style']}}}text-underline-style", "solid")
    propriedades.set(f"{{{NS_ODT['style']}}}text-underline-width", "auto")
    propriedades.set(f"{{{NS_ODT['style']}}}text-underline-color", "font-color")

    return nome


def odt_criar_mapa_estilos(root_styles):
    mapa = {}
    for style in root_styles.iter(f"{{{NS_ODT['style']}}}style"):
        nome = style.attrib.get(f"{{{NS_ODT['style']}}}name")
        if nome:
            mapa[nome] = style
    return mapa


def odt_estilo_tem_sublinhado(node, mapa_estilos):
    if node is None:
        return False
    nome = node.attrib.get(f"{{{NS_ODT['text']}}}style-name")
    visitados = set()
    while nome and nome not in visitados:
        visitados.add(nome)
        estilo = mapa_estilos.get(nome)
        if estilo is None:
            break
        propriedades = estilo.find(f"{{{NS_ODT['style']}}}text-properties")
        if propriedades is not None:
            decoracao = propriedades.attrib.get(
                f"{{{NS_ODT['fo']}}}text-decoration", "").lower()
            if "underline" in decoracao:
                return True
            moderno = propriedades.attrib.get(
                f"{{{NS_ODT['style']}}}text-underline-style", "").lower()
            if moderno and moderno != "none":
                return True
        nome = estilo.attrib.get(f"{{{NS_ODT['style']}}}parent-style-name")
    return False


def odt_criar_span_sublinhado(texto, estilo_sublinhado):
    span = ET.Element(
        f"{{{NS_ODT['text']}}}span",
        {f"{{{NS_ODT['text']}}}style-name": estilo_sublinhado}
    )
    span.text = texto
    return span


def odt_sublinhar_texto_do_node(item, inicio_local, fim_local, estilo_sublinhado):
    node = item["node"]
    texto = node.text or ""
    antes = texto[:inicio_local]
    selecionado = texto[inicio_local:fim_local]
    depois = texto[fim_local:]
    if not selecionado:
        return False
    node.text = antes
    span = odt_criar_span_sublinhado(selecionado, estilo_sublinhado)
    if depois:
        span.tail = depois
    node.insert(0, span)
    return True


def odt_sublinhar_tail(item, inicio_local, fim_local, estilo_sublinhado):
    node = item["node"]
    pai = item["pai"]
    if pai is None:
        return False
    texto = node.tail or ""
    antes = texto[:inicio_local]
    selecionado = texto[inicio_local:fim_local]
    depois = texto[fim_local:]
    if not selecionado:
        return False
    node.tail = antes
    span = odt_criar_span_sublinhado(selecionado, estilo_sublinhado)
    if depois:
        span.tail = depois
    indice = list(pai).index(node)
    pai.insert(indice + 1, span)
    return True


def odt_analisar_cabecalho(paragrafo, match, estilo_sublinhado, mapa_estilos):
    mapa = odt_construir_mapa(paragrafo)
    inicio, fim = match.start(), match.end()

    afetados = [
        item for item in mapa
        if not (item["fim"] <= inicio or item["inicio"] >= fim)
        and item["tipo"] in ("text", "tail")
    ]
    if not afetados:
        return "sem"

    partes_sublinhadas = partes_totais = 0
    for item in afetados:
        trecho_inicio = max(inicio, item["inicio"])
        trecho_fim = min(fim, item["fim"])
        if trecho_fim <= trecho_inicio:
            continue
        partes_totais += 1
        if odt_estilo_tem_sublinhado(item["node"], mapa_estilos):
            partes_sublinhadas += 1

    if partes_totais > 0 and partes_sublinhadas == partes_totais:
        return "completo"

    estado = "parcial" if partes_sublinhadas > 0 else "sem"

    alterou = False
    mapa_atual = odt_construir_mapa(paragrafo)
    for item in reversed(mapa_atual):
        if item["tipo"] not in ("text", "tail"):
            continue
        item_inicio, item_fim = item["inicio"], item["fim"]
        if item_fim <= inicio or item_inicio >= fim:
            continue
        trecho_inicio = max(inicio, item_inicio)
        trecho_fim = min(fim, item_fim)
        if trecho_fim <= trecho_inicio:
            continue
        if odt_estilo_tem_sublinhado(item["node"], mapa_estilos):
            continue

        inicio_local = trecho_inicio - item_inicio
        fim_local = trecho_fim - item_inicio

        if item["tipo"] == "text":
            ok = odt_sublinhar_texto_do_node(item, inicio_local, fim_local, estilo_sublinhado)
        else:
            ok = odt_sublinhar_tail(item, inicio_local, fim_local, estilo_sublinhado)
        if ok:
            alterou = True

    return "alterado" if alterou else estado


def odt_processar_paragrafo(paragrafo, estilo_sublinhado, mapa_estilos):
    mapa = odt_construir_mapa(paragrafo)
    texto = "".join(item["texto"] for item in mapa)
    matches = list(PADRAO.finditer(texto))
    if not matches:
        return (0, 0, 0, 0, 0)

    encontrados = len(matches)
    completos = parciais = sem_sublinhado = alterados = 0

    for match in reversed(matches):
        estado = odt_analisar_cabecalho(paragrafo, match, estilo_sublinhado, mapa_estilos)
        if estado == "completo":
            completos += 1
        elif estado == "parcial":
            parciais += 1
            alterados += 1
        elif estado == "sem":
            sem_sublinhado += 1
            alterados += 1
        elif estado == "alterado":
            alterados += 1

    return (encontrados, completos, parciais, sem_sublinhado, alterados)


def processar_odt(arquivo_entrada, arquivo_saida):
    arquivo_entrada = Path(arquivo_entrada)
    arquivo_saida = Path(arquivo_saida)

    with tempfile.TemporaryDirectory() as pasta_temp:
        pasta_temp = Path(pasta_temp)

        with zipfile.ZipFile(arquivo_entrada, "r") as zip_in:
            zip_in.extractall(pasta_temp)

        content_xml = pasta_temp / "content.xml"
        styles_xml = pasta_temp / "styles.xml"

        if not content_xml.exists():
            raise RuntimeError(
                "O arquivo não possui content.xml. Ele não parece ser um ODT válido."
            )

        tree_content = ET.parse(content_xml)
        root_content = tree_content.getroot()

        if styles_xml.exists():
            tree_styles = ET.parse(styles_xml)
            root_styles = tree_styles.getroot()
        else:
            root_styles = ET.Element(f"{{{NS_ODT['office']}}}document-styles")
            tree_styles = ET.ElementTree(root_styles)

        estilos = root_styles.find(f"{{{NS_ODT['office']}}}styles")
        if estilos is None:
            estilos = ET.SubElement(root_styles, f"{{{NS_ODT['office']}}}styles")

        estilo_sublinhado = odt_criar_estilo_sublinhado(estilos)
        mapa_estilos = odt_criar_mapa_estilos(root_styles)

        encontrados = completos = parciais = sem_sublinhado = alterados = 0

        for paragrafo in root_content.iter(f"{{{NS_ODT['text']}}}p"):
            e, c, p, s, a = odt_processar_paragrafo(paragrafo, estilo_sublinhado, mapa_estilos)
            encontrados += e
            completos += c
            parciais += p
            sem_sublinhado += s
            alterados += a

        tree_content.write(content_xml, encoding="UTF-8", xml_declaration=True)
        tree_styles.write(styles_xml, encoding="UTF-8", xml_declaration=True)

        _reempacotar_zip(pasta_temp, arquivo_saida)

    return {
        "encontrados": encontrados, "completos": completos, "parciais": parciais,
        "sem_sublinhado": sem_sublinhado, "alterados": alterados,
    }


# ============================================================
# ============================================================
#   BLOCO DOCX
# ============================================================
# ============================================================
#
# Diferença estrutural importante em relação ao ODT:
#
#   No ODF, um <text:span> pode conter outro <text:span> dentro
#   dele - por isso, na V4/V5, o trecho sublinhado era inserido
#   como FILHO do próprio nó de texto sendo dividido.
#
#   No formato do Word (WordprocessingML), o texto de um parágrafo
#   é uma sequência de "runs" (<w:r>) IRMÃOS - um <w:r> não pode
#   conter outro <w:r> dentro. Cada run carrega sua própria
#   formatação em <w:rPr> (run properties) e seu texto dentro de
#   <w:t>. Por isso, para sublinhar só um TRECHO de um run, esse
#   run precisa ser dividido em até 3 runs IRMÃOS (antes / trecho
#   sublinhado / depois), cada um copiando a formatação original.

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_XML = "http://www.w3.org/XML/1998/namespace"

ET.register_namespace("w", NS_W)


def _w(tag):
    return f"{{{NS_W}}}{tag}"


def docx_coletar_runs(paragrafo):
    """
    Retorna a lista de (run, pai_do_run) em ordem de leitura,
    percorrendo toda a árvore do parágrafo (isso cobre runs dentro
    de <w:hyperlink>, <w:ins> (texto inserido em controle de
    alterações), <w:smartTag> etc., não só filhos diretos do <w:p>).
    """
    runs = []

    def visitar(node):
        for filho in list(node):
            if filho.tag == _w("r"):
                runs.append((filho, node))
            else:
                visitar(filho)

    visitar(paragrafo)
    return runs


def docx_texto_do_run(run):
    """
    Retorna (texto, simples).
    "simples" = True quando o run é o caso comum (só um <w:t>),
    que é seguro de dividir preservando 100% da estrutura.
    Runs com tabulação/quebra de linha/múltiplos elementos são
    lidos normalmente (para não perder texto na detecção do
    cabeçalho), mas não são divididos (limitação conhecida,
    documentada no topo do arquivo).
    """
    conteudo = [filho for filho in run if filho.tag != _w("rPr")]
    texto = ""
    for filho in conteudo:
        if filho.tag == _w("t"):
            texto += filho.text or ""
        elif filho.tag == _w("tab"):
            texto += "\t"
        elif filho.tag in (_w("br"), _w("cr")):
            texto += "\n"
        # outros elementos (bookmarks, campos etc.) não geram texto

    simples = len(conteudo) == 1 and conteudo[0].tag == _w("t")
    return texto, simples


def docx_construir_mapa(paragrafo):
    partes = []
    posicao = 0
    for run, pai in docx_coletar_runs(paragrafo):
        texto, simples = docx_texto_do_run(run)
        if not texto:
            continue
        partes.append({
            "run": run, "pai": pai, "texto": texto,
            "inicio": posicao, "fim": posicao + len(texto),
            "simples": simples,
        })
        posicao += len(texto)
    return partes


def docx_criar_mapa_estilos(root_styles):
    mapa = {}
    if root_styles is None:
        return mapa
    for style in root_styles.iter(_w("style")):
        style_id = style.attrib.get(_w("styleId"))
        if style_id:
            mapa[style_id] = style
    return mapa


def docx_estilo_tem_sublinhado(nome_estilo, mapa_estilos, visitados=None):
    if visitados is None:
        visitados = set()
    if not nome_estilo or nome_estilo in visitados:
        return False
    visitados.add(nome_estilo)

    estilo = mapa_estilos.get(nome_estilo)
    if estilo is None:
        return False

    rPr = estilo.find(_w("rPr"))
    if rPr is not None:
        u = rPr.find(_w("u"))
        if u is not None:
            val = u.attrib.get(_w("val"), "single").lower()
            if val != "none":
                return True

    based_on = estilo.find(_w("basedOn"))
    if based_on is not None:
        return docx_estilo_tem_sublinhado(
            based_on.attrib.get(_w("val")), mapa_estilos, visitados
        )
    return False


def docx_run_tem_sublinhado(run, mapa_estilos):
    """Sublinhado direto no run, ou herdado de um estilo de caractere (rStyle)."""
    rPr = run.find(_w("rPr"))
    if rPr is None:
        return False

    u = rPr.find(_w("u"))
    if u is not None:
        val = u.attrib.get(_w("val"), "single").lower()
        if val != "none":
            return True

    rStyle = rPr.find(_w("rStyle"))
    if rStyle is not None:
        return docx_estilo_tem_sublinhado(rStyle.attrib.get(_w("val")), mapa_estilos)

    return False


def docx_clonar_rpr(run):
    rPr = run.find(_w("rPr"))
    return copy.deepcopy(rPr) if rPr is not None else None


def docx_criar_run(texto, rpr_modelo):
    run = ET.Element(_w("r"))
    if rpr_modelo is not None:
        run.append(copy.deepcopy(rpr_modelo))
    t = ET.SubElement(run, _w("t"))
    t.set(f"{{{NS_XML}}}space", "preserve")
    t.text = texto
    return run


def docx_garantir_sublinhado(run):
    rPr = run.find(_w("rPr"))
    if rPr is None:
        rPr = ET.Element(_w("rPr"))
        run.insert(0, rPr)
    u = rPr.find(_w("u"))
    if u is None:
        u = ET.SubElement(rPr, _w("u"))
    u.set(_w("val"), "single")


def docx_dividir_run(item, inicio_local, fim_local):
    """
    Substitui o run original por até 3 runs irmãos (antes / meio
    sublinhado / depois), copiando a formatação original em cada
    um. Só é chamado quando item["simples"] é True.
    """
    run = item["run"]
    pai = item["pai"]
    texto = item["texto"]

    antes = texto[:inicio_local]
    selecionado = texto[inicio_local:fim_local]
    depois = texto[fim_local:]

    if not selecionado:
        return False

    rpr_modelo = docx_clonar_rpr(run)

    novos = []
    if antes:
        novos.append(docx_criar_run(antes, rpr_modelo))

    run_meio = docx_criar_run(selecionado, rpr_modelo)
    docx_garantir_sublinhado(run_meio)
    novos.append(run_meio)

    if depois:
        novos.append(docx_criar_run(depois, rpr_modelo))

    indice = list(pai).index(run)
    pai.remove(run)
    for deslocamento, novo in enumerate(novos):
        pai.insert(indice + deslocamento, novo)

    return True


def docx_analisar_cabecalho(paragrafo, match, mapa_estilos):
    mapa = docx_construir_mapa(paragrafo)
    inicio, fim = match.start(), match.end()

    afetados = [
        item for item in mapa
        if not (item["fim"] <= inicio or item["inicio"] >= fim)
    ]
    if not afetados:
        return "sem"

    partes_sublinhadas = partes_totais = 0
    for item in afetados:
        trecho_inicio = max(inicio, item["inicio"])
        trecho_fim = min(fim, item["fim"])
        if trecho_fim <= trecho_inicio:
            continue
        partes_totais += 1
        if docx_run_tem_sublinhado(item["run"], mapa_estilos):
            partes_sublinhadas += 1

    if partes_totais > 0 and partes_sublinhadas == partes_totais:
        return "completo"

    estado = "parcial" if partes_sublinhadas > 0 else "sem"

    alterou = False
    mapa_atual = docx_construir_mapa(paragrafo)
    for item in reversed(mapa_atual):
        item_inicio, item_fim = item["inicio"], item["fim"]
        if item_fim <= inicio or item_inicio >= fim:
            continue
        trecho_inicio = max(inicio, item_inicio)
        trecho_fim = min(fim, item_fim)
        if trecho_fim <= trecho_inicio:
            continue
        if docx_run_tem_sublinhado(item["run"], mapa_estilos):
            continue
        if not item["simples"]:
            # Limitação conhecida: run com tabulação/quebra/mistura
            # de elementos não é dividido automaticamente.
            continue

        inicio_local = trecho_inicio - item_inicio
        fim_local = trecho_fim - item_inicio

        if docx_dividir_run(item, inicio_local, fim_local):
            alterou = True

    return "alterado" if alterou else estado


def docx_processar_paragrafo(paragrafo, mapa_estilos):
    mapa = docx_construir_mapa(paragrafo)
    texto = "".join(item["texto"] for item in mapa)
    matches = list(PADRAO.finditer(texto))
    if not matches:
        return (0, 0, 0, 0, 0)

    encontrados = len(matches)
    completos = parciais = sem_sublinhado = alterados = 0

    for match in reversed(matches):
        estado = docx_analisar_cabecalho(paragrafo, match, mapa_estilos)
        if estado == "completo":
            completos += 1
        elif estado == "parcial":
            parciais += 1
            alterados += 1
        elif estado == "sem":
            sem_sublinhado += 1
            alterados += 1
        elif estado == "alterado":
            alterados += 1

    return (encontrados, completos, parciais, sem_sublinhado, alterados)


def processar_docx(arquivo_entrada, arquivo_saida):
    arquivo_entrada = Path(arquivo_entrada)
    arquivo_saida = Path(arquivo_saida)

    with tempfile.TemporaryDirectory() as pasta_temp:
        pasta_temp = Path(pasta_temp)

        with zipfile.ZipFile(arquivo_entrada, "r") as zip_in:
            zip_in.extractall(pasta_temp)

        document_xml = pasta_temp / "word" / "document.xml"
        styles_xml = pasta_temp / "word" / "styles.xml"

        if not document_xml.exists():
            raise RuntimeError(
                "O arquivo não possui word/document.xml. "
                "Ele não parece ser um DOCX válido."
            )

        tree_documento = ET.parse(document_xml)
        root_documento = tree_documento.getroot()

        root_styles = None
        if styles_xml.exists():
            root_styles = ET.parse(styles_xml).getroot()

        mapa_estilos = docx_criar_mapa_estilos(root_styles)

        encontrados = completos = parciais = sem_sublinhado = alterados = 0

        for paragrafo in root_documento.iter(_w("p")):
            e, c, p, s, a = docx_processar_paragrafo(paragrafo, mapa_estilos)
            encontrados += e
            completos += c
            parciais += p
            sem_sublinhado += s
            alterados += a

        tree_documento.write(document_xml, encoding="UTF-8", xml_declaration=True)

        _reempacotar_zip(pasta_temp, arquivo_saida)

    return {
        "encontrados": encontrados, "completos": completos, "parciais": parciais,
        "sem_sublinhado": sem_sublinhado, "alterados": alterados,
    }


# ============================================================
# ============================================================
#   BLOCO DOC (Word 97-2003, via conversão pelo LibreOffice)
# ============================================================
# ============================================================

def localizar_soffice():
    """
    Procura o executável do LibreOffice (soffice) no PATH e em
    locais padrão de instalação por sistema operacional. Retorna
    o caminho, ou None se não encontrar.
    """
    encontrado = shutil.which("soffice") or shutil.which("libreoffice")
    if encontrado:
        return encontrado

    sistema = platform.system()

    candidatos = []
    if sistema == "Windows":
        candidatos = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
    elif sistema == "Darwin":
        candidatos = [
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        ]
    else:
        candidatos = [
            "/usr/bin/soffice",
            "/usr/bin/libreoffice",
            "/opt/libreoffice/program/soffice",
        ]

    for candidato in candidatos:
        if Path(candidato).exists():
            return candidato

    return None


def converter_com_libreoffice(soffice, arquivo, formato_saida, pasta_saida):
    """
    Converte "arquivo" para "formato_saida" (ex.: "docx" ou "doc")
    usando o LibreOffice em modo headless. Retorna o caminho do
    arquivo convertido.
    """
    comando = [
        soffice, "--headless", "--norestore",
        "--convert-to", formato_saida,
        "--outdir", str(pasta_saida), str(arquivo),
    ]

    resultado = subprocess.run(
        comando, capture_output=True, text=True, timeout=180
    )

    esperado = pasta_saida / f"{Path(arquivo).stem}.{formato_saida}"

    if resultado.returncode != 0 or not esperado.exists():
        detalhe = (resultado.stderr or resultado.stdout or "").strip()
        raise RuntimeError(
            "Falha ao converter o arquivo com o LibreOffice.\n"
            f"{detalhe}"
        )

    return esperado


def processar_doc(arquivo_entrada, arquivo_saida):
    soffice = localizar_soffice()

    if not soffice:
        raise RuntimeError(
            "Arquivos .doc (Word 97-2003) exigem o LibreOffice instalado "
            "nesta máquina para serem processados (ele é usado, em segundo "
            "plano, para converter .doc <-> .docx).\n"
            "Baixe gratuitamente em: https://www.libreoffice.org/download/\n"
            "Depois de instalar, rode o programa novamente."
        )

    arquivo_entrada = Path(arquivo_entrada)
    arquivo_saida = Path(arquivo_saida)

    with tempfile.TemporaryDirectory() as pasta_temp:
        pasta_temp = Path(pasta_temp)

        # 1) .doc -> .docx
        docx_convertido = converter_com_libreoffice(
            soffice, arquivo_entrada, "docx", pasta_temp
        )

        # 2) processa o .docx normalmente
        docx_formatado = pasta_temp / f"{arquivo_entrada.stem}_formatado.docx"
        estatisticas = processar_docx(docx_convertido, docx_formatado)

        # 3) .docx -> .doc
        pasta_saida_doc = pasta_temp / "saida_doc"
        pasta_saida_doc.mkdir()
        doc_final = converter_com_libreoffice(
            soffice, docx_formatado, "doc", pasta_saida_doc
        )

        shutil.copyfile(doc_final, arquivo_saida)

    return estatisticas


# ============================================================
# UTILITÁRIO COMUM: reempacotar pasta extraída em .odt/.docx (zip)
# ============================================================

def _reempacotar_zip(pasta_temp, arquivo_saida):
    with zipfile.ZipFile(arquivo_saida, "w") as zip_out:
        mimetype = pasta_temp / "mimetype"

        # O "mimetype" só existe em ODT, e precisa ser o primeiro
        # arquivo do zip, sem compressão (regra do formato ODF).
        if mimetype.exists():
            zip_out.write(mimetype, "mimetype", compress_type=zipfile.ZIP_STORED)

        for arquivo in pasta_temp.rglob("*"):
            if not arquivo.is_file():
                continue
            relativo = arquivo.relative_to(pasta_temp)
            if relativo.as_posix() == "mimetype":
                continue
            zip_out.write(arquivo, relativo.as_posix(), compress_type=zipfile.ZIP_DEFLATED)


# ============================================================
# DISPATCH POR EXTENSÃO + RELATÓRIO
# ============================================================

PROCESSADORES = {
    ".odt": processar_odt,
    ".docx": processar_docx,
    ".doc": processar_doc,
}


def processar_arquivo(arquivo_entrada, arquivo_saida):
    extensao = Path(arquivo_entrada).suffix.lower()
    processador = PROCESSADORES.get(extensao)

    if processador is None:
        raise RuntimeError(
            f"Extensão não suportada: {extensao}. "
            "Formatos aceitos: .odt, .docx, .doc"
        )

    estatisticas = processador(arquivo_entrada, arquivo_saida)

    print()
    print("=" * 65)
    print("                    RESULTADO")
    print("=" * 65)
    print()
    print(f"  Cabeçalhos encontrados       : {estatisticas['encontrados']}")
    print(f"  Já totalmente sublinhados    : {estatisticas['completos']}")
    print(f"  Parcialmente sublinhados     : {estatisticas['parciais']}")
    print(f"  Sem sublinhado               : {estatisticas['sem_sublinhado']}")
    print()
    print(f"  Cabeçalhos alterados         : {estatisticas['alterados']}")
    print()
    print(f"  Arquivo gerado               : {Path(arquivo_saida).name}")
    print(f"  Local                        : {Path(arquivo_saida).parent}")
    print()
    print("=" * 65)

    return estatisticas


# ============================================================
# ARQUIVO ARRASTADO / CAMINHO MANUAL
# ============================================================

def obter_arquivo():
    argumentos = sys.argv[1:]

    if argumentos:
        candidato = " ".join(argumentos).strip('" ')
        if Path(candidato).exists():
            return candidato
        for argumento in argumentos:
            candidato = argumento.strip('" ')
            if Path(candidato).exists():
                return candidato

    print()
    print("Nenhum arquivo foi arrastado.")
    print()

    entrada = input("Digite ou cole o caminho do arquivo (.odt, .docx ou .doc):\n> ")
    return entrada.strip('" ')


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 65)
    print("       PROCESSADOR DE CABEÇALHOS - WHATSAPP V6")
    print("=" * 65)
    print()
    print("Formatos aceitos: .odt, .docx e .doc")
    print()
    print("O programa procura cabeçalhos nos formatos:")
    print("[10/08/2000 11:17:20] Nome:")
    print("[10/08/2000, 11:17:20] Nome:")
    print("10/08/2000 11:17:20 Nome:        (sem colchetes)")
    print("10/08/2000, 11:17:20 Nome:       (sem colchetes)")
    print()
    print("Ele preserva o que já está sublinhado e")
    print("completa somente o que estiver faltando.")
    print()

    arquivo = obter_arquivo()

    if not arquivo:
        print()
        print("[ ERRO ] Nenhum arquivo informado.")
        input("\nPressione ENTER para fechar...")
        return

    caminho_entrada = Path(arquivo)

    if not caminho_entrada.exists():
        print()
        print("[ ERRO ] Arquivo não encontrado:")
        print(caminho_entrada)
        input("\nPressione ENTER para fechar...")
        return

    if caminho_entrada.suffix.lower() not in PROCESSADORES:
        print()
        print("[ ERRO ] Formato não suportado. Use .odt, .docx ou .doc.")
        input("\nPressione ENTER para fechar...")
        return

    caminho_saida = caminho_entrada.with_name(
        f"{caminho_entrada.stem}_formatado{caminho_entrada.suffix}"
    )

    if caminho_saida.resolve() == caminho_entrada.resolve():
        print()
        print("[ ERRO ] O arquivo de saída é igual ao arquivo de entrada.")
        input("\nPressione ENTER para fechar...")
        return

    try:
        print("Processando o documento...")
        processar_arquivo(caminho_entrada, caminho_saida)
        print()
        print("[ SUCESSO ] Processamento concluído.")

    except zipfile.BadZipFile:
        print()
        print("[ ERRO ] O arquivo não é um .odt/.docx válido (zip corrompido).")

    except PermissionError:
        print()
        print("[ ERRO ] Não foi possível acessar o arquivo.")
        print("Verifique se ele está aberto em outro programa.")

    except Exception as erro:
        print()
        print("[ ERRO ] Ocorreu uma falha:")
        print()
        print(str(erro))

    print()
    print("=" * 65)

    input("Pressione ENTER para fechar...")


if __name__ == "__main__":
    main()
