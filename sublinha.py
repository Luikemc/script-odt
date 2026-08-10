import re
import sys
import copy
import shutil
import zipfile
import tempfile
import platform
import subprocess
import traceback
from pathlib import Path
from xml.etree import ElementTree as ET

# ============================================================
# SUBLINHA CABECALHOS - V8
# ============================================================
#
# Novidade da V8: arrastar o arquivo para DENTRO da janela (além
# de arrastar sobre o ícone do programa, que já funcionava desde a
# V7). Isso usa uma biblioteca extra chamada "tkinterdnd2", que não
# vem junto com o Python:
#
#     pip install tkinterdnd2
#
# Se ela não estiver instalada, o programa continua funcionando
# normalmente (botão "Escolher arquivo" e arrastar sobre o ícone),
# só a área de "solte o arquivo aqui" dentro da janela fica
# desativada, com um aviso explicando como habilitá-la.
#
# Objetivo:
#   Encontrar cabeçalhos de conversas no formato:
#
#   [10/08/2000 11:17:20] Lucas Ribeiro:
#   [10/08/2000, 11:17:20] Lucas Ribeiro:
#   [ 10/08/2000 , 11:17:20 ] Lucas Ribeiro:
#   10/08/2000 11:17:20 Lucas Ribeiro:          <- sem colchetes
#   10/08/2000, 11:17:20 Lucas Ribeiro:         <- sem colchetes
#
#   e sublinhar SOMENTE até o ":" do nome.
#
# ------------------------------------------------------------
# O QUE MUDOU DA V6 PARA A V7
# ------------------------------------------------------------
#
#   BUG 1 CORRIGIDO — cabeçalho com quebra de linha no meio:
#     Na V6, a regex excluía \n e \r do trecho do nome
#     ("[^:\n\r]+?"). Se o nome ficasse dividido por uma quebra
#     de linha (ex: "Lucas<quebra>Ribeiro:"), o cabeçalho INTEIRO
#     deixava de casar com o padrão — e como ele simplesmente não
#     é encontrado, nada é sublinhado e nada aparece no relatório
#     como problema. É uma perda total e silenciosa do cabeçalho,
#     bem diferente da "limitação cosmética" que o script dizia
#     ter. Na V7 a regex não exclui mais \n/\r do nome, então o
#     cabeçalho volta a ser encontrado e sublinhado normalmente.
#     (Continua existindo uma limitação bem menor, e essa sim
#     cosmética: o próprio caractere da quebra de linha, assim
#     como um TAB, nunca recebe o traço de sublinhado nele mesmo
#     — só o texto antes e depois dele. Isso é uma limitação do
#     formato, não um bug.)
#
#   BUG 2 CORRIGIDO — TAB/quebra de linha misturado num mesmo
#   "run" do DOCX/DOC fazia o trecho inteiro ser pulado:
#     Na V6, um "run" do Word que misturasse texto normal com um
#     TAB (ou quebra de linha) no meio era marcado como "não
#     simples" e simplesmente ignorado por completo — nem o texto
#     antes, nem o texto depois do TAB recebiam sublinhado. Isso
#     era pior do que o comportamento do ODT (que sublinha em
#     volta do TAB normalmente). Na V7, a função que divide um
#     "run" foi reescrita para lidar com QUALQUER combinação de
#     texto + tab + quebra de linha + outros elementos dentro do
#     mesmo run, preservando tudo o que não é texto (marcadores,
#     etc.) e sublinhando corretamente as partes de texto ao redor
#     do TAB/quebra — igual ao que já acontecia no ODT.
#
#   BUG 3 CORRIGIDO — colchete "solto" (] sem [ correspondente)
#   sendo sublinhado junto com o nome:
#     Na V6, um texto malformado como "10/08/2000 11:17:20] Nome:"
#     (sem colchete de abertura) fazia a regex casar pela variante
#     "sem colchetes" e, como o "]" não é excluído do trecho do
#     nome, ele acabava sendo incluído e sublinhado junto. Na V7,
#     a variante "sem colchetes" só casa se NÃO houver um "]"
#     logo em seguida (evita confundir um colchete de fechamento
#     órfão com parte do cabeçalho), e o "]" também foi excluído
#     do conjunto de caracteres aceitos no nome, como reforço.
#
# ------------------------------------------------------------
# Formatos aceitos (mantido da V6): .odt, .docx e .doc
# ------------------------------------------------------------
#
#        .odt  -> processado diretamente (é um zip com XML).
#        .docx -> processado diretamente (também é um zip com XML,
#                 mas com um esquema (schema) totalmente diferente
#                 do ODF: word/document.xml, elementos <w:p>, <w:r>,
#                 <w:t> etc.).
#        .doc  -> Word 97-2003 (binário/OLE2). É convertido nos
#                 bastidores via LibreOffice (.doc -> .docx ->
#                 processa -> .doc). Exige o LibreOffice instalado
#                 (gratuito: https://www.libreoffice.org/download/).
#
# Características mantidas:
#   - preserva o arquivo original;
#   - mantém sublinhados existentes;
#   - completa sublinhados parciais;
#   - entende texto dividido em vários trechos (spans/runs);
#   - pode ser executado várias vezes;
#   - gera relatório detalhado;
#   - cria automaticamente *_formatado.<extensao original>;
#   - interface gráfica simples (Tkinter) além do modo linha de
#     comando / arquivo arrastado.
#
# ============================================================

# ------------------------------------------------------------
# PADRÃO DO CABEÇALHO (comum a ODT e DOCX)
# ------------------------------------------------------------
#
# Ou casa o bloco "[data hora]" completo (com os dois colchetes),
# ou casa "data hora" sem colchete nenhum — e, nesse segundo caso,
# só casa se não houver um "]" logo em seguida (isso evita tratar
# um colchete de fechamento órfão como se fosse parte do
# cabeçalho: bug 3 da V6).
#
# O trecho do nome aceita QUALQUER caractere, inclusive quebras de
# linha, exceto ":" (fim do nome) e "]" (evita vazar um colchete
# solto para dentro do nome: bug 3 da V6). Antes (V6) ele também
# excluía \n e \r, o que causava o bug 1.

PADRAO = re.compile(
    r'(?:'
    r'\[\s*\d{2}/\d{2}/\d{4}\s*(?:,\s*)?\d{2}:\d{2}:\d{2}\s*\]'
    r'|'
    r'\d{2}/\d{2}/\d{4}\s*(?:,\s*)?\d{2}:\d{2}:\d{2}(?!\s*\])'
    r')'
    r'\s*[^:\]]+?:'
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
#   dele. No formato do Word (WordprocessingML), o texto de um
#   parágrafo é uma sequência de "runs" (<w:r>) IRMÃOS - um <w:r>
#   não pode conter outro <w:r> dentro. Cada run carrega sua
#   própria formatação em <w:rPr> (run properties) e seu conteúdo
#   dentro de elementos como <w:t> (texto), <w:tab/> (tabulação) e
#   <w:br/>/<w:cr/> (quebra de linha). Por isso, para sublinhar só
#   um TRECHO de um run, esse run precisa ser dividido em até 3
#   runs IRMÃOS (antes / trecho sublinhado / depois), cada um
#   copiando a formatação original.
#
#   Na V6, essa divisão só era feita para o caso simples (run com
#   um único <w:t> e nada mais). Um run que misturasse texto com
#   <w:tab/> ou <w:br/> no meio era pulado por inteiro (bug 2). Na
#   V7, a divisão do run foi generalizada: ela entende uma
#   sequência qualquer de <w:t>/<w:tab/>/<w:br/>/<w:cr/> (e
#   preserva sem alteração qualquer outro elemento raro que
#   apareça no run, como marcadores internos do Word, para não
#   perder nada do documento), e sabe dividir tudo isso em até 3
#   runs preservando a formatação original em cada pedaço.

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


def docx_conteudo_do_run(run):
    """
    Descreve o CONTEÚDO de um run (tudo, exceto <w:rPr>) como uma
    lista ordenada de pedaços:

        {"elemento": <nó XML original>, "tipo": "t"/"tab"/"br"/"outro",
         "texto": <string usada para posicionar e para underline>}

    "texto" é o que entra no mapa de caracteres do parágrafo:
      - "t"     -> o texto de dentro de <w:t>
      - "tab"   -> "\t" (um caractere, representando a tabulação)
      - "br"/"cr" -> "\n" (um caractere, representando a quebra)
      - "outro" -> "" (elementos sem representação textual, como
                   marcadores/bookmarks; não ocupam posição, mas
                   são preservados na reconstrução do run)
    """
    pedacos = []
    for filho in run:
        if filho.tag == _w("rPr"):
            continue
        if filho.tag == _w("t"):
            pedacos.append({"elemento": filho, "tipo": "t", "texto": filho.text or ""})
        elif filho.tag == _w("tab"):
            pedacos.append({"elemento": filho, "tipo": "tab", "texto": "\t"})
        elif filho.tag in (_w("br"), _w("cr")):
            pedacos.append({"elemento": filho, "tipo": "br", "texto": "\n"})
        else:
            # Elemento sem texto (bookmark, campo, etc.): preservado
            # na reconstrução, mas não ocupa posição no mapa.
            pedacos.append({"elemento": filho, "tipo": "outro", "texto": ""})
    return pedacos


def docx_texto_do_run(run):
    return "".join(p["texto"] for p in docx_conteudo_do_run(run))


def docx_construir_mapa(paragrafo):
    partes = []
    posicao = 0
    for run, pai in docx_coletar_runs(paragrafo):
        texto = docx_texto_do_run(run)
        if not texto:
            continue
        partes.append({
            "run": run, "pai": pai, "texto": texto,
            "inicio": posicao, "fim": posicao + len(texto),
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


def docx_montar_run(pedacos, rpr_modelo):
    """Constrói um novo <w:r> a partir de uma lista de pedaços (na
    mesma ordem em que devem aparecer), copiando a formatação
    original. Elementos "outro" e "tab"/"br" são reaproveitados via
    deepcopy do nó original (preserva quaisquer atributos que
    tenham); pedaços de texto ("t") viram um <w:t xml:space="preserve">
    novo, para garantir que espaços nas bordas do corte não se
    percam."""
    run = ET.Element(_w("r"))
    if rpr_modelo is not None:
        run.append(copy.deepcopy(rpr_modelo))
    for pedaco in pedacos:
        if pedaco["tipo"] == "t":
            t = ET.SubElement(run, _w("t"))
            t.set(f"{{{NS_XML}}}space", "preserve")
            t.text = pedaco["texto"]
        else:
            run.append(copy.deepcopy(pedaco["elemento"]))
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
    um. Funciona com qualquer combinação de texto + tab + quebra de
    linha + outros elementos dentro do run (bug 2 da V6: antes só
    funcionava para um run com um único <w:t>).
    """
    run = item["run"]
    pai = item["pai"]

    if fim_local <= inicio_local:
        return False

    rpr_modelo = docx_clonar_rpr(run)

    antes, meio, depois = [], [], []
    pos = 0
    for pedaco in docx_conteudo_do_run(run):
        tamanho = len(pedaco["texto"])

        if tamanho == 0:
            # Elemento sem posição própria (bookmark, campo etc.):
            # colocado no bloco correspondente à posição atual.
            if pos < inicio_local:
                antes.append(pedaco)
            elif pos < fim_local:
                meio.append(pedaco)
            else:
                depois.append(pedaco)
            continue

        elem_inicio, elem_fim = pos, pos + tamanho

        if elem_fim <= inicio_local:
            antes.append(pedaco)
        elif elem_inicio >= fim_local:
            depois.append(pedaco)
        elif pedaco["tipo"] == "t":
            corte1 = max(0, inicio_local - elem_inicio)
            corte2 = min(tamanho, fim_local - elem_inicio)
            texto = pedaco["texto"]
            parte_antes, parte_meio, parte_depois = (
                texto[:corte1], texto[corte1:corte2], texto[corte2:]
            )
            if parte_antes:
                antes.append({"tipo": "t", "texto": parte_antes})
            if parte_meio:
                meio.append({"tipo": "t", "texto": parte_meio})
            if parte_depois:
                depois.append({"tipo": "t", "texto": parte_depois})
        else:
            # tab/br: elemento atômico (não dá para cortar no meio
            # de uma tabulação/quebra) — como ele se sobrepõe à
            # seleção, entra inteiro no bloco do meio.
            meio.append(pedaco)

        pos = elem_fim

    if not meio:
        return False

    novos = []
    if antes:
        novos.append(docx_montar_run(antes, rpr_modelo))
    run_meio = docx_montar_run(meio, rpr_modelo)
    docx_garantir_sublinhado(run_meio)
    novos.append(run_meio)
    if depois:
        novos.append(docx_montar_run(depois, rpr_modelo))

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


def formatar_relatorio(estatisticas, arquivo_saida):
    linhas = []
    linhas.append("=" * 65)
    linhas.append("                    RESULTADO")
    linhas.append("=" * 65)
    linhas.append("")
    linhas.append(f"  Cabeçalhos encontrados       : {estatisticas['encontrados']}")
    linhas.append(f"  Já totalmente sublinhados    : {estatisticas['completos']}")
    linhas.append(f"  Parcialmente sublinhados     : {estatisticas['parciais']}")
    linhas.append(f"  Sem sublinhado               : {estatisticas['sem_sublinhado']}")
    linhas.append("")
    linhas.append(f"  Cabeçalhos alterados         : {estatisticas['alterados']}")
    linhas.append("")
    linhas.append(f"  Arquivo gerado               : {Path(arquivo_saida).name}")
    linhas.append(f"  Local                        : {Path(arquivo_saida).parent}")
    linhas.append("")
    linhas.append("=" * 65)
    return "\n".join(linhas)


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
    print(formatar_relatorio(estatisticas, arquivo_saida))
    return estatisticas


def caminho_saida_padrao(caminho_entrada):
    caminho_entrada = Path(caminho_entrada)
    return caminho_entrada.with_name(
        f"{caminho_entrada.stem}_formatado{caminho_entrada.suffix}"
    )


# ============================================================
# ============================================================
#   INTERFACE GRÁFICA (Tkinter)
# ============================================================
# ============================================================
#
# Bem simples de propósito: uma área para arrastar o arquivo (ou,
# se preferir, um botão para escolher o arquivo), que já processa
# na hora, e uma caixa de texto mostrando o que aconteceu.
#
# A área de "arrastar e soltar" depende da biblioteca opcional
# "tkinterdnd2" (pip install tkinterdnd2). Se ela não estiver
# instalada, a interface toda funciona normalmente do mesmo jeito,
# só que sem a opção de soltar o arquivo dentro da janela — nesse
# caso ainda dá para usar o botão, ou arrastar o arquivo por cima
# do ícone do programa/.exe antes de abrir (isso nunca dependeu de
# biblioteca nenhuma).

def _tentar_importar_dnd():
    """Tenta importar o tkinterdnd2. Retorna (TkinterDnD, DND_FILES)
    ou (None, None) se não estiver instalado."""
    try:
        from tkinterdnd2 import TkinterDnD, DND_FILES
        return TkinterDnD, DND_FILES
    except Exception:
        return None, None


def _extrair_caminhos_soltos(janela, dados_evento):
    """
    Converte o texto bruto que o tkinterdnd2 entrega no evento de
    soltar arquivo em uma lista de caminhos. Caminhos com espaço no
    nome vêm entre chaves, ex: "{C:/pasta/meu arquivo.docx}" — por
    isso usamos janela.tk.splitlist, que entende esse formato,
    em vez de um simples .split().
    """
    try:
        return list(janela.tk.splitlist(dados_evento))
    except Exception:
        return [dados_evento]


def rodar_gui(arquivo_inicial=None):
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    TkinterDnD, DND_FILES = _tentar_importar_dnd()
    arrastar_disponivel = TkinterDnD is not None

    janela = TkinterDnD.Tk() if arrastar_disponivel else tk.Tk()
    janela.title("Sublinhar Cabeçalhos - WhatsApp")
    janela.geometry("640x520")
    janela.minsize(520, 420)

    tk.Label(
        janela,
        text="Sublinhar cabeçalhos de conversa (.odt, .docx, .doc)",
        font=("Segoe UI", 13, "bold"),
        pady=10,
    ).pack()

    tk.Label(
        janela,
        text=(
            "Procura por cabeçalhos como \"[10/08/2000 11:17:20] Nome:\"\n"
            "e sublinha até o nome. Mantém o que já está sublinhado e\n"
            "gera um novo arquivo \"..._formatado\" — o original não é alterado."
        ),
        justify="center",
        fg="#444444",
    ).pack(pady=(0, 10))

    # ------------------------------------------------------------
    # Área de arrastar-e-soltar
    # ------------------------------------------------------------
    if arrastar_disponivel:
        texto_area = "Arraste aqui o arquivo (.odt, .docx ou .doc)"
        cor_fundo = "#eef6ff"
        cor_borda = "#4a90d9"
    else:
        texto_area = (
            "Arrastar-e-soltar não está disponível nesta instalação.\n"
            "Use o botão abaixo, ou instale com: pip install tkinterdnd2"
        )
        cor_fundo = "#f2f2f2"
        cor_borda = "#999999"

    area_soltar = tk.Label(
        janela,
        text=texto_area,
        font=("Segoe UI", 11),
        bg=cor_fundo,
        fg="#333333",
        relief="ridge",
        bd=2,
        height=4,
        wraplength=520,
        justify="center",
    )
    area_soltar.pack(fill="x", padx=16, pady=(0, 10))

    caixa_log = scrolledtext.ScrolledText(
        janela, height=14, font=("Consolas", 10), state="disabled", wrap="word"
    )
    caixa_log.pack(fill="both", expand=True, padx=12, pady=(0, 10))

    barra_botoes = tk.Frame(janela)
    barra_botoes.pack(pady=(0, 12))

    def escrever_log(texto):
        caixa_log.configure(state="normal")
        caixa_log.insert("end", texto + "\n")
        caixa_log.see("end")
        caixa_log.configure(state="disabled")

    def limpar_log():
        caixa_log.configure(state="normal")
        caixa_log.delete("1.0", "end")
        caixa_log.configure(state="disabled")

    ultimo_arquivo_gerado = {"caminho": None}

    def processar(caminho):
        caminho_entrada = Path(caminho)

        if not caminho_entrada.exists():
            messagebox.showerror("Erro", f"Arquivo não encontrado:\n{caminho_entrada}")
            return

        if caminho_entrada.suffix.lower() not in PROCESSADORES:
            messagebox.showerror(
                "Formato não suportado",
                "Use um arquivo .odt, .docx ou .doc.",
            )
            return

        caminho_saida = caminho_saida_padrao(caminho_entrada)
        if caminho_saida.resolve() == caminho_entrada.resolve():
            messagebox.showerror("Erro", "O arquivo de saída ficaria igual ao de entrada.")
            return

        limpar_log()
        escrever_log(f"Processando: {caminho_entrada.name}")
        escrever_log("Aguarde...\n")
        janela.update_idletasks()

        try:
            estatisticas = PROCESSADORES[caminho_entrada.suffix.lower()](
                caminho_entrada, caminho_saida
            )
            escrever_log(formatar_relatorio(estatisticas, caminho_saida))
            escrever_log("\n[ SUCESSO ] Processamento concluído.")
            ultimo_arquivo_gerado["caminho"] = caminho_saida
            botao_abrir_pasta.configure(state="normal")
        except zipfile.BadZipFile:
            escrever_log("[ ERRO ] O arquivo não é um .odt/.docx válido (zip corrompido).")
        except PermissionError:
            escrever_log(
                "[ ERRO ] Não foi possível acessar o arquivo.\n"
                "Verifique se ele está aberto em outro programa."
            )
        except Exception as erro:
            escrever_log("[ ERRO ] Ocorreu uma falha:\n")
            escrever_log(str(erro))
            escrever_log("")
            escrever_log(traceback.format_exc())

    def processar_varios(caminhos):
        # Se soltarem mais de um arquivo de uma vez, processa cada
        # um (gera um "..._formatado" para cada), um após o outro.
        for caminho in caminhos:
            processar(caminho)

    def escolher_arquivo():
        caminho = filedialog.askopenfilename(
            title="Escolha o arquivo da conversa",
            filetypes=[
                ("Documentos suportados", "*.odt *.docx *.doc"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        if caminho:
            processar(caminho)

    def abrir_pasta():
        caminho = ultimo_arquivo_gerado["caminho"]
        if not caminho:
            return
        pasta = str(Path(caminho).parent)
        sistema = platform.system()
        try:
            if sistema == "Windows":
                subprocess.run(["explorer", pasta])
            elif sistema == "Darwin":
                subprocess.run(["open", pasta])
            else:
                subprocess.run(["xdg-open", pasta])
        except Exception:
            messagebox.showinfo("Local do arquivo", pasta)

    # Ativa o "soltar arquivo" na área dedicada (e também na caixa
    # de log, para dar uma margem de erro maior a quem soltar um
    # pouco fora da área marcada).
    if arrastar_disponivel:
        def ao_soltar(evento):
            caminhos = _extrair_caminhos_soltos(janela, evento.data)
            if caminhos:
                processar_varios(caminhos)

        for alvo in (area_soltar, caixa_log):
            alvo.drop_target_register(DND_FILES)
            alvo.dnd_bind("<<Drop>>", ao_soltar)

    tk.Button(
        barra_botoes,
        text="Escolher arquivo e processar",
        font=("Segoe UI", 11, "bold"),
        command=escolher_arquivo,
        padx=14, pady=8,
    ).pack(side="left", padx=6)

    botao_abrir_pasta = tk.Button(
        barra_botoes,
        text="Abrir pasta do resultado",
        command=abrir_pasta,
        state="disabled",
        padx=14, pady=8,
    )
    botao_abrir_pasta.pack(side="left", padx=6)

    escrever_log(
        "Pronto. Arraste um arquivo para a área acima (ou clique em "
        "\"Escolher arquivo e processar\").\nFormatos aceitos: .odt, .docx e .doc"
    )
    if not arrastar_disponivel:
        escrever_log(
            "\n[ AVISO ] Arrastar-e-soltar dentro da janela está desativado "
            "porque a biblioteca \"tkinterdnd2\" não foi encontrada.\n"
            "Para habilitar: pip install tkinterdnd2  (e rode o programa de novo).\n"
            "Enquanto isso, use o botão acima, ou arraste o arquivo por cima "
            "do ícone do programa antes de abrir."
        )

    if arquivo_inicial:
        janela.after(200, lambda: processar(arquivo_inicial))

    janela.mainloop()


# ============================================================
# MAIN
# ============================================================
#
# - Sem argumentos: abre a interface gráfica.
# - Com um arquivo arrastado/passado por linha de comando: abre a
#   interface gráfica já processando esse arquivo (mantém o antigo
#   comportamento de "arrastar o arquivo para o programa", só que
#   agora com uma janela em vez do modo texto).

def main():
    argumentos = sys.argv[1:]
    arquivo_inicial = None

    if argumentos:
        candidato = " ".join(argumentos).strip('" ')
        if Path(candidato).exists():
            arquivo_inicial = candidato
        else:
            for argumento in argumentos:
                candidato = argumento.strip('" ')
                if Path(candidato).exists():
                    arquivo_inicial = candidato
                    break

    try:
        rodar_gui(arquivo_inicial)
    except ImportError:
        print("[ ERRO ] Tkinter não está disponível nesta instalação do Python.")
        input("\nPressione ENTER para fechar...")


if __name__ == "__main__":
    main()
