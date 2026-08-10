import re
import sys
import zipfile
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET

# ============================================================
# SUBLINHA ODT - V3
# ============================================================
#
# Objetivo:
#   Encontrar cabeçalhos de conversas no formato:
#
#   [10/08/2000 11:17:20] Lucas Ribeiro:
#   [10/08/2000, 11:17:20] Lucas Ribeiro:
#   [ 10/08/2000 , 11:17:20 ] Lucas Ribeiro:
#
#   e sublinhar SOMENTE até o ":" do nome.
#
# Características da V3:
#   - preserva o arquivo original;
#   - mantém sublinhados existentes;
#   - completa sublinhados parciais;
#   - entende texto dividido em vários <text:span>;
#   - não depende de o cabeçalho estar em um único nó XML;
#   - pode ser executada várias vezes;
#   - gera relatório detalhado;
#   - aceita arquivo arrastado para o .exe;
#   - cria automaticamente *_formatado.odt.
#
# ============================================================

PADRAO = re.compile(
    r'\[\s*'
    r'\d{2}/\d{2}/\d{4}'
    r'\s*(?:,\s*)?'
    r'\d{2}:\d{2}:\d{2}'
    r'\s*\]\s*'
    r'[^:\n\r]+?'
    r':'
)

NS = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
}

for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


# ============================================================
# TEXTO / ESTRUTURA XML
# ============================================================

def texto_especial(node):
    """Retorna o texto representado por elementos especiais ODT."""
    if node.tag == f"{{{NS['text']}}}s":
        quantidade = int(
            node.attrib.get(
                f"{{{NS['text']}}}c",
                "1"
            )
        )
        return " " * quantidade

    if node.tag == f"{{{NS['text']}}}tab":
        return "\t"

    if node.tag == f"{{{NS['text']}}}line-break":
        return "\n"

    return ""


def construir_mapa(paragrafo):
    """
    Constrói uma visão linear do texto do parágrafo.

    Exemplo:

        XML:
          span("Lucas")
          span(" Ribeiro:")

        Mapa:
          0..5   -> primeiro span
          5..14  -> segundo span

    Isso permite localizar um cabeçalho mesmo quando ele está
    dividido em vários elementos XML.
    """

    partes = []
    posicao = 0

    def adicionar(node, texto, tipo):
        nonlocal posicao

        if not texto:
            return

        partes.append({
            "node": node,
            "texto": texto,
            "inicio": posicao,
            "fim": posicao + len(texto),
            "tipo": tipo,
        })

        posicao += len(texto)

    def visitar(node):

        if node.text:
            adicionar(
                node,
                node.text,
                "text"
            )

        for filho in list(node):

            especial = texto_especial(filho)

            if especial:
                adicionar(
                    filho,
                    especial,
                    "special"
                )
            else:
                visitar(filho)

            # Texto que vem depois de um filho XML.
            if filho.tail:
                adicionar(
                    filho,
                    filho.tail,
                    "tail"
                )

    visitar(paragrafo)

    return partes


# ============================================================
# ESTILOS
# ============================================================

def criar_estilo_sublinhado(styles_root):
    """
    Cria o estilo utilizado pelo programa.

    Se já existir, reutiliza.
    """

    nome = "SublinhadoAutomatico"

    for style in styles_root.iter(
        f"{{{NS['style']}}}style"
    ):
        if style.attrib.get(
            f"{{{NS['style']}}}name"
        ) == nome:
            return nome

    estilo = ET.SubElement(
        styles_root,
        f"{{{NS['style']}}}style",
        {
            f"{{{NS['style']}}}name": nome,
            f"{{{NS['style']}}}family": "text",
        }
    )

    propriedades = ET.SubElement(
        estilo,
        f"{{{NS['style']}}}text-properties"
    )

    propriedades.set(
        f"{{{NS['fo']}}}text-decoration",
        "underline"
    )

    return nome


def criar_mapa_estilos(root_styles):
    """
    Cria um mapa:
        nome_do_estilo -> elemento XML

    Assim conseguimos descobrir se um span usa um estilo
    que já possui sublinhado.
    """

    mapa = {}

    for style in root_styles.iter(
        f"{{{NS['style']}}}style"
    ):
        nome = style.attrib.get(
            f"{{{NS['style']}}}name"
        )

        if nome:
            mapa[nome] = style

    return mapa


def estilo_tem_sublinhado(node, mapa_estilos):
    """
    Verifica se um nó está sublinhado.

    Considera:
      - estilo direto do span;
      - estilo pai;
      - text-decoration do estilo.
    """

    if node is None:
        return False

    nome = node.attrib.get(
        f"{{{NS['text']}}}style-name"
    )

    visitados = set()

    while nome and nome not in visitados:

        visitados.add(nome)

        estilo = mapa_estilos.get(nome)

        if estilo is None:
            break

        propriedades = estilo.find(
            f"{{{NS['style']}}}text-properties"
        )

        if propriedades is not None:

            decoracao = propriedades.attrib.get(
                f"{{{NS['fo']}}}text-decoration",
                ""
            ).lower()

            if "underline" in decoracao:
                return True

        nome = estilo.attrib.get(
            f"{{{NS['style']}}}parent-style-name"
        )

    return False


# ============================================================
# XML - PAIS
# ============================================================

def encontrar_pai(raiz, filho):
    for elemento in raiz.iter():

        if filho in list(elemento):
            return elemento

    return None


# ============================================================
# SUBLINHADO
# ============================================================

def criar_span_sublinhado(
    texto,
    estilo_sublinhado
):
    span = ET.Element(
        f"{{{NS['text']}}}span",
        {
            f"{{{NS['text']}}}style-name":
                estilo_sublinhado
        }
    )

    span.text = texto

    return span


def sublinhar_texto_do_node(
    paragrafo,
    item,
    inicio_local,
    fim_local,
    estilo_sublinhado
):
    """
    Sublinha um pedaço do .text de um elemento XML.

    Exemplo:

        "ABCDEF"
           ^^^^

    vira:

        "AB"
        <span>CD</span>
        "EF"
    """

    node = item["node"]

    texto = node.text or ""

    antes = texto[:inicio_local]
    selecionado = texto[
        inicio_local:fim_local
    ]
    depois = texto[fim_local:]

    if not selecionado:
        return False

    node.text = antes

    span = criar_span_sublinhado(
        selecionado,
        estilo_sublinhado
    )

    pai = encontrar_pai(
        paragrafo,
        node
    )

    if pai is None:
        return False

    indice = list(pai).index(node)

    pai.insert(
        indice + 1,
        span
    )

    if depois:
        span.tail = depois

    return True


def sublinhar_tail(
    paragrafo,
    item,
    inicio_local,
    fim_local,
    estilo_sublinhado
):
    """
    Faz o mesmo que a função anterior, mas quando o texto
    está no .tail de um elemento XML.
    """

    node = item["node"]

    texto = node.tail or ""

    antes = texto[:inicio_local]
    selecionado = texto[
        inicio_local:fim_local
    ]
    depois = texto[fim_local:]

    if not selecionado:
        return False

    node.tail = antes

    span = criar_span_sublinhado(
        selecionado,
        estilo_sublinhado
    )

    pai = encontrar_pai(
        paragrafo,
        node
    )

    if pai is None:
        return False

    indice = list(pai).index(node)

    pai.insert(
        indice + 1,
        span
    )

    if depois:
        span.tail = depois

    return True


# ============================================================
# PROCESSAMENTO DE UM CABEÇALHO
# ============================================================

def analisar_cabecalho(
    paragrafo,
    match,
    estilo_sublinhado,
    mapa_estilos
):
    """
    Analisa UM cabeçalho.

    Retorna:

        "completo"
        "parcial"
        "sem"
        "alterado"
    """

    mapa = construir_mapa(paragrafo)

    inicio = match.start()
    fim = match.end()

    afetados = []

    for item in mapa:

        if item["fim"] <= inicio:
            continue

        if item["inicio"] >= fim:
            continue

        if item["tipo"] not in (
            "text",
            "tail"
        ):
            continue

        afetados.append(item)

    if not afetados:
        return "sem"

    partes_sublinhadas = 0
    partes_totais = 0

    for item in afetados:

        trecho_inicio = max(
            inicio,
            item["inicio"]
        )

        trecho_fim = min(
            fim,
            item["fim"]
        )

        if trecho_fim <= trecho_inicio:
            continue

        partes_totais += 1

        if estilo_tem_sublinhado(
            item["node"],
            mapa_estilos
        ):
            partes_sublinhadas += 1

    if partes_sublinhadas == partes_totais:
        return "completo"

    if partes_sublinhadas > 0:
        estado = "parcial"
    else:
        estado = "sem"

    # --------------------------------------------------------
    # Segunda etapa:
    # sublinha somente as partes que ainda não possuem
    # sublinhado.
    # --------------------------------------------------------

    alterou = False

    # Recria o mapa porque inserções anteriores podem alterar
    # a estrutura XML.
    mapa_atual = construir_mapa(paragrafo)

    for item in reversed(mapa_atual):

        if item["tipo"] not in (
            "text",
            "tail"
        ):
            continue

        item_inicio = item["inicio"]
        item_fim = item["fim"]

        if item_fim <= inicio:
            continue

        if item_inicio >= fim:
            continue

        trecho_inicio = max(
            inicio,
            item_inicio
        )

        trecho_fim = min(
            fim,
            item_fim
        )

        if trecho_fim <= trecho_inicio:
            continue

        # Se esse nó já está sublinhado,
        # preservamos exatamente como está.
        if estilo_tem_sublinhado(
            item["node"],
            mapa_estilos
        ):
            continue

        inicio_local = (
            trecho_inicio - item_inicio
        )

        fim_local = (
            trecho_fim - item_inicio
        )

        if item["tipo"] == "text":

            ok = sublinhar_texto_do_node(
                paragrafo,
                item,
                inicio_local,
                fim_local,
                estilo_sublinhado
            )

        else:

            ok = sublinhar_tail(
                paragrafo,
                item,
                inicio_local,
                fim_local,
                estilo_sublinhado
            )

        if ok:
            alterou = True

    if alterou:
        return "alterado"

    return estado


# ============================================================
# PROCESSAMENTO DO PARÁGRAFO
# ============================================================

def processar_paragrafo(
    paragrafo,
    estilo_sublinhado,
    mapa_estilos
):
    """
    Retorna:

        encontrados
        completos
        parciais
        sem_sublinhado
        alterados
    """

    mapa = construir_mapa(
        paragrafo
    )

    texto = "".join(
        item["texto"]
        for item in mapa
    )

    matches = list(
        PADRAO.finditer(texto)
    )

    if not matches:
        return (
            0,
            0,
            0,
            0,
            0
        )

    encontrados = len(matches)
    completos = 0
    parciais = 0
    sem_sublinhado = 0
    alterados = 0

    # --------------------------------------------------------
    # Importante:
    # processamos de trás para frente.
    # --------------------------------------------------------

    for match in reversed(matches):

        estado = analisar_cabecalho(
            paragrafo,
            match,
            estilo_sublinhado,
            mapa_estilos
        )

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

    return (
        encontrados,
        completos,
        parciais,
        sem_sublinhado,
        alterados
    )


# ============================================================
# PROCESSAMENTO DO ODT
# ============================================================

def processar_odt(
    arquivo_entrada,
    arquivo_saida
):

    arquivo_entrada = Path(
        arquivo_entrada
    )

    arquivo_saida = Path(
        arquivo_saida
    )

    with tempfile.TemporaryDirectory() as pasta_temp:

        pasta_temp = Path(
            pasta_temp
        )

        # ----------------------------------------------------
        # Extrai ODT
        # ----------------------------------------------------

        with zipfile.ZipFile(
            arquivo_entrada,
            "r"
        ) as zip_in:

            zip_in.extractall(
                pasta_temp
            )

        content_xml = (
            pasta_temp / "content.xml"
        )

        styles_xml = (
            pasta_temp / "styles.xml"
        )

        if not content_xml.exists():
            raise RuntimeError(
                "O arquivo não possui content.xml. "
                "Ele não parece ser um ODT válido."
            )

        # ----------------------------------------------------
        # XML do conteúdo
        # ----------------------------------------------------

        tree_content = ET.parse(
            content_xml
        )

        root_content = (
            tree_content.getroot()
        )

        # ----------------------------------------------------
        # XML dos estilos
        # ----------------------------------------------------

        if styles_xml.exists():

            tree_styles = ET.parse(
                styles_xml
            )

            root_styles = (
                tree_styles.getroot()
            )

        else:

            root_styles = ET.Element(
                f"{{{NS['office']}}}"
                "document-styles"
            )

            tree_styles = ET.ElementTree(
                root_styles
            )

        estilos = root_styles.find(
            f"{{{NS['office']}}}styles"
        )

        if estilos is None:

            estilos = ET.SubElement(
                root_styles,
                f"{{{NS['office']}}}styles"
            )

        # ----------------------------------------------------
        # Estilo do programa
        # ----------------------------------------------------

        estilo_sublinhado = (
            criar_estilo_sublinhado(
                estilos
            )
        )

        mapa_estilos = (
            criar_mapa_estilos(
                root_styles
            )
        )

        # ----------------------------------------------------
        # Contadores
        # ----------------------------------------------------

        encontrados = 0
        completos = 0
        parciais = 0
        sem_sublinhado = 0
        alterados = 0

        # ----------------------------------------------------
        # Processa todos os parágrafos
        # ----------------------------------------------------

        paragrafos = root_content.iter(
            f"{{{NS['text']}}}p"
        )

        for paragrafo in paragrafos:

            (
                e,
                c,
                p,
                s,
                a
            ) = processar_paragrafo(
                paragrafo,
                estilo_sublinhado,
                mapa_estilos
            )

            encontrados += e
            completos += c
            parciais += p
            sem_sublinhado += s
            alterados += a

        # ----------------------------------------------------
        # Salva XML
        # ----------------------------------------------------

        tree_content.write(
            content_xml,
            encoding="UTF-8",
            xml_declaration=True
        )

        tree_styles.write(
            styles_xml,
            encoding="UTF-8",
            xml_declaration=True
        )

        # ----------------------------------------------------
        # Recria o ODT
        # ----------------------------------------------------

        with zipfile.ZipFile(
            arquivo_saida,
            "w"
        ) as zip_out:

            mimetype = (
                pasta_temp / "mimetype"
            )

            if mimetype.exists():

                zip_out.write(
                    mimetype,
                    "mimetype",
                    compress_type=zipfile.ZIP_STORED
                )

            for arquivo in pasta_temp.rglob("*"):

                if not arquivo.is_file():
                    continue

                relativo = (
                    arquivo.relative_to(
                        pasta_temp
                    )
                )

                if relativo.as_posix() == "mimetype":
                    continue

                zip_out.write(
                    arquivo,
                    relativo.as_posix(),
                    compress_type=zipfile.ZIP_DEFLATED
                )

    # --------------------------------------------------------
    # Resultado
    # --------------------------------------------------------

    print()
    print("=" * 65)
    print("                    RESULTADO")
    print("=" * 65)
    print()
    print(
        f"  Cabeçalhos encontrados       : {encontrados}"
    )
    print(
        f"  Já totalmente sublinhados    : {completos}"
    )
    print(
        f"  Parcialmente sublinhados     : {parciais}"
    )
    print(
        f"  Sem sublinhado               : {sem_sublinhado}"
    )
    print()
    print(
        f"  Cabeçalhos alterados         : {alterados}"
    )
    print()
    print(
        f"  Arquivo gerado               : "
        f"{arquivo_saida.name}"
    )
    print(
        f"  Local                        : "
        f"{arquivo_saida.parent}"
    )
    print()
    print("=" * 65)

    return {
        "encontrados": encontrados,
        "completos": completos,
        "parciais": parciais,
        "sem_sublinhado": sem_sublinhado,
        "alterados": alterados,
    }


# ============================================================
# ARQUIVO ARRASTADO / CAMINHO MANUAL
# ============================================================

def obter_arquivo():

    argumentos = sys.argv[1:]

    # --------------------------------------------------------
    # 1. Arquivo arrastado para o programa
    # --------------------------------------------------------

    if argumentos:

        # Normalmente o Windows envia o caminho como um único
        # argumento, mesmo quando possui espaços.
        candidato = " ".join(
            argumentos
        ).strip('" ')

        if Path(candidato).exists():
            return candidato

        # Segurança: tenta cada argumento individualmente.
        for argumento in argumentos:

            candidato = argumento.strip(
                '" '
            )

            if Path(candidato).exists():
                return candidato

    # --------------------------------------------------------
    # 2. Caminho digitado manualmente
    # --------------------------------------------------------

    print()
    print(
        "Nenhum arquivo foi arrastado."
    )
    print()

    entrada = input(
        "Digite ou cole o caminho do arquivo .odt:\n"
        "> "
    )

    return entrada.strip(
        '" '
    )


# ============================================================
# MAIN
# ============================================================

def main():

    print("=" * 65)
    print(
        "          PROCESSADOR ODT - WHATSAPP V3"
    )
    print("=" * 65)
    print()
    print(
        "O programa procura cabeçalhos no formato:"
    )
    print(
        "[10/08/2000 11:17:20] Nome:"
    )
    print(
        "[10/08/2000, 11:17:20] Nome:"
    )
    print()
    print(
        "Ele preserva o que já está sublinhado e"
    )
    print(
        "completa somente o que estiver faltando."
    )
    print()

    arquivo = obter_arquivo()

    if not arquivo:

        print()
        print(
            "[ ERRO ] Nenhum arquivo informado."
        )

        input(
            "\nPressione ENTER para fechar..."
        )

        return

    caminho_entrada = Path(
        arquivo
    )

    # --------------------------------------------------------
    # Validação
    # --------------------------------------------------------

    if not caminho_entrada.exists():

        print()
        print(
            "[ ERRO ] Arquivo não encontrado:"
        )
        print(
            caminho_entrada
        )

        input(
            "\nPressione ENTER para fechar..."
        )

        return

    if caminho_entrada.suffix.lower() != ".odt":

        print()
        print(
            "[ ERRO ] O arquivo precisa ser .odt."
        )

        input(
            "\nPressione ENTER para fechar..."
        )

        return

    # --------------------------------------------------------
    # Saída
    # --------------------------------------------------------

    caminho_saida = (
        caminho_entrada.with_name(
            f"{caminho_entrada.stem}_formatado"
            f"{caminho_entrada.suffix}"
        )
    )

    # Evita tentar processar o arquivo sobre ele mesmo.
    if caminho_saida.resolve() == caminho_entrada.resolve():

        print()
        print(
            "[ ERRO ] O arquivo de saída é igual ao arquivo de entrada."
        )

        input(
            "\nPressione ENTER para fechar..."
        )

        return

    # --------------------------------------------------------
    # Execução
    # --------------------------------------------------------

    try:

        print(
            "Processando o documento..."
        )

        processar_odt(
            caminho_entrada,
            caminho_saida
        )

        print()
        print(
            "[ SUCESSO ] Processamento concluído."
        )

    except zipfile.BadZipFile:

        print()
        print(
            "[ ERRO ] O arquivo não é um ODT válido."
        )

    except PermissionError:

        print()
        print(
            "[ ERRO ] Não foi possível acessar o arquivo."
        )
        print(
            "Verifique se ele está aberto no LibreOffice."
        )

    except Exception as erro:

        print()
        print(
            "[ ERRO ] Ocorreu uma falha:"
        )
        print()
        print(
            str(erro)
        )

    print()
    print("=" * 65)

    input(
        "Pressione ENTER para fechar..."
    )


if __name__ == "__main__":
    main()
