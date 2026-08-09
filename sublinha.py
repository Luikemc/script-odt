import re
import sys
import zipfile
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET


# ============================================================
# CONFIGURAÇÃO
# ============================================================

# Aceita:
# [01/03/2025, 11:10:24] Nome:
# [ 01/03/2025 , 11:10:24] Nome:
# [01/03/2025, 11:10:24] Nome: qualquer coisa
#
# O sublinhado será aplicado somente até o ":" do nome.

PADRAO = re.compile(
    r'\[\s*'
    r'\d{2}/\d{2}/\d{4}'
    r'\s*,\s*'
    r'\d{2}:\d{2}:\d{2}'
    r'\s*\]\s*'
    r'[^:\n\r]+'
    r':'
)


# Namespaces usados pelo ODT
NS = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
}

for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


# ============================================================
# FUNÇÕES
# ============================================================

def texto_do_elemento(elemento):
    """
    Retorna todo o texto existente dentro de um elemento.
    """
    partes = []

    for node in elemento.iter():
        if node.tag == f"{{{NS['text']}}}s":
            quantidade = int(
                node.attrib.get(
                    f"{{{NS['text']}}}c",
                    "1"
                )
            )
            partes.append(" " * quantidade)

        elif node.tag == f"{{{NS['text']}}}tab":
            partes.append("\t")

        elif node.tag == f"{{{NS['text']}}}line-break":
            partes.append("\n")

        elif node.text:
            partes.append(node.text)

    return "".join(partes)


def criar_estilo_sublinhado(styles_root):
    """
    Cria um estilo ODT específico para sublinhado.
    """

    nome_estilo = "SublinhadoAutomatico"

    # Verifica se já existe
    for style in styles_root.findall(
        f".//{{{NS['style']}}}style"
    ):
        if style.attrib.get(
            f"{{{NS['style']}}}name"
        ) == nome_estilo:
            return nome_estilo

    estilo = ET.Element(
        f"{{{NS['style']}}}style",
        {
            f"{{{NS['style']}}}name": nome_estilo,
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

    styles_root.append(estilo)

    return nome_estilo


def processar_paragrafo(paragrafo, estilo_sublinhado):
    """
    Procura cabeçalhos de mensagens dentro do parágrafo.
    """

    texto = texto_do_elemento(paragrafo)

    matches = list(PADRAO.finditer(texto))

    if not matches:
        return 0

    nos_texto = []

    for node in paragrafo.iter():
        if node.text:
            nos_texto.append({
                "node": node,
                "inicio": None,
                "fim": None,
                "texto": node.text
            })

    posicao = 0

    for item in nos_texto:
        tamanho = len(item["texto"])
        item["inicio"] = posicao
        item["fim"] = posicao + tamanho
        posicao += tamanho

    total = 0

    for match in reversed(matches):

        inicio = match.start()
        fim = match.end()

        afetados = []

        for item in nos_texto:
            if item["fim"] <= inicio:
                continue

            if item["inicio"] >= fim:
                continue

            afetados.append(item)

        if not afetados:
            continue

        if len(afetados) == 1:

            item = afetados[0]
            node = item["node"]

            inicio_local = inicio - item["inicio"]
            fim_local = fim - item["inicio"]

            texto_original = node.text or ""

            antes = texto_original[:inicio_local]
            selecionado = texto_original[
                inicio_local:fim_local
            ]
            depois = texto_original[fim_local:]

            node.text = antes

            span = ET.Element(
                f"{{{NS['text']}}}span",
                {
                    f"{{{NS['text']}}}style-name":
                        estilo_sublinhado
                }
            )

            span.text = selecionado

            pai = encontrar_pai(paragrafo, node)

            if pai is not None:

                indice = list(pai).index(node)

                pai.insert(indice + 1, span)

                if depois:
                    span.tail = depois

            total += 1

    return total


def encontrar_pai(raiz, filho):
    """
    Encontra o elemento pai de um determinado nó.
    """

    for elemento in raiz.iter():

        for filho_atual in list(elemento):

            if filho_atual is filho:
                return elemento

    return None


def processar_odt(arquivo_entrada, arquivo_saida):
    """
    Abre o ODT, encontra os cabeçalhos WhatsApp
    e aplica sublinhado.
    """

    arquivo_entrada = Path(arquivo_entrada)
    arquivo_saida = Path(arquivo_saida)

    with tempfile.TemporaryDirectory() as temp:

        temp = Path(temp)

        with zipfile.ZipFile(
            arquivo_entrada,
            "r"
        ) as zip_in:

            zip_in.extractall(temp)

        content_xml = temp / "content.xml"
        styles_xml = temp / "styles.xml"

        tree_content = ET.parse(content_xml)
        root_content = tree_content.getroot()

        if styles_xml.exists():

            tree_styles = ET.parse(styles_xml)
            root_styles = tree_styles.getroot()

        else:

            root_styles = ET.Element(
                f"{{{NS['office']}}}document-styles"
            )

            tree_styles = ET.ElementTree(root_styles)

        estilos = root_styles.find(
            f"{{{NS['office']}}}styles"
        )

        if estilos is None:

            estilos = ET.SubElement(
                root_styles,
                f"{{{NS['office']}}}styles"
            )

        estilo_sublinhado = criar_estilo_sublinhado(
            estilos
        )

        total = 0

        elementos = root_content.iter(
            f"{{{NS['text']}}}p"
        )

        for paragrafo in elementos:

            total += processar_paragrafo(
                paragrafo,
                estilo_sublinhado
            )

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

        with zipfile.ZipFile(
            arquivo_saida,
            "w",
            zipfile.ZIP_DEFLATED
        ) as zip_out:

            mimetype = temp / "mimetype"

            if mimetype.exists():

                zip_out.write(
                    mimetype,
                    "mimetype",
                    compress_type=zipfile.ZIP_STORED
                )

            for arquivo in temp.rglob("*"):

                if not arquivo.is_file():
                    continue

                relativo = arquivo.relative_to(temp)

                if relativo.as_posix() == "mimetype":
                    continue

                zip_out.write(
                    arquivo,
                    relativo.as_posix()
                )

    print(f"[ SUCESSO ] Arquivo gerado: {arquivo_saida.name}")
    print(f"            Cabeçalhos formatados: {total}")
    print("-" * 50)


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":

    raw_args = sys.argv[1:]

    # Trata caminhos divididos por espaços no Windows
    if raw_args and not Path(" ".join(raw_args)).exists():
        caminho_reconstruido = " ".join(raw_args)
        if Path(caminho_reconstruido).exists():
            arquivos_recebidos = [caminho_reconstruido]
        else:
            arquivos_recebidos = raw_args
    else:
        arquivos_recebidos = [" ".join(raw_args)] if len(raw_args) > 1 and Path(" ".join(raw_args)).exists() else raw_args

    print("========================================")
    print(" SUBLINHAR CABEÇALHOS WHATSAPP (ODT)")
    print("========================================")

    if not arquivos_recebidos:
        print("\nINSTRUÇÃO:")
        print("Arraste e solte um ou mais arquivos .odt em cima deste programa.\n")
        input("Pressione ENTER para fechar...")
        sys.exit(0)

    processados = 0

    for caminho in arquivos_recebidos:
        caminho_limpo = caminho.strip('"\'')
        entrada = Path(caminho_limpo)

        if not entrada.exists():
            print(f"\n[ ERRO ] Arquivo não encontrado: {entrada}")
            continue

        if entrada.suffix.lower() != ".odt":
            print(f"\n[ ERRO ] Ignorado ({entrada.name}): Não é um arquivo .odt")
            continue

        saida = entrada.with_name(entrada.stem + "_sublinhado.odt")

        try:
            processar_odt(entrada, saida)
            processados += 1
        except Exception as e:
            print(f"\n[ ERRO ] Falha ao processar {entrada.name}: {e}")

    if processados == 0:
        print("\nNenhum arquivo válido foi processado.")
        print("Verifique se o arquivo arrastado é um .odt válido.")

    print("\nProcessamento concluído!")
    input("Pressione ENTER para fechar...")
