import re
import sys
import zipfile
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET

PADRAO = re.compile(
    r'\[\s*\d{2}/\d{2}/\d{4}\s*,\s*\d{2}:\d{2}:\d{2}\s*\]\s*[^:\n\r]+?:'
)

NS = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
}
for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


def encontrar_pai(raiz, filho):
    for elemento in raiz.iter():
        if filho in list(elemento):
            return elemento
    return None


def criar_estilo_sublinhado(styles_root):
    nome = "SublinhadoAutomatico"
    for style in styles_root.iter(f"{{{NS['style']}}}style"):
        if style.attrib.get(f"{{{NS['style']}}}name") == nome:
            return nome
    estilo = ET.SubElement(
        styles_root,
        f"{{{NS['style']}}}style",
        {
            f"{{{NS['style']}}}name": nome,
            f"{{{NS['style']}}}family": "text",
        },
    )
    props = ET.SubElement(
        estilo, f"{{{NS['style']}}}text-properties"
    )
    props.set(f"{{{NS['fo']}}}text-decoration", "underline")
    return nome


def montar_mapa(paragrafo):
    partes = []
    pos = 0

    def add(node, texto, kind):
        nonlocal pos
        if texto:
            partes.append({
                "node": node, "texto": texto,
                "inicio": pos, "fim": pos + len(texto),
                "kind": kind
            })
            pos += len(texto)

    def visit(node):
        if node.text:
            add(node, node.text, "text")
        for child in list(node):
            if child.tag == f"{{{NS['text']}}}s":
                n = int(child.attrib.get(f"{{{NS['text']}}}c", "1"))
                add(child, " " * n, "special")
            elif child.tag == f"{{{NS['text']}}}tab":
                add(child, "\t", "special")
            elif child.tag == f"{{{NS['text']}}}line-break":
                add(child, "\n", "special")
            else:
                visit(child)
            if child.tail:
                add(child, child.tail, "tail")

    visit(paragrafo)
    return partes


def estilo_sublinhado(node, estilos):
    if node is None:
        return False
    nome = node.attrib.get(f"{{{NS['text']}}}style-name")
    vistos = set()
    while nome and nome not in vistos:
        vistos.add(nome)
        style = estilos.get(nome)
        if style is None:
            break
        props = style.find(f"{{{NS['style']}}}text-properties")
        if props is not None:
            decor = props.attrib.get(f"{{{NS['fo']}}}text-decoration", "")
            if "underline" in decor.lower():
                return True
        nome = style.attrib.get(f"{{{NS['style']}}}parent-style-name")
    return False


def aplicar_span(paragrafo, item, a, b, estilo):
    node = item["node"]
    texto = node.text or ""
    antes, alvo, depois = texto[:a], texto[a:b], texto[b:]
    if not alvo:
        return False
    node.text = antes
    span = ET.Element(
        f"{{{NS['text']}}}span",
        {f"{{{NS['text']}}}style-name": estilo}
    )
    span.text = alvo
    pai = encontrar_pai(paragrafo, node)
    if pai is None:
        return False
    idx = list(pai).index(node)
    pai.insert(idx + 1, span)
    if depois:
        span.tail = depois
    return True


def aplicar_tail(paragrafo, item, a, b, estilo):
    node = item["node"]
    texto = node.tail or ""
    antes, alvo, depois = texto[:a], texto[a:b], texto[b:]
    if not alvo:
        return False
    node.tail = antes
    span = ET.Element(
        f"{{{NS['text']}}}span",
        {f"{{{NS['text']}}}style-name": estilo}
    )
    span.text = alvo
    pai = encontrar_pai(paragrafo, node)
    if pai is None:
        return False
    idx = list(pai).index(node)
    pai.insert(idx + 1, span)
    if depois:
        span.tail = depois
    return True


def processar_paragrafo(paragrafo, estilo, estilos):
    mapa = montar_mapa(paragrafo)
    texto = "".join(x["texto"] for x in mapa)
    matches = list(PADRAO.finditer(texto))
    encontrados = len(matches)
    ja = novos = 0

    for match in reversed(matches):
        ini, fim = match.span()
        afetados = [
            x for x in mapa
            if x["fim"] > ini and x["inicio"] < fim
            and x["kind"] in ("text", "tail")
        ]
        if not afetados:
            continue

        if all(estilo_sublinhado(x["node"], estilos) for x in afetados):
            ja += 1
            continue

        alterou = False
        mapa_atual = montar_mapa(paragrafo)

        for item in reversed(mapa_atual):
            if item["kind"] not in ("text", "tail"):
                continue
            if item["fim"] <= ini or item["inicio"] >= fim:
                continue

            a = max(ini, item["inicio"]) - item["inicio"]
            b = min(fim, item["fim"]) - item["inicio"]
            if b <= a:
                continue

            ok = (
                aplicar_span(paragrafo, item, a, b, estilo)
                if item["kind"] == "text"
                else aplicar_tail(paragrafo, item, a, b, estilo)
            )
            alterou = alterou or ok

        if alterou:
            novos += 1

    return encontrados, ja, novos


def processar_odt(entrada, saida):
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        with zipfile.ZipFile(entrada, "r") as z:
            z.extractall(tmp)

        content = tmp / "content.xml"
        styles_file = tmp / "styles.xml"
        tc = ET.parse(content)
        root = tc.getroot()
        ts = ET.parse(styles_file)
        sr = ts.getroot()

        styles = sr.find(f"{{{NS['office']}}}styles")
        if styles is None:
            styles = ET.SubElement(sr, f"{{{NS['office']}}}styles")

        estilo = criar_estilo_sublinhado(styles)
        estilos = {}
        for s in sr.iter(f"{{{NS['style']}}}style"):
            n = s.attrib.get(f"{{{NS['style']}}}name")
            if n:
                estilos[n] = s

        encontrados = ja = novos = 0
        for p in root.iter(f"{{{NS['text']}}}p"):
            e, j, n = processar_paragrafo(p, estilo, estilos)
            encontrados += e
            ja += j
            novos += n

        tc.write(content, encoding="UTF-8", xml_declaration=True)
        ts.write(styles_file, encoding="UTF-8", xml_declaration=True)

        with zipfile.ZipFile(saida, "w") as z:
            mime = tmp / "mimetype"
            if mime.exists():
                z.write(mime, "mimetype", compress_type=zipfile.ZIP_STORED)
            for f in tmp.rglob("*"):
                if f.is_file() and f.name != "mimetype":
                    z.write(f, f.relative_to(tmp).as_posix(),
                            compress_type=zipfile.ZIP_DEFLATED)

    print("\n" + "=" * 60)
    print("RESULTADO")
    print("=" * 60)
    print(f"Cabeçalhos encontrados : {encontrados}")
    print(f"Já estavam sublinhados : {ja}")
    print(f"Novos sublinhados      : {novos}")
    print(f"Arquivo gerado         : {saida.name}")
    print("=" * 60)


def main():
    print("=" * 60)
    print("PROCESSADOR DE ARQUIVOS ODT - WHATSAPP")
    print("=" * 60)

    args = sys.argv[1:]
    entrada = None
    if args:
        candidato = " ".join(args).strip('" ')
        if Path(candidato).exists():
            entrada = candidato
        else:
            for a in args:
                if Path(a.strip('" ')).exists():
                    entrada = a.strip('" ')
                    break

    if not entrada:
        entrada = input("\nDigite ou cole o caminho do arquivo .odt:\n> ").strip('" ')

    p = Path(entrada)
    if not p.exists():
        print("\n[ERRO] Arquivo não encontrado.")
    elif p.suffix.lower() != ".odt":
        print("\n[ERRO] O arquivo precisa ser .odt.")
    else:
        saida = p.with_name(f"{p.stem}_formatado{p.suffix}")
        try:
            print("\nProcessando...")
            processar_odt(p, saida)
        except Exception as e:
            print(f"\n[ERRO] {e}")

    input("\nPressione ENTER para fechar...")


if __name__ == "__main__":
    main()
