#!/usr/bin/env python3
"""Injeta src/customUI14.xml (o ribbon do Prumo) num .ppam, no próprio arquivo.

Uso: python3 build/build.py dist/Prumo.ppam
     python3 build/build.py dist/Prumo.pptm   (gera dist/Prumo.ppam ao lado)
O SaveAs do Mac não tem o formato .ppam, então o VBA salva .pptm e aqui ele vira
suplemento: muda só o tipo de conteúdo da parte principal.
Pode rodar de novo a cada mudança no ribbon; a injeção anterior é substituída.
"""
import re
import sys
from typing import Optional
import zipfile
from pathlib import Path
from xml.dom import minidom

REL_TYPE = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
PART = "customUI/customUI14.xml"
CT_PPTM = "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
CT_PPAM = "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml"
UI_SRC = Path(__file__).resolve().parent.parent / "src" / "customUI14.xml"
ICONS = UI_SRC.parent / "icons"  # ícones próprios: image="<nome>" no XML = icons/<nome>.png
IMG_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"


def inject(target: Path, ui: bytes, out: Optional[Path] = None, icons: Path = ICONS) -> None:
    out = out or target
    minidom.parseString(ui)  # falha cedo se o XML do ribbon estiver quebrado
    with zipfile.ZipFile(target) as z:
        items = {n: z.read(n) for n in z.namelist() if not n.startswith("customUI/")}  # refaz a parte do ribbon

    names = sorted(set(re.findall(r'\bimage="([^"]+)"', ui.decode("utf-8"))))
    for n in names:
        items[f"customUI/images/{n}.png"] = (icons / f"{n}.png").read_bytes()  # PNG faltando = erro aqui
    if names:
        rel = "".join(f'<Relationship Id="{n}" Type="{IMG_REL}" Target="images/{n}.png"/>' for n in names)
        items["customUI/_rels/customUI14.xml.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rel}</Relationships>'
        ).encode("utf-8")

    # o .ppam é distribuído: sai o rótulo de confidencialidade da empresa (custom.xml) e o nome de quem montou
    items.pop("docProps/custom.xml", None)
    if "docProps/core.xml" in items:
        core = items["docProps/core.xml"].decode("utf-8")
        core = re.sub(r"<(dc:creator|cp:lastModifiedBy)>[^<]*</\1>", r"<\1></\1>", core)
        items["docProps/core.xml"] = core.encode("utf-8")

    rels = items["_rels/.rels"].decode("utf-8")
    rels = re.sub(r'<Relationship\b[^>]*Target="docProps/custom.xml"[^>]*/>', "", rels)
    rels = re.sub(r"<Relationship\b[^>]*Type=\"%s\"[^>]*/>" % re.escape(REL_TYPE), "", rels)
    rels = rels.replace(
        "</Relationships>",
        f'<Relationship Id="prumoUI" Type="{REL_TYPE}" Target="{PART}"/></Relationships>',
    )
    items["_rels/.rels"] = rels.encode("utf-8")

    ct = items["[Content_Types].xml"].decode("utf-8")
    if 'Extension="xml"' not in ct:
        ct = ct.replace("</Types>", '<Default Extension="xml" ContentType="application/xml"/></Types>')
    if names and 'Extension="png"' not in ct:
        ct = ct.replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>')
    ct = re.sub(r'<Override\b[^>]*PartName="/docProps/custom.xml"[^>]*/>', "", ct)
    if out.suffix == ".ppam":
        ct = ct.replace(CT_PPTM, CT_PPAM)
    items["[Content_Types].xml"] = ct.encode("utf-8")

    items[PART] = ui
    tmp = out.with_name(out.name + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in items.items():
            z.writestr(name, data)
    tmp.replace(out)  # troca atômica: o arquivo original nunca fica pela metade


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    path = Path(sys.argv[1])
    if not path.is_file():
        sys.exit(f"arquivo não encontrado: {path}")
    out = path.with_suffix(".ppam")
    inject(path, UI_SRC.read_bytes(), out)
    print(f"ribbon injetado em {out}")
