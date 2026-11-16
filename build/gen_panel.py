#!/usr/bin/env python3
"""Gera src/modPanel.bas (um Sub sem parâmetros por ação do Prumo_Run) e win/panel.tsv (os botões do painel do Windows).

O "run VB macro" do AppleScript no Mac não passa parâmetros (aceita e não executa),
então o painel lateral chama PrumoP_<id>. Rodar sempre que mudar um Case em modRibbon.bas
ou um botão em panel/PrumoPainel.swift: o painel do Mac é a fonte dos dois painéis.
"""
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "src"
SWIFT = ROOT / "panel" / "PrumoPainel.swift"

# Ícones do painel do Windows: imageMso do Office, ou "txt:" + caractere da fonte Segoe UI Symbol.
# (os SF Symbols do Mac só podem ser usados em plataformas da Apple)
WIN_ICON = {
    "PageNumbers": "HeaderFooterInsert", "PageNumbersDelete": "txt:#̸", "AgendaUpdate": "txt:☰",
    "PasteOnSlides": "Paste", "SaveSelected": "FileSaveAs", "SlidesAsPictures": "PictureInsertFromFile",
    "CleanAnimations": "txt:✨", "BlackWhite": "txt:◐",
    "AlignLeft": "ObjectsAlignLeft", "AlignCenter": "ObjectsAlignCenterHorizontal", "AlignRight": "ObjectsAlignRight",
    "AlignTop": "ObjectsAlignTop", "AlignMiddle": "ObjectsAlignMiddleVertical", "AlignBottom": "ObjectsAlignBottom",
    "DockLeft": "txt:⇤", "DockRight": "txt:⇥", "DockUp": "txt:↥", "DockDown": "txt:↧",
    "StackH": "txt:▥", "StackV": "txt:▤",
    "DistH": "AlignDistributeHorizontally", "DistV": "AlignDistributeVertically",
    "Swap": "txt:⇄", "DupRight": "Copy", "CornerTL": "txt:◰", "MemSave": "txt:📌", "MemBoth": "txt:📍",
    "AlignInTable": "TableInsert", "Matrix": "txt:▦", "Circle": "txt:◌", "Golden": "txt:φ",
    "SameWidth": "txt:↔", "SameHeight": "txt:↕", "SameSize": "txt:⤢", "WidthPlus": "txt:±",
    "StretchLeft": "txt:⇠", "StretchRight": "txt:⇢", "StretchUp": "txt:⇡", "StretchDown": "txt:⇣",
    "FillLeft": "txt:⇦", "FillRight": "txt:⇨", "FillUp": "txt:⇧", "FillDown": "txt:⇩",
    "MagicResizer": "txt:%", "ProcessChain": "txt:»", "MatchAdjust": "txt:◇",
    "LastFill": "ShapeFillColorPicker", "LastLine": "ShapeOutlineColorPicker", "LastFont": "FontColorPicker",
    "CopyFill": "txt:▧", "CopyLine": "txt:▢", "CopyFont": "txt:A", "ToRGB": "txt:#",
    "LangPT": "", "LangEN": "", "Spelling": "Spelling", "MarginPlus": "IndentIncrease",
    "FitToText": "txt:⇕", "WrapText": "txt:↵", "SplitText": "txt:✂", "MergeText": "txt:⊞",
    "Footnote": "txt:¹", "ReplaceTexts": "ReplaceDialog",
    "SlidePrev": "txt:‹", "GoToSlide": "txt:↷", "SlideNext": "txt:›",
    "HideShapes": "txt:⊘", "UnhideShapes": "txt:👁", "MasterShapes": "txt:▣", "SelectSimilar": "txt:≈",
    "CommentInsert": "ReviewNewComment", "CommentNext": "ReviewNextComment", "CommentDeleteAll": "txt:✕",
    "StickerDraft": "txt:■", "StickerWip": "txt:■", "StickerReview": "txt:■", "StickerDone": "txt:■",
    "StickerNext": "txt:➜", "StickerDeleteAll": "txt:✕",
}
TINTS = {"systemGray": "#8E8E93", "systemYellow": "#E0A800", "systemBlue": "#0A64D8", "systemGreen": "#2E9E44"}


def case_ids() -> list:
    """Todos os ids dos Case de Prumo_Run, inclusive os de linhas com vários ids (Case "A", "B": ...)."""
    ids = []
    for line in (SRC / "modRibbon.bas").read_text(encoding="ascii").splitlines():
        m = re.match(r'\s*Case ((?:"\w+"(?:, )?)+):', line)
        if m:
            ids += re.findall(r'"(\w+)"', m.group(1))
    return ids


def panel_module() -> str:
    ids = case_ids()
    lines = [
        'Attribute VB_Name = "modPanel"',
        "Option Explicit",
        "",
        "' GERADO por build/gen_panel.py: nao editar a mao.",
        "' O painel lateral chama estes Subs sem parametros (run VB macro do Mac nao passa parametros).",
    ]
    lines += [f'Public Sub PrumoP_{i}(): Prumo_Panel "{i}": End Sub' for i in ids]
    return "\r\n".join(lines) + "\r\n"


def panel_tsv() -> str:
    """Seções e botões do painel do Mac, uma linha por registro: S(eção), I(tem), M(enu da setinha)."""
    out = []
    for line in SWIFT.read_text(encoding="utf-8").splitlines():
        if m := re.search(r'Section\(title: "([^"]+)"', line):
            out.append(["S", m.group(1), "0"])
        elif "colorBar: true" in line:
            next(r for r in reversed(out) if r[0] == "S")[2] = "1"
        if m := re.search(r'\bItem\((.*)', line):
            f = dict(re.findall(r'(\w+): "([^"]*)"', m.group(1)))
            tint = re.search(r'tint: \.(\w+)', line)
            out.append(["I", f["id"], f["tip"], f.get("title", ""), f.get("colors", ""), WIN_ICON[f["id"]],
                        TINTS[tint.group(1)] if tint else ""])
        for mid, label in re.findall(r'\bM\("(\w+)", "([^"]*)"\)', line):
            out.append(["M", mid, label])
    return "".join("\t".join(r) + "\n" for r in out)


if __name__ == "__main__":
    (SRC / "modPanel.bas").write_bytes(panel_module().encode("ascii"))
    (ROOT / "win" / "panel.tsv").write_bytes(panel_tsv().encode("utf-8"))
    print("src/modPanel.bas e win/panel.tsv gerados")
