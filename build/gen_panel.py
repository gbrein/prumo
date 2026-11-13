#!/usr/bin/env python3
"""Gera src/modPanel.bas: um Sub sem parâmetros por ação do Prumo_Run.

O "run VB macro" do AppleScript no Mac não passa parâmetros (aceita e não executa),
então o painel lateral chama PrumoP_<id>. Rodar sempre que mudar um Case em modRibbon.bas.
"""
import re
from pathlib import Path

SRC = Path(__file__).resolve().parent.parent / "src"


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


if __name__ == "__main__":
    (SRC / "modPanel.bas").write_bytes(panel_module().encode("ascii"))
    print("src/modPanel.bas gerado")
