"""Testes do que dá para checar sem o PowerPoint: injeção do ribbon e ids do ribbon x VBA.
Rodar: python3 tests/test_build.py
"""
import re
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "build"))
from build import CT_PPAM, CT_PPTM, PART, REL_TYPE, UI_SRC, inject  # noqa: E402


def test_inject_idempotente():
    with tempfile.TemporaryDirectory() as d:
        f = Path(d) / "x.ppam"
        with zipfile.ZipFile(f, "w") as z:
            z.writestr("[Content_Types].xml", '<Types xmlns="t"><Override PartName="/docProps/custom.xml" ContentType="c"/></Types>')
            z.writestr("_rels/.rels", '<Relationships xmlns="r"><Relationship Id="rId1" Type="other" Target="ppt/presentation.xml"/><Relationship Id="rId5" Type="c" Target="docProps/custom.xml"/></Relationships>')
            z.writestr("ppt/vbaProject.bin", b"\x00vba")
            z.writestr("docProps/custom.xml", "<Properties>MSIP_Label_x</Properties>")
            z.writestr("docProps/core.xml", "<cp:coreProperties><dc:creator>Fulano</dc:creator><cp:lastModifiedBy>Fulano</cp:lastModifiedBy></cp:coreProperties>")
        ui = UI_SRC.read_bytes()
        inject(f, ui)
        inject(f, ui)
        with zipfile.ZipFile(f) as z:
            rels = z.read("_rels/.rels").decode()
            assert rels.count(REL_TYPE) == 1, rels
            assert 'Target="ppt/presentation.xml"' in rels
            assert z.read(PART) == ui
            assert z.read("ppt/vbaProject.bin") == b"\x00vba"
            assert 'Extension="xml"' in z.read("[Content_Types].xml").decode()
            assert z.namelist().count(PART) == 1
            imgs = set(re.findall(r'\bimage="([^"]+)"', ui.decode()))
            assert imgs, "ribbon sem ícones próprios"
            uirels = z.read("customUI/_rels/customUI14.xml.rels").decode()
            for n in imgs:
                assert z.read(f"customUI/images/{n}.png")[:4] == b"\x89PNG", n
                assert f'Id="{n}"' in uirels, n
            assert 'Extension="png"' in z.read("[Content_Types].xml").decode()
            assert "docProps/custom.xml" not in z.namelist() and "custom.xml" not in rels
            assert "custom.xml" not in z.read("[Content_Types].xml").decode()
            assert "Fulano" not in z.read("docProps/core.xml").decode()


def test_pptm_vira_ppam():
    with tempfile.TemporaryDirectory() as d:
        src, out = Path(d) / "x.pptm", Path(d) / "x.ppam"
        with zipfile.ZipFile(src, "w") as z:
            z.writestr("[Content_Types].xml", f'<Types xmlns="t"><Override PartName="/ppt/presentation.xml" ContentType="{CT_PPTM}"/></Types>')
            z.writestr("_rels/.rels", '<Relationships xmlns="r"></Relationships>')
        inject(src, UI_SRC.read_bytes(), out)
        with zipfile.ZipFile(out) as z:
            ct = z.read("[Content_Types].xml").decode()
            assert CT_PPAM in ct and CT_PPTM not in ct, ct
        with zipfile.ZipFile(src) as z:
            assert CT_PPTM in z.read("[Content_Types].xml").decode()  # origem intacta


def test_todo_botao_tem_acao():
    ui = UI_SRC.read_text(encoding="utf-8")
    ids = set(re.findall(r'<button id="(\w+)"[^>]*onAction="Prumo_Action"', ui))
    from gen_panel import case_ids
    cases = case_ids()
    assert ids, "nenhum botão encontrado"
    assert len(cases) == len(set(cases)), "Case repetido no VBA"
    assert ids <= set(cases), f"sem Case no VBA: {ids - set(cases)}"


def test_modpanel_em_dia():
    from gen_panel import panel_module
    assert (ROOT / "src" / "modPanel.bas").read_bytes().decode("ascii") == panel_module(), "rode build/gen_panel.py"


def test_painel_so_chama_acoes_que_existem():
    swift = (ROOT / "panel" / "PrumoPainel.swift").read_text(encoding="utf-8")
    ids = re.findall(r'Item\(id: "(\w+)"', swift) + re.findall(r'M\("(\w+)"', swift)
    from gen_panel import case_ids
    cases = set(case_ids())
    assert ids, "nenhum botão no painel"
    assert len(ids) == len(set(ids)), "id repetido no painel"
    assert set(ids) <= cases, f"sem Case no VBA: {set(ids) - cases}"


def test_bas_sao_ascii_e_crlf():
    # o editor VBA do Mac importa .bas em codificação legada: acentos quebram
    for bas in (ROOT / "src").glob("*.bas"):
        raw = bas.read_bytes()
        raw.decode("ascii")
        assert b"\n" not in raw.replace(b"\r\n", b""), f"{bas.name} tem LF solto"


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("ok", name)
