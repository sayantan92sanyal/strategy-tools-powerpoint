#!/usr/bin/env python3
"""Build dist/StrategyTools.ppam from src/.

Usage: python3 scripts/build.py [--pptm]   (--pptm also writes a test .pptm)
"""
import os
import shutil
import subprocess
import re
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from ovba import build_vba_project  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "src"
DIST = ROOT / "dist"

ADDIN_CT = "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml"
PPTM_CT = "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
PPTX_CT = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
VBA_REL = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
UI14_REL = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"


def load_modules(folder=SRC):
    mods = []
    for path in sorted(folder.glob("*.bas")):
        text = path.read_text(encoding="utf-8")
        m = re.match(r'Attribute VB_Name = "([^"]+)"', text)
        if not m:
            raise SystemExit(f"{path.name}: first line must be Attribute VB_Name = \"...\"")
        mods.append((m.group(1), text))
    only = os.environ.get("ST_ONLY")
    if only:
        mods = [m for m in mods if m[0] in only.split(",")]
    return mods


def add_rel(xml: str, rid: str, rtype: str, target: str) -> str:
    rel = f'<Relationship Id="{rid}" Type="{rtype}" Target="{target}"/>'
    return xml.replace("</Relationships>", rel + "</Relationships>")


def package(out_path: Path, main_ct: str, vba_bin: bytes, with_ui: bool):
    with zipfile.ZipFile(SRC / "base.pptx") as base, \
            zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as out:
        for item in base.infolist():
            data = base.read(item.filename)
            name = item.filename
            if name == "[Content_Types].xml":
                xml = data.decode("utf-8").replace(PPTX_CT, main_ct)
                xml = xml.replace(
                    "<Default Extension=\"xml\"",
                    '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>'
                    "<Default Extension=\"xml\"", 1)
                data = xml.encode("utf-8")
            elif name == "_rels/.rels" and with_ui:
                data = add_rel(data.decode("utf-8"), "rIdUI14", UI14_REL,
                               "customUI/customUI14.xml").encode("utf-8")
            elif name == "ppt/_rels/presentation.xml.rels":
                data = add_rel(data.decode("utf-8"), "rIdVBA", VBA_REL,
                               "vbaProject.bin").encode("utf-8")
            elif name == "docProps/thumbnail.jpeg":
                continue
            elif name == "_rels/.rels":
                pass
            if name == "_rels/.rels":
                data = re.sub(rb'<Relationship [^>]*thumbnail[^>]*/>', b"", data)
            out.writestr(item.filename, data)
        out.writestr("ppt/vbaProject.bin", vba_bin)
        if with_ui:
            out.writestr("customUI/customUI14.xml", (SRC / "customUI14.xml").read_bytes())


def main():
    DIST.mkdir(exist_ok=True)
    mods = load_modules()
    vba_bin = build_vba_project(mods)
    (DIST / "vbaProject.bin").write_bytes(vba_bin)
    package(DIST / "StrategyTools.ppam", ADDIN_CT, vba_bin, with_ui=(SRC / "customUI14.xml").exists())
    if "--pptm" in sys.argv:
        test_bin = build_vba_project(mods + load_modules(ROOT / "tests"))
        package(DIST / "StrategyTools-test.pptm", PPTM_CT, test_bin, with_ui=False)
    bundle(DIST / "StrategyTools.ppam")
    print(f"Built {DIST / 'StrategyTools.ppam'} with modules: {', '.join(m for m, _ in mods)}")
    print(f"Shareable package: {DIST / 'Strategy Tools.zip'}")


def bundle(ppam: Path):
    """dist/Strategy Tools/ (installer folder) and dist/Strategy Tools.zip."""
    folder = DIST / "Strategy Tools"
    shutil.rmtree(folder, ignore_errors=True)
    folder.mkdir()
    shutil.copy(ppam, folder)
    shutil.copy(SRC / "shortcuts.html", folder)
    for f in (SRC / "installer").iterdir():
        shutil.copy(f, folder)
        if f.suffix == ".command":
            (folder / f.name).chmod(0o755)
    z = DIST / "Strategy Tools.zip"
    z.unlink(missing_ok=True)
    subprocess.run(["ditto", "-c", "-k", "--keepParent", str(folder), str(z)], check=True)


if __name__ == "__main__":
    main()
