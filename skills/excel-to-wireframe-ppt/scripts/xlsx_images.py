# -*- coding: utf-8 -*-
"""Excel에 삽입된 와이어프레임 이미지를 추출한다.

가장 깨지기 쉬운 부분이라 2단 방어로 간다. openpyxl이 놓치는 이미지가
xl/media에 남아 있으면 zip 폴백이 주워 담는다.
"""
from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from common import Warnings

RASTER = {"png", "jpg", "jpeg", "gif", "bmp"}
XDR_NS = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"


def _image_bytes(img) -> bytes:
    ref = img.ref
    if isinstance(ref, (str, Path)):
        return Path(ref).read_bytes()
    if isinstance(ref, (bytes, bytearray)):
        return bytes(ref)
    if hasattr(ref, "read"):
        pos = ref.tell() if hasattr(ref, "tell") else None
        try:
            if hasattr(ref, "seek"):
                ref.seek(0)
            return ref.read()
        finally:
            if pos is not None and hasattr(ref, "seek"):
                ref.seek(pos)
    buf = io.BytesIO()
    img.image.save(buf, format=(img.format or "PNG").upper())
    return buf.getvalue()


def collect_openpyxl_images(wb, warns: Warnings | None = None) -> list[dict]:
    out = []
    for ws in wb.worksheets:
        for img in getattr(ws, "_images", []):
            anchor = getattr(img, "anchor", None)
            frm = getattr(anchor, "_from", None)
            row = int(getattr(frm, "row", -1)) if frm is not None else -1
            col = int(getattr(frm, "col", -1)) if frm is not None else -1
            try:
                data = _image_bytes(img)
            except Exception as exc:
                if warns is not None:
                    warns.add(
                        None,
                        "image-convert-failed",
                        "%s 시트의 이미지를 읽지 못했습니다 (%s)" % (ws.title, exc),
                    )
                continue
            ext = (getattr(img, "format", None) or "png").lower()
            out.append(
                {"sheet": ws.title, "row": row, "col": col, "data": data, "ext": ext}
            )
    return out


def _drawing_anchor_map(z: zipfile.ZipFile) -> dict[str, tuple[str, int, int]]:
    """xl/drawings/*.xml에서 (미디어 파일명 → 시트 미상, row, col)을 만든다.

    시트 귀속은 drawing → sheet 관계를 거꾸로 타야 해서 비용이 크다. 폴백 경로에서는
    순서 기반 배분으로 충분하므로 앵커만 뽑는다.
    """
    result: dict[str, tuple[str, int, int]] = {}
    for name in z.namelist():
        if not re.match(r"xl/drawings/drawing\d+\.xml$", name):
            continue
        rels_name = "xl/drawings/_rels/%s.rels" % Path(name).name
        rid_to_media: dict[str, str] = {}
        if rels_name in z.namelist():
            rels = ET.fromstring(z.read(rels_name))
            for rel in rels:
                target = rel.get("Target", "")
                rid_to_media[rel.get("Id", "")] = Path(target).name
        root = ET.fromstring(z.read(name))
        for anchor in root:
            frm = anchor.find("%sfrom" % XDR_NS)
            row = col = -1
            if frm is not None:
                row_el = frm.find("%srow" % XDR_NS)
                col_el = frm.find("%scol" % XDR_NS)
                row = int(row_el.text) if row_el is not None else -1
                col = int(col_el.text) if col_el is not None else -1
            for blip in anchor.iter():
                embed = blip.get(
                    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
                )
                if embed and embed in rid_to_media:
                    result[rid_to_media[embed]] = ("", row, col)
    return result


def _natural_key(name: str) -> list:
    """`image10` > `image2`로 취급하는 lexicographic 정렬을 피하려고 숫자를 int로 쪼갠다."""
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", name)]


def collect_zip_images(xlsx_path: Path) -> list[dict]:
    out = []
    with zipfile.ZipFile(xlsx_path) as z:
        anchors = _drawing_anchor_map(z)
        media = sorted(
            (n for n in z.namelist() if n.startswith("xl/media/")),
            key=lambda n: _natural_key(Path(n).name),
        )
        for name in media:
            base = Path(name).name
            sheet, row, col = anchors.get(base, ("", -1, -1))
            out.append(
                {
                    "sheet": sheet,
                    "row": row,
                    "col": col,
                    "data": z.read(name),
                    "ext": Path(name).suffix.lstrip(".").lower(),
                }
            )
    return out


def _to_png(data: bytes, ext: str, warns: Warnings, screen_id: str) -> tuple[bytes, str]:
    """EMF/WMF는 PPT에서 안 보이는 환경이 있어 PNG 변환을 시도한다."""
    if ext in RASTER:
        return data, "png" if ext == "png" else ext
    try:
        from PIL import Image

        with Image.open(io.BytesIO(data)) as im:
            buf = io.BytesIO()
            im.convert("RGB").save(buf, format="PNG")
            return buf.getvalue(), "png"
    except Exception as exc:
        warns.add(screen_id, "image-convert-failed",
                  "%s 이미지를 PNG로 변환하지 못해 원본을 사용합니다 (%s)" % (ext, exc))
        return data, ext


def _assign(mapping: dict, screens: list[dict], found: list[dict]) -> dict[str, list[dict]]:
    layout = mapping["excel"].get("layout", "sheet-per-screen")
    by_screen: dict[str, list[dict]] = {s["id"]: [] for s in screens}
    if layout == "sheet-per-screen":
        sheet_to_id = {s.get("sheet"): s["id"] for s in screens}
        leftovers = []
        for f in found:
            sid = sheet_to_id.get(f["sheet"])
            if sid:
                by_screen[sid].append(f)
            else:
                leftovers.append(f)
        # 시트명을 못 구한 폴백 결과는 앵커(row, col) 순으로 이미지 없는 화면에 배분한다.
        # 앵커가 전부 -1(미상)이거나 서로 같으면 sorted()의 안정성 덕에
        # collect_zip_images가 만들어 둔 natural 파일명 순서가 그대로 유지된다.
        leftovers.sort(key=lambda f: (f["row"], f["col"]))
        empty = [s["id"] for s in screens if not by_screen[s["id"]]]
        for sid, f in zip(empty, leftovers):
            by_screen[sid].append(f)
    else:
        ordered = sorted(found, key=lambda f: (f["row"], f["col"]))
        for i, f in enumerate(ordered):
            if i < len(screens):
                by_screen[screens[i]["id"]].append(f)
    return by_screen


def extract_images(
    xlsx_path: Path,
    wb,
    mapping: dict,
    screens: list[dict],
    out_dir: Path,
    warns: Warnings,
) -> None:
    found = collect_openpyxl_images(wb, warns)
    with zipfile.ZipFile(xlsx_path) as z:
        media_count = sum(1 for n in z.namelist() if n.startswith("xl/media/"))
    if media_count > len(found):
        found = collect_zip_images(Path(xlsx_path))

    by_screen = _assign(mapping, screens, found)
    img_dir = Path(out_dir) / "images"

    for scr in screens:
        items = by_screen.get(scr["id"], [])
        if not items:
            warns.add(scr["id"], "no-image", "이 화면에 연결된 이미지를 찾지 못했습니다")
            continue
        for i, f in enumerate(items):
            data, ext = _to_png(f["data"], f["ext"], warns, scr["id"])
            suffix = "" if len(items) == 1 else "-%d" % (i + 1)
            fname = "%s%s.%s" % (scr["id"], suffix, ext)
            img_dir.mkdir(parents=True, exist_ok=True)
            (img_dir / fname).write_bytes(data)
            scr["images"].append("images/%s" % fname)
