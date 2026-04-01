"""
.xlsx 안의 그림을 시트에서 읽는 순서(행→열, 시트 순)로 저장하고
각 파일에 순차적인 타임스탬프를 붙인다.

  python extract_xlsx_images_ordered.py "book.xlsx" [출력폴더]

Windows 탐색기: '만든 날짜'=생성, '수정한 날짜'=수정.
os.utime() 만 쓰면 생성 시각이 같게 남는 경우가 많아, Windows 에서는 SetFileTime 으로
생성·접근·수정을 함께 맞춘다(기본은 생성=수정, 연속 촬영과 비슷).

OOXML 에 촬영 시각은 없으므로 순서는 앵커 위치로만 정한다.
정렬의 기준은 항상 파일명 접두(0001_) 가 가장 확실하다.
"""
from __future__ import annotations

import argparse
import os
import sys
import time
import zipfile
import xml.etree.ElementTree as ET

if sys.platform == "win32":
    import ctypes
    from ctypes import wintypes

    _EPOCH_AS_FILETIME = 116_444_736_000_000_000

    class _FILETIME(ctypes.Structure):
        _fields_ = [
            ("dwLowDateTime", wintypes.DWORD),
            ("dwHighDateTime", wintypes.DWORD),
        ]

    _kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    def _unix_ts_to_filetime(ts: float) -> _FILETIME:
        v = int(ts * 10_000_000) + _EPOCH_AS_FILETIME
        return _FILETIME(v & 0xFFFFFFFF, v >> 32)

    def _set_windows_file_times(path: str, created: float, accessed: float, modified: float) -> None:
        handle = _kernel32.CreateFileW(
            os.path.abspath(path),
            0x40000000,
            0,
            None,
            3,
            0x80,
            None,
        )
        if handle == wintypes.HANDLE(-1).value:
            raise ctypes.WinError(ctypes.get_last_error())
        try:
            c = _unix_ts_to_filetime(created)
            a = _unix_ts_to_filetime(accessed)
            m = _unix_ts_to_filetime(modified)
            if not _kernel32.SetFileTime(
                handle, ctypes.byref(c), ctypes.byref(a), ctypes.byref(m)
            ):
                raise ctypes.WinError(ctypes.get_last_error())
        finally:
            _kernel32.CloseHandle(handle)


def apply_sequential_timestamps(
    path: str, ts: float, *, modified_delta: float = 0.0
) -> None:
    """ts를 '만든 날짜'로 두고, 수정·접근은 ts+modified_delta."""
    mod = ts + modified_delta
    if sys.platform == "win32":
        try:
            _set_windows_file_times(path, created=ts, accessed=mod, modified=mod)
            return
        except OSError:
            pass
    os.utime(path, (mod, mod))


REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _ns(tag_local: str, ns: str) -> str:
    return f"{{{ns}}}{tag_local}"


def _load_rels(z: zipfile.ZipFile, rel_path: str) -> dict[str, str]:
    raw = z.read(rel_path)
    root = ET.fromstring(raw)
    out: dict[str, str] = {}
    for el in root.findall(f".//{{{REL_NS}}}Relationship"):
        rid = el.get("Id")
        tgt = el.get("Target")
        if rid and tgt:
            out[rid] = tgt.replace("\\", "/")
    return out


def _resolve_xl_path(target: str) -> str:
    """../drawings/drawing1.xml → xl/drawings/drawing1.xml"""
    t = target.strip()
    if t.startswith("/"):
        t = t.lstrip("/")
        if not t.startswith("xl/"):
            t = "xl/" + t
        return t
    if t.startswith("../"):
        t = t[3:]
    return "xl/" + t


def _workbook_sheet_paths(z: zipfile.ZipFile) -> list[tuple[str, str]]:
    """[( xl/worksheets/sheet1.xml, sheet name ), ...] 워크북 순서."""
    wb_root = ET.fromstring(z.read("xl/workbook.xml"))
    wb_rels = _load_rels(z, "xl/_rels/workbook.xml.rels")
    sheets: list[tuple[str, str, str]] = []
    for sh in wb_root.findall(f".//{{{MAIN_NS}}}sheet"):
        name = sh.get("name") or ""
        rid = sh.get(f"{{{R_NS}}}id")
        if not rid or rid not in wb_rels:
            continue
        path = _resolve_xl_path(wb_rels[rid])
        if not path.startswith("xl/worksheets/"):
            continue
        sheet_id = sh.get("sheetId") or "0"
        sheets.append((int(sheet_id) if sheet_id.isdigit() else 0, path, name))
    sheets.sort(key=lambda x: x[0])
    return [(p, n) for _, p, n in sheets]


def _drawing_path_for_sheet(z: zipfile.ZipFile, sheet_xml_path: str) -> str | None:
    rels = f"{os.path.dirname(sheet_xml_path)}/_rels/{os.path.basename(sheet_xml_path)}.rels"
    try:
        m = _load_rels(z, rels)
    except KeyError:
        return None
    for rid, tgt in m.items():
        if "drawing" in tgt and tgt.endswith(".xml"):
            return _resolve_xl_path(tgt)
    return None


def _int0(el, tag: str, ns: str) -> int:
    c = el.find(_ns(tag, ns))
    if c is None or c.text is None:
        return 0
    try:
        return int(c.text)
    except ValueError:
        return 0


def _parse_pic_embeds_from_drawing(
    drawing_xml: bytes,
) -> list[tuple[tuple[int, int, int, int], str]]:
    """
    ([row, col, rowOff, colOff], embed_rId) — 앵커당 최대 하나의 그림 blip.
    읽기 순서 정렬용 튜플: row, col, rowOff, colOff
    """
    root = ET.fromstring(drawing_xml)
    out: list[tuple[tuple[int, int, int, int], str]] = []
    for anchor_tag in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
        for anc in root.findall(f".//{{{XDR_NS}}}{anchor_tag}"):
            row = col = row_off = col_off = 0
            if anchor_tag in ("twoCellAnchor", "oneCellAnchor"):
                frm = anc.find(f"{{{XDR_NS}}}from")
                if frm is not None:
                    col = _int0(frm, "col", XDR_NS)
                    row = _int0(frm, "row", XDR_NS)
                    row_off = _int0(frm, "rowOff", XDR_NS)
                    col_off = _int0(frm, "colOff", XDR_NS)
            else:
                sp = anc.find(f".//{{{XDR_NS}}}sp")
                xfrm = None
                if sp is not None:
                    xfrm = sp.find(f".//{{{A_NS}}}xfrm")
                if xfrm is None:
                    xfrm = anc.find(f".//{{{A_NS}}}xfrm")
                if xfrm is not None:
                    off = xfrm.find(f"{{{A_NS}}}off")
                    if off is not None:
                        try:
                            row = int(off.get("x") or 0)
                            col = int(off.get("y") or 0)
                        except ValueError:
                            pass
            for blip in anc.findall(f".//{{{A_NS}}}blip"):
                embed = blip.get(_ns("embed", R_NS))
                if not embed:
                    embed = blip.get("embed")  # 일부 파일
                if embed:
                    pos = (row, col, row_off, col_off)
                    out.append((pos, embed))
                    break
    return out


def collect_ordered_media_paths(z: zipfile.ZipFile) -> list[str]:
    """
    워크북 시트 순서 → 각 시트에서 행·열·오프셋 순.
    같은 xl/media/ 파일이 여러 셀에 붙어 있으면, 가장 위·왼쪽(및 앞 시트)에 먼저 나온 위치만 사용.
    """
    sheet_paths = _workbook_sheet_paths(z)
    # (sheet_idx, row, col, rowOff, colOff, internal_path)
    entries: list[tuple[int, int, int, int, int, str]] = []

    for si, (sheet_path, _name) in enumerate(sheet_paths):
        dpath = _drawing_path_for_sheet(z, sheet_path)
        if not dpath or dpath not in z.namelist():
            continue
        pos_rids = _parse_pic_embeds_from_drawing(z.read(dpath))
        d_rels_path = f"{os.path.dirname(dpath)}/_rels/{os.path.basename(dpath)}.rels"
        d_rels = _load_rels(z, d_rels_path) if d_rels_path in z.namelist() else {}

        for (row, col, row_off, col_off), rid in pos_rids:
            target = d_rels.get(rid)
            if not target or "media/" not in target:
                continue
            internal = _resolve_xl_path(target)
            entries.append((si, row, col, row_off, col_off, internal))

    entries.sort(key=lambda t: (t[0], t[1], t[2], t[3], t[4]))
    seen: set[str] = set()
    ordered: list[str] = []
    for *_, internal_path in entries:
        if internal_path in seen:
            continue
        seen.add(internal_path)
        ordered.append(internal_path)
    return ordered


def extract_ordered(
    xlsx_path: str,
    out_dir: str,
    step_seconds: float = 1.0,
    base_time: float | None = None,
    modified_delta: float = 0.0,
) -> list[str]:
    os.makedirs(out_dir, exist_ok=True)
    written: list[str] = []
    if base_time is None:
        base_time = time.time() - (3600 * 24)

    with zipfile.ZipFile(xlsx_path, "r") as z:
        ordered = collect_ordered_media_paths(z)
        for i, internal in enumerate(ordered):
            data = z.read(internal)
            base = os.path.basename(internal)
            name, ext = os.path.splitext(base)
            out_name = f"{i + 1:04d}_{base}"
            out_path = os.path.join(out_dir, out_name)
            with open(out_path, "wb") as f:
                f.write(data)
            ts = base_time + i * step_seconds
            apply_sequential_timestamps(out_path, ts, modified_delta=modified_delta)
            written.append(out_path)
    return written


def main() -> None:
    ap = argparse.ArgumentParser(description="xlsx 그림을 시트 순서로 추출 + 순차 타임스탬프")
    ap.add_argument("xlsx", help=".xlsx 경로")
    ap.add_argument(
        "out_dir",
        nargs="?",
        default="",
        help="출력 폴더 (기본: 파일명_ordered_images)",
    )
    ap.add_argument(
        "--step",
        type=float,
        default=1.0,
        help="파일 간 시각 간격(초), 기본 1",
    )
    ap.add_argument(
        "--modified-after",
        type=float,
        default=0.0,
        metavar="SEC",
        help="수정 시각만 생성보다 SEC초 늦게(0이면 생성=수정, 미보정 사진과 유사)",
    )
    args = ap.parse_args()

    path = os.path.abspath(args.xlsx)
    if not os.path.isfile(path):
        print("파일 없음:", path)
        sys.exit(1)
    out = args.out_dir.strip()
    if not out:
        out = os.path.join(
            os.path.dirname(path),
            os.path.splitext(os.path.basename(path))[0] + "_ordered_images",
        )
    out = os.path.abspath(out)

    n = len(
        extract_ordered(
            path,
            out,
            step_seconds=args.step,
            modified_delta=args.modified_after,
        )
    )
    print(f"저장: {n}개 → {out}")
    if sys.platform == "win32":
        print("Windows: 만든 날짜·수정한 날짜 모두 설정(기본은 같은 시각, --modified-after 로 수정만 늦출 수 있음).")
    else:
        print("비-Windows: 수정 시각 위주(os.utime); 탐색기 '만든 날짜'는 OS 에 따라 다를 수 있음.")


if __name__ == "__main__":
    main()
