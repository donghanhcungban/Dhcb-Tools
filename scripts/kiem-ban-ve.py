#!/usr/bin/env python3
"""Kiểm bản vẽ AutoCAD phục vụ bóc khối lượng — bốn lệnh chỉ ĐỌC, xuất Excel.

Làm thay các Lisp tiện ích hay dùng (RD, DIMSUM, TLEN, TKB) nhưng chạy được ngầm, không phải nạp Lisp lạ:

- ``dim-de``   DIM bị **gõ đè chữ**: chữ hiển thị khác số đo thật (như lệnh RD). Đây là bẫy khi kiểm khối lượng
  bằng DIM — số nhìn thấy không phải chiều dài vẽ.
- ``dim-cong`` **cộng chuỗi DIM** thẳng hàng nối tiếp nhau (như DIMSUM): mỗi chuỗi một dòng, tổng + từng số,
  để so nhanh với tổng đốt trong bảng khối lượng.
- ``pl-dai``   **tổng chiều dài** Line + Polyline theo layer (như TLEN) — ống uPVC, ống đồng vẽ bằng nét.
- ``block``    **đếm block** theo tên (như TKB), tách theo vùng/tầng.

Vùng: ``--vung 1F=x0,y0,x1,y1`` (lặp lại cho nhiều tầng). Không có vùng = cả bản vẽ.

Nguồn dữ liệu (chọn một):
    (mặc định)        bản vẽ đang mở, qua Bridge 8766
    --dwg FILE.dwg    đọc NGẦM bản sao bằng accoreconsole (AutoCAD không cần mở)
    --entities/--inserts/--geometry JSON đã lưu từ query entities / inserts / entity_geometry

Ví dụ:
    python scripts/kiem-ban-ve.py dim-de   --dwg "PODIUM.dwg" -o dim-de.xlsx
    python scripts/kiem-ban-ve.py dim-cong --vung 2F=-190000,25000,-15000,125000 -o dim-cong.xlsx
    python scripts/kiem-ban-ve.py pl-dai   --layer "uPVC|DP" -o pl.xlsx
    python scripts/kiem-ban-ve.py block    --ten "MFD|FCU|Mieng gio" --vung 1F=... --vung 2F=... -o block.xlsx

Mã thoát 0 = đã ghi file · 1 = không có đối tượng nào khớp · 2 = đầu vào hỏng / không đọc được bản vẽ.
"""

from __future__ import annotations

import argparse
import collections
import importlib.util
import json
import math
import os
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
_spec = importlib.util.spec_from_file_location("boc_ong_gio", HERE / "boc-khoi-luong-ong-gio.py")
_bog = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_bog)
write_xlsx, HEAD, TOTAL = _bog.write_xlsx, _bog.HEAD, _bog.TOTAL

PAGE = 10000          # giới hạn một trang của query entities/inserts
MAX_PAGES = 10        # chế độ ngầm hỏi trước đủ số trang trong MỘT phiên accoreconsole
NUM = re.compile(r"[-+]?\d+(?:[.,]\d+)?")


# ----------------------------------------------------------------------------- vùng
def parse_region(text: str) -> tuple:
    name, sep, box = text.partition("=")
    parts = [float(v) for v in box.split(",")] if sep else []
    if not name.strip() or len(parts) != 4:
        raise ValueError(f"--vung cần dạng TÊN=xmin,ymin,xmax,ymax, nhận: {text!r}")
    return (name.strip(), (min(parts[0], parts[2]), min(parts[1], parts[3]), max(parts[0], parts[2]), max(parts[1], parts[3])))


def region_of(x: float, y: float, regions: list):
    """Tên vùng chứa điểm; không có vùng nào thì 'Toàn bản vẽ'; ngoài mọi vùng thì None."""
    if not regions:
        return "Toàn bản vẽ"
    for name, (x0, y0, x1, y1) in regions:
        if x0 <= x <= x1 and y0 <= y <= y1:
            return name
    return None


# ----------------------------------------------------------------------------- đọc dữ liệu
def _result(r: dict) -> dict:
    return r.get("result", r) if isinstance(r, dict) else {}


def paged(run, query: str, key: str, params: dict, eager: bool) -> list:
    """Đọc hết các trang. eager: gửi trước MAX_PAGES trang trong một lượt (chế độ ngầm, mỗi lượt là một phiên)."""
    if eager:
        results = run([(query, {**params, "limit": PAGE, "offset": i * PAGE}) for i in range(MAX_PAGES)])
        out = []
        for r in map(_result, results):
            if key not in r:
                raise RuntimeError(r.get("summary") or r.get("error") or f"query {query} không trả dữ liệu")
            out += r[key]
            if not r.get("hasMore"):
                return out
        raise RuntimeError(f"query {query}: hơn {PAGE * MAX_PAGES} đối tượng — tách bản vẽ hoặc dùng Bridge")
    out, offset = [], 0
    while offset is not None:
        r = _result(run([(query, {**params, "limit": PAGE, "offset": offset})])[0])
        if key not in r:
            raise RuntimeError(r.get("summary") or r.get("error") or f"query {query} không trả dữ liệu")
        out += r[key]
        offset = r.get("nextOffset") if r.get("hasMore") else None
    return out


def geometry(run, handles: list, chunk: int) -> list:
    """Hộp bao (extentsMin/Max) cho polyline — query entities không trả đỉnh."""
    reqs = [("entity_geometry", {"handles": handles[i:i + chunk]}) for i in range(0, len(handles), chunk)]
    out = []
    for r in map(_result, run(reqs) if reqs else []):
        out += r.get("entities", [])
    return out


def live_runner(agent, app: str = "autocad"):
    return lambda reqs: [agent.request(app, "POST", "/query", {"query": q, "params": p}, timeout=300) for q, p in reqs]


def offline_runner(dwg: str, mcp):
    return lambda reqs: mcp._offline_batch(os.path.abspath(dwg), reqs)


def load_json(path: str, key: str) -> list:
    data = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    data = data.get("result", data) if isinstance(data, dict) else data
    return data[key] if isinstance(data, dict) else data


# ----------------------------------------------------------------------------- dim-de
def dim_overrides(entities: list, regions: list, tol: float = 1.0) -> list:
    """DIM có chữ gõ đè (dimText khác rỗng/'<>'). Có '<>' trong chữ = giữ số thật kèm tiền/hậu tố → bỏ qua."""
    rows = []
    for e in entities:
        if "Dimension" not in str(e.get("type", "")) or e.get("measurement") is None:
            continue
        t = (e.get("dimText") or "").strip()
        if not t or "<>" in t:
            continue
        p = e.get("textPosition") or e.get("dimLinePoint") or {}
        reg = region_of(p.get("x", 0), p.get("y", 0), regions)
        if reg is None:
            continue
        meas = float(e["measurement"])
        clean = _bog.clean_mtext(t)
        m = NUM.search(clean)
        if m:
            shown = float(m.group(0).replace(",", "."))
            diff = shown - meas
            verdict = "LỆCH" if abs(diff) > tol else "khớp (gõ lại cùng số)"
        else:
            shown, diff, verdict = None, None, "chữ không phải số"
        rows.append({"vung": reg, "handle": e.get("handle"), "layer": e.get("layer"), "chu": clean, "so_ghi": shown,
                     "so_do": round(meas, 1), "lech": None if diff is None else round(diff, 1), "ket_luan": verdict,
                     "x": round(p.get("x", 0)), "y": round(p.get("y", 0))})
    return sorted(rows, key=lambda r: (r["vung"], r["ket_luan"] != "LỆCH", r["x"], r["y"]))


# ----------------------------------------------------------------------------- dim-cong
def _dim_axis(e: dict):
    a, b, d = e.get("xLine1Point"), e.get("xLine2Point"), e.get("dimLinePoint")
    if not (a and b and d) or e.get("measurement") is None:
        return None
    ang = math.radians(float(e.get("rotationDeg") or 0.0))
    u = (math.cos(ang), math.sin(ang))
    t1, t2 = a["x"] * u[0] + a["y"] * u[1], b["x"] * u[0] + b["y"] * u[1]
    off = -d["x"] * u[1] + d["y"] * u[0]           # vị trí đường kích thước, vuông góc với hướng đo
    return round(math.degrees(ang) % 180, 1), off, min(t1, t2), max(t1, t2)


def dim_chains(entities: list, regions: list, tol: float = 5.0, off_tol: float = 20.0) -> list:
    """Gom DIM cùng hướng, cùng đường kích thước, đầu nọ nối đuôi kia (≤ tol) thành chuỗi."""
    groups = collections.defaultdict(list)
    for e in entities:
        if "Dimension" not in str(e.get("type", "")):
            continue
        ax = _dim_axis(e)
        if ax is None:
            continue
        p = e["dimLinePoint"]
        reg = region_of(p["x"], p["y"], regions)
        if reg is None:
            continue
        ang, off, lo, hi = ax
        groups[(reg, ang)].append((off, lo, hi, float(e["measurement"]), e))
    chains = []
    for (reg, ang), items in groups.items():
        items.sort(key=lambda it: (it[0], it[1]))
        lanes = []
        for it in items:                     # gom theo đường kích thước (độ lệch vuông góc gần nhau)
            if lanes and abs(lanes[-1][-1][0] - it[0]) <= off_tol:
                lanes[-1].append(it)
            else:
                lanes.append([it])
        for lane in lanes:
            lane.sort(key=lambda it: it[1])
            cur = [lane[0]]
            for it in lane[1:]:
                if abs(it[1] - cur[-1][2]) <= tol:
                    cur.append(it)
                else:
                    chains.append((reg, ang, cur))
                    cur = [it]
            chains.append((reg, ang, cur))
    rows = []
    for reg, ang, ch in chains:
        if len(ch) < 2:
            continue
        p = ch[0][4]["dimLinePoint"]
        rows.append({"vung": reg, "huong": ang, "so_dim": len(ch), "tong": round(sum(c[3] for c in ch)),
                     "chieu_dai": round(ch[-1][2] - ch[0][1]), "cac_so": " + ".join(str(round(c[3])) for c in ch),
                     "x": round(p["x"]), "y": round(p["y"]), "handle_dau": ch[0][4].get("handle")})
    return sorted(rows, key=lambda r: (r["vung"], -r["tong"]))


# ----------------------------------------------------------------------------- pl-dai
def _center(e: dict, boxes: dict):
    if e.get("start") and e.get("end"):
        return (e["start"]["x"] + e["end"]["x"]) / 2, (e["start"]["y"] + e["end"]["y"]) / 2
    g = boxes.get(e.get("handle"))
    if g and g.get("extentsMin") and g.get("extentsMax"):
        return (g["extentsMin"]["x"] + g["extentsMax"]["x"]) / 2, (g["extentsMin"]["y"] + g["extentsMax"]["y"]) / 2
    return None


def line_lengths(entities: list, regions: list, layer_re: str = "", boxes: dict | None = None) -> list:
    """Tổng chiều dài Line + Polyline theo (vùng, layer). Polyline cần hộp bao (entity_geometry) khi có vùng."""
    pat = re.compile(layer_re, re.I) if layer_re else None
    acc = collections.defaultdict(lambda: [0, 0.0, 0, 0.0])
    boxes = boxes or {}
    for e in entities:
        kind = e.get("type")
        if kind not in ("Line", "Polyline", "Polyline2d", "Polyline3d") or e.get("length") is None:
            continue
        if pat and not pat.search(str(e.get("layer", ""))):
            continue
        if regions:
            c = _center(e, boxes)
            reg = region_of(c[0], c[1], regions) if c else None
            if reg is None:
                continue
        else:
            reg = "Toàn bản vẽ"
        slot = acc[(reg, e.get("layer", ""))]
        if kind == "Line":
            slot[0] += 1
            slot[1] += float(e["length"])
        else:
            slot[2] += 1
            slot[3] += float(e["length"])
    return [{"vung": r, "layer": lay, "so_line": v[0], "dai_line_m": round(v[1] / 1000, 3), "so_pline": v[2],
             "dai_pline_m": round(v[3] / 1000, 3), "tong_m": round((v[1] + v[3]) / 1000, 3)}
            for (r, lay), v in sorted(acc.items())]


# ----------------------------------------------------------------------------- block
def block_counts(inserts: list, regions: list, name_re: str = "", by_layer: bool = False) -> list:
    pat = re.compile(name_re, re.I) if name_re else None
    acc = collections.Counter()
    for i in inserts:
        name = str(i.get("blockName", ""))
        if pat and not pat.search(name):
            continue
        reg = region_of(float(i.get("x", 0)), float(i.get("y", 0)), regions)
        if reg is None:
            continue
        acc[(reg, name, i.get("layer", "") if by_layer else "")] += 1
    return [{"vung": r, "block": n, "layer": lay, "so_luong": c} for (r, n, lay), c in sorted(acc.items())]


# ----------------------------------------------------------------------------- Excel
COLUMNS = {
    "dim-de": [("Vùng", "vung", 12), ("Kết luận", "ket_luan", 22), ("Chữ hiển thị", "chu", 18), ("Số ghi", "so_ghi", 10),
               ("Số đo thật", "so_do", 11), ("Lệch", "lech", 9), ("Layer", "layer", 20), ("Handle", "handle", 10),
               ("X", "x", 11), ("Y", "y", 11)],
    "dim-cong": [("Vùng", "vung", 12), ("Hướng (°)", "huong", 9), ("Số DIM", "so_dim", 8), ("Tổng DIM (mm)", "tong", 13),
                 ("Từ đầu đến cuối (mm)", "chieu_dai", 18), ("Các số", "cac_so", 60), ("X", "x", 11), ("Y", "y", 11),
                 ("Handle DIM đầu", "handle_dau", 14)],
    "pl-dai": [("Vùng", "vung", 12), ("Layer", "layer", 26), ("Số Line", "so_line", 9), ("Line (m)", "dai_line_m", 11),
               ("Số Polyline", "so_pline", 11), ("Polyline (m)", "dai_pline_m", 12), ("Tổng (m)", "tong_m", 11)],
    "block": [("Vùng", "vung", 12), ("Tên block", "block", 60), ("Layer", "layer", 20), ("Số lượng", "so_luong", 10)],
}
TOTALS = {"dim-cong": ("tong",), "pl-dai": ("dai_line_m", "dai_pline_m", "tong_m"), "block": ("so_luong",)}
TITLES = {"dim-de": "DIM gõ đè chữ", "dim-cong": "Cộng chuỗi DIM", "pl-dai": "Chiều dài Line + Polyline",
          "block": "Thống kê block"}


def build_sheet(cmd: str, rows: list, source: str, note: str) -> list:
    cols = COLUMNS[cmd]
    out = [[f"{TITLES[cmd]} — {source}"], [note], [c[0] for c in cols]]
    out += [[r.get(k) for _, k, _ in cols] for r in rows]
    styles = {3: HEAD}
    if cmd in TOTALS and rows:
        keys = [k for _, k, _ in cols]
        total = ["Cộng"] + [None] * (len(cols) - 1)
        for k in TOTALS[cmd]:
            total[keys.index(k)] = round(sum(r[k] or 0 for r in rows), 3)
        out.append(total)
        styles[len(out)] = TOTAL
    return [(TITLES[cmd][:31], out, [c[2] for c in cols], styles)]


# ----------------------------------------------------------------------------- main
def main(argv=None, agent=None, mcp=None) -> int:
    ap = argparse.ArgumentParser(description="Kiểm bản vẽ AutoCAD cho bóc khối lượng (chỉ đọc) → Excel")
    ap.add_argument("lenh", choices=sorted(COLUMNS), help="dim-de | dim-cong | pl-dai | block")
    ap.add_argument("-o", "--output", required=True, help="file .xlsx")
    ap.add_argument("--vung", action="append", default=[], help="TÊN=xmin,ymin,xmax,ymax (lặp lại cho nhiều vùng)")
    ap.add_argument("--dwg", help="đọc ngầm bản sao DWG bằng accoreconsole")
    ap.add_argument("--entities", help="JSON query entities")
    ap.add_argument("--inserts", help="JSON query inserts")
    ap.add_argument("--geometry", help="JSON entity_geometry (hộp bao polyline cho pl-dai có vùng)")
    ap.add_argument("--layer", default="", help="pl-dai: regex lọc layer")
    ap.add_argument("--ten", default="", help="block: regex lọc tên block")
    ap.add_argument("--theo-layer", action="store_true", help="block: tách thêm theo layer")
    ap.add_argument("--tol", type=float, default=None, help="dim-de: lệch cho phép (mặc định 1); dim-cong: khe nối (mặc định 5)")
    ap.add_argument("--source", default="", help="tên bản vẽ ghi vào bảng")
    args = ap.parse_args(argv)
    cmd = args.lenh
    try:
        regions = [parse_region(v) for v in args.vung]
        if args.entities or args.inserts:
            need = "inserts" if cmd == "block" else "entities"
            path = getattr(args, need)
            if not path:
                raise ValueError(f"lệnh {cmd} cần --{need}")
            data = load_json(path, need)
            boxes = {g["handle"]: g for g in load_json(args.geometry, "entities")} if args.geometry else {}
            source = args.source or Path(path).name
        else:
            if args.dwg:
                if mcp is None:  # pragma: no cover — nhánh chạy thật, test truyền mcp giả
                    sys.path.insert(0, str(HERE))
                    saved, sys.argv = sys.argv, ["dhcb_mcp_server.py", "autocad", "--read-only"]
                    try:                     # module đọc sys.argv ngay lúc import
                        import dhcb_mcp_server as mcp
                    finally:
                        sys.argv = saved
                run, eager, chunk = offline_runner(args.dwg, mcp), True, 500
                source = args.source or Path(args.dwg).name
            else:
                if agent is None:  # pragma: no cover — nhánh chạy thật, test truyền agent giả
                    sys.path.insert(0, str(HERE))
                    import dhcb_agent as agent
                run, eager, chunk = live_runner(agent), False, 25
                info = _result(run([("drawing_info", {})])[0])
                source = args.source or Path(str(info.get("filename") or "bản vẽ đang mở").replace("\\", "/")).name
            if cmd == "block":
                data = paged(run, "inserts", "inserts", {}, eager)
            else:
                data = paged(run, "entities", "entities", {}, eager)
            boxes = {}
            if cmd == "pl-dai" and regions:
                pat = re.compile(args.layer, re.I) if args.layer else None
                handles = [e["handle"] for e in data if str(e.get("type", "")).startswith("Polyline")
                           and (not pat or pat.search(str(e.get("layer", ""))))]
                boxes = {g["handle"]: g for g in geometry(run, handles, chunk)}
    except (OSError, ValueError, RuntimeError, KeyError) as ex:
        print(f"Lỗi đầu vào: {ex}", file=sys.stderr)
        return 2

    if cmd == "dim-de":
        tol = 1.0 if args.tol is None else args.tol
        rows = dim_overrides(data, regions, tol)
        note = (f"Chữ DIM gõ đè (khác '<>'). LỆCH = số ghi khác số đo quá {tol:g} mm — đừng lấy số này để kiểm khối lượng. "
                "Chữ có '<>' (giữ số thật kèm tiền/hậu tố) không tính.")
        n_dim = sum(1 for e in data if "Dimension" in str(e.get("type", "")))
        summary = f"{n_dim} DIM, {len(rows)} DIM gõ đè, {sum(r['ket_luan'] == 'LỆCH' for r in rows)} LỆCH so với số đo"
        if not n_dim:
            print("Không có DIM nào trong phạm vi.", file=sys.stderr)
            return 1
    elif cmd == "dim-cong":
        tol = 5.0 if args.tol is None else args.tol
        rows = dim_chains(data, regions, tol)
        note = f"Chuỗi = DIM cùng hướng, cùng đường kích thước, đầu nối đuôi (khe ≤ {tol:g} mm). Chỉ liệt kê chuỗi ≥ 2 DIM."
        summary = f"{len(rows)} chuỗi DIM"
    elif cmd == "pl-dai":
        rows = line_lengths(data, regions, args.layer, boxes)
        note = "Tổng chiều dài Line + Polyline theo layer" + (f" (lọc layer: {args.layer})" if args.layer else "") + \
               ("; polyline gán vùng theo tâm hộp bao" if regions else "")
        summary = f"{len(rows)} layer, {sum(r['tong_m'] for r in rows):.2f} m"
    else:
        rows = block_counts(data, regions, args.ten, args.theo_layer)
        note = "Đếm block theo tên" + (f" (lọc: {args.ten})" if args.ten else "") + ("; tách theo layer" if args.theo_layer else "")
        summary = f"{len(rows)} loại block, {sum(r['so_luong'] for r in rows)} cái"
    if not rows and cmd != "dim-de":
        print(f"Không có đối tượng nào khớp ({summary}).", file=sys.stderr)
        return 1
    write_xlsx(Path(args.output), build_sheet(cmd, rows, source, note))
    print(f"Đã ghi {args.output}: {summary}.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
