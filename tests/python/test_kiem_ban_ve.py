"""Test cho scripts/kiem-ban-ve.py — 4 lệnh kiểm bản vẽ (dim-de, dim-cong, pl-dai, block), không cần AutoCAD."""

from __future__ import annotations

import importlib.util
import io
import json
import tempfile
import unittest
import zipfile
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path

SCRIPT = Path(__file__).resolve().parents[2] / "scripts" / "kiem-ban-ve.py"
_spec = importlib.util.spec_from_file_location("kiem_ban_ve", SCRIPT)
kbv = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(kbv)

R1 = ("1F", (0, 0, 10000, 10000))
R2 = ("2F", (0, 20000, 10000, 30000))


def dim(handle, x1, x2, y=0, off=500, meas=None, text="", rot=0.0, kind="RotatedDimension"):
    """DIM ngang đo từ x1 tới x2 trên đường y+off (rot=90 thì đo theo trục y)."""
    if rot == 90:
        a, b, d = {"x": y, "y": x1}, {"x": y, "y": x2}, {"x": y - off, "y": (x1 + x2) / 2}
    else:
        a, b, d = {"x": x1, "y": y}, {"x": x2, "y": y}, {"x": (x1 + x2) / 2, "y": y + off}
    return {"handle": handle, "type": kind, "layer": "0 DIM", "dimText": text, "xLine1Point": a, "xLine2Point": b,
            "dimLinePoint": d, "textPosition": d, "rotationDeg": rot,
            "measurement": abs(x2 - x1) if meas is None else meas}


def line(handle, x1, y1, x2, y2, layer="0 DP"):
    return {"handle": handle, "type": "Line", "layer": layer, "start": {"x": x1, "y": y1}, "end": {"x": x2, "y": y2},
            "length": abs(x2 - x1) + abs(y2 - y1)}


def pline(handle, length, layer="0 DP"):
    return {"handle": handle, "type": "Polyline", "layer": layer, "length": length}


def ins(name, x, y, layer="0"):
    return {"blockName": name, "x": x, "y": y, "layer": layer}


def sheet_rows(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read("xl/worksheets/sheet1.xml").decode("utf-8")


class Agent:
    """Bridge giả: trả theo tên query, ghi lại từng lần gọi."""

    def __init__(self, answers):
        self.answers, self.calls = answers, []

    def request(self, app, method, path, payload, timeout=0):
        self.calls.append(payload)
        ans = self.answers[payload["query"]]
        return ans(payload["params"]) if callable(ans) else ans


class Mcp:
    def __init__(self, fn):
        self.fn, self.calls = fn, []

    def _offline_batch(self, dwg, reqs):
        self.calls.append((dwg, reqs))
        return [self.fn(q, p) for q, p in reqs]


class RegionTests(unittest.TestCase):
    def test_parse_region(self):
        self.assertEqual(("1F", (0.0, -5.0, 10.0, 20.0)), kbv.parse_region(" 1F =10,20,0,-5"))
        for bad in ("1F", "=1,2,3,4", "1F=1,2,3"):
            with self.subTest(bad=bad), self.assertRaises(ValueError):
                kbv.parse_region(bad)

    def test_region_of(self):
        self.assertEqual("Toàn bản vẽ", kbv.region_of(5, 5, []))
        self.assertEqual("2F", kbv.region_of(5, 25000, [R1, R2]))
        self.assertIsNone(kbv.region_of(50000, 5, [R1, R2]))


class FetchTests(unittest.TestCase):
    def test_result_unwrap(self):
        self.assertEqual({"a": 1}, kbv._result({"result": {"a": 1}}))
        self.assertEqual({"a": 1}, kbv._result({"a": 1}))
        self.assertEqual({}, kbv._result(None))

    def test_paged_live_and_error(self):
        pages = {0: {"entities": [1, 2], "hasMore": True, "nextOffset": 2}, 2: {"result": {"entities": [3], "hasMore": False}}}
        run = lambda reqs: [pages[p["offset"]] for _, p in reqs]  # noqa: E731
        self.assertEqual([1, 2, 3], kbv.paged(run, "entities", "entities", {}, eager=False))
        with self.assertRaisesRegex(RuntimeError, "hỏng"):
            kbv.paged(lambda reqs: [{"summary": "hỏng"}], "entities", "entities", {}, eager=False)

    def test_paged_eager(self):
        def run(reqs):
            self.assertEqual(kbv.MAX_PAGES, len(reqs))
            return [{"inserts": [i], "hasMore": i < 1} for i in range(len(reqs))]
        self.assertEqual([0, 1], kbv.paged(run, "inserts", "inserts", {}, eager=True))
        with self.assertRaisesRegex(RuntimeError, "query inserts không trả"):
            kbv.paged(lambda reqs: [{}] * len(reqs), "inserts", "inserts", {}, eager=True)
        with self.assertRaisesRegex(RuntimeError, "hơn"):
            kbv.paged(lambda reqs: [{"inserts": [], "hasMore": True}] * len(reqs), "inserts", "inserts", {}, eager=True)

    def test_geometry_chunks(self):
        seen = []

        def run(reqs):
            seen.extend(reqs)
            return [{"entities": [{"handle": h} for h in p["handles"]]} for _, p in reqs]
        self.assertEqual([], kbv.geometry(run, [], 2))
        self.assertEqual(3, len(kbv.geometry(run, ["A", "B", "C"], 2)))
        self.assertEqual(2, len(seen))

    def test_runners_and_load_json(self):
        agent = Agent({"entities": {"entities": []}})
        self.assertEqual([{"entities": []}], kbv.live_runner(agent)([("entities", {"limit": 1})]))
        mcp = Mcp(lambda q, p: {"q": q})
        self.assertEqual([{"q": "inserts"}], kbv.offline_runner("a.dwg", mcp)([("inserts", {})]))
        self.assertTrue(mcp.calls[0][0].endswith("a.dwg"))
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "x.json"
            for data, want in (({"result": {"entities": [1]}}, [1]), ({"entities": [2]}, [2]), ([3], [3])):
                p.write_text(json.dumps(data), encoding="utf-8")
                self.assertEqual(want, kbv.load_json(str(p), "entities"))


class DimTests(unittest.TestCase):
    def test_overrides(self):
        ents = [
            dim("A", 0, 1110, text="840"),                    # gõ đè, lệch
            dim("B", 0, 600, text="{\\W0.8;600}"),            # gõ lại đúng số
            dim("C", 0, 600, text="ĐOẠN BÙ"),                 # chữ không phải số
            dim("D", 0, 600, text="<> mm"),                   # giữ số thật
            dim("E", 0, 600, text=""),                        # không gõ đè
            dim("F", 0, 600, text="500", y=50000),            # ngoài vùng
            {"type": "RotatedDimension", "measurement": None, "dimText": "1"},
            {"type": "Line", "dimText": "1", "measurement": 5},
        ]
        rows = kbv.dim_overrides(ents, [R1])
        self.assertEqual("A", rows[0]["handle"])              # LỆCH đứng đầu
        by = {r["handle"]: r for r in rows}
        self.assertEqual({"A", "B", "C"}, set(by))
        self.assertEqual(("LỆCH", -270.0), (by["A"]["ket_luan"], by["A"]["lech"]))
        self.assertEqual("chữ không phải số", by["C"]["ket_luan"])
        self.assertIsNone(by["C"]["lech"])
        self.assertTrue(by["B"]["ket_luan"].startswith("khớp"))

    def test_axis_none(self):
        self.assertIsNone(kbv._dim_axis({"measurement": 5}))
        self.assertIsNone(kbv._dim_axis({**dim("A", 0, 1), "measurement": None}))

    def test_chains(self):
        ents = [
            dim("A1", 0, 1110), dim("A2", 1110, 2220), dim("A3", 2222, 2822),       # một chuỗi (khe 2 mm)
            dim("B1", 5000, 6000),                                                   # cùng đường, đứt → chuỗi lẻ (bỏ)
            dim("C1", 0, 400, off=3000), dim("C2", 400, 900, off=3005),              # đường khác (lệch 2995)
            dim("V1", 0, 1180, y=8000, rot=90), dim("V2", 1180, 2360, y=8000, rot=90),
            dim("X", 0, 1000, y=50000), dim("Y", 1000, 2000, y=50000),              # ngoài vùng
            {"type": "AlignedDimension", "measurement": 1},                           # thiếu điểm
            line("L", 0, 0, 1, 0),
        ]
        rows = kbv.dim_chains(ents, [R1])
        by = {r["handle_dau"]: r for r in rows}
        self.assertEqual({"A1", "C1", "V1"}, set(by))
        self.assertEqual((3, 2820, "1110 + 1110 + 600"), (by["A1"]["so_dim"], by["A1"]["tong"], by["A1"]["cac_so"]))
        self.assertEqual(900, by["C1"]["tong"])
        self.assertEqual((90.0, 2360), (by["V1"]["huong"], by["V1"]["tong"]))


class LengthBlockTests(unittest.TestCase):
    def test_line_lengths(self):
        ents = [line("L1", 0, 0, 1000, 0), line("L2", 0, 0, 0, 500, layer="0 GAS"), pline("P1", 2000),
                pline("P2", 3000), pline("P3", 700), {"type": "Arc", "length": 5}, {"type": "Line", "length": None},
                line("L3", 0, 0, 100, 0, layer="0 KHAC"), line("L4", 0, 50000, 100, 50000)]
        boxes = {"P1": {"extentsMin": {"x": 0, "y": 0}, "extentsMax": {"x": 10, "y": 10}},
                 "P2": {"extentsMin": {"x": 0, "y": 20000}, "extentsMax": {"x": 10, "y": 20010}}}
        rows = kbv.line_lengths(ents, [R1, R2], "DP|GAS", boxes)       # P3 không có hộp bao → bỏ
        got = {(r["vung"], r["layer"]): r for r in rows}
        self.assertEqual((1, 1.0, 1, 2.0, 3.0), tuple(got[("1F", "0 DP")][k] for k in
                                                      ("so_line", "dai_line_m", "so_pline", "dai_pline_m", "tong_m")))
        self.assertEqual(3.0, got[("2F", "0 DP")]["tong_m"])
        self.assertEqual(0.5, got[("1F", "0 GAS")]["tong_m"])
        all_rows = kbv.line_lengths(ents, [])
        self.assertEqual({"Toàn bản vẽ"}, {r["vung"] for r in all_rows})
        self.assertIsNone(kbv._center({"handle": "Z"}, {}))

    def test_block_counts(self):
        data = [ins("MFD 700", 1, 1), ins("MFD 700", 2, 2, "L2"), ins("FCU", 1, 25000), ins("Cot", 1, 1), ins("MFD 700", 1, 99999)]
        self.assertEqual([("1F", "MFD 700", 2), ("2F", "FCU", 1)],
                         [(r["vung"], r["block"], r["so_luong"]) for r in kbv.block_counts(data, [R1, R2], "mfd|fcu")])
        by_layer = kbv.block_counts(data, [], "MFD", by_layer=True)
        self.assertEqual({"0": 2, "L2": 1}, {r["layer"]: r["so_luong"] for r in by_layer})


class SheetTests(unittest.TestCase):
    def test_build_sheet_totals(self):
        (name, rows, widths, styles), = kbv.build_sheet("block", [{"vung": "1F", "block": "A", "layer": "", "so_luong": 2},
                                                                  {"vung": "1F", "block": "B", "layer": "", "so_luong": 3}], "x.dwg", "n")
        self.assertEqual(["Cộng", None, None, 5], rows[-1])
        self.assertEqual(kbv.TOTAL, styles[len(rows)])
        (_, rows, _, styles), = kbv.build_sheet("dim-de", [], "x.dwg", "n")
        self.assertEqual(3, len(rows))
        (_, rows, _, _), = kbv.build_sheet("pl-dai", [], "x.dwg", "n")
        self.assertEqual(3, len(rows))


class MainTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dir = Path(self.tmp.name)

    def run_main(self, argv, **kw):
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            code = kbv.main(argv, **kw)
        return code, out.getvalue(), err.getvalue()

    def jfile(self, name, data):
        p = self.dir / name
        p.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        return str(p)

    def test_files_mode_all_commands(self):
        ents = self.jfile("e.json", {"entities": [dim("A", 0, 1110, text="840"), dim("B", 1110, 2220), line("L", 0, 0, 1000, 0),
                                                  pline("P", 500)]})
        geo = self.jfile("g.json", {"entities": [{"handle": "P", "extentsMin": {"x": 0, "y": 0}, "extentsMax": {"x": 1, "y": 1}}]})
        insf = self.jfile("i.json", {"inserts": [ins("MFD", 1, 1)]})
        out = self.dir / "o.xlsx"
        code, msg, _ = self.run_main(["dim-de", "--entities", ents, "-o", str(out), "--tol", "2"])
        self.assertEqual((0, True), (code, "1 LỆCH" in msg))
        self.assertIn("840", sheet_rows(out))
        code, msg, _ = self.run_main(["dim-cong", "--entities", ents, "-o", str(out), "--tol", "1"])
        self.assertEqual((0, True), (code, "1 chuỗi" in msg))
        code, msg, _ = self.run_main(["pl-dai", "--entities", ents, "--geometry", geo, "--vung", "1F=0,0,9000,9000",
                                      "--layer", "DP", "-o", str(out), "--source", "T.dwg"])
        self.assertEqual((0, True), (code, "1.50 m" in msg))
        self.assertIn("T.dwg", sheet_rows(out))
        code, msg, _ = self.run_main(["block", "--inserts", insf, "--ten", "MFD", "--theo-layer", "-o", str(out)])
        self.assertEqual((0, True), (code, "1 cái" in msg))
        code, msg, _ = self.run_main(["dim-de", "--entities", self.jfile("n.json", {"entities": [dim("C", 0, 5)]}), "-o", str(out)])
        self.assertEqual((0, True), (code, "0 DIM gõ đè" in msg))

    def test_input_errors_and_empty(self):
        out = str(self.dir / "o.xlsx")
        ents = self.jfile("e.json", {"entities": [line("L", 0, 0, 1, 0)]})
        self.assertEqual(2, self.run_main(["block", "--entities", ents, "-o", out])[0])
        self.assertEqual(2, self.run_main(["block", "--vung", "sai", "--inserts", ents, "-o", out])[0])
        self.assertEqual(2, self.run_main(["block", "--inserts", str(self.dir / "khong.json"), "-o", out])[0])
        code, _, err = self.run_main(["dim-de", "--entities", ents, "-o", out])
        self.assertEqual((1, True), (code, "Không có DIM" in err))
        code, _, err = self.run_main(["pl-dai", "--entities", ents, "--layer", "KHONG", "-o", out])
        self.assertEqual((1, True), (code, "Không có đối tượng" in err))

    def test_offline_mode(self):
        data = {"entities": [line("L", 0, 0, 2000, 0), pline("P", 1000), pline("Q", 9, layer="0 KHAC")],
                "inserts": [ins("FCU", 1, 1)]}

        def fn(q, p):
            if q == "entity_geometry":
                return {"entities": [{"handle": h, "extentsMin": {"x": 0, "y": 0}, "extentsMax": {"x": 2, "y": 2}} for h in p["handles"]]}
            return {q: data[q] if p["offset"] == 0 else [], "hasMore": False}
        mcp = Mcp(fn)
        out = str(self.dir / "o.xlsx")
        code, msg, _ = self.run_main(["pl-dai", "--dwg", "x.dwg", "--vung", "1F=0,0,9000,9000", "--layer", "DP", "-o", out], mcp=mcp)
        self.assertEqual((0, True), (code, "3.00 m" in msg))
        geo_reqs = [r for _, reqs in mcp.calls for r in reqs if r[0] == "entity_geometry"]
        self.assertEqual([["P"]], [p["handles"] for _, p in geo_reqs])
        code, msg, _ = self.run_main(["block", "--dwg", "x.dwg", "-o", out], mcp=mcp)
        self.assertEqual(0, code)
        self.assertIn("x.dwg", sheet_rows(Path(out)))

    def test_live_mode(self):
        agent = Agent({"drawing_info": {"filename": "C:\\Du an\\PODIUM.dwg"},
                       "entities": {"entities": [dim("A", 0, 100), dim("B", 100, 300), pline("P", 100)], "hasMore": False},
                       "entity_geometry": {"entities": []}})
        out = str(self.dir / "o.xlsx")
        code, msg, _ = self.run_main(["dim-cong", "-o", out], agent=agent)
        self.assertEqual(0, code)
        self.assertIn("PODIUM.dwg", sheet_rows(Path(out)))
        code, _, err = self.run_main(["pl-dai", "--vung", "1F=0,0,10,10", "-o", out], agent=agent)
        self.assertEqual(1, code)                       # polyline không có hộp bao → không gán được vùng
        self.assertEqual("entity_geometry", agent.calls[-1]["query"])
        agent2 = Agent({"drawing_info": {}, "inserts": {"error": "Bridge tắt"}})
        code, _, err = self.run_main(["block", "-o", out], agent=agent2)
        self.assertEqual((2, True), (code, "Bridge tắt" in err))


if __name__ == "__main__":
    unittest.main()
