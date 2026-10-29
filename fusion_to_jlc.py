#!/usr/bin/env python3
"""Convert Fusion 360 Electronics BOM / Pick-and-Place exports to JLCPCB format.

Usage:
  python fusion_to_jlc.py        (no arguments: convert everything in "Fusion In" into "JLC Out")
  python fusion_to_jlc.py bom BOM.txt [-o out.xlsx]
  python fusion_to_jlc.py pnp PnP_front.txt [PnP_back.txt ...] [-o out.xlsx] [--layer Top|Bottom]
  python fusion_to_jlc.py auto FILE [FILE ...]      (guess type from content)

Output format is chosen by the -o extension (.xlsx or .csv); default is .xlsx
next to the input file.

An optional LCSC/JLCPCB part number is taken from a Fusion BOM attribute column
named LCSC, LCSC_PART, JLCPCB_PART, JLCPCB, etc. when present.
"""
import argparse
import csv
import re
import sys
from pathlib import Path

BOM_HEADER = ["Comment", "Designator", "Footprint", "JLCPCB Part #（optional）"]
CPL_HEADER = ["Designator", "Mid X", "Mid Y", "Layer", "Rotation"]
LCSC_COLUMNS = ["VENDOR_PART_NUMBER", "LCSC", "LCSC_PART", "LCSC_PART_NUMBER", "LCSC_PART_#", "JLCPCB", "JLCPCB_PART", "JLCPCB_PART_NUMBER"]


def read_text(path):
    raw = Path(path).read_bytes()
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return raw.decode("utf-16").splitlines()
    for enc in ("utf-8-sig", "cp1252"):
        try:
            return raw.decode(enc).splitlines()
        except UnicodeError:
            continue
    return raw.decode("latin-1").splitlines()


BOM_HEADER_RE = re.compile(r"\s*(Qty\s+Value\b|Reference Designators\s+Quantity\b)")
MULTI_WORD_COLUMNS = r"Reference Designators|Manufacturer Part Number|Vendor Part Number|Top/Bottom"
COLUMN_ALIASES = {"REFERENCE_DESIGNATORS": "Parts", "QUANTITY": "Qty", "TOP/BOTTOM": "Side"}


def parse_bom(path):
    """Parse Fusion's fixed-width partlist into a list of dicts (supports both Fusion layouts)."""
    lines = read_text(path)
    hdr_idx = next((i for i, l in enumerate(lines) if BOM_HEADER_RE.match(l)), None)
    if hdr_idx is None:
        raise ValueError(f"{path}: could not find a Fusion partlist header line")
    header = lines[hdr_idx]
    cols = []
    for m in re.finditer(rf"(?i){MULTI_WORD_COLUMNS}|\S+", header):
        name = re.sub(r"\s+", "_", m.group()).upper() if " " in m.group() else m.group()
        cols.append((COLUMN_ALIASES.get(name.upper(), name), m.start()))
    rows = []
    for line in lines[hdr_idx + 1:]:
        if not line.strip():
            continue
        row = {}
        for i, (name, start) in enumerate(cols):
            end = cols[i + 1][1] if i + 1 < len(cols) else None
            row[name] = line[start:end].strip()
        rows.append(row)
    return rows


def convert_bom(path):
    out = []
    for r in parse_bom(path):
        parts = [p.strip() for p in r.get("Parts", "").split(",") if p.strip()]
        if not parts:
            continue
        mpn = r.get("MANUFACTURER_PART_NUMBER", "")
        mpn = "" if mpn == "-" else mpn
        comment = mpn or r.get("Value") or r.get("Device", "")
        lcsc = next((r[c] for c in LCSC_COLUMNS if r.get(c)), "")
        out.append([comment, ",".join(parts), r.get("Package", ""), lcsc])
    return out


def guess_layer(path):
    name = Path(path).stem.lower()
    if re.search(r"back|bottom|bot\b|_b$", name):
        return "Bottom"
    return "Top"


def convert_pnp(path, layer=None):
    layer = layer or guess_layer(path)
    out = []
    for line in read_text(path):
        if not line.strip():
            continue
        f = line.split("\t") if "\t" in line else line.split()
        if len(f) < 4:
            continue
        try:
            x, y, rot = float(f[1]), float(f[2]), float(f[3])
        except ValueError:
            continue  # header or junk line
        out.append([f[0].strip(), f"{x:g}mm", f"{y:g}mm", layer, round(rot % 360, 2) if rot % 1 else int(rot % 360)])
    return out


def write_table(path, header, rows):
    path = Path(path)
    if path.suffix.lower() == ".csv":
        with open(path, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(header)
            w.writerows(rows)
    else:
        try:
            import openpyxl
        except ImportError:
            sys.exit("openpyxl is required for .xlsx output (pip install openpyxl) or use -o file.csv")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(header)
        for r in rows:
            ws.append(r)
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = min(60, max(len(str(c.value or "")) for c in col) + 2)
        wb.save(path)
    print(f"Wrote {path} ({len(rows)} rows)")


def looks_like_bom(path):
    return any(BOM_HEADER_RE.match(l) for l in read_text(path)[:10])


def default_out(inp, kind, suffix=".xlsx"):
    p = Path(inp)
    return p.with_name(f"JLC_{kind}_{p.stem}{suffix}")


def run_folders():
    """Convert everything in 'Fusion In' and write results to 'JLC Out'."""
    base = Path(__file__).resolve().parent
    src, dst = base / "Fusion In", base / "JLC Out"
    src.mkdir(exist_ok=True)
    dst.mkdir(exist_ok=True)
    files = sorted(p for p in src.iterdir() if p.suffix.lower() == ".txt")
    if not files:
        print(f"No .txt files found in '{src}'. Drop your Fusion exports there and run again.")
        return 1
    cpl_groups = {}
    for p in files:
        try:
            if looks_like_bom(p):
                write_table(dst / f"JLC_BOM_{p.stem}.xlsx", BOM_HEADER, convert_bom(p))
            else:
                # front/back files of one board are merged into a single CPL
                board = re.sub(r"(?i)[_\- ]?(front|back|top|bottom)$", "", p.stem)
                cpl_groups.setdefault(board, []).append(p)
        except Exception as e:
            print(f"SKIPPED {p.name}: {e}")
    for board, paths in cpl_groups.items():
        rows = [r for p in paths for r in convert_pnp(p)]
        if not rows:
            print(f"SKIPPED {', '.join(p.name for p in paths)}: no pick-and-place rows recognized")
            continue
        write_table(dst / f"JLC_CPL_{board}.xlsx", CPL_HEADER, rows)
    print(f"Done. Results are in '{dst}'.")
    return 0


def main():
    if len(sys.argv) == 1:
        sys.exit(run_folders())
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("kind", choices=["bom", "pnp", "cpl", "auto"])
    ap.add_argument("inputs", nargs="+")
    ap.add_argument("-o", "--output", help="output file (.xlsx or .csv); only valid with one input, or for merged pnp")
    ap.add_argument("--layer", choices=["Top", "Bottom"], help="override layer for pnp (default: guessed from filename)")
    a = ap.parse_args()

    if a.kind in ("pnp", "cpl") and a.output and len(a.inputs) > 1:
        rows = []
        for p in a.inputs:
            rows += convert_pnp(p, a.layer)
        write_table(a.output, CPL_HEADER, rows)
        return
    if a.output and len(a.inputs) > 1:
        ap.error("-o with multiple inputs is only supported for pnp (merged)")

    for p in a.inputs:
        kind = a.kind if a.kind != "auto" else ("bom" if looks_like_bom(p) else "pnp")
        out = a.output or default_out(p, "BOM" if kind == "bom" else "CPL")
        if kind == "bom":
            write_table(out, BOM_HEADER, convert_bom(p))
        else:
            write_table(out, CPL_HEADER, convert_pnp(p, a.layer))


if __name__ == "__main__":
    main()

