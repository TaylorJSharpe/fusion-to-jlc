# fusion-to-jlc

Converts Fusion 360 Electronics text exports into the files JLCPCB's SMT assembly service asks for:

| Fusion export | JLCPCB file | Command |
|---|---|---|
| Partlist / BOM (`<board>.txt`) | BOM (`Comment, Designator, Footprint, JLCPCB Part #`) | `bom` |
| Pick and Place (`PnP_<board>_front.txt`) | CPL (`Designator, Mid X, Mid Y, Layer, Rotation`) | `cpl` (alias `pnp`) |

## Requirements

- Python 3.8+
- `openpyxl` for `.xlsx` output: `pip install openpyxl` (not needed if you write `.csv`)

## Quick start

1. Put your Fusion exports (the BOM `.txt` and the PnP `.txt` file(s)) in the `Fusion In` folder. *Use the `Vendor Part Number` column to add JLCPCB parts for easy import!*
2. Double-click `Convert.bat` (or run `python fusion_to_jlc.py` with no arguments).
3. Pick up the results from the `JLC Out` folder: `JLC_BOM_<name>.xlsx` and `JLC_CPL_<name>.xlsx`.

Details:
- Each `.txt` is identified automatically as a BOM or a pick-and-place file; other file types are ignored.
- Front and back PnP files of the same board (names ending in `front`/`back`, `top`/`bottom`) are merged into one CPL with the correct Layer values.
- Existing files in `JLC Out` with the same name are overwritten. Files in `Fusion In` are never modified or deleted.
- Both folders are created on first run if missing.

## Manual usage (single files)

Run from the folder containing `fusion_to_jlc.py`:

```
python fusion_to_jlc.py bom  "Example Fusion Exports\DAS_POWER_BOARD_V5.txt"
python fusion_to_jlc.py cpl  "Example Fusion Exports\PnP_DAS_POWER_BOARD_V5_front.txt"
```

Outputs are written next to the input file as `JLC_BOM_<input name>.xlsx` and `JLC_CPL_<input name>.xlsx`.

### Options

- `-o FILE` sets the output path. The extension picks the format: `.xlsx` or `.csv`.
- `--layer Top|Bottom` sets the CPL layer. Without it, the layer is guessed from the file name: `back`, `bottom`, `bot` or a trailing `_b` gives `Bottom`; anything else gives `Top`.
- `auto` can replace `bom`/`cpl`. It detects each file's type from its content, and you can pass several files at once:
  `python fusion_to_jlc.py auto bom.txt pnp_front.txt`

### Boards with parts on both sides

Pass both PnP files with a single `-o` to get one merged CPL, with each file's layer guessed from its name:

```
python fusion_to_jlc.py cpl PnP_board_front.txt PnP_board_back.txt -o JLC_CPL_board.xlsx
```

If a file name doesn't contain `back`/`bottom`, convert it separately with `--layer Bottom`.

## What the conversion does

**BOM**
- Reads Fusion's fixed-width partlist by column position (values may contain spaces).
- `Comment` is the Fusion `MANUFACTURER_PART_NUMBER` when set, otherwise `Value` (falls back to `Device` if empty).
- `Designator` is the `Parts` list with spaces removed (`J1, J2` becomes `J1,J2`).
- `Footprint` is the Fusion `Package`.
- Rows with no parts listed are skipped.

**CPL**
- Fusion columns are: designator, X, Y, rotation, value, package.
- X and Y are written with an `mm` suffix (e.g. `59.69mm`).
- Rotation is copied from Fusion, normalized to 0-360.

## Things to do before uploading

- **JLCPCB Part #** is filled from the Fusion `Vendor Part Number` column (also accepted as `VENDOR_PART_NUMBER`). If that is empty, an `LCSC`, `LCSC_PART`, `JLCPCB` or `JLCPCB_PART` attribute is used. Otherwise it stays blank. Your example export has no such column, so add it as a part attribute in Fusion.
- The manufacturer part number goes in the **Comment** column instead (see below).
- **Check rotations.** Some packages need a rotation offset in JLCPCB's assembly preview. Adjust there or in the output file.
- **Check designators.** Fusion names like `6C` or `TO_P3` pass through unchanged. Rename them if JLCPCB flags them.
- Make sure BOM designators and CPL designators match; JLCPCB matches the two files by designator.
