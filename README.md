# Origin Project Scripts

Scripts for generating Origin/OriginPro `.opju` projects from measurement exports.

## Scripts

- `make_origin_from_csv.py`: builds Origin projects from B1500-style CSV exports.
- `make_origin_from_xls.py`: builds Origin projects from `.xls` workbooks whose curves are stored in `Data` and `Append1`, `Append2`, ... sheets.

Both scripts split output projects automatically after 11 graph pages:

- first 11 graphs: `<folder>.opju`
- next 11 graphs: `<folder>_part2.opju`
- then `<folder>_part3.opju`, and so on

This avoids the graph/window limit in OriginPro Learning Edition.

## Usage

Run from this folder on Windows with Origin/OriginPro installed:

```powershell
python .\make_origin_from_csv.py
python .\make_origin_from_xls.py
```

Useful options:

```powershell
python .\make_origin_from_csv.py --dry-run
python .\make_origin_from_xls.py --dry-run
python .\make_origin_from_xls.py --max-graphs-per-project 0
```

## Notes

- Generated `.opju` files and raw measurement data are intentionally ignored by Git.
- The scripts use Origin COM automation through `Origin.ApplicationSI`.
