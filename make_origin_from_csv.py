#!/usr/bin/env python
"""Build Origin projects from B1500 CSV exports in this folder tree."""

from __future__ import annotations

import argparse
import csv
import math
import pathlib
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Iterable, Sequence


ROOT = pathlib.Path(__file__).resolve().parent
DEFAULT_TEMPLATE = ROOT / "5.22.opju"
DEFAULT_MAX_GRAPHS_PER_PROJECT = 11


@dataclass
class Curve:
    name: str
    vs: list[float]
    abs_is: list[float]
    record_time: datetime | None = None
    source_order: int = 0


@dataclass
class CsvData:
    path: pathlib.Path
    curves: list[Curve]


def parse_float(value: str) -> float | None:
    text = (value or "").strip()
    if not text:
        return None
    try:
        number = float(text)
    except ValueError:
        return None
    if math.isnan(number):
        return None
    return number


def clean_cells(row: Sequence[str]) -> list[str]:
    return [cell.strip() for cell in row]


def read_rows(path: pathlib.Path) -> list[list[str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return [clean_cells(row) for row in csv.reader(handle)]


def parse_record_time(value: str) -> datetime | None:
    text = (value or "").strip()
    if not text:
        return None
    for fmt in ("%m/%d/%Y %H:%M:%S", "%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return None


def finish_tagged_group(
    path: pathlib.Path,
    group_index: int,
    headers: list[str],
    values: list[list[str]],
    record_time: datetime | None,
) -> Curve | None:
    try:
        vs_idx = headers.index("Vs")
        abs_is_idx = headers.index("absIs")
    except ValueError as exc:
        raise ValueError(f"{path} group {group_index} is missing Vs or absIs") from exc

    vs: list[float] = []
    abs_is: list[float] = []
    for row in values:
        if len(row) <= max(vs_idx, abs_is_idx):
            continue
        x = parse_float(row[vs_idx])
        y = parse_float(row[abs_is_idx])
        if x is None or y is None:
            continue
        vs.append(x)
        abs_is.append(y)

    if not vs:
        return None
    return Curve(name=f"group{group_index}", vs=vs, abs_is=abs_is, record_time=record_time, source_order=group_index)


def parse_tagged(rows: list[list[str]], path: pathlib.Path) -> list[Curve]:
    curves: list[Curve] = []
    headers: list[str] | None = None
    values: list[list[str]] = []
    group_index = 0
    pending_record_time: datetime | None = None
    active_record_time: datetime | None = None

    for row in rows:
        if not row:
            continue
        tag = row[0]
        if len(row) > 2 and row[1] == "TestRecord.RecordTime":
            pending_record_time = parse_record_time(row[2])
        if tag == "DataName":
            if headers is not None:
                curve = finish_tagged_group(path, group_index, headers, values, active_record_time)
                if curve:
                    curves.append(curve)
            group_index += 1
            active_record_time = pending_record_time
            headers = row[1:]
            values = []
        elif tag == "DataValue" and headers is not None:
            values.append(row[1:])

    if headers is not None:
        curve = finish_tagged_group(path, group_index, headers, values, active_record_time)
        if curve:
            curves.append(curve)
    curves.sort(key=lambda curve: (curve.record_time is None, curve.record_time or datetime.max, curve.source_order))
    for index, curve in enumerate(curves, start=1):
        curve.name = f"group{index}"
    return curves


def parse_plain(rows: list[list[str]], path: pathlib.Path) -> list[Curve]:
    for index, row in enumerate(rows):
        if "Vs" not in row or "absIs" not in row:
            continue
        vs_idx = row.index("Vs")
        abs_is_idx = row.index("absIs")
        pairs: list[tuple[float, float]] = []
        for data_row in rows[index + 1 :]:
            if len(data_row) <= max(vs_idx, abs_is_idx):
                continue
            x = parse_float(data_row[vs_idx])
            y = parse_float(data_row[abs_is_idx])
            if x is None or y is None:
                continue
            pairs.append((x, y))
        if not pairs:
            continue

        curves: list[Curve] = []
        current_vs: list[float] = []
        current_abs_is: list[float] = []
        for x, y in pairs:
            if current_vs and x < current_vs[-1]:
                curves.append(Curve(name=f"group{len(curves) + 1}", vs=current_vs, abs_is=current_abs_is, source_order=len(curves) + 1))
                current_vs = []
                current_abs_is = []
            current_vs.append(x)
            current_abs_is.append(y)
        if current_vs:
            curves.append(Curve(name=f"group{len(curves) + 1}", vs=current_vs, abs_is=current_abs_is, source_order=len(curves) + 1))
        if curves:
            return curves
    return []


def parse_csv(path: pathlib.Path) -> CsvData:
    rows = read_rows(path)
    curves = parse_tagged(rows, path)
    if not curves:
        curves = parse_plain(rows, path)
    return CsvData(path=path, curves=curves)


def leaf_csv_folders(root: pathlib.Path) -> list[pathlib.Path]:
    folders = sorted({path.parent for path in root.rglob("*.csv")})
    return [folder for folder in folders if not folder.name.startswith("_")]


def project_chunks(items: list[CsvData], max_graphs: int) -> list[list[CsvData]]:
    if max_graphs <= 0 or len(items) <= max_graphs:
        return [items]
    return [items[index : index + max_graphs] for index in range(0, len(items), max_graphs)]


def chunk_output_path(folder: pathlib.Path, chunk_index: int, total_chunks: int) -> pathlib.Path:
    if total_chunks == 1 or chunk_index == 1:
        return folder / f"{folder.name}.opju"
    return folder / f"{folder.name}_part{chunk_index}.opju"


def remove_stale_part_projects(folder: pathlib.Path, total_chunks: int) -> None:
    valid_names = {
        chunk_output_path(folder, chunk_index, total_chunks).name
        for chunk_index in range(1, total_chunks + 1)
    }
    for path in folder.glob(f"{folder.name}_part*.opju"):
        if path.name not in valid_names:
            path.unlink()


def safe_origin_name(text: str, prefix: str) -> str:
    ascii_text = re.sub(r"[^A-Za-z0-9_]", "_", text)
    ascii_text = re.sub(r"_+", "_", ascii_text).strip("_")
    if not ascii_text or not ascii_text[0].isalpha():
        ascii_text = f"{prefix}_{ascii_text}"
    return ascii_text[:24]


def column_values(values: Iterable[float]) -> list[object]:
    return [float(value) for value in values]


def lt_quote(text: str) -> str:
    return text.replace("\\", "\\\\").replace('"', '\\"')


def lt_text_quote(text: str) -> str:
    return text.replace('"', '\\"')


def graph_legend_text(curves: list[Curve]) -> str:
    rows = [f"\\l({index}) absIs {curve.name}" for index, curve in enumerate(curves, start=1)]
    return "\r\n".join(rows)


def make_project(
    folder: pathlib.Path,
    csv_files: list[CsvData],
    output: pathlib.Path,
    visible: bool,
    make_graphs: bool,
) -> None:
    import win32com.client as win32

    app = win32.Dispatch("Origin.ApplicationSI")
    app.Visible = visible
    app.BeginSession()

    try:
        app.NewProject()
        app.Execute(f'pe_mkdir "{folder.name}";')
        app.Execute(f'pe_cd "{folder.name}";')
        book_short = safe_origin_name(f"data_{folder.name}", "data")
        app.Execute(f'newbook name:="{book_short}" option:=lsname;')
        actual_book = app.LTStr("page.name$")
        app.Execute(f'page.longname$="{folder.name} data";')

        for file_index, csv_data in enumerate(csv_files, start=1):
            stem = csv_data.path.stem
            print(f"    sheet: {stem} ({len(csv_data.curves)} curve groups)", flush=True)
            graph_short = safe_origin_name(f"g_{stem}_{file_index}", "g")
            ncols = max(2, len(csv_data.curves) + 1)
            nrows = max(len(curve.vs) for curve in csv_data.curves)

            app.Execute(f'win -a {actual_book};')
            if file_index > 1:
                sheet_short = safe_origin_name(stem, f"s{file_index}")
                app.Execute(f'newsheet name:="{sheet_short}";')
            else:
                sheet_short = safe_origin_name(stem, f"s{file_index}")
                app.Execute(f'wks.name$="{sheet_short}";')
            sheet_index = file_index
            app.Execute(f'win -a {actual_book}; page.active={sheet_index};')
            worksheet = app.FindWorksheet(actual_book)
            if not worksheet:
                raise RuntimeError(f"Could not create worksheet for {csv_data.path}")
            worksheet.Activate()
            app.Execute(f'wks.longname$="{lt_quote(stem)}";')
            worksheet.Cols = ncols
            worksheet.Rows = nrows
            worksheet.ClearData(0, int(worksheet.Cols) - 1)

            columns = worksheet.Columns
            x_values = csv_data.curves[0].vs
            app.Execute('wks.col1.lname$="Vs";')
            app.Execute('wks.col1.unit$="V";')
            app.Execute('wks.col1.type=4;')
            columns(0).SetData(column_values(x_values), 0)

            for curve_index, curve in enumerate(csv_data.curves):
                y_col = curve_index + 1
                app.Execute(f'wks.col{y_col + 1}.lname$="{lt_quote(f"absIs {curve.name}")}";')
                app.Execute(f'wks.col{y_col + 1}.unit$="A";')
                app.Execute(f'wks.col{y_col + 1}.comment$="";')
                app.Execute(f"wks.col{y_col + 1}.type=1;")
                columns(y_col).SetData(column_values(curve.abs_is), 0)

            if make_graphs:
                print(f"    graph: {stem}", flush=True)
                app.Execute(f'plotxy iy:=[{actual_book}]{sheet_index}!(1,2) plot:=200;')
                app.Execute(f'page.name$="{graph_short}";')
                app.Execute(f'page.longname$="{stem}";')
                app.Execute('layer.y.type=2;')

                for curve_index in range(1, len(csv_data.curves)):
                    y_col = curve_index + 2
                    app.Execute(f'plotxy iy:=[{actual_book}]{sheet_index}!(1,{y_col}) plot:=200 ogl:=1;')

                app.Execute("layer -g;")
                app.Execute("layer -s 1;")
                for plot_index in range(1, len(csv_data.curves) + 1):
                    app.Execute(f"layer.plot{plot_index}.line.width=3;")
                app.Execute("layer.x.thickness=3;")
                app.Execute("layer.y.thickness=3;")
                app.Execute("layer.framewidth=3;")
                app.Execute("layer.x.linewidth=3;")
                app.Execute("layer.y.linewidth=3;")
                app.Execute('xb.text$="\\b(Vs (V))";')
                app.Execute('yl.text$="\\b(absIs (A))";')
                app.Execute("label -b xb;")
                app.Execute("label -b yl;")
                app.Execute("label -b legend;")
                app.Execute("xb.font.bold=1;")
                app.Execute("yl.font.bold=1;")
                app.Execute("legend.font.bold=1;")
                app.Execute("layer.x.label.bold=1;")
                app.Execute("layer.y.label.bold=1;")
                app.Execute("legend -s;")
                app.Execute(f'legend.text$="{lt_text_quote(graph_legend_text(csv_data.curves))}";')
                app.Execute("legend.show=1;")
                app.Execute("label -b legend;")
                app.Execute("legend.x=1.55;")
                app.Execute("legend.y=1E-11;")
                app.Execute("layer -a;")

        if not app.Save(str(output)):
            raise RuntimeError(f"Origin failed to save {output}")
    finally:
        try:
            app.Exit()
        except Exception:
            pass


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", default=str(ROOT), help="Root folder containing CSV subfolders.")
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE), help="Template project used as visual reference.")
    parser.add_argument("--folder", help="Only process this CSV folder.")
    parser.add_argument("--visible", action="store_true", help="Show Origin during automation.")
    parser.add_argument("--no-graphs", action="store_true", help="Create workbooks only, for Origin COM debugging.")
    parser.add_argument(
        "--max-graphs-per-project",
        type=int,
        default=DEFAULT_MAX_GRAPHS_PER_PROJECT,
        help="Split projects after this many graph pages. Use 0 to disable splitting.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Only parse and report CSV groups.")
    args = parser.parse_args()

    root = pathlib.Path(args.root).resolve()
    template = pathlib.Path(args.template).resolve()
    folders = [pathlib.Path(args.folder).resolve()] if args.folder else leaf_csv_folders(root)
    if not folders:
        raise FileNotFoundError(f"No CSV files found under {root}")

    print(f"Root: {root}")
    print(f"Template reference: {template}")
    for folder in folders:
        parsed = [parse_csv(path) for path in sorted(folder.glob("*.csv"))]
        parsed = [item for item in parsed if item.curves]
        skipped = [path.name for path in sorted(folder.glob("*.csv")) if path not in {item.path for item in parsed}]
        chunks = project_chunks(parsed, args.max_graphs_per_project)
        if len(chunks) > 1:
            targets = ", ".join(chunk_output_path(folder, index, len(chunks)).name for index in range(1, len(chunks) + 1))
        else:
            targets = f"{folder.name}.opju"
        summary = ", ".join(f"{item.path.name}:{len(item.curves)}" for item in parsed)
        print(f"{folder.relative_to(root)} -> {targets} | {summary}")
        if skipped:
            print(f"  skipped empty/unrecognized: {', '.join(skipped)}")
        if args.dry_run:
            continue
        if parsed:
            remove_stale_part_projects(folder, len(chunks))
            for chunk_index, chunk in enumerate(chunks, start=1):
                output = chunk_output_path(folder, chunk_index, len(chunks))
                if len(chunks) > 1:
                    print(f"  part {chunk_index}/{len(chunks)}: {len(chunk)} graph pages -> {output.name}")
                make_project(folder, chunk, output, args.visible, not args.no_graphs)
                print(f"  saved: {output}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
