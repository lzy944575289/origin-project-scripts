#!/usr/bin/env python
"""Build Origin projects from AV/AI XLS exports in selected folders."""

from __future__ import annotations

import argparse
import math
import pathlib
import re
import warnings
from dataclasses import dataclass
from typing import Iterable

import pandas as pd


ROOT = pathlib.Path(__file__).resolve().parent
DEFAULT_FOLDERS = (ROOT / "20260611", ROOT / "20260612")
DEFAULT_MAX_GRAPHS_PER_PROJECT = 11


@dataclass
class Curve:
    name: str
    av: list[float]
    abs_ai: list[float]


@dataclass
class XlsData:
    path: pathlib.Path
    curves: list[Curve]


def parse_float(value: object) -> float | None:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if math.isnan(number):
        return None
    return number


def append_sort_key(sheet_name: str) -> tuple[int, int]:
    if sheet_name.lower() == "data":
        return (0, 0)
    match = re.fullmatch(r"append(\d+)", sheet_name, flags=re.IGNORECASE)
    if match:
        return (1, int(match.group(1)))
    return (2, 0)


def relevant_sheets(path: pathlib.Path) -> list[str]:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        book = pd.ExcelFile(path, engine="xlrd")
    sheets = [
        name
        for name in book.sheet_names
        if name.lower() == "data" or re.fullmatch(r"append\d+", name, flags=re.IGNORECASE)
    ]
    return sorted(sheets, key=append_sort_key)


def read_curve(path: pathlib.Path, sheet_name: str) -> Curve | None:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        frame = pd.read_excel(path, sheet_name=sheet_name, engine="xlrd")

    if "AV" not in frame.columns or "AI" not in frame.columns:
        raise ValueError(f"{path.name} sheet {sheet_name} is missing AV or AI")

    av: list[float] = []
    abs_ai: list[float] = []
    for raw_x, raw_y in zip(frame["AV"], frame["AI"]):
        x = parse_float(raw_x)
        y = parse_float(raw_y)
        if x is None or y is None:
            continue
        av.append(x)
        abs_ai.append(abs(y))

    if not av:
        return None
    return Curve(name=sheet_name, av=av, abs_ai=abs_ai)


def parse_xls(path: pathlib.Path) -> XlsData:
    curves: list[Curve] = []
    for sheet_name in relevant_sheets(path):
        curve = read_curve(path, sheet_name)
        if curve:
            curves.append(curve)
    return XlsData(path=path, curves=curves)


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
    rows = [f"\\l({index}) {curve.name}" for index, curve in enumerate(curves, start=1)]
    return "\r\n".join(rows)


def make_project(
    folder: pathlib.Path,
    xls_files: list[XlsData],
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

        for file_index, xls_data in enumerate(xls_files, start=1):
            stem = xls_data.path.stem
            print(f"    sheet: {stem} ({len(xls_data.curves)} curve groups)", flush=True)
            graph_short = safe_origin_name(f"g_{stem}_{file_index}", "g")
            ncols = max(2, len(xls_data.curves) + 1)
            nrows = max(len(curve.av) for curve in xls_data.curves)

            app.Execute(f"win -a {actual_book};")
            if file_index > 1:
                sheet_short = safe_origin_name(stem, f"s{file_index}")
                app.Execute(f'newsheet name:="{sheet_short}";')
            else:
                sheet_short = safe_origin_name(stem, f"s{file_index}")
                app.Execute(f'wks.name$="{sheet_short}";')
            sheet_index = file_index
            app.Execute(f"win -a {actual_book}; page.active={sheet_index};")
            worksheet = app.FindWorksheet(actual_book)
            if not worksheet:
                raise RuntimeError(f"Could not create worksheet for {xls_data.path}")
            worksheet.Activate()
            app.Execute(f'wks.longname$="{lt_quote(stem)}";')
            worksheet.Cols = ncols
            worksheet.Rows = nrows
            worksheet.ClearData(0, int(worksheet.Cols) - 1)

            columns = worksheet.Columns
            x_values = xls_data.curves[0].av
            app.Execute('wks.col1.lname$="AV";')
            app.Execute('wks.col1.unit$="V";')
            app.Execute("wks.col1.type=4;")
            columns(0).SetData(column_values(x_values), 0)

            for curve_index, curve in enumerate(xls_data.curves):
                y_col = curve_index + 1
                app.Execute(f'wks.col{y_col + 1}.lname$="{lt_quote(f"absAI {curve.name}")}";')
                app.Execute(f'wks.col{y_col + 1}.unit$="A";')
                app.Execute(f'wks.col{y_col + 1}.comment$="{lt_quote(curve.name)}";')
                app.Execute(f"wks.col{y_col + 1}.type=1;")
                columns(y_col).SetData(column_values(curve.abs_ai), 0)

            if make_graphs:
                print(f"    graph: {stem}", flush=True)
                app.Execute(f"plotxy iy:=[{actual_book}]{sheet_index}!(1,2) plot:=200;")
                app.Execute(f'page.name$="{graph_short}";')
                app.Execute(f'page.longname$="{lt_quote(stem)}";')
                app.Execute("layer.y.type=2;")

                for curve_index in range(1, len(xls_data.curves)):
                    y_col = curve_index + 2
                    app.Execute(f"plotxy iy:=[{actual_book}]{sheet_index}!(1,{y_col}) plot:=200 ogl:=1;")

                app.Execute("layer -g;")
                app.Execute("layer -s 1;")
                for plot_index in range(1, len(xls_data.curves) + 1):
                    app.Execute(f"layer.plot{plot_index}.line.width=3;")
                app.Execute("layer.x.thickness=3;")
                app.Execute("layer.y.thickness=3;")
                app.Execute("layer.framewidth=3;")
                app.Execute("layer.x.linewidth=3;")
                app.Execute("layer.y.linewidth=3;")
                app.Execute('xb.text$="\\b(AV (V))";')
                app.Execute('yl.text$="\\b(absAI (A))";')
                app.Execute("label -b xb;")
                app.Execute("label -b yl;")
                app.Execute("label -b legend;")
                app.Execute("xb.font.bold=1;")
                app.Execute("yl.font.bold=1;")
                app.Execute("legend.font.bold=1;")
                app.Execute("layer.x.label.bold=1;")
                app.Execute("layer.y.label.bold=1;")
                app.Execute("legend -s;")
                app.Execute(f'legend.text$="{lt_text_quote(graph_legend_text(xls_data.curves))}";')
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


def folder_xls_files(folder: pathlib.Path) -> list[pathlib.Path]:
    return sorted(
        path
        for path in folder.glob("*.xls")
        if not path.name.startswith("~$") and path.is_file()
    )


def project_chunks(items: list[XlsData], max_graphs: int) -> list[list[XlsData]]:
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


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--folder",
        action="append",
        help="Folder containing XLS files. Can be passed more than once. Defaults to 20260611 and 20260612.",
    )
    parser.add_argument("--visible", action="store_true", help="Show Origin during automation.")
    parser.add_argument("--no-graphs", action="store_true", help="Create workbooks only, for Origin COM debugging.")
    parser.add_argument(
        "--max-graphs-per-project",
        type=int,
        default=DEFAULT_MAX_GRAPHS_PER_PROJECT,
        help="Split projects after this many graph pages. Use 0 to disable splitting.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Only parse and report XLS sheets.")
    args = parser.parse_args()

    folders = [pathlib.Path(item).resolve() for item in args.folder] if args.folder else list(DEFAULT_FOLDERS)
    for folder in folders:
        if not folder.exists():
            raise FileNotFoundError(folder)

        parsed = [parse_xls(path) for path in folder_xls_files(folder)]
        parsed = [item for item in parsed if item.curves]
        skipped = [path.name for path in folder_xls_files(folder) if path not in {item.path for item in parsed}]
        summary = ", ".join(f"{item.path.name}:{len(item.curves)}" for item in parsed)
        chunks = project_chunks(parsed, args.max_graphs_per_project)
        if len(chunks) > 1:
            targets = ", ".join(chunk_output_path(folder, index, len(chunks)).name for index in range(1, len(chunks) + 1))
        else:
            targets = f"{folder.name}.opju"
        print(f"{folder} -> {targets} | {summary}")
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
