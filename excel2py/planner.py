import argparse
import json
import re
from pathlib import Path

import networkx as nx
from openpyxl.utils import cell as xl_cell

from excel2py.formula_parser import extract_dependencies_from_formula
from excel2py.loader import *

CELL_REF = r"\$?[A-Za-z]{1,3}\$?\d+"
RANGE_REF = rf"{CELL_REF}:{CELL_REF}"
COL_RANGE_REF = r"\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}"
ROW_RANGE_REF = r"\$?\d+:\$?\d+"
ADDR_REF = rf"(?:{RANGE_REF}|{CELL_REF}|{COL_RANGE_REF}|{ROW_RANGE_REF})"
SHEET_TOKEN = r"(?:'[^']+'|[A-Za-z0-9_ .\[\]-]+)"
QUALIFIED_REF_PATTERN = re.compile(rf"(?P<sheet>{SHEET_TOKEN})!(?P<addr>{ADDR_REF})")
UNQUALIFIED_REF_PATTERN = re.compile(rf"(?<![A-Za-z0-9_\.])(?P<addr>{ADDR_REF})(?![A-Za-z0-9_])")
NAME_TOKEN_PATTERN = re.compile(r"[A-Za-z_\\][A-Za-z0-9_.\\]*")


def _planner_paths(workbook_path: Path, artifacts_root: str | Path) -> dict[str, Path]:
    workbook_key = workbook_path.stem
    workbook_dir = Path(artifacts_root) / workbook_key
    raw_dir = workbook_dir / "raw"
    derived_dir = workbook_dir / "derived"
    return {
        "raw_dir": raw_dir,
        "derived_dir": derived_dir,
        "cells_path": raw_dir / "cells.jsonl",
        "metadata_path": raw_dir / "workbook_meta.json",
        "formulas_path": derived_dir / "formulas.jsonl",
        "dependencies_path": derived_dir / "dependencies.jsonl",
        "calc_order_path": derived_dir / "calc_order.json",
    }


def export_formula_rows(
    excel_file: str | Path = DEFAULT_WORKBOOK,
    artifacts_root: str | Path = DEFAULT_ARTIFACTS_ROOT,
) -> dict[str, object]:
    workbook_path = Path(excel_file)
    paths = _planner_paths(workbook_path, artifacts_root)
    paths["derived_dir"].mkdir(parents=True, exist_ok=True)

    metadata = json.loads(paths["metadata_path"].read_text(encoding="utf-8"))
    record_fields = metadata["cells_jsonl_record_fields"]
    formula_idx = record_fields.index("formula")

    formula_count = 0
    with paths["cells_path"].open("r", encoding="utf-8") as source:
        with paths["formulas_path"].open("w", encoding="utf-8", newline="\n") as target:
            for line in source:
                record = json.loads(line)
                if record[formula_idx] is None:
                    continue
                target.write(json.dumps(record, ensure_ascii=True) + "\n")
                formula_count += 1

    return {
        "workbook_path": str(workbook_path.resolve()),
        "raw_dir": str(paths["raw_dir"].resolve()),
        "derived_dir": str(paths["derived_dir"].resolve()),
        "formulas_path": str(paths["formulas_path"].resolve()),
        "formula_count": formula_count,
    }


def _strip_quoted_strings(text: str) -> str:
    chars = list(text)
    in_quote = False
    for idx, char in enumerate(chars):
        if char == '"':
            in_quote = not in_quote
            chars[idx] = " "
            continue
        if in_quote:
            chars[idx] = " "
    return "".join(chars)


def _normalize_sheet_token(sheet_token: str) -> str:
    token = sheet_token.strip()
    if token.startswith("'") and token.endswith("'") and len(token) >= 2:
        token = token[1:-1].replace("''", "'")
    return token


def _sheet_idx_from_token(sheet_token: str, sheet_idx_by_name: dict[str, int]) -> int | None:
    normalized = _normalize_sheet_token(sheet_token)
    if normalized in sheet_idx_by_name:
        return sheet_idx_by_name[normalized]

    no_book_prefix = re.sub(r"^\[[^\]]+\]", "", normalized)
    if no_book_prefix in sheet_idx_by_name:
        return sheet_idx_by_name[no_book_prefix]

    return None


def _extract_dependencies(
    formula: str,
    current_sheet_idx: int,
    sheet_idx_by_name: dict[str, int],
    defined_name_by_upper: dict[str, str],
) -> list[list[object]]:
    formula_body = formula[1:] if formula.startswith("=") else formula
    sanitized = _strip_quoted_strings(formula_body)
    chars = list(sanitized)

    dependencies: list[list[object]] = []
    seen: set[tuple[object, ...]] = set()

    def add_ref(sheet_idx: int, addr: str) -> None:
        key = ("ref", sheet_idx, addr)
        if key in seen:
            return
        seen.add(key)
        dependencies.append([sheet_idx, addr])

    def add_external_ref(ref: str) -> None:
        key = ("ext", ref)
        if key in seen:
            return
        seen.add(key)
        dependencies.append(["ext", ref])

    def add_named_range(name: str) -> None:
        key = ("name", name)
        if key in seen:
            return
        seen.add(key)
        dependencies.append(["name", name])

    for match in QUALIFIED_REF_PATTERN.finditer(sanitized):
        start, end = match.span()
        for idx in range(start, end):
            chars[idx] = " "

        target_idx = _sheet_idx_from_token(match.group("sheet"), sheet_idx_by_name)
        addr = match.group("addr")
        if target_idx is None:
            add_external_ref(f"{_normalize_sheet_token(match.group('sheet'))}!{addr}")
        else:
            add_ref(target_idx, addr)

    remaining = "".join(chars)
    remaining_chars = list(remaining)

    for match in UNQUALIFIED_REF_PATTERN.finditer(remaining):
        start, end = match.span()
        for idx in range(start, end):
            remaining_chars[idx] = " "
        add_ref(current_sheet_idx, match.group("addr"))

    named_range_scan_text = "".join(remaining_chars)
    for token in NAME_TOKEN_PATTERN.findall(named_range_scan_text):
        name = defined_name_by_upper.get(token.upper())
        if name is not None:
            add_named_range(name)

    return dependencies


def export_formula_dependencies(
    excel_file: str | Path = DEFAULT_WORKBOOK,
    artifacts_root: str | Path = DEFAULT_ARTIFACTS_ROOT,
) -> dict[str, object]:
    workbook_path = Path(excel_file)
    paths = _planner_paths(workbook_path, artifacts_root)
    paths["derived_dir"].mkdir(parents=True, exist_ok=True)

    metadata = json.loads(paths["metadata_path"].read_text(encoding="utf-8"))
    formula_record_fields = metadata["cells_jsonl_record_fields"]
    sheet_idx_idx = formula_record_fields.index("sheet_idx")
    addr_idx = formula_record_fields.index("addr")
    formula_idx = formula_record_fields.index("formula")

    sheet_idx_by_name = {name: idx for idx, name in enumerate(metadata["sheet_names"])}
    defined_name_by_upper = {
        item["name"].upper(): item["name"] for item in metadata.get("named_ranges", [])
    }

    dependency_record_count = 0
    dependency_edge_count = 0
    parse_error_count = 0
    parse_error_samples: list[dict[str, object]] = []
    max_error_samples = 20
    with paths["formulas_path"].open("r", encoding="utf-8") as source:
        with paths["dependencies_path"].open("w", encoding="utf-8", newline="\n") as target:
            for line in source:
                record = json.loads(line)
                formula = record[formula_idx]
                if formula is None:
                    continue

                dependencies, parse_error = extract_dependencies_from_formula(
                    formula=formula,
                    current_sheet_idx=record[sheet_idx_idx],
                    sheet_idx_by_name=sheet_idx_by_name,
                    defined_name_by_upper=defined_name_by_upper,
                )
                if dependencies is None:
                    dependencies = []
                if parse_error is not None:
                    parse_error_count += 1
                    if len(parse_error_samples) < max_error_samples:
                        parse_error_samples.append(
                            {
                                "sheet_idx": record[sheet_idx_idx],
                                "addr": record[addr_idx],
                                "formula": formula,
                                "error": parse_error,
                            }
                        )

                out_record = [
                    record[sheet_idx_idx],
                    record[addr_idx],
                    dependencies,
                ]
                target.write(json.dumps(out_record, ensure_ascii=True) + "\n")
                dependency_record_count += 1
                dependency_edge_count += len(dependencies)

    return {
        "workbook_path": str(workbook_path.resolve()),
        "derived_dir": str(paths["derived_dir"].resolve()),
        "dependencies_path": str(paths["dependencies_path"].resolve()),
        "dependency_record_count": dependency_record_count,
        "dependency_edge_count": dependency_edge_count,
        "dependency_record_fields": ["sheet_idx", "addr", "dependencies"],
        "dependency_entry_variants": [
            "[sheet_idx, addr]",
            "['name', range_name]",
            "['ext', external_ref]",
        ],
        "parse_error_count": parse_error_count,
        "parse_error_samples": parse_error_samples,
    }


def _normalize_addr(addr: str) -> str:
    return addr.replace("$", "").upper()


def _parse_cell_addr(addr: str) -> tuple[int, int] | None:
    normalized = _normalize_addr(addr)
    try:
        row, col = xl_cell.coordinate_to_tuple(normalized)
        return row, col
    except ValueError:
        return None


def _addr_kind(addr: str) -> str:
    normalized = _normalize_addr(addr)
    if ":" not in normalized:
        parsed = _parse_cell_addr(normalized)
        return "single" if parsed is not None else "other"

    left, right = normalized.split(":", 1)
    if re.fullmatch(r"[A-Z]{1,3}", left) and re.fullmatch(r"[A-Z]{1,3}", right):
        return "col_range"
    if re.fullmatch(r"\d+", left) and re.fullmatch(r"\d+", right):
        return "row_range"

    try:
        xl_cell.range_boundaries(normalized)
        return "rect_range"
    except ValueError:
        return "other"


def _formula_nodes_by_sheet(formulas_path: Path) -> tuple[set[tuple[int, str]], dict[int, list[tuple[str, int, int]]]]:
    formula_nodes: set[tuple[int, str]] = set()
    by_sheet: dict[int, list[tuple[str, int, int]]] = {}

    with formulas_path.open("r", encoding="utf-8") as source:
        for line in source:
            sheet_idx, addr, *_rest = json.loads(line)
            normalized_addr = _normalize_addr(addr)
            formula_nodes.add((sheet_idx, normalized_addr))
            parsed = _parse_cell_addr(normalized_addr)
            if parsed is None:
                continue
            row, col = parsed
            by_sheet.setdefault(sheet_idx, []).append((normalized_addr, row, col))

    return formula_nodes, by_sheet


def _resolve_ref_to_formula_nodes(
    sheet_idx: int,
    addr: str,
    formula_nodes: set[tuple[int, str]],
    formula_coords_by_sheet: dict[int, list[tuple[str, int, int]]],
) -> list[tuple[int, str]]:
    normalized = _normalize_addr(addr)
    kind = _addr_kind(normalized)

    if kind == "single":
        return [(sheet_idx, normalized)] if (sheet_idx, normalized) in formula_nodes else []

    coords = formula_coords_by_sheet.get(sheet_idx, [])
    if not coords:
        return []

    if kind == "rect_range":
        min_col, min_row, max_col, max_row = xl_cell.range_boundaries(normalized)
        return [
            (sheet_idx, node_addr)
            for node_addr, row, col in coords
            if min_row <= row <= max_row and min_col <= col <= max_col
        ]

    if kind == "col_range":
        left, right = normalized.split(":", 1)
        min_col = xl_cell.column_index_from_string(left)
        max_col = xl_cell.column_index_from_string(right)
        if min_col > max_col:
            min_col, max_col = max_col, min_col
        return [
            (sheet_idx, node_addr)
            for node_addr, _row, col in coords
            if min_col <= col <= max_col
        ]

    if kind == "row_range":
        left, right = normalized.split(":", 1)
        min_row = int(left)
        max_row = int(right)
        if min_row > max_row:
            min_row, max_row = max_row, min_row
        return [
            (sheet_idx, node_addr)
            for node_addr, row, _col in coords
            if min_row <= row <= max_row
        ]

    return []


def build_calc_order(
    excel_file: str | Path = DEFAULT_WORKBOOK,
    artifacts_root: str | Path = DEFAULT_ARTIFACTS_ROOT,
    parse_diagnostics: dict[str, object] | None = None,
) -> dict[str, object]:
    workbook_path = Path(excel_file)
    paths = _planner_paths(workbook_path, artifacts_root)

    formula_nodes, formula_coords_by_sheet = _formula_nodes_by_sheet(paths["formulas_path"])

    graph = nx.DiGraph()
    graph.add_nodes_from(formula_nodes)

    input_ref_count = 0
    unresolved_named_range_count = 0
    unresolved_external_ref_count = 0

    with paths["dependencies_path"].open("r", encoding="utf-8") as source:
        for line in source:
            sheet_idx, addr, dependencies = json.loads(line)
            target = (sheet_idx, _normalize_addr(addr))
            for dependency in dependencies:
                if (
                    isinstance(dependency, list)
                    and len(dependency) == 2
                    and isinstance(dependency[0], int)
                    and isinstance(dependency[1], str)
                ):
                    dep_sheet_idx, dep_addr = dependency
                    resolved_nodes = _resolve_ref_to_formula_nodes(
                        dep_sheet_idx,
                        dep_addr,
                        formula_nodes,
                        formula_coords_by_sheet,
                    )
                    if not resolved_nodes:
                        input_ref_count += 1
                    for dep_node in resolved_nodes:
                        graph.add_edge(dep_node, target)
                    continue

                if (
                    isinstance(dependency, list)
                    and len(dependency) == 2
                    and dependency[0] == "name"
                ):
                    unresolved_named_range_count += 1
                    continue

                if (
                    isinstance(dependency, list)
                    and len(dependency) == 2
                    and dependency[0] == "ext"
                ):
                    unresolved_external_ref_count += 1
                    continue

    cycle_components = []
    for component in nx.strongly_connected_components(graph):
        if len(component) > 1:
            cycle_components.append(component)
        elif len(component) == 1:
            only = next(iter(component))
            if graph.has_edge(only, only):
                cycle_components.append(component)

    cycle_nodes = set().union(*cycle_components) if cycle_components else set()
    acyclic_graph = graph.copy()
    acyclic_graph.remove_nodes_from(cycle_nodes)
    calc_order = [[node[0], node[1]] for node in nx.topological_sort(acyclic_graph)]
    cycles = [
        sorted([[node[0], node[1]] for node in component], key=lambda item: (item[0], item[1]))
        for component in cycle_components
    ]

    calc_order_payload = {
        "workbook_path": str(workbook_path.resolve()),
        "calc_order_record_fields": ["sheet_idx", "addr"],
        "calc_order": calc_order,
        "has_cycles": len(cycles) > 0,
        "cycles": cycles,
        "parse_diagnostics": {
            "parse_error_count": 0 if parse_diagnostics is None else parse_diagnostics.get("parse_error_count", 0),
            "parse_error_samples": [] if parse_diagnostics is None else parse_diagnostics.get("parse_error_samples", []),
        },
        "stats": {
            "formula_node_count": graph.number_of_nodes(),
            "formula_edge_count": graph.number_of_edges(),
            "calc_order_count": len(calc_order),
            "cycle_group_count": len(cycles),
            "cycle_node_count": len(cycle_nodes),
            "input_ref_count": input_ref_count,
            "unresolved_named_range_count": unresolved_named_range_count,
            "unresolved_external_ref_count": unresolved_external_ref_count,
            "parse_error_count": 0 if parse_diagnostics is None else parse_diagnostics.get("parse_error_count", 0),
        },
    }
    paths["calc_order_path"].write_text(
        json.dumps(calc_order_payload, indent=2, ensure_ascii=True) + "\n",
        encoding="utf-8",
    )

    return {
        "workbook_path": str(workbook_path.resolve()),
        "derived_dir": str(paths["derived_dir"].resolve()),
        "calc_order_path": str(paths["calc_order_path"].resolve()),
        "formula_node_count": graph.number_of_nodes(),
        "formula_edge_count": graph.number_of_edges(),
        "calc_order_count": len(calc_order),
        "cycle_group_count": len(cycles),
        "cycle_node_count": len(cycle_nodes),
        "input_ref_count": input_ref_count,
        "unresolved_named_range_count": unresolved_named_range_count,
        "unresolved_external_ref_count": unresolved_external_ref_count,
        "parse_error_count": 0 if parse_diagnostics is None else parse_diagnostics.get("parse_error_count", 0),
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Build derived planning artifacts from raw workbook artifacts."
    )
    parser.add_argument(
        "excel_file",
        nargs="?",
        default=str(DEFAULT_WORKBOOK),
        help=f"Path (or base name) of the workbook. Default: {DEFAULT_WORKBOOK.name}",
    )
    parser.add_argument(
        "--artifacts-root",
        default=str(DEFAULT_ARTIFACTS_ROOT),
        help=f"Root folder for generated artifacts. Default: {DEFAULT_ARTIFACTS_ROOT}",
    )
    args = parser.parse_args()

    result = export_formula_rows(
        excel_file=args.excel_file,
        artifacts_root=args.artifacts_root,
    )
    dep_result = export_formula_dependencies(
        excel_file=args.excel_file,
        artifacts_root=args.artifacts_root,
    )
    order_result = build_calc_order(
        excel_file=args.excel_file,
        artifacts_root=args.artifacts_root,
        parse_diagnostics=dep_result,
    )

    print(f"Workbook: {result['workbook_path']}")
    print(f"Raw directory: {result['raw_dir']}")
    print(f"Derived directory: {result['derived_dir']}")
    print(f"Formulas JSONL: {result['formulas_path']}")
    print(f"Formula rows: {result['formula_count']}")
    print(f"Dependencies JSONL: {dep_result['dependencies_path']}")
    print(f"Dependency records: {dep_result['dependency_record_count']}")
    print(f"Dependency edges: {dep_result['dependency_edge_count']}")
    print(f"Parse errors: {dep_result['parse_error_count']}")
    print(f"Calc Order JSON: {order_result['calc_order_path']}")
    print(f"Formula nodes: {order_result['formula_node_count']}")
    print(f"Graph edges: {order_result['formula_edge_count']}")
    print(f"Calc order rows: {order_result['calc_order_count']}")
    print(f"Cycle groups: {order_result['cycle_group_count']}")
    print(f"Cycle nodes: {order_result['cycle_node_count']}")
    print(f"Input refs: {order_result['input_ref_count']}")


if __name__ == "__main__":
    main()
