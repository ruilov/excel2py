import argparse
import json
import re
from pathlib import Path

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

    for match in UNQUALIFIED_REF_PATTERN.finditer(remaining):
        add_ref(current_sheet_idx, match.group("addr"))

    for token in NAME_TOKEN_PATTERN.findall(remaining):
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
    with paths["formulas_path"].open("r", encoding="utf-8") as source:
        with paths["dependencies_path"].open("w", encoding="utf-8", newline="\n") as target:
            for line in source:
                record = json.loads(line)
                formula = record[formula_idx]
                if formula is None:
                    continue

                dependencies = _extract_dependencies(
                    formula=formula,
                    current_sheet_idx=record[sheet_idx_idx],
                    sheet_idx_by_name=sheet_idx_by_name,
                    defined_name_by_upper=defined_name_by_upper,
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

    print(f"Workbook: {result['workbook_path']}")
    print(f"Raw directory: {result['raw_dir']}")
    print(f"Derived directory: {result['derived_dir']}")
    print(f"Formulas JSONL: {result['formulas_path']}")
    print(f"Formula rows: {result['formula_count']}")
    print(f"Dependencies JSONL: {dep_result['dependencies_path']}")
    print(f"Dependency records: {dep_result['dependency_record_count']}")
    print(f"Dependency edges: {dep_result['dependency_edge_count']}")


if __name__ == "__main__":
    main()
