import argparse
import json
import re
from pathlib import Path

from excel2py.loader import *


def _translator_paths(workbook_path: Path, artifacts_root: str | Path) -> dict[str, Path]:
    workbook_key = workbook_path.stem
    workbook_dir = Path(artifacts_root) / workbook_key
    raw_dir = workbook_dir / "raw"
    derived_dir = workbook_dir / "derived"
    generated_dir = workbook_dir / "generated"
    return {
        "raw_dir": raw_dir,
        "derived_dir": derived_dir,
        "generated_dir": generated_dir,
        "metadata_path": raw_dir / "workbook_meta.json",
        "formulas_path": derived_dir / "formulas.jsonl",
        "dependencies_path": derived_dir / "dependencies.jsonl",
        "calc_order_path": derived_dir / "calc_order.json",
        "output_path": generated_dir / f"{workbook_key}_literal.py",
    }


def _normalize_addr(addr: str) -> str:
    return addr.replace("$", "").upper()


def _safe_ident(value: str) -> str:
    return re.sub(r"[^a-zA-Z0-9_]", "_", value)


def _cell_func_name(sheet_idx: int, addr: str) -> str:
    return f"calc_s{sheet_idx}_{_safe_ident(_normalize_addr(addr).lower())}"


def _dep_arg_name(dep: list[object], position: int) -> str:
    if len(dep) == 2 and isinstance(dep[0], int):
        return f"d_s{dep[0]}_{_safe_ident(str(dep[1]).lower())}_{position}"
    if len(dep) == 2 and dep[0] == "name":
        return f"n_{_safe_ident(str(dep[1]).lower())}_{position}"
    if len(dep) == 2 and dep[0] == "ext":
        return f"x_{_safe_ident(str(dep[1]).lower())}_{position}"
    return f"dep_{position}"


def _load_jsonl(path: Path) -> list[object]:
    rows = []
    with path.open("r", encoding="utf-8") as handle:
        for line in handle:
            rows.append(json.loads(line))
    return rows


def _emit_function(
    lines: list[str],
    sheet_names: list[str],
    sheet_idx: int,
    addr: str,
    formula: str,
    dependencies: list[list[object]],
) -> None:
    func_name = _cell_func_name(sheet_idx, addr)
    arg_names = [_dep_arg_name(dep, idx) for idx, dep in enumerate(dependencies)]
    args_src = ", ".join(arg_names)

    lines.append(f"def {func_name}({args_src}) -> object:")
    lines.append(f"    # from {sheet_names[sheet_idx]}!{addr}")
    lines.append(f"    # excel: {formula}")
    lines.append(
        f"    raise NotImplementedError(\"Formula lowering not implemented yet for {sheet_names[sheet_idx]}!{addr}\")"
    )
    lines.append("")


def _emit_run_model(
    lines: list[str],
    calc_order: list[list[object]],
    cycle_groups: list[list[list[object]]],
    dependency_map: dict[tuple[int, str], list[list[object]]],
) -> None:
    lines.append("def run_model(inputs: dict[tuple[int, str], object], max_iterations: int = 100, tolerance: float = 1e-9) -> dict[tuple[int, str], object]:")
    lines.append("    cells = dict(inputs)")
    lines.append("")
    lines.append("    # Acyclic formula evaluations")
    for sheet_idx, addr in calc_order:
        key = (sheet_idx, _normalize_addr(addr))
        deps = dependency_map.get(key, [])
        dep_calls = [
            f"resolve_dependency(cells, {repr(dep)}, sheet_dimensions=SHEET_DIMENSIONS)"
            for dep in deps
        ]
        func_name = _cell_func_name(sheet_idx, addr)
        lines.append(f"    cells[{repr(key)}] = {func_name}(")
        if dep_calls:
            for dep_call in dep_calls:
                lines.append(f"        {dep_call},")
        lines.append("    )")
    lines.append("")
    lines.append("    # Cyclic groups (iterative blocks)")
    for cycle in cycle_groups:
        lines.append("    for _iteration in range(max_iterations):")
        lines.append("        max_delta = 0.0")
        for sheet_idx, addr in cycle:
            key = (sheet_idx, _normalize_addr(addr))
            deps = dependency_map.get(key, [])
            dep_calls = [
                f"resolve_dependency(cells, {repr(dep)}, sheet_dimensions=SHEET_DIMENSIONS)"
                for dep in deps
            ]
            func_name = _cell_func_name(sheet_idx, addr)
            lines.append(f"        _old_value = cells.get({repr(key)})")
            lines.append(f"        _new_value = {func_name}(")
            if dep_calls:
                for dep_call in dep_calls:
                    lines.append(f"            {dep_call},")
            lines.append("        )")
            lines.append(f"        cells[{repr(key)}] = _new_value")
            lines.append("        if isinstance(_old_value, (int, float)) and isinstance(_new_value, (int, float)):")
            lines.append("            _delta = abs(_new_value - _old_value)")
            lines.append("            if _delta > max_delta:")
            lines.append("                max_delta = _delta")
        lines.append("        if max_delta <= tolerance:")
        lines.append("            break")
    lines.append("")
    lines.append("    return cells")
    lines.append("")


def emit_literal_skeleton(
    excel_file: str | Path = DEFAULT_WORKBOOK,
    artifacts_root: str | Path = DEFAULT_ARTIFACTS_ROOT,
    output_path: str | Path | None = None,
    limit_formulas: int | None = None,
) -> dict[str, object]:
    workbook_path = Path(excel_file)
    paths = _translator_paths(workbook_path, artifacts_root)
    paths["generated_dir"].mkdir(parents=True, exist_ok=True)

    metadata = json.loads(paths["metadata_path"].read_text(encoding="utf-8"))
    calc_order_payload = json.loads(paths["calc_order_path"].read_text(encoding="utf-8"))

    formula_rows = _load_jsonl(paths["formulas_path"])
    dependency_rows = _load_jsonl(paths["dependencies_path"])

    formula_map = {
        (row[0], _normalize_addr(row[1])): row[3]
        for row in formula_rows
    }
    dependency_map = {
        (row[0], _normalize_addr(row[1])): row[2]
        for row in dependency_rows
    }

    calc_order = [[row[0], _normalize_addr(row[1])] for row in calc_order_payload["calc_order"]]
    cycle_groups = [
        [[row[0], _normalize_addr(row[1])] for row in group]
        for group in calc_order_payload["cycles"]
    ]

    if limit_formulas is not None:
        calc_order = calc_order[:limit_formulas]

    sheet_dimensions = {
        idx: metadata["sheet_dimensions"][sheet_name]
        for idx, sheet_name in enumerate(metadata["sheet_names"])
    }

    output = Path(output_path) if output_path is not None else paths["output_path"]
    lines: list[str] = []
    lines.append("# Auto-generated deterministic literal skeleton.")
    lines.append("# This file is intentionally explicit for later AI refactoring.")
    lines.append("from excel2py.runtime_helpers import *")
    lines.append("")
    lines.append(f"SHEET_NAMES = {repr(metadata['sheet_names'])}")
    lines.append(f"SHEET_DIMENSIONS = {repr(sheet_dimensions)}")
    lines.append(f"CALC_ORDER = {repr(calc_order)}")
    lines.append(f"CYCLE_GROUPS = {repr(cycle_groups)}")
    lines.append("")

    emitted_formula_keys: set[tuple[int, str]] = set()
    for sheet_idx, addr in calc_order:
        key = (sheet_idx, addr)
        emitted_formula_keys.add(key)
        formula = formula_map.get(key)
        if formula is None:
            continue
        dependencies = dependency_map.get(key, [])
        _emit_function(
            lines=lines,
            sheet_names=metadata["sheet_names"],
            sheet_idx=sheet_idx,
            addr=addr,
            formula=formula,
            dependencies=dependencies,
        )

    for cycle in cycle_groups:
        for sheet_idx, addr in cycle:
            key = (sheet_idx, addr)
            if key in emitted_formula_keys:
                continue
            emitted_formula_keys.add(key)
            formula = formula_map.get(key)
            if formula is None:
                continue
            dependencies = dependency_map.get(key, [])
            _emit_function(
                lines=lines,
                sheet_names=metadata["sheet_names"],
                sheet_idx=sheet_idx,
                addr=addr,
                formula=formula,
                dependencies=dependencies,
            )

    _emit_run_model(
        lines=lines,
        calc_order=calc_order,
        cycle_groups=cycle_groups,
        dependency_map=dependency_map,
    )

    lines.append("def main() -> None:")
    lines.append("    raise SystemExit(\"Generated literal skeleton is not executable until formula lowering is implemented.\")")
    lines.append("")
    lines.append("if __name__ == \"__main__\":")
    lines.append("    main()")
    lines.append("")

    output.write_text("\n".join(lines), encoding="utf-8", newline="\n")
    return {
        "workbook_path": str(workbook_path.resolve()),
        "output_path": str(output.resolve()),
        "formula_function_count": len(emitted_formula_keys),
        "calc_order_count": len(calc_order),
        "cycle_group_count": len(cycle_groups),
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Emit deterministic literal Python skeleton from planner artifacts."
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
    parser.add_argument(
        "--output",
        default=None,
        help="Optional output .py path. Defaults to artifacts/<workbook>/generated/<workbook>_literal.py",
    )
    parser.add_argument(
        "--limit-formulas",
        type=int,
        default=None,
        help="Optional limit for emitted acyclic formulas (debug helper).",
    )
    args = parser.parse_args()

    result = emit_literal_skeleton(
        excel_file=args.excel_file,
        artifacts_root=args.artifacts_root,
        output_path=args.output,
        limit_formulas=args.limit_formulas,
    )

    print(f"Workbook: {result['workbook_path']}")
    print(f"Output script: {result['output_path']}")
    print(f"Formula functions: {result['formula_function_count']}")
    print(f"Acyclic formulas emitted: {result['calc_order_count']}")
    print(f"Cycle groups: {result['cycle_group_count']}")


if __name__ == "__main__":
    main()
