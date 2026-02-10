import importlib.util
import json
import math
import uuid
from numbers import Real
from pathlib import Path

from excel2py.loader import *
from excel2py.planner import *
from excel2py.runtime_helpers import *
from excel2py.translator import *


FLOAT_REL_TOL = 1e-6
FLOAT_ABS_TOL = 1e-6
MAX_DIAGNOSTIC_ROWS = 40


def _is_number(value: object) -> bool:
    return isinstance(value, Real) and not isinstance(value, bool)


def _load_generated_module(script_path: Path):
    module_name = f"excel2py_generated_{script_path.stem}_{uuid.uuid4().hex}"
    spec = importlib.util.spec_from_file_location(module_name, script_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Unable to import generated script: {script_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _load_parity_dataset(raw_dir: Path) -> dict[str, object]:
    metadata = json.loads((raw_dir / "workbook_meta.json").read_text(encoding="utf-8"))
    fields = metadata["cells_jsonl_record_fields"]
    field_index = {name: idx for idx, name in enumerate(fields)}

    sheet_idx_idx = field_index["sheet_idx"]
    addr_idx = field_index["addr"]
    formula_idx = field_index["formula"]
    value_idx = field_index["value"]

    inputs: dict[tuple[int, str], object] = {}
    expected: dict[tuple[int, str], object] = {}
    formulas: dict[tuple[int, str], str] = {}

    with (raw_dir / "cells.jsonl").open("r", encoding="utf-8") as handle:
        for line in handle:
            row = json.loads(line)
            key = (row[sheet_idx_idx], normalize_addr(str(row[addr_idx])))
            formula = row[formula_idx]
            value = row[value_idx]

            if formula is None:
                inputs[key] = value
                continue

            expected[key] = value
            formulas[key] = str(formula)

    return {
        "sheet_names": metadata["sheet_names"],
        "inputs": inputs,
        "expected": expected,
        "formulas": formulas,
    }


def _load_cycle_keys(derived_dir: Path) -> list[tuple[int, str]]:
    payload = json.loads((derived_dir / "calc_order.json").read_text(encoding="utf-8"))
    cycle_keys: list[tuple[int, str]] = []
    for group in payload.get("cycles", []):
        for sheet_idx, addr in group:
            cycle_keys.append((int(sheet_idx), normalize_addr(str(addr))))
    return cycle_keys


def _values_match(expected: object, actual: object) -> tuple[bool, str]:
    if _is_number(expected) and _is_number(actual):
        expected_value = float(expected)
        actual_value = float(actual)
        if math.isclose(
            expected_value,
            actual_value,
            rel_tol=FLOAT_REL_TOL,
            abs_tol=FLOAT_ABS_TOL,
        ):
            return True, ""
        return (
            False,
            f"numeric mismatch (expected={expected_value}, actual={actual_value}, "
            f"abs_diff={abs(expected_value - actual_value)})",
        )

    if expected == actual:
        return True, ""

    return False, f"value mismatch (expected={expected!r}, actual={actual!r})"


def _format_diagnostics(
    missing_keys: list[tuple[int, str]],
    mismatches: list[dict[str, object]],
    sheet_names: list[str],
) -> str:
    lines = []
    lines.append(
        f"Parity check failed: missing={len(missing_keys)}, mismatched={len(mismatches)}"
    )

    if missing_keys:
        lines.append("Missing calculated cells:")
        for sheet_idx, addr in missing_keys[:MAX_DIAGNOSTIC_ROWS]:
            lines.append(f"- {sheet_names[sheet_idx]}!{addr}")

    if mismatches:
        lines.append("Mismatched cells:")
        for item in mismatches[:MAX_DIAGNOSTIC_ROWS]:
            lines.append(
                "- "
                f"{item['sheet_name']}!{item['addr']} | "
                f"{item['reason']} | "
                f"formula={item['formula']}"
            )

    return "\n".join(lines)


def test_full_cell_parity() -> None:
    excel_file = DEFAULT_WORKBOOK
    artifacts_root = DEFAULT_ARTIFACTS_ROOT

    print("[parity] Preparing artifacts...", flush=True)
    export_workbook_artifacts(excel_file=excel_file, artifacts_root=artifacts_root)
    export_formula_rows(excel_file=excel_file, artifacts_root=artifacts_root)
    dep_result = export_formula_dependencies(excel_file=excel_file, artifacts_root=artifacts_root)
    build_calc_order(
        excel_file=excel_file,
        artifacts_root=artifacts_root,
        parse_diagnostics=dep_result,
    )
    translation_result = emit_literal_skeleton(
        excel_file=excel_file,
        artifacts_root=artifacts_root,
    )

    workbook_key = Path(excel_file).stem
    raw_dir = Path(artifacts_root) / workbook_key / "raw"
    derived_dir = Path(artifacts_root) / workbook_key / "derived"
    dataset = _load_parity_dataset(raw_dir=raw_dir)
    cycle_keys = _load_cycle_keys(derived_dir=derived_dir)
    print(
        f"[parity] Loaded inputs={len(dataset['inputs'])}, expected_formula_cells={len(dataset['expected'])}",
        flush=True,
    )

    for key in cycle_keys:
        if key in dataset["expected"]:
            dataset["inputs"][key] = dataset["expected"][key]

    print("[parity] Running generated model...", flush=True)
    generated_module = _load_generated_module(Path(translation_result["output_path"]))
    calculated = generated_module.run_model(dataset["inputs"])

    print("[parity] Comparing results...", flush=True)
    missing_keys: list[tuple[int, str]] = []
    mismatches: list[dict[str, object]] = []
    for key, expected_value in dataset["expected"].items():
        if key not in calculated:
            missing_keys.append(key)
            continue

        actual_value = calculated[key]
        matches, reason = _values_match(expected_value, actual_value)
        if matches:
            continue

        sheet_idx, addr = key
        mismatches.append(
            {
                "sheet_name": dataset["sheet_names"][sheet_idx],
                "addr": addr,
                "formula": dataset["formulas"].get(key, ""),
                "reason": reason,
            }
        )

    print(
        f"[parity] Done. missing={len(missing_keys)} mismatched={len(mismatches)}",
        flush=True,
    )
    assert not missing_keys and not mismatches, _format_diagnostics(
        missing_keys=missing_keys,
        mismatches=mismatches,
        sheet_names=dataset["sheet_names"],
    )
