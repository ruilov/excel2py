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


def _is_number(value: object) -> bool:
    return isinstance(value, Real) and not isinstance(value, bool)


def _values_close(
    left: object,
    right: object,
    rel_tol: float = 1e-9,
    abs_tol: float = 1e-9,
) -> bool:
    if _is_number(left) and _is_number(right):
        return math.isclose(float(left), float(right), rel_tol=rel_tol, abs_tol=abs_tol)
    return left == right


def _load_generated_module(script_path: Path):
    module_name = f"excel2py_generated_{script_path.stem}_{uuid.uuid4().hex}"
    spec = importlib.util.spec_from_file_location(module_name, script_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Unable to import generated script: {script_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _load_dataset(raw_dir: Path) -> dict[str, object]:
    metadata = json.loads((raw_dir / "workbook_meta.json").read_text(encoding="utf-8"))
    fields = metadata["cells_jsonl_record_fields"]
    field_index = {name: idx for idx, name in enumerate(fields)}

    sheet_idx_idx = field_index["sheet_idx"]
    addr_idx = field_index["addr"]
    formula_idx = field_index["formula"]
    value_idx = field_index["value"]

    inputs: dict[tuple[int, str], object] = {}
    expected: dict[tuple[int, str], object] = {}

    with (raw_dir / "cells.jsonl").open("r", encoding="utf-8") as handle:
        for line in handle:
            row = json.loads(line)
            key = (row[sheet_idx_idx], normalize_addr(str(row[addr_idx])))
            formula = row[formula_idx]
            value = row[value_idx]

            if formula is None:
                inputs[key] = value
            else:
                expected[key] = value

    return {"inputs": inputs, "expected": expected}


def _load_cycle_keys(derived_dir: Path) -> list[tuple[int, str]]:
    payload = json.loads((derived_dir / "calc_order.json").read_text(encoding="utf-8"))
    cycle_keys: list[tuple[int, str]] = []
    for group in payload.get("cycles", []):
        for sheet_idx, addr in group:
            cycle_keys.append((int(sheet_idx), normalize_addr(str(addr))))
    return cycle_keys


def test_circular_convergence_and_iteration_parameters() -> None:
    excel_file = DEFAULT_WORKBOOK
    artifacts_root = DEFAULT_ARTIFACTS_ROOT

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
    dataset = _load_dataset(raw_dir=raw_dir)
    cycle_keys = _load_cycle_keys(derived_dir=derived_dir)
    assert cycle_keys, "Expected at least one cycle group in this workbook."

    seeded_inputs = dict(dataset["inputs"])
    for key in cycle_keys:
        if key in dataset["expected"]:
            seeded_inputs[key] = dataset["expected"][key]

    generated_module = _load_generated_module(Path(translation_result["output_path"]))
    output_iter0 = generated_module.run_model(
        dict(seeded_inputs),
        max_iterations=0,
        tolerance=1e-9,
    )
    output_iter1 = generated_module.run_model(
        dict(seeded_inputs),
        max_iterations=1,
        tolerance=1e-9,
    )
    output_iter10 = generated_module.run_model(
        dict(seeded_inputs),
        max_iterations=10,
        tolerance=1e-9,
    )

    cycle_changed_after_one_iter = [
        key
        for key in cycle_keys
        if not _values_close(output_iter0.get(key), output_iter1.get(key), rel_tol=1e-12, abs_tol=1e-12)
    ]
    assert cycle_changed_after_one_iter, "Cycle cells did not change between 0 and 1 iterations."

    cycle_not_converged_by_10 = [
        key
        for key in cycle_keys
        if not _values_close(output_iter1.get(key), output_iter10.get(key), rel_tol=1e-9, abs_tol=1e-9)
    ]
    assert not cycle_not_converged_by_10, (
        "Cycle cells are not stable between 1 and 10 iterations: "
        f"{cycle_not_converged_by_10[:10]}"
    )

