import argparse
import json
from pathlib import Path

from excel2py.loader import *


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


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Build derived/formulas.jsonl from raw workbook artifacts."
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

    print(f"Workbook: {result['workbook_path']}")
    print(f"Raw directory: {result['raw_dir']}")
    print(f"Derived directory: {result['derived_dir']}")
    print(f"Formulas JSONL: {result['formulas_path']}")
    print(f"Formula rows: {result['formula_count']}")


if __name__ == "__main__":
    main()
