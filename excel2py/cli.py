import argparse

from excel2py.loader import (
    DEFAULT_ARTIFACTS_ROOT,
    DEFAULT_WORKBOOK,
    export_workbook_artifacts,
)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Export workbook metadata and cell contents into JSON artifacts."
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

    result = export_workbook_artifacts(
        excel_file=args.excel_file,
        artifacts_root=args.artifacts_root,
    )

    print(f"Workbook: {result['workbook_path']}")
    print(f"Artifact directory: {result['artifact_dir']}")
    print(f"Sheets: {result['sheet_count']}")
    print(f"Cells exported: {result['cell_count']}")
    print(f"Formula cells: {result['formula_cell_count']}")
    print(f"Manifest: {result['manifest_path']}")
    print(f"Metadata: {result['metadata_path']}")
    print(f"Cells JSONL: {result['cells_path']}")


if __name__ == "__main__":
    main()
