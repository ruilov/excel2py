import argparse
import sys

from excel2py.loader import (
    DEFAULT_ARTIFACTS_ROOT,
    DEFAULT_WORKBOOK,
    export_workbook_artifacts,
)
from excel2py.refactor_ai import prepare_ai_refactor_inputs


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Excel2Py CLI."
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    raw_parser = subparsers.add_parser(
        "raw",
        help="Export workbook metadata and cell contents into JSON artifacts.",
    )
    raw_parser.add_argument(
        "excel_file",
        nargs="?",
        default=str(DEFAULT_WORKBOOK),
        help=f"Path (or base name) of the workbook. Default: {DEFAULT_WORKBOOK.name}",
    )
    raw_parser.add_argument(
        "--artifacts-root",
        default=str(DEFAULT_ARTIFACTS_ROOT),
        help=f"Root folder for generated artifacts. Default: {DEFAULT_ARTIFACTS_ROOT}",
    )
    raw_parser.set_defaults(handler=_run_raw)

    ai_parser = subparsers.add_parser(
        "prepare-ai",
        help="Generate deterministic AI refactor context + prompt packets.",
    )
    ai_parser.add_argument(
        "excel_file",
        nargs="?",
        default=str(DEFAULT_WORKBOOK),
        help=f"Path (or base name) of workbook. Default: {DEFAULT_WORKBOOK.name}",
    )
    ai_parser.add_argument(
        "--artifacts-root",
        default=str(DEFAULT_ARTIFACTS_ROOT),
        help=f"Root folder for generated artifacts. Default: {DEFAULT_ARTIFACTS_ROOT}",
    )
    ai_parser.add_argument(
        "--max-cluster-prompts",
        type=int,
        default=200,
        help="Maximum number of per-cluster prompt files to emit.",
    )
    ai_parser.add_argument(
        "--include-singletons",
        action="store_true",
        help="Include single-cell clusters when emitting per-cluster prompts.",
    )
    ai_parser.set_defaults(handler=_run_prepare_ai)

    return parser


def _run_raw(args: argparse.Namespace) -> None:
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


def _run_prepare_ai(args: argparse.Namespace) -> None:
    result = prepare_ai_refactor_inputs(
        excel_file=args.excel_file,
        artifacts_root=args.artifacts_root,
        max_cluster_prompts=args.max_cluster_prompts,
        include_singletons=args.include_singletons,
    )

    print(f"Workbook: {result['workbook_path']}")
    print(f"AI output dir: {result['ai_dir']}")
    print(f"Global context: {result['global_context_path']}")
    print(f"Clusters JSON: {result['clusters_path']}")
    print(f"Cluster packets JSONL: {result['cluster_packets_path']}")
    print(f"Clusters: {result['cluster_count']}")
    print(f"Repeated clusters: {result['repeated_cluster_count']}")
    print(f"Global prompt: {result['prompt_outputs']['global_prompt_path']}")
    print(f"Cluster prompts: {result['prompt_outputs']['cluster_prompt_count']}")
    print(f"Plan prompt: {result['prompt_outputs']['plan_prompt_path']}")
    print(f"Codegen prompt: {result['prompt_outputs']['codegen_prompt_path']}")
    print(f"Self-check prompt: {result['prompt_outputs']['self_check_prompt_path']}")


def main(argv: list[str] | None = None) -> None:
    if argv is None:
        argv = sys.argv[1:]

    # Backward compatibility:
    # `python -m excel2py.cli excel_model.xlsx` => `raw excel_model.xlsx`
    if len(argv) == 0:
        argv = ["raw"]
    elif argv[0] not in {"raw", "prepare-ai"}:
        argv = ["raw", *argv]

    parser = _build_parser()
    args = parser.parse_args(argv)
    args.handler(args)


if __name__ == "__main__":
    main()
