import argparse
import json
import re
from pathlib import Path

from lark import Token
from lark import Tree

from excel2py.formula_parser import parse_formula
from excel2py.loader import *


QUALIFIED_REF_SPLIT = re.compile(r"^(?P<sheet>.+)!(?P<addr>.+)$")


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


def _normalize_sheet_token(sheet_token: str) -> str:
    token = sheet_token.strip()
    if token.startswith("'") and token.endswith("'") and len(token) >= 2:
        token = token[1:-1].replace("''", "'")
    return token


def _sheet_idx_from_token(sheet_token: str, sheet_idx_by_name: dict[str, int]) -> int | None:
    normalized = _normalize_sheet_token(sheet_token)
    if normalized in sheet_idx_by_name:
        return sheet_idx_by_name[normalized]

    without_book_prefix = re.sub(r"^\[[^\]]+\]", "", normalized)
    if without_book_prefix in sheet_idx_by_name:
        return sheet_idx_by_name[without_book_prefix]

    return None


def _split_qualified_ref(ref_text: str) -> tuple[str | None, str | None]:
    match = QUALIFIED_REF_SPLIT.match(ref_text)
    if match is None:
        return None, None
    return match.group("sheet"), match.group("addr")


def _dependency_key(dependency: list[object]) -> tuple[object, ...] | None:
    if len(dependency) == 2 and isinstance(dependency[0], int):
        return ("ref", dependency[0], _normalize_addr(str(dependency[1])))
    if len(dependency) == 2 and dependency[0] == "name":
        return ("name", str(dependency[1]))
    if len(dependency) == 2 and dependency[0] == "ext":
        return ("ext", str(dependency[1]))
    return None


def _string_literal_from_excel_token(token_value: str) -> str:
    unquoted = token_value[1:-1]
    return repr(unquoted.replace('""', '"'))


class _FormulaLowerer:
    def __init__(
        self,
        current_sheet_idx: int,
        sheet_idx_by_name: dict[str, int],
        defined_name_by_upper: dict[str, str],
        dep_arg_by_key: dict[tuple[object, ...], str],
    ) -> None:
        self.current_sheet_idx = current_sheet_idx
        self.sheet_idx_by_name = sheet_idx_by_name
        self.defined_name_by_upper = defined_name_by_upper
        self.dep_arg_by_key = dep_arg_by_key
        self.errors: list[str] = []

    def lower(self, node: Tree | Token) -> str:
        if isinstance(node, Token):
            return self._lower_token(node)

        data = node.data
        if data == "reference":
            return self._lower_reference(node)

        if data == "function_call":
            return self._lower_function_call(node)

        if data == "arg":
            if not node.children:
                return "None"
            return self.lower(node.children[0])

        if data == "arg_list":
            args = self._lower_arg_list(node)
            return "[" + ", ".join(args) + "]"

        if data == "concat_expr":
            return self._lower_binary_chain(
                node.children,
                combine=lambda left, op, right: f"xl_concat({left}, {right})",
            )

        if data == "comparison_expr":
            return self._lower_binary_chain(
                node.children,
                combine=lambda left, op, right: f"xl_compare({repr(op)}, {left}, {right})",
            )

        if data == "additive_expr":
            return self._lower_binary_chain(
                node.children,
                combine=lambda left, op, right: f"xl_add({left}, {right})"
                if op == "+"
                else f"xl_sub({left}, {right})",
            )

        if data == "multiplicative_expr":
            return self._lower_binary_chain(
                node.children,
                combine=lambda left, op, right: f"xl_mul({left}, {right})"
                if op == "*"
                else f"xl_div({left}, {right})",
            )

        if data == "power_expr":
            return self._lower_binary_chain(
                node.children,
                combine=lambda left, op, right: f"xl_pow({left}, {right})",
            )

        if data == "unary_expr":
            if len(node.children) == 2 and isinstance(node.children[0], Token):
                op = node.children[0].value
                inner = self.lower(node.children[1])
                if op == "+":
                    return f"xl_pos({inner})"
                return f"xl_neg({inner})"
            if len(node.children) == 1:
                return self.lower(node.children[0])

        if data == "postfix_expr":
            if len(node.children) == 2:
                value = self.lower(node.children[0])
                return f"xl_percent({value})"
            if len(node.children) == 1:
                return self.lower(node.children[0])

        if data == "array_literal":
            return self._lower_array_literal(node)

        if data == "array_rows":
            rows = [self._lower_array_row(child) for child in node.children if isinstance(child, Tree)]
            return "[" + ", ".join(rows) + "]"

        if data == "array_row":
            return self._lower_array_row(node)

        if len(node.children) == 1:
            return self.lower(node.children[0])

        self.errors.append(f"Unsupported AST node: {data}")
        return "None"

    def _lower_token(self, token: Token) -> str:
        if token.type == "NUMBER":
            return token.value

        if token.type == "STRING":
            return _string_literal_from_excel_token(token.value)

        if token.type == "BOOL":
            return "True" if token.value.upper() == "TRUE" else "False"

        if token.type == "ERROR":
            return f"xl_error({repr(token.value)})"

        if token.type == "NAME":
            range_name = self.defined_name_by_upper.get(token.value.upper())
            if range_name is None:
                self.errors.append(f"Unresolved name token: {token.value}")
                return "None"

            key = ("name", range_name)
            arg_name = self.dep_arg_by_key.get(key)
            if arg_name is None:
                self.errors.append(f"Named range missing from dependency list: {range_name}")
                return "None"
            return arg_name

        if token.type in {
            "COMPOP",
            "CONCAT_OP",
            "ADD_OP",
            "MUL_OP",
            "POW_OP",
            "UNARY_OP",
            "PERCENT_OP",
            "COMMA",
            "SEMI",
        }:
            return token.value

        self.errors.append(f"Unsupported token type: {token.type}")
        return "None"

    def _lower_reference(self, node: Tree) -> str:
        if not node.children or not isinstance(node.children[0], Token):
            self.errors.append("Malformed reference node")
            return "None"

        token = node.children[0]
        key: tuple[object, ...] | None = None

        if token.type == "REF_QUALIFIED":
            sheet_text, addr_text = _split_qualified_ref(token.value)
            if sheet_text is None or addr_text is None:
                key = ("ext", token.value)
            else:
                target_sheet_idx = _sheet_idx_from_token(sheet_text, self.sheet_idx_by_name)
                normalized_addr = _normalize_addr(addr_text)
                if target_sheet_idx is None:
                    normalized_sheet = _normalize_sheet_token(sheet_text)
                    key = ("ext", f"{normalized_sheet}!{normalized_addr}")
                else:
                    key = ("ref", target_sheet_idx, normalized_addr)

        elif token.type in {"REF_CELL", "REF_CELL_RANGE", "REF_COL_RANGE", "REF_ROW_RANGE"}:
            key = ("ref", self.current_sheet_idx, _normalize_addr(token.value))

        if key is None:
            self.errors.append(f"Unsupported reference token: {token.type}")
            return "None"

        arg_name = self.dep_arg_by_key.get(key)
        if arg_name is None:
            self.errors.append(f"Reference missing from dependency list: {token.value}")
            return "None"
        return arg_name

    def _lower_function_call(self, node: Tree) -> str:
        if not node.children or not isinstance(node.children[0], Token):
            self.errors.append("Malformed function_call node")
            return "None"

        function_name = node.children[0].value
        args: list[str] = []

        if len(node.children) > 1 and isinstance(node.children[1], Tree) and node.children[1].data == "arg_list":
            args = self._lower_arg_list(node.children[1])

        if args:
            return f"xl_call({repr(function_name)}, " + ", ".join(args) + ")"
        return f"xl_call({repr(function_name)})"

    def _lower_arg_list(self, arg_list: Tree) -> list[str]:
        args: list[str] = []
        for child in arg_list.children:
            if isinstance(child, Token) and child.type in {"COMMA", "SEMI"}:
                continue
            args.append(self.lower(child))
        return args

    def _lower_array_row(self, row_node: Tree) -> str:
        values = [self.lower(child) for child in row_node.children]
        return "[" + ", ".join(values) + "]"

    def _lower_array_literal(self, node: Tree) -> str:
        if not node.children:
            return "[]"

        rows_node = node.children[0]
        if not isinstance(rows_node, Tree) or rows_node.data != "array_rows":
            self.errors.append("Malformed array literal")
            return "[]"

        rows = [self._lower_array_row(child) for child in rows_node.children if isinstance(child, Tree)]
        if len(rows) == 1:
            # Keep single-row arrays simple and readable.
            return rows[0]
        return "[" + ", ".join(rows) + "]"

    def _lower_binary_chain(self, children: list[Tree | Token], combine) -> str:
        if not children:
            self.errors.append("Empty binary chain")
            return "None"

        left = self.lower(children[0])
        index = 1
        while index + 1 < len(children):
            op_token = children[index]
            right_node = children[index + 1]

            if not isinstance(op_token, Token):
                self.errors.append("Malformed binary chain operator")
                return "None"

            right = self.lower(right_node)
            left = combine(left, op_token.value, right)
            index += 2

        if index != len(children):
            self.errors.append("Malformed binary chain tail")

        return left


def _lower_formula_expression(
    formula: str,
    current_sheet_idx: int,
    sheet_idx_by_name: dict[str, int],
    defined_name_by_upper: dict[str, str],
    dep_arg_by_key: dict[tuple[object, ...], str],
) -> tuple[str | None, str | None]:
    try:
        tree = parse_formula(formula)
    except Exception as exc:  # noqa: BLE001
        return None, f"Formula parse failed: {exc}"

    lowerer = _FormulaLowerer(
        current_sheet_idx=current_sheet_idx,
        sheet_idx_by_name=sheet_idx_by_name,
        defined_name_by_upper=defined_name_by_upper,
        dep_arg_by_key=dep_arg_by_key,
    )
    expression = lowerer.lower(tree)

    if lowerer.errors:
        return None, "; ".join(lowerer.errors)

    return expression, None


def _emit_function(
    lines: list[str],
    sheet_names: list[str],
    sheet_idx_by_name: dict[str, int],
    defined_name_by_upper: dict[str, str],
    sheet_idx: int,
    addr: str,
    formula: str,
    dependencies: list[list[object]],
) -> bool:
    func_name = _cell_func_name(sheet_idx, addr)
    arg_names = [_dep_arg_name(dep, idx) for idx, dep in enumerate(dependencies)]
    args_src = ", ".join(arg_names)

    dep_arg_by_key: dict[tuple[object, ...], str] = {}
    for idx, dependency in enumerate(dependencies):
        dep_key = _dependency_key(dependency)
        if dep_key is not None and dep_key not in dep_arg_by_key:
            dep_arg_by_key[dep_key] = arg_names[idx]

    lowered_expr, lower_error = _lower_formula_expression(
        formula=formula,
        current_sheet_idx=sheet_idx,
        sheet_idx_by_name=sheet_idx_by_name,
        defined_name_by_upper=defined_name_by_upper,
        dep_arg_by_key=dep_arg_by_key,
    )

    lines.append(f"def {func_name}({args_src}) -> object:")
    lines.append(f"    # from {sheet_names[sheet_idx]}!{addr}")
    lines.append(f"    # excel: {formula}")

    if lowered_expr is None:
        lines.append(f"    raise NotImplementedError({repr(f'Formula lowering failed for {sheet_names[sheet_idx]}!{addr}: {lower_error}')} )")
        lines.append("")
        return False

    lines.append(f"    return {lowered_expr}")
    lines.append("")
    return True


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
    sheet_idx_by_name = {name: idx for idx, name in enumerate(metadata["sheet_names"])}
    defined_name_by_upper = {
        item["name"].upper(): item["name"] for item in metadata.get("named_ranges", [])
    }

    output = Path(output_path) if output_path is not None else paths["output_path"]
    lines: list[str] = []
    lines.append("# Auto-generated deterministic literal Python from planner artifacts.")
    lines.append("# This file is intentionally explicit for later AI refactoring.")
    lines.append("from excel2py.runtime_helpers import *")
    lines.append("")
    lines.append(f"SHEET_NAMES = {repr(metadata['sheet_names'])}")
    lines.append(f"SHEET_DIMENSIONS = {repr(sheet_dimensions)}")
    lines.append(f"CALC_ORDER = {repr(calc_order)}")
    lines.append(f"CYCLE_GROUPS = {repr(cycle_groups)}")
    lines.append("")

    emitted_formula_keys: set[tuple[int, str]] = set()
    lowered_formula_count = 0
    lowering_failed_count = 0
    for sheet_idx, addr in calc_order:
        key = (sheet_idx, addr)
        emitted_formula_keys.add(key)
        formula = formula_map.get(key)
        if formula is None:
            continue
        dependencies = dependency_map.get(key, [])
        lowered = _emit_function(
            lines=lines,
            sheet_names=metadata["sheet_names"],
            sheet_idx_by_name=sheet_idx_by_name,
            defined_name_by_upper=defined_name_by_upper,
            sheet_idx=sheet_idx,
            addr=addr,
            formula=formula,
            dependencies=dependencies,
        )
        if lowered:
            lowered_formula_count += 1
        else:
            lowering_failed_count += 1

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
            lowered = _emit_function(
                lines=lines,
                sheet_names=metadata["sheet_names"],
                sheet_idx_by_name=sheet_idx_by_name,
                defined_name_by_upper=defined_name_by_upper,
                sheet_idx=sheet_idx,
                addr=addr,
                formula=formula,
                dependencies=dependencies,
            )
            if lowered:
                lowered_formula_count += 1
            else:
                lowering_failed_count += 1

    _emit_run_model(
        lines=lines,
        calc_order=calc_order,
        cycle_groups=cycle_groups,
        dependency_map=dependency_map,
    )

    lines.append("def main() -> None:")
    lines.append("    print('Generated module loaded. Call run_model(inputs) to evaluate.')")
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
        "lowered_formula_count": lowered_formula_count,
        "lowering_failed_count": lowering_failed_count,
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Emit deterministic literal Python from planner artifacts."
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
    print(f"Lowered formulas: {result['lowered_formula_count']}")
    print(f"Lowering failures: {result['lowering_failed_count']}")
    print(f"Acyclic formulas emitted: {result['calc_order_count']}")
    print(f"Cycle groups: {result['cycle_group_count']}")


if __name__ == "__main__":
    main()
