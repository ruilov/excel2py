import lark
import re


QUALIFIED_REF_SPLIT = re.compile(r"^(?P<sheet>.+)!(?P<addr>.+)$")


FORMULA_GRAMMAR = r"""
?formula: "=" expr
       | expr

?expr: concat_expr

?concat_expr: comparison_expr (CONCAT_OP comparison_expr)*

?comparison_expr: additive_expr (COMPOP additive_expr)*

?additive_expr: multiplicative_expr (ADD_OP multiplicative_expr)*

?multiplicative_expr: power_expr (MUL_OP power_expr)*

?power_expr: unary_expr (POW_OP unary_expr)*

?unary_expr: UNARY_OP unary_expr
           | postfix_expr

?postfix_expr: primary PERCENT_OP
             | primary

?primary: function_call
        | reference
        | array_literal
        | NUMBER
        | BOOL
        | ERROR
        | STRING
        | NAME
        | "(" expr ")"

function_call: NAME "(" [arg_list] ")"

arg_list: arg (ARG_SEP arg)*
?arg: expr?

array_literal: "{" [array_rows] "}"
array_rows: array_row (";" array_row)*
array_row: expr? ("," expr?)*

reference: REF_QUALIFIED
         | REF_CELL_RANGE
         | REF_COL_RANGE
         | REF_ROW_RANGE
         | REF_CELL

COMPOP: "<>" | "<=" | ">=" | "<" | ">" | "="
CONCAT_OP: "&"
ADD_OP: "+" | "-"
MUL_OP: "*" | "/"
POW_OP: "^"
UNARY_OP: "+" | "-"
PERCENT_OP: "%"
ARG_SEP: "," | ";"

BOOL.2: "TRUE"i | "FALSE"i
ERROR.2: /#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|CALC!|SPILL!|FIELD!|GETTING_DATA|CONNECT!|BLOCKED!|UNKNOWN!|BUSY!)/

REF_QUALIFIED.3: /(?:'([^']|'')+'|(?:\[[^\]]+\])?[A-Za-z0-9_. \-]+)!\$?(?:[A-Za-z]{1,3}\$?\d+|[A-Za-z]{1,3}|\d+)(?::\$?(?:[A-Za-z]{1,3}\$?\d+|[A-Za-z]{1,3}|\d+))?/
REF_CELL_RANGE.2: /\$?[A-Za-z]{1,3}\$?\d+:\$?[A-Za-z]{1,3}\$?\d+/
REF_COL_RANGE.2: /\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}/
REF_ROW_RANGE.2: /\$?\d+:\$?\d+/
REF_CELL.2: /\$?[A-Za-z]{1,3}\$?\d+/

NAME: /[A-Za-z_\\][A-Za-z0-9_.\\]*/
NUMBER: /(?:\d+\.\d*|\d+|\.\d+)(?:[Ee][+\-]?\d+)?/
STRING: /"([^"]|"")*"/

%import common.WS_INLINE
%ignore WS_INLINE
"""


_PARSER = lark.Lark(
    FORMULA_GRAMMAR,
    start="formula",
    parser="earley",
    lexer="dynamic",
)


def parse_formula(formula: str) -> lark.Tree:
    return _PARSER.parse(formula.strip())


def try_parse_formula(formula: str) -> tuple[lark.Tree | None, str | None]:
    try:
        return parse_formula(formula), None
    except lark.exceptions.LarkError as exc:
        return None, str(exc)


def _normalize_addr(addr: str) -> str:
    return addr.replace("$", "").upper()


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


class _DependencyExtractor:
    def __init__(
        self,
        current_sheet_idx: int,
        sheet_idx_by_name: dict[str, int],
        defined_name_by_upper: dict[str, str],
    ) -> None:
        self.current_sheet_idx = current_sheet_idx
        self.sheet_idx_by_name = sheet_idx_by_name
        self.defined_name_by_upper = defined_name_by_upper
        self.dependencies: list[list[object]] = []
        self.seen: set[tuple[object, ...]] = set()

    def add_ref(self, sheet_idx: int, addr: str) -> None:
        normalized_addr = _normalize_addr(addr)
        key = ("ref", sheet_idx, normalized_addr)
        if key in self.seen:
            return
        self.seen.add(key)
        self.dependencies.append([sheet_idx, normalized_addr])

    def add_external_ref(self, ref_text: str) -> None:
        key = ("ext", ref_text)
        if key in self.seen:
            return
        self.seen.add(key)
        self.dependencies.append(["ext", ref_text])

    def add_named_range(self, range_name: str) -> None:
        key = ("name", range_name)
        if key in self.seen:
            return
        self.seen.add(key)
        self.dependencies.append(["name", range_name])

    def handle_reference_token(self, token: lark.Token) -> None:
        if token.type == "REF_QUALIFIED":
            sheet_text, addr_text = _split_qualified_ref(token.value)
            if sheet_text is None or addr_text is None:
                self.add_external_ref(token.value)
                return

            target_sheet_idx = _sheet_idx_from_token(sheet_text, self.sheet_idx_by_name)
            normalized_addr = _normalize_addr(addr_text)
            if target_sheet_idx is None:
                normalized_sheet = _normalize_sheet_token(sheet_text)
                self.add_external_ref(f"{normalized_sheet}!{normalized_addr}")
            else:
                self.add_ref(target_sheet_idx, normalized_addr)
            return

        if token.type in {"REF_CELL", "REF_CELL_RANGE", "REF_COL_RANGE", "REF_ROW_RANGE"}:
            self.add_ref(self.current_sheet_idx, token.value)

    def visit(self, node: lark.Tree | lark.Token) -> None:
        if isinstance(node, lark.Tree):
            if node.data == "reference":
                for child in node.children:
                    if isinstance(child, lark.Token):
                        self.handle_reference_token(child)
                return

            if node.data == "function_call":
                for idx, child in enumerate(node.children):
                    if idx == 0 and isinstance(child, lark.Token) and child.type == "NAME":
                        continue
                    self.visit(child)
                return

            for child in node.children:
                self.visit(child)
            return

        if isinstance(node, lark.Token) and node.type == "NAME":
            range_name = self.defined_name_by_upper.get(node.value.upper())
            if range_name is not None:
                self.add_named_range(range_name)


def extract_dependencies_from_tree(
    tree: lark.Tree,
    current_sheet_idx: int,
    sheet_idx_by_name: dict[str, int],
    defined_name_by_upper: dict[str, str],
) -> list[list[object]]:
    extractor = _DependencyExtractor(
        current_sheet_idx=current_sheet_idx,
        sheet_idx_by_name=sheet_idx_by_name,
        defined_name_by_upper=defined_name_by_upper,
    )
    extractor.visit(tree)
    return extractor.dependencies


def extract_dependencies_from_formula(
    formula: str,
    current_sheet_idx: int,
    sheet_idx_by_name: dict[str, int],
    defined_name_by_upper: dict[str, str],
) -> tuple[list[list[object]] | None, str | None]:
    tree, parse_error = try_parse_formula(formula)
    if tree is None:
        return None, parse_error

    dependencies = extract_dependencies_from_tree(
        tree=tree,
        current_sheet_idx=current_sheet_idx,
        sheet_idx_by_name=sheet_idx_by_name,
        defined_name_by_upper=defined_name_by_upper,
    )
    return dependencies, None
