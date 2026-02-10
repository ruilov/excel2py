from openpyxl.utils import cell as xl_cell


def normalize_addr(addr: str) -> str:
    return addr.replace("$", "").upper()


def cell_key(sheet_idx: int, addr: str) -> tuple[int, str]:
    return (sheet_idx, normalize_addr(addr))


def get_cell(cells: dict[tuple[int, str], object], sheet_idx: int, addr: str) -> object:
    return cells.get(cell_key(sheet_idx, addr))


def set_cell(cells: dict[tuple[int, str], object], sheet_idx: int, addr: str, value: object) -> None:
    cells[cell_key(sheet_idx, addr)] = value


def _addr_kind(addr: str) -> str:
    normalized = normalize_addr(addr)
    if ":" not in normalized:
        return "single"

    left, right = normalized.split(":", 1)
    if left.isdigit() and right.isdigit():
        return "row_range"
    if left.isalpha() and right.isalpha():
        return "col_range"
    return "rect_range"


def get_range(
    cells: dict[tuple[int, str], object],
    sheet_idx: int,
    addr: str,
    sheet_dimension: str | None = None,
) -> list[object] | list[list[object]]:
    normalized = normalize_addr(addr)
    kind = _addr_kind(normalized)

    if kind == "single":
        return [get_cell(cells, sheet_idx, normalized)]

    if kind == "rect_range":
        min_col, min_row, max_col, max_row = xl_cell.range_boundaries(normalized)
        rows: list[list[object]] = []
        for row in range(min_row, max_row + 1):
            current_row: list[object] = []
            for col in range(min_col, max_col + 1):
                coord = f"{xl_cell.get_column_letter(col)}{row}"
                current_row.append(get_cell(cells, sheet_idx, coord))
            rows.append(current_row)

        if len(rows) == 1:
            return rows[0]
        if all(len(row) == 1 for row in rows):
            return [row[0] for row in rows]
        return rows

    if kind == "row_range":
        if sheet_dimension is None:
            raise ValueError("sheet_dimension is required for row ranges")
        left, right = normalized.split(":", 1)
        min_row = int(left)
        max_row = int(right)
        if min_row > max_row:
            min_row, max_row = max_row, min_row

        min_col, _min_dim_row, max_col, _max_dim_row = xl_cell.range_boundaries(sheet_dimension)
        rows: list[list[object]] = []
        for row in range(min_row, max_row + 1):
            current_row: list[object] = []
            for col in range(min_col, max_col + 1):
                coord = f"{xl_cell.get_column_letter(col)}{row}"
                current_row.append(get_cell(cells, sheet_idx, coord))
            rows.append(current_row)
        return rows

    if kind == "col_range":
        if sheet_dimension is None:
            raise ValueError("sheet_dimension is required for column ranges")
        left, right = normalized.split(":", 1)
        min_col = xl_cell.column_index_from_string(left)
        max_col = xl_cell.column_index_from_string(right)
        if min_col > max_col:
            min_col, max_col = max_col, min_col

        _min_dim_col, min_row, _max_dim_col, max_row = xl_cell.range_boundaries(sheet_dimension)
        cols: list[list[object]] = []
        for col in range(min_col, max_col + 1):
            current_col: list[object] = []
            for row in range(min_row, max_row + 1):
                coord = f"{xl_cell.get_column_letter(col)}{row}"
                current_col.append(get_cell(cells, sheet_idx, coord))
            cols.append(current_col)
        return cols

    raise ValueError(f"Unsupported address reference: {addr}")


def resolve_dependency(
    cells: dict[tuple[int, str], object],
    dependency: list[object],
    sheet_dimensions: dict[int, str] | None = None,
) -> object:
    if len(dependency) == 2 and isinstance(dependency[0], int):
        sheet_idx = dependency[0]
        addr = str(dependency[1])
        if ":" in addr:
            sheet_dimension = None if sheet_dimensions is None else sheet_dimensions.get(sheet_idx)
            return get_range(cells, sheet_idx, addr, sheet_dimension=sheet_dimension)
        return get_cell(cells, sheet_idx, addr)

    if len(dependency) == 2 and dependency[0] == "name":
        raise NotImplementedError(f"Named range dependency not implemented: {dependency[1]}")

    if len(dependency) == 2 and dependency[0] == "ext":
        raise NotImplementedError(f"External reference dependency not implemented: {dependency[1]}")

    raise ValueError(f"Unsupported dependency shape: {dependency}")


def xl_if(condition: object, value_if_true: object, value_if_false: object = False) -> object:
    return value_if_true if bool(condition) else value_if_false


def xl_iferror(value: object, fallback: object) -> object:
    return fallback if isinstance(value, Exception) else value


def xl_index(array: list[object] | list[list[object]], row_num: int, col_num: int | None = None) -> object:
    if row_num is None:
        raise ValueError("row_num cannot be None")
    row_index = int(row_num) - 1

    if isinstance(array, list) and array and isinstance(array[0], list):
        col_index = 0 if col_num is None else int(col_num) - 1
        return array[row_index][col_index]

    if isinstance(array, list):
        return array[row_index]

    raise ValueError("Unsupported array shape for INDEX")


def xl_match(lookup_value: object, lookup_array: list[object], match_type: int = 1) -> int:
    if match_type == 0:
        for idx, current in enumerate(lookup_array, start=1):
            if current == lookup_value:
                return idx
        raise ValueError("MATCH exact lookup failed")
    raise NotImplementedError("Only MATCH(..., ..., 0) is implemented in v0")
