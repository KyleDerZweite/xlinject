from __future__ import annotations


def column_name_to_index(column_name: str) -> int:
    name = column_name.strip().upper()
    if not name.isalpha():
        raise ValueError(f"Invalid column name: {column_name}")

    result = 0
    for character in name:
        result = result * 26 + (ord(character) - ord("A") + 1)
    return result


def column_index_to_name(index: int) -> str:
    if index <= 0:
        raise ValueError("Column index must be positive")

    chars: list[str] = []
    current = index
    while current > 0:
        current, remainder = divmod(current - 1, 26)
        chars.append(chr(ord("A") + remainder))
    return "".join(reversed(chars))


def split_cell_reference(cell_ref: str) -> tuple[str, int]:
    cleaned = cell_ref.strip().upper()
    if not cleaned:
        raise ValueError("Cell reference cannot be empty")

    i = 0
    while i < len(cleaned) and cleaned[i].isalpha():
        i += 1

    column = cleaned[:i]
    row_text = cleaned[i:]

    if not column or not row_text or not row_text.isdigit():
        raise ValueError(f"Invalid cell reference: {cell_ref}")

    row = int(row_text)
    if row <= 0:
        raise ValueError(f"Invalid row in cell reference: {cell_ref}")

    return column, row


def build_cell_reference(column_name: str, row: int) -> str:
    if row <= 0:
        raise ValueError("Row must be positive")
    return f"{column_name.strip().upper()}{row}"


def parse_single_column_range(range_ref: str) -> tuple[str, int, int]:
    cleaned = range_ref.strip().upper()
    if ":" not in cleaned:
        column, row = split_cell_reference(cleaned)
        return column, row, row

    start_ref, end_ref = cleaned.split(":", 1)
    start_column, start_row = split_cell_reference(start_ref)
    end_column, end_row = split_cell_reference(end_ref)

    if start_column != end_column:
        raise ValueError(
            f"Range must stay in one column for this operation: {range_ref}"
        )

    if end_row < start_row:
        raise ValueError(f"Range end row must be >= start row: {range_ref}")

    return start_column, start_row, end_row
