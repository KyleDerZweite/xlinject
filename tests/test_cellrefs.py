from xlinject.cellrefs import (
    build_cell_reference,
    column_index_to_name,
    column_name_to_index,
    parse_single_column_range,
    split_cell_reference,
)


def test_column_name_to_index_roundtrip() -> None:
    assert column_name_to_index("A") == 1
    assert column_name_to_index("Z") == 26
    assert column_name_to_index("AA") == 27
    assert column_index_to_name(1) == "A"
    assert column_index_to_name(26) == "Z"
    assert column_index_to_name(27) == "AA"


def test_split_and_build_cell_reference() -> None:
    column, row = split_cell_reference("C45")
    assert column == "C"
    assert row == 45
    assert build_cell_reference(column, row) == "C45"


def test_parse_single_column_range() -> None:
    assert parse_single_column_range("B2:B121") == ("B", 2, 121)
    assert parse_single_column_range("D10") == ("D", 10, 10)
