from importlib.metadata import PackageNotFoundError, version

from xlinject.injector import (
	ReplaceReport,
	ValidationRule,
	WriteReport,
	extract_validation_rules,
	replace_sentinel_in_column_range,
	validate_cell_values,
	write_cells,
	write_numeric_cells,
)
from xlinject.highlevel import (
	apply_recalc_policy,
	build_column_cell_map,
	inject_cells,
	inject_cells_mixed,
	merge_cell_maps,
	normalize_numeric_value,
	remove_calc_chain,
	to_excel_serial,
)

try:
	__version__ = version("xlinject")
except PackageNotFoundError:
	__version__ = "0.0.0"

__all__ = [
	"ReplaceReport",
	"ValidationRule",
	"WriteReport",
	"extract_validation_rules",
	"replace_sentinel_in_column_range",
	"validate_cell_values",
	"write_cells",
	"write_numeric_cells",
	"build_column_cell_map",
	"inject_cells",
	"inject_cells_mixed",
	"apply_recalc_policy",
	"merge_cell_maps",
	"normalize_numeric_value",
	"remove_calc_chain",
	"to_excel_serial",
	"__version__",
]
