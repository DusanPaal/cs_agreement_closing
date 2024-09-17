# pylint: disable = C0103, E0110, E1101

"""
The module provides an interface for generating
processing output reports for end users.
"""

from pandas import ExcelWriter, DataFrame, Series
from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet
from xlsxwriter.format import Format

def _get_col_width(vals: Series, col_name: str, add_width: int = 0) -> int:
    """Returns the width of a column calculated as the maximum number
    of characters contained in column name and column values plus additional
    points provided with the 'add_width' argument (default 0 points).
    """

    alpha = 1 # additional offset factor

    if col_name.isnumeric():
        return 14 + add_width

    if col_name == "Agreement":
        return 11 + add_width

    if col_name in ("Valid_From", "Valid_To"):
        return 11 + add_width

    if col_name == "Payments":
        return 12 + add_width

    data_vals = vals.astype("string").dropna().str.len()
    data_vals = list(data_vals)
    data_vals.append(len(str(col_name)))
    width = max(data_vals) + alpha + add_width

    return width

def _write_to_excel(wrtr: ExcelWriter, data: DataFrame, sht_name: str) -> Workbook:
    """Writes data contained in a DataFrame objet to an excel file."""

    data.columns = data.columns.str.replace("_", " ", regex = False)
    data.to_excel(wrtr, index = False, sheet_name = sht_name)

    # replace spaces in column names back with underscores
    # for a better field manupulation further in the code
    data.columns = data.columns.str.replace(" ", "_", regex = False)

    return wrtr.book

def _generate_formats(report: Workbook) -> dict:
    """Generates formats to apply to the columns of the report sheet."""

    formats = {}

    formats["general"] = report.add_format({
        "align": "center"
    })

    formats["money"] = report.add_format({
        "num_format": "#,##0.00",
        "align": "center"
    })

    formats["header"] = report.add_format({
        "align": "center",
        "bg_color": "#F06B00",
        "font_color": "white",
        "bold": True
    })

    return formats

def _col_to_rng(data: DataFrame, first_col: str, last_col: str = None, row: int = -1) -> str:
    """Generates excel data range notation (e.g. 'A1:D1', 'B2:G2'). If 'last_col' is None,
    then only single-column range will be generated (e.g. 'A:A', 'B1:B1'). if 'row' is '-1',
    then the generated range will span all the column(s) rows (e.g. 'A:A', 'E:E').
    """

    if row < -1:
        raise ValueError(f"Argument 'row' has incorrect value: {row}")

    if isinstance(first_col, str):
        first_col_idx = data.columns.get_loc(first_col)
    elif isinstance(first_col, int):
        first_col_idx = first_col
    else:
        assert False, "Argument 'first_col' has invalid type!"

    first_col_idx += 1
    prim_lett_idx = first_col_idx // 26
    sec_lett_idx = first_col_idx % 26

    lett_a = chr(ord('@') + prim_lett_idx) if prim_lett_idx != 0 else ""
    lett_b = chr(ord('@') + sec_lett_idx) if sec_lett_idx != 0 else ""
    lett = "".join([lett_a, lett_b])

    if last_col is None:
        last_lett = lett
    else:

        if isinstance(last_col, str):
            last_col_idx =  data.columns.get_loc(last_col)
        elif isinstance(last_col, int):
            last_col_idx = last_col
        else:
            assert False, "Argument 'last_col' has invalid type!"

        last_col_idx += 1
        prim_lett_idx = last_col_idx // 26
        sec_lett_idx = last_col_idx % 26

        lett_a = chr(ord('@') + prim_lett_idx) if prim_lett_idx != 0 else ""
        lett_b = chr(ord('@') + sec_lett_idx) if sec_lett_idx != 0 else ""
        last_lett = "".join([lett_a, lett_b])

    if row == -1:
        rng = ":".join([lett, last_lett])
    else:
        rng = ":".join([f"{lett}{row}", f"{last_lett}{row}"])

    return rng

def _format_header(data: DataFrame, sht: Worksheet, fmt: Format) -> None:
    """Applies visual formatting to the report header."""

    first_row = _col_to_rng(data,  data.columns[0], data.columns[-1], row = 1)
    sht.conditional_format(first_row, {"type": "no_errors", "format": fmt})

def _format_data(data: DataFrame, sht: Worksheet, formats: dict) -> None:
    """Applies column-specific visual formats to the report data."""

    for idx, col in enumerate(data.columns):

        col_width = _get_col_width(data[col], col)

        if col in ("Open_Value", "Open_Accruals"):
            col_fmt = formats["money"]
        else:
            col_fmt = formats["general"]

        # apply new column format
        sht.set_column(idx, idx, col_width, col_fmt)

def create(file_path: str, data: DataFrame, sht_name: str) -> None:
    """Generates an .xlsx report fom the outcome of the agreement closing.

    Params:
    -------
    file_path:
        Path to the report file.

    data:
        Data containing the result of agreement closing.

    sht_name:
        Name of the report sheet where data is written.
    """

    if not file_path.lower().endswith(".xlsx"):
        raise ValueError("A file path to an .xlsx file is expected!")

    # print all and cleared items to separate sheets of a workbook
    with ExcelWriter(file_path, engine = "xlsxwriter") as wrtr:

        report = _write_to_excel(wrtr, data, sht_name)
        sht = wrtr.sheets[sht_name]
        formats = _generate_formats(report)
        _format_header(data, sht, formats["header"])
        _format_data(data, sht, formats)
