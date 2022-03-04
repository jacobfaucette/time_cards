"""Classes and functions for importing data."""

from app.constants import BAQ_TRANSLATIONS, FORM_TRANSLATIONS, REQUIRED_COLUMNS
from app.helpers import LabeledWidgetGroup, translate_dict_key, translate_dict_keys
from app.ops import OP_TRANSLATIONS
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from tkinter import messagebox

import math
import pathlib
import re
import tkinter as tk


class FormValues:
    """Container for GUI form values."""

    def __init__(
        self,
        assembly: str,
        bkt_hrs: str,
        bkt_qty: str,
        exp_vel: str,
        job_num: str,
        part_name: str,
        part_num: str,
        part_qty: str,
        pro_date: str,
        ops: list[str],
        **_,  # discard unrecognized columns
    ):
        self.assembly = assembly
        self.bkt_hrs = bkt_hrs
        self.bkt_qty = bkt_qty
        self.exp_vel = exp_vel
        self.job_num = job_num
        self.part_name = part_name
        self.part_num = part_num
        self.part_qty = part_qty
        self.pro_date = datetime.strftime(pro_date, "%m/%d/%Y")
        self.ops = ops

    def fill_labeled_widget_group(self, widget_group: LabeledWidgetGroup) -> None:
        """Fill a group of widgets contained in a `LabeledWidgetGroup`.

        parameters:
        - widget_group <LabeledWidgetGroup>: widgets to be filled out
        """

        def set_entry_value(form: str, value: str):
            """Clear then set the value of (form) to (value)."""
            form.delete(0, tk.END)
            form.insert(0, value)

        for label, widget in widget_group.widgets.items():
            actual_key = self.__dict__.get(
                translate_dict_key(label, (BAQ_TRANSLATIONS, FORM_TRANSLATIONS))
            )

            if actual_key is None:
                continue

            set_entry_value(widget, actual_key)

    def fill_text_entry(self, widget: tk.Text) -> None:
        """Fill a tkinter `Text` widget."""
        widget.delete(1.0, tk.END)
        widget.insert(tk.END, ",".join(self.ops))


def cell_values(data: list) -> list[str]:
    """Get cell values from row or column."""
    return [i.value for i in data]


def form_values_from_excel(ws: Worksheet) -> FormValues:
    """Get values from an Excel sheet and output them as a FormValues object."""

    # get rows and columns, discard empty rows
    rows = list(filter(lambda i: i[0].value is not None, ws.rows))
    columns = list(col[1 : len(rows)] for col in ws.columns)

    headers = cell_values(rows.pop(0))

    # display error if required columns are missing
    if not all((i in headers for i in REQUIRED_COLUMNS)):
        messagebox.showerror(
            "Missing/Invalid Columns",
            "This file has missing or incorrectly named columns."
            " Please make sure the following columns are present:\n\n- "
            + "\n- ".join(sorted(REQUIRED_COLUMNS)),
        )

        return

    values = {k: v for k, v in zip(headers, cell_values(rows[0]))}

    # discard unused columns and translate remaining
    values = dict(filter(lambda x: x[0] in REQUIRED_COLUMNS, values.items()))
    values = translate_dict_keys(values, (BAQ_TRANSLATIONS,))

    # column-specific (hard-coded) values

    bkt_hrs = round(sum(cell_values(columns[headers.index("TotalEstHours")])) / values["part_qty"], 2)
    total_hrs = sum(cell_values(columns[headers.index("HrsPerPiece")]))

    bkt_qty = max(1, math.floor(25 / total_hrs))
    bkt_qty = min(bkt_qty, values["part_qty"])

    exp_vel = math.ceil(bkt_hrs / 10.2)

    ops = cell_values(columns[headers.index("Op Dtl Desc")])
    ops = [translate_op(i) for i in ops]

    # remove "-ing" suffix from operations (saves space)
    ops = [re.sub("ing$", "", i.lower()) for i in ops]

    return FormValues(
        **values,
        bkt_hrs=bkt_hrs,
        bkt_qty=bkt_qty,
        exp_vel=exp_vel,
        ops=ops,
    )


def load_excel(file_path: pathlib.Path) -> Worksheet:
    wb = load_workbook(filename=file_path.absolute())
    return wb.active


def translate_op(op: str) -> str:
    """Translate and shorten an operation description.

    If no translation is found, returns the original input.
    """
    pattern = "^[\d\.]+"
    result = re.match(pattern, op)

    return OP_TRANSLATIONS.get(result.group()) or op
