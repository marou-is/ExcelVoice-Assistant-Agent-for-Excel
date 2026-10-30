import pythoncom
import pywintypes
import win32com.client as win32

from .vba_generator import ENTRY_POINT

VBEXT_CT_STD_MODULE = 1
MAX_HEADER_COLUMNS = 30

TRUST_HELP = (
    "Excel refused access to the VBA project. Enable it in Excel: File > Options > "
    "Trust Center > Trust Center Settings > Macro Settings > check "
    "'Trust access to the VBA project object model', then try again."
)


class ExcelError(Exception):
    pass


def com_error_message(error: pywintypes.com_error) -> str:
    excepinfo = error.excepinfo if len(error.args) > 2 else None
    if excepinfo and excepinfo[2]:
        return str(excepinfo[2]).strip()
    return str(error.strerror or error)


class ExcelBridge:
    """COM connection to Excel. Use as a context manager inside the thread that calls it."""

    def __enter__(self):
        pythoncom.CoInitialize()
        self.app = self._connect()
        return self

    def __exit__(self, *exc):
        self.app = None
        pythoncom.CoUninitialize()

    @staticmethod
    def _connect():
        try:
            app = win32.GetActiveObject("Excel.Application")
        except pywintypes.com_error:
            app = win32.Dispatch("Excel.Application")
        app.Visible = True
        if app.Workbooks.Count == 0:
            app.Workbooks.Add()
        return app

    def describe_workbook(self) -> str:
        app = self.app
        wb = app.ActiveWorkbook
        ws = app.ActiveSheet
        lines = [
            f"Workbook: {wb.Name}",
            "Sheets: " + ", ".join(sheet.Name for sheet in wb.Worksheets),
            f"Active sheet: {ws.Name}",
        ]

        if app.WorksheetFunction.CountA(ws.Cells) == 0:
            lines.append("Active sheet is empty.")
        else:
            used = ws.UsedRange
            lines.append(f"Used range: {used.Address}")
            header = used.Rows(1).Value
            if header:
                values = header[0] if isinstance(header, tuple) else (header,)
                values = [str(v) for v in values[:MAX_HEADER_COLUMNS] if v is not None]
                lines.append("First row of used range: " + ", ".join(values))

        for sheet in wb.Worksheets:
            for table in sheet.ListObjects:
                columns = ", ".join(col.Name for col in table.ListColumns)
                lines.append(
                    f"Table '{table.Name}' on sheet '{sheet.Name}' at {table.Range.Address}: "
                    f"columns [{columns}], {table.ListRows.Count} data rows"
                )

        try:
            lines.append(f"Current selection: {app.Selection.Address}")
        except (pywintypes.com_error, AttributeError):
            pass
        return "\n".join(lines)

    def run_vba(self, code: str) -> None:
        wb = self.app.ActiveWorkbook
        try:
            components = wb.VBProject.VBComponents
        except pywintypes.com_error as error:
            raise ExcelError(TRUST_HELP) from error

        module = components.Add(VBEXT_CT_STD_MODULE)
        try:
            module.CodeModule.AddFromString(code)
            self.app.Run(f"'{wb.Name}'!{module.Name}.{ENTRY_POINT}")
        except pywintypes.com_error as error:
            raise ExcelError(com_error_message(error)) from error
        finally:
            components.Remove(module)
