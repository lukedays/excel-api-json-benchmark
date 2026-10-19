"""Mede o tempo de levar uma tabela JSON de uma API até uma tabela nomeada do Excel, por método."""
import itertools
import json
import os
import statistics
import sys
import tempfile
import time
import traceback
import urllib.request
import winreg
from pathlib import Path

import win32api
import win32com.client as win32
import win32con
import win32event
import win32process

API = "http://localhost:8765/table?rows={rows}&run={run}"
RUN_IDS = itertools.count()  # parâmetro único por chamada, para nenhum método se beneficiar de cache HTTP
SIZES = [int(s) for s in os.environ.get("BENCH_SIZES", "1000,10000,100000").split(",")]
RUNS = int(os.environ.get("BENCH_RUNS", "3"))
METHODS = set(os.environ.get("BENCH_METHODS", "py,vbaflat,vbajson,cs,pqrecords,pqrows").split(","))
OUTPUT_FILE = os.environ.get("BENCH_OUT", "results.json")
ROOT = Path(__file__).parent
CS_ADDIN = ROOT / "csaddin" / "bin" / "Release" / "net10.0-windows" / "CsAddIn-AddIn64.xll"
SCRIPTING_RUNTIME_GUID = "{420B2830-E718-11CF-893D-00A0C9054228}"
TABLE_NAME = "tblVendas"
FRIENDLY_NAMES = json.loads((ROOT / "columns.json").read_text(encoding="utf-8"))
FRIENDLY_NAMES_VBA = "|".join(f"{key}={label}" for key, label in FRIENDLY_NAMES.items())


def m_text(value):
    return '"' + value.replace('"', '""') + '"'


M_RENAMES = "{" + ", ".join("{" + m_text(k) + ", " + m_text(v) + "}" for k, v in FRIENDLY_NAMES.items()) + "}"


def m_source(rows):
    query = f'[rows="{rows}", run="{next(RUN_IDS)}"]'
    return f'Json.Document(Web.Contents("http://localhost:8765", [RelativePath="table", Query={query}]))'


def pq_from_records(rows):
    return (
        f"let Src = {m_source(rows)}, T = Table.FromRecords(Src)"
        f" in Table.RenameColumns(T, {M_RENAMES}, MissingField.Ignore)"
    )


def pq_from_rows(rows):
    return (
        f"let Src = {m_source(rows)}, Keys = Record.FieldNames(Src{{0}}),"
        " T = Table.FromRows(List.Transform(Src, Record.FieldValues), Keys)"
        f" in Table.RenameColumns(T, {M_RENAMES}, MissingField.Ignore)"
    )


def set_vbom_access(enabled):
    """Liga/desliga 'Confiar no acesso ao modelo de objeto do VBA' (necessário para injetar o módulo)."""
    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\16.0\Excel\Security")
    if enabled:
        winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
    else:
        try:
            winreg.DeleteValue(key, "AccessVBOM")
        except FileNotFoundError:
            pass


def excel_session(work):
    """Roda work(excel) numa instância nova do Excel e devolve o resultado.

    AccessVBOM fica ligado só durante a sessão. O Excel regrava as opções de segurança ao terminar,
    então ela só é desligada depois que o processo saiu, o que exige soltar toda referência COM antes.
    """
    set_vbom_access(True)
    try:
        result, pid = _run_and_quit(work)
    except BaseException as error:
        traceback.clear_frames(error.__traceback__)  # os frames seguram referências COM
        if getattr(error, "pid", None) is not None:
            _wait_for_exit(error.pid)
        set_vbom_access(False)
        raise
    _wait_for_exit(pid)
    set_vbom_access(False)
    return result


def _run_and_quit(work):
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    _, pid = win32process.GetWindowThreadProcessId(excel.Hwnd)
    try:
        return work(excel), pid
    except BaseException as error:
        error.pid = pid
        raise
    finally:
        for workbook in list(excel.Workbooks):
            workbook.Close(SaveChanges=False)
        excel.Quit()


def _wait_for_exit(pid):
    """Espera o Excel sair; se não sair (add-in .NET segurando referências COM), encerra o processo."""
    process = win32api.OpenProcess(win32con.SYNCHRONIZE | win32con.PROCESS_TERMINATE, False, pid)
    if win32event.WaitForSingleObject(process, 15000) == win32event.WAIT_TIMEOUT:
        win32api.TerminateProcess(process, 1)
        win32event.WaitForSingleObject(process, 5000)
    win32api.CloseHandle(process)


def import_vba_module(workbook, bas_path):
    """Importa um .bas regravado em cp1252, que é o encoding que o editor VBA espera."""
    source = Path(bas_path).read_text(encoding="utf-8")
    with tempfile.NamedTemporaryFile(
        "w", suffix=".bas", encoding="cp1252", errors="replace", newline="\r\n", delete=False
    ) as tmp:
        tmp.write(source)
    workbook.VBProject.VBComponents.Import(tmp.name)
    os.unlink(tmp.name)


def open_bench_workbook(excel):
    workbook = excel.Workbooks.Add()
    workbook.VBProject.References.AddFromGuid(SCRIPTING_RUNTIME_GUID, 1, 0)
    import_vba_module(workbook, ROOT / "vendor" / "JsonConverter.bas")
    import_vba_module(workbook, ROOT / "vba" / "BenchFetch.bas")
    import_vba_module(workbook, ROOT / "vba" / "PqSetup.bas")
    if not excel.RegisterXLL(str(CS_ADDIN)):
        raise RuntimeError(f"Excel recusou o add-in {CS_ADDIN}")
    return workbook


def fresh_sheet(workbook, name):
    for sheet in list(workbook.Worksheets):
        if sheet.Name == name:
            sheet.Delete()
    sheet = workbook.Worksheets.Add()
    sheet.Name = name
    return sheet


def find_table(sheet, name):
    for table in sheet.ListObjects:
        if table.Name == name:
            return table
    return None


def write_to_table(sheet, name, matrix):
    """Grava a matriz (linha 0 = cabeçalhos) numa tabela nomeada, crescendo ou encolhendo para caber exatamente."""
    row_count, col_count = len(matrix), len(matrix[0])
    table = find_table(sheet, name)
    anchor = sheet.Range("A1") if table is None else table.Range.Cells(1, 1)
    new_range = anchor.GetResize(max(row_count, 2), col_count)  # tabela exige ao menos uma linha de dados

    if table is None:
        new_range.GetResize(row_count, col_count).Value = matrix
        table = sheet.ListObjects.Add(1, new_range, None, 1)
        table.Name = name
        return

    # Cabeçalho e corpo vão em gravações separadas: um bloco único sobre a tabela inteira a desfaz.
    old_range = table.Range
    old_rows, old_cols = old_range.Rows.Count, old_range.Columns.Count
    table.Resize(new_range)
    table.HeaderRowRange.Value = [matrix[0]]
    if row_count == 1:
        table.DataBodyRange.ClearContents()
    else:
        table.DataBodyRange.Value = matrix[1:]
    if old_rows > new_range.Rows.Count:
        old_range.GetOffset(new_range.Rows.Count, 0).GetResize(old_rows - new_range.Rows.Count, old_cols).Clear()
    if old_cols > col_count:
        old_range.GetOffset(0, col_count).GetResize(old_rows, old_cols - col_count).Clear()


def load_python_com(workbook, rows):
    with urllib.request.urlopen(API.format(rows=rows, run=next(RUN_IDS))) as response:
        records = json.loads(response.read())
    keys = list(records[0].keys())
    matrix = [[FRIENDLY_NAMES.get(k, k) for k in keys]] + [[rec.get(k) for k in keys] for rec in records]
    write_to_table(workbook.Worksheets("py"), TABLE_NAME, matrix)


def load_vba(macro):
    def run(workbook, rows):
        workbook.Application.Run(
            f"'{workbook.Name}'!{macro}", API.format(rows=rows, run=next(RUN_IDS)), "vba", TABLE_NAME, FRIENDLY_NAMES_VBA
        )
    return run


def load_csharp_addin(command):
    def run(workbook, rows):
        workbook.Application.Run(command, API.format(rows=rows, run=next(RUN_IDS)), "cs", TABLE_NAME, FRIENDLY_NAMES_VBA)
    return run


def make_power_query_loader(build_formula):
    def run(workbook, rows):
        workbook.Queries("Q").Formula = build_formula(rows)
        find_table(workbook.Worksheets("pq"), TABLE_NAME).QueryTable.Refresh(False)
    return run


def setup_power_query(workbook, build_formula):
    """Cria a consulta e a tabela de saída; devolve o tempo da primeira carga (motor do Power Query frio)."""
    for query in list(workbook.Queries):
        query.Delete()
    sheet = fresh_sheet(workbook, "pq")
    workbook.Queries.Add("Q", build_formula(10))
    workbook.Application.Run(f"'{workbook.Name}'!PqSetup.CreateQueryTable", "pq", "Q", TABLE_NAME)
    table = find_table(sheet, TABLE_NAME)
    started = time.perf_counter()
    table.QueryTable.Refresh(False)
    return time.perf_counter() - started


def check_table(sheet, rows):
    """Confere linhas da tabela, ausência de sobras abaixo dela e cabeçalhos amigáveis."""
    table = find_table(sheet, TABLE_NAME)
    if table is None:
        raise RuntimeError(f"{sheet.Name}: tabela {TABLE_NAME} não existe")
    last_used_row = sheet.Cells.Find("*", sheet.Cells(1, 1), -4123, 2, 1, 2).Row  # xlFormulas, xlPart, xlByRows, xlPrevious
    headers = [cell.Value for cell in table.HeaderRowRange]
    if table.ListRows.Count != rows or last_used_row != rows + 1:
        raise RuntimeError(
            f"{sheet.Name}: esperava {rows} linhas, tabela tem {table.ListRows.Count}, última usada {last_used_row}"
        )
    if headers != list(FRIENDLY_NAMES.values()):
        raise RuntimeError(f"{sheet.Name}: cabeçalhos inesperados {headers}")


def time_runs(loader, workbook, rows):
    timings = []
    for _ in range(RUNS):
        started = time.perf_counter()
        loader(workbook, rows)
        timings.append(time.perf_counter() - started)
    return timings


def measure_sizes(label, loader, workbook, sheet, results):
    """Roda os tamanhos em ordem crescente (a tabela cresce) e depois volta ao menor (a tabela encolhe)."""
    for rows in SIZES:
        timings = time_runs(loader, workbook, rows)
        check_table(sheet, rows)
        results.append((label, rows, timings))
        print(f"{label:46} {rows:>7} {statistics.median(timings):8.3f}s", flush=True)
    started = time.perf_counter()
    loader(workbook, SIZES[0])
    shrink = time.perf_counter() - started
    check_table(sheet, SIZES[0])
    results.append((label, f"{SIZES[-1]}->{SIZES[0]}", [shrink]))
    print(f"{label:46} encolhe {SIZES[-1]}->{SIZES[0]} {shrink:8.3f}s", flush=True)


def run_methods(excel):
    excel.ScreenUpdating = False
    workbook = open_bench_workbook(excel)
    excel.Calculation = -4135  # manual
    results = []

    # Aquece o servidor (gera e cacheia o JSON de cada tamanho).
    for rows in SIZES:
        urllib.request.urlopen(API.format(rows=rows, run=next(RUN_IDS))).read()

    methods = [
        ("py", "Python (urllib+json) -> tabela via COM", "py", load_python_com),
        ("vbaflat", "VBA parser próprio -> tabela", "vba", load_vba("LoadWithFlatParser")),
        ("vbajson", "VBA + VBA-JSON -> tabela", "vba", load_vba("LoadWithVbaJson")),
        ("cs", "Add-in C# (Excel-DNA) -> tabela", "cs", load_csharp_addin("CsLoadTable")),
    ]
    for key, label, sheet_name, loader in methods:
        if key in METHODS:
            measure_sizes(label, loader, workbook, fresh_sheet(workbook, sheet_name), results)

    for key, label, build_formula in [
        ("pqrecords", "Power Query Table.FromRecords", pq_from_records),
        ("pqrows", "Power Query Table.FromRows", pq_from_rows),
    ]:
        if key not in METHODS:
            continue
        cold = setup_power_query(workbook, build_formula)
        print(f"{label:46} primeira carga (frio) {cold:8.3f}s", flush=True)
        measure_sizes(label, make_power_query_loader(build_formula), workbook, workbook.Worksheets("pq"), results)
    return results


def main():
    results = excel_session(run_methods)
    out = ROOT / OUTPUT_FILE
    out.write_text(json.dumps([{"method": m, "rows": r, "runs": t} for m, r, t in results], indent=2))
    print(f"\nresultados em {out}")


if __name__ == "__main__":
    sys.exit(main())
