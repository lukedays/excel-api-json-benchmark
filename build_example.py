"""Gera exemplo\\PowerQueryApi.xlsm (consultas + tabelas prontas + módulo ApiTabelas) e mede as macros."""
import os
import sys
import time
from pathlib import Path

import win32com.client as win32

from bench import excel_session, find_table, import_vba_module

ROOT = Path(__file__).parent
OUTPUT = ROOT / "exemplo" / "PowerQueryApi.xlsm"
XL_OPEN_XML_WORKBOOK_MACRO_ENABLED = 52
TABLES = [("Vendas", "tblVendas"), ("VendasMes", "tblVendasMes"), ("VendasDia", "tblVendasDia")]
LOADS = [("tblVendas", 100000), ("tblVendasMes", 30000), ("tblVendasDia", 1000)]
API_DELAY_MS = int(os.environ.get("API_DELAY_MS", "0"))  # latência simulada da API em cada chamada


def build_workbook(excel):
    workbook = excel.Workbooks.Add()
    import_vba_module(workbook, ROOT / "vba" / "ApiTabelas.bas")
    first_sheet = workbook.Worksheets(1)
    for index, (sheet_name, table_name) in enumerate(TABLES):
        sheet = first_sheet if index == 0 else workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
        sheet.Name = sheet_name
        excel.Run(f"'{workbook.Name}'!ApiTabelas.CriarTabelaApi", table_name, sheet.Range("A1"))
    OUTPUT.parent.mkdir(exist_ok=True)
    workbook.SaveAs(str(OUTPUT), XL_OPEN_XML_WORKBOOK_MACRO_ENABLED)
    return workbook


def timed(label, action):
    started = time.perf_counter()
    action()
    print(f"{label:52} {time.perf_counter() - started:7.3f}s", flush=True)


def check_rows(workbook, table_name, rows):
    for sheet in workbook.Worksheets:
        table = find_table(sheet, table_name)
        if table is not None:
            if table.ListRows.Count != rows:
                raise RuntimeError(f"{table_name}: esperava {rows}, tem {table.ListRows.Count}")
            return
    raise RuntimeError(f"{table_name} não existe")


def measure(excel, workbook):
    run = lambda macro, *args: excel.Run(f"'{workbook.Name}'!ApiTabelas.{macro}", *args)
    timed("CarregarTabela tblVendas 100000 (1a, motor frio)", lambda: run("CarregarTabela", "tblVendas", "rows", "100000"))
    timed("CarregarTabela tblVendas 100000 (mesma fórmula)", lambda: run("CarregarTabela", "tblVendas", "rows", "100000"))
    timed("CarregarTabela tblVendas 500 (encolhe)", lambda: run("CarregarTabela", "tblVendas", "rows", "500"))
    check_rows(workbook, "tblVendas", 500)

    def sequential():
        for table_name, rows in LOADS:
            run("CarregarTabela", table_name, "rows", str(rows), "delay", str(API_DELAY_MS))

    # O pywin32 converte listas aninhadas em matriz 2D, não em array de arrays; por isso a chamada
    # paralela vai por um módulo temporário, removido antes de salvar.
    loads_vba = ", ".join(f'Array("{table_name}", "rows", "{rows}", "delay", "{API_DELAY_MS}")' for table_name, rows in LOADS)
    test_module = workbook.VBProject.VBComponents.Add(1)
    test_module.CodeModule.AddFromString(f"Public Sub TesteParalelo()\nCarregarVarias Array({loads_vba})\nEnd Sub")

    def parallel():
        excel.Run(f"'{workbook.Name}'!{test_module.Name}.TesteParalelo")

    for label, action in [
        (f"3 tabelas uma a uma (latência {API_DELAY_MS} ms)", sequential),
        (f"3 tabelas em paralelo (latência {API_DELAY_MS} ms)", parallel),
    ]:
        for table_name, _ in LOADS:
            run("CarregarTabela", table_name, "rows", "10")
        timed(label, action)
        for table_name, rows in LOADS:
            check_rows(workbook, table_name, rows)
    workbook.VBProject.VBComponents.Remove(test_module)


def build_and_measure(excel):
    workbook = build_workbook(excel)
    measure(excel, workbook)
    for table_name, _ in LOADS:
        excel.Run(f"'{workbook.Name}'!ApiTabelas.CarregarTabela", table_name, "rows", "10")
    workbook.Save()


def main():
    excel_session(build_and_measure)
    print(f"gerado {OUTPUT}")


if __name__ == "__main__":
    sys.exit(main())
