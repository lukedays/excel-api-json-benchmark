"""Mede o add-in C# num Excel visível: fases da carga real e variantes de gravação (CsBenchWrite)."""
import os
import sys
import tempfile
import time
from pathlib import Path

import json
import urllib.request

from bench import CS_ADDIN, FRIENDLY_NAMES_VBA, TABLE_NAME, check_table, excel_session, find_table

ROWS = int(os.environ.get("BENCH_ROWS", "100000"))
SEQUENCE = [int(s) for s in os.environ.get("BENCH_SEQUENCE", "1000,10000,100000,100000,1000").split(",")]
COMMAND = os.environ.get("BENCH_COMMAND", "CsLoadTable")
URL = "http://localhost:8765/table?rows={rows}&run={run}"


def run_load_sequence(excel, workbook):
    """Carga real numa tabela nomeada, crescendo e encolhendo, com o tempo de cada fase."""
    workbook.Worksheets(1).Name = "cs"
    for run, rows in enumerate(SEQUENCE):
        started = time.perf_counter()
        excel.Run(COMMAND, URL.format(rows=rows, run=run), "cs", TABLE_NAME, FRIENDLY_NAMES_VBA)
        total = time.perf_counter() - started
        print(f"{COMMAND} {rows:>7} total {total:6.3f}  {excel.Run('CsLastTimings')}", flush=True)
        check_table(workbook.Worksheets("cs"), rows)
        check_edge_rows(workbook.Worksheets("cs"), rows)


def check_edge_rows(sheet, rows):
    """Compara a primeira e a última linha da tabela com o JSON da API."""
    records = json.loads(urllib.request.urlopen(URL.format(rows=rows, run="check")).read())
    body = find_table(sheet, TABLE_NAME).DataBodyRange
    for index in (0, rows - 1):
        expected = [float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else v for v in records[index].values()]
        actual = list(body.Rows(index + 1).Value2[0])
        if actual != expected:
            raise RuntimeError(f"linha {index + 1}: esperava {expected}, veio {actual}")


def run_write_variants(excel):
    report = Path(tempfile.gettempdir()) / "cs_write_bench.txt"
    excel.Run("CsBenchWrite", URL.format(rows=ROWS, run="w"), str(report))
    print(report.read_text(), flush=True)


def run_all(excel):
    excel.Visible = True
    if not excel.RegisterXLL(str(CS_ADDIN)):
        raise RuntimeError(f"Excel recusou o add-in {CS_ADDIN}")
    workbook = excel.Workbooks.Add()
    run_load_sequence(excel, workbook)
    if os.environ.get("BENCH_VARIANTS") == "1":
        run_write_variants(excel)


if __name__ == "__main__":
    excel_session(run_all)
    sys.exit(0)
