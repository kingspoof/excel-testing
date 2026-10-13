# Excel tester

Tests an Excel template using a YAML specification. For each test it:

1. copies the template (to a temporary file, deleted afterwards),
2. writes the `inputs` into the copy,
3. recalculates the formulas,
4. runs the `checks`: a cell equals a fixed value, another cell in the copy, or a cell in another workbook.

## Usage

```bash
pip install -r requirements.txt
```

Open `run_tests.ipynb`, set `TESTS`, and run all cells. From Python:

```python
from excel_tester import ExcelTester
tester = ExcelTester()          # engine="formulas" (default) or "libreoffice"
tester.print_report(tester.run("tests.yaml"))
```

The Excel to test is set with `template:` at the top of the YAML. See `examples/tests.yaml` for the format (regenerate the example workbooks with `python examples/make_examples.py`).

## Formula engine: read this

openpyxl cannot calculate formulas, so something else has to:

| engine | needs | notes |
|---|---|---|
| `formulas` (default) | `pip` only | Fast. Most common functions work. Unsupported functions (e.g. `OFFSET`) show up as `#NAME?`. **Spill/dynamic-array functions (`UNIQUE`, `FILTER`, `SORT`, …) silently return only the first value.** |
| `libreoffice` | LibreOffice installed | Very close to Excel. About 3s per test. |

Before trusting the `formulas` engine with a real template, run the same tests once with `engine="libreoffice"` (or check the results against Excel) to confirm that both give the same results.

Values read from a reference workbook (`file:`) are the values saved by Excel. If that file contains formulas, it must have been saved in Excel.
