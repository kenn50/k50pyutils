# k50pyutils

Personal Python utilities.
This package currently includes an Excel helper for sending **pandas** DataFrames to **Excel** (via **xlwings**) and reading them back.

## Installation

Install directly from GitHub:

```bash
pip install --upgrade git+https://github.com/kenn50/k50pyutils.git
```

---

## Quickstart

By convention, import the Excel helper as **`EW`** and instantiate it as **`ew`**:

```python
import pandas as pd
from k50pyutils import ExcelWriter as EW

ew = EW("test.xlsx")  # open or attach to workbook
```

### Write a DataFrame (includes index)

```python
df = pd.DataFrame(
    {"A": [10, 20], "B": [30, 40]},
    index=["row1", "row2"]
)

# Writes to sheet "Data" (creates it at the END if missing),
# clears A1’s expanded region, writes df (with index), formats as an Excel Table, autofits columns.
ew(df, sheet_name="Data")
```

### Read a DataFrame back (first Excel Table in sheet)

```python
# Returns the first Excel Table in "Data" as a pandas DataFrame.
df_back = ew.get("Data")
```

If you omit `sheet_name` in either call, the **first sheet** (index 0) is used.



## Import Conventions

We strongly recommend the following convention in your projects:

```python
from k50pyutils import ExcelWriter as EW

ew = EW("test.xlsx")
```

This keeps your code concise and consistent across notebooks and scripts.

---

## Requirements

* Python 3.9+
* `pandas`
* `xlwings` (requires Excel on Windows/Mac)

---

Happy scrolling 📊🧭
