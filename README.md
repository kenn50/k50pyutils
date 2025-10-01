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

---

## Tutorial

### 1) Writing DataFrames to Excel

```python
from k50pyutils import ExcelWriter as EW
import pandas as pd

ew = EW("test.xlsx")

# Example dataset
df = pd.DataFrame(
    {
        "category": ["A", "B", "A", "C"],
        "score": [71.5, 66.2, 88.0, 73.1],
    },
    index=["r1", "r2", "r3", "r4"],
)

# Send to Excel
ew(df, sheet_name="Scores")
```

What happens under the hood:

* If `sheet_name` is **None** → uses the **first sheet**.
* If `sheet_name` is provided and exists → uses it.
* If it doesn’t exist → **creates** a new sheet **at the end** of the workbook.
* Clears the existing data starting at **A1** (using `.expand()`).
* Writes the DataFrame **including the index**.
* Wraps the written range as an **Excel Table**.
* **Autofits** columns.

### 2) Reading DataFrames from Excel (`get`)

```python
# Read the first Table from the sheet back into pandas
df_back = ew.get("Scores")
print(df_back)
```

Notes:

* `get(sheet_name=None)` uses the **first sheet** if not specified.
* It returns the **first Excel Table** on that sheet as a DataFrame.
* If there are **no Table objects** on that sheet, it raises:

  ```
  ValueError: No tables found in sheet '...'
  ```

Tip: Since `__call__` (write) always formats your output as a Table, you can reliably round-trip with `get()`.

---

## API

### `ExcelWriter(path: str)`

Class for writing and reading pandas DataFrames from Excel workbooks.

```python
from k50pyutils import ExcelWriter as EW
ew = EW("test.xlsx")
```

### `ew(df: pd.DataFrame, sheet_name: str | None = None) -> None`

Write a DataFrame to Excel.

* **sheet_name=None** → write to the first sheet.
* If the sheet doesn’t exist, it is **created at the end**.
* Always writes to **A1**.
* **Includes the DataFrame’s index** in Excel.
* Clears the previous output region starting at A1.
* Formats the output as an **Excel Table** and autofits columns.

### `ew.get(sheet_name: str | None = None) -> pd.DataFrame`

Return the **first Excel Table** in the sheet as a pandas DataFrame.

* **sheet_name=None** → read from the first sheet.
* **Raises** `ValueError` if the sheet has no Table objects.

---

## Examples

### Work with multiple sheets using the same instance

```python
from k50pyutils import ExcelWriter as EW
import pandas as pd

ew = EW("test.xlsx")

df1 = pd.DataFrame({"A": [1, 2, 3]}, index=["a", "b", "c"])
df2 = pd.DataFrame({"X": [10, 20], "Y": [30, 40]}, index=["r1", "r2"])

ew(df1, sheet_name="Input")
ew(df2, sheet_name="Results")

df2_back = ew.get("Results")
```

### “Exploration mode” convenience

```python
def explore(df, name="Preview"):
    ew = EW("explore.xlsx")
    ew(df, sheet_name=name)

# Usage
explore(df, name="After_Filtering")
```

---

## Troubleshooting

### `ValueError: The truth value of an array with more than one element is ambiguous...`

Cause: One or more DataFrame columns contain **lists/arrays/tuples** per cell. Excel can’t write these directly.

**Fix**: Convert those columns to strings (or expand them into multiple columns) before writing:

```python
import numpy as np

for col in df.columns:
    if df[col].apply(lambda x: isinstance(x, (list, tuple, np.ndarray))).any():
        df[col] = df[col].astype(str)

ew(df, sheet_name="Data")
```

### COM errors when selecting sheets by name

If a sheet name doesn’t exist and you try to index it directly via xlwings, COM errors can surface. `EW` avoids this by **checking existing sheet names** first and creating the sheet if needed (at the end).

### `ValueError: No tables found in sheet '...'`

You called `get()` on a sheet that has no Excel Table objects. Make sure you wrote via `ew(...)` first (or manually format a range as a Table in Excel).

---

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
