import duckdb
import pandas as pd
import inspect

def sql_query(query: str):
    # inspect caller’s variables
    caller_locals = inspect.currentframe().f_back.f_locals
    caller_globals = inspect.currentframe().f_back.f_globals

    con = duckdb.connect()

    # register all pandas DataFrames in caller’s scope
    for name, val in {**caller_globals, **caller_locals}.items():
        if isinstance(val, pd.DataFrame):
            con.register(name, val)

    return con.execute(query).df()

def _df_sql(self, query: str, table_name="self"):
    """
    Instance method: runs a query where `self` is always registered as 'self',
    and all other DataFrames in caller's scope are also available.
    """
    # Look at caller’s variables
    caller_locals = inspect.currentframe().f_back.f_locals
    caller_globals = inspect.currentframe().f_back.f_globals

    con = duckdb.connect()

    # Register self as "self" (or custom table_name)
    con.register(table_name, self)

    # Register other DataFrames from caller's scope
    for name, val in {**caller_globals, **caller_locals}.items():
        if isinstance(val, pd.DataFrame):
            con.register(name, val)

    return con.execute(query).df()


# Patch onto pandas
pd.DataFrame.sql = _df_sql