from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Literal

import pandas as pd


def load_sqlserver_odbc_connection_string() -> str:
    """
    Load the SQL Server ODBC string from (in order):
    1) Environment variable SQLSERVER_ODBC_CONN_STR
    2) Streamlit secrets: [sqlserver] connection_string in .streamlit/secrets.toml

    Do not put credentials in application code or in the Streamlit UI.
    """
    env = os.environ.get("SQLSERVER_ODBC_CONN_STR", "").strip()
    if env:
        return env
    try:
        import streamlit as st

        sec = st.secrets.get("sqlserver", {})
        s = str(sec.get("connection_string", "")).strip()
        if s:
            return s
    except Exception:
        pass
    raise RuntimeError(
        "SQL Server is not configured. Set environment variable SQLSERVER_ODBC_CONN_STR, "
        "or add to .streamlit/secrets.toml:\n\n"
        "[sqlserver]\n"
        'connection_string = "DRIVER={ODBC Driver 18 for SQL Server};SERVER=...;DATABASE=...;UID=...;PWD=...;Encrypt=yes;TrustServerCertificate=yes;"'
    )


DbType = Literal["sqlite", "sqlserver"]


@dataclass(frozen=True)
class DbConfig:
    db_type: DbType
    # For sqlite: path to .db file
    sqlite_path: str | None = None
    # For sqlserver: full ODBC connection string
    sqlserver_conn_str: str | None = None


def read_students_table(
    config: DbConfig,
    *,
    table: str = "student",
    exclude_columns: tuple[str, ...] = ("Password", "password"),
) -> pd.DataFrame:
    """
    Returns a DataFrame from a DB table.

    Expected columns for marksheet PDF:
    - Student Name (or Name/Student)
    - Subject columns (English, Physics, ...)
    """
    if config.db_type == "sqlite":
        if not config.sqlite_path:
            raise ValueError("sqlite_path is required for sqlite.")
        import sqlite3

        conn = sqlite3.connect(config.sqlite_path)
        try:
            df = pd.read_sql_query(f"SELECT * FROM {table}", conn)
            for col in exclude_columns:
                if col in df.columns:
                    df = df.drop(columns=[col])
            return df
        finally:
            conn.close()

    if config.db_type == "sqlserver":
        if not config.sqlserver_conn_str:
            raise ValueError("sqlserver_conn_str is required for sqlserver.")
        try:
            import pyodbc  # type: ignore
        except Exception as e:
            raise RuntimeError(
                "pyodbc is required for SQL Server. Install it with: pip install pyodbc"
            ) from e

        conn = pyodbc.connect(config.sqlserver_conn_str)
        try:
            df = pd.read_sql_query(f"SELECT * FROM {table}", conn)
            for col in exclude_columns:
                if col in df.columns:
                    df = df.drop(columns=[col])
            return df
        finally:
            conn.close()

    raise ValueError(f"Unsupported db_type: {config.db_type}")

