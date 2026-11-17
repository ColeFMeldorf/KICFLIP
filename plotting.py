# Plotting code for Kicflip
import pandas as pd
import numpy as np
from datetime import datetime
from matplotlib import pyplot as plt



def _read_raw_table(path, header=9):
    """Read a table whose header is on row `header` (0-indexed).

    Some exporters save plain text (tab- or comma-separated) or HTML with an
    .xls extension, which xlrd rejects ("Expected BOF record"). Pick the reader
    from the file's actual contents instead of its extension.
    """
    with open(path, "rb") as f:
        start = f.read(2048)

    if start.startswith(b"\xd0\xcf\x11\xe0"):  # real legacy .xls (OLE2)
        return pd.read_excel(path, header=header, engine="xlrd")
    if start.startswith(b"PK"):  # .xlsx saved as .xls
        return pd.read_excel(path, header=header, engine="openpyxl")
    if start.lstrip().lower().startswith((b"<html", b"<!doctype", b"<table", b"<?xml")):
        raise ValueError(
            f"{path} is an HTML/XML spreadsheet, not a real .xls. "
            "Open it in Excel and use Save As > Excel Workbook (.xlsx) or CSV."
        )

    # Plain text export. Detect encoding and delimiter.
    if start.startswith((b"\xff\xfe", b"\xfe\xff")):
        encoding = "utf-16"
    elif start.startswith(b"\xef\xbb\xbf"):
        encoding = "utf-8-sig"
    else:
        encoding = "latin-1"  # never fails to decode
    with open(path, "r", encoding=encoding, newline="") as f:
        lines = f.read().splitlines()
    header_line = lines[header] if len(lines) > header else ""
    sep = "\t" if header_line.count("\t") >= header_line.count(",") else ","

    df = pd.read_csv(
        path,
        sep=sep,
        header=header,
        encoding=encoding,
        skip_blank_lines=False,
        engine="python",
        index_col=False,
    )
    return df


def _text_columns_to_values(df):
    """Text exports give strings; convert to the types xlrd would have produced."""
    for c in df.columns:
        if df[c].dtype != object:
            continue
        s = df[c].astype("string").str.strip()
        nums = pd.to_numeric(s, errors="coerce").astype(float)
        if nums.notna().sum() == s.notna().sum():
            df[c] = nums
            continue
        if c.strip().lower() == "time":
            times = pd.to_datetime(s, errors="coerce", format="mixed")  # HH:MM[:SS] [AM/PM]
            if times.notna().any():
                df[c] = times.dt.time
    return df


def _load_xls(path):
    df = _read_raw_table(path, header=9)  # note that XLSX files are 1 indexed in the viewer
    # drop row 10
    df = df.drop(df.index[1])
    df = df.drop(df.index[0])
    df = df.replace("---", np.nan)
    df = df.replace("<>", np.nan)
    _strip_df_col_whitespace(df)
    _text_columns_to_values(df)
    return df

def _strip_df_col_whitespace(df):
    for c in df.columns:
        df.rename(columns={c: c.strip()}, inplace=True)

# def _time_col_to_timedelta(df, col_name):
#     # Convert the time column to timedelta
#     df[col_name] = pd.to_timedelta(df[col_name])
#     return df``

def _convert_time_to_Minutes(df, col_name):
    # Convert the time column to Minutes
    df["Minutes"] = df[col_name].apply(
        lambda t: t.hour * 60 + t.minute + t.second / 60
    )
    day_delta = np.zeros_like(df[col_name])
    day_wrap = np.where(np.diff(df["Minutes"]) < 0)[0] + 1
    for i, index in enumerate(day_wrap):
        day_delta[index:] = i + 1
    df["Minutes"] += day_delta * 24 * 60

    return df

def make_plot(df, kwds, outpath, title = None):
    plt.figure(figsize=(10, 6), dpi = 200)

    df = _convert_time_to_Minutes(df, "Time")
    df["Minutes"] = df["Minutes"].astype(float)

    if isinstance(kwds, str):
        kwds = [kwds]
    if title is None:
        title = kwds[0]
    for kwd in kwds:
        plt.scatter(df["Minutes"], df[kwd],marker='o', linestyle='-', label = kwd)
    plt.grid()
    plt.xlabel("Time (minutes)")
    plt.ylabel("Value")
    plt.legend()
    plt.title(title, fontsize=20)
    plt.savefig(outpath + title + "_plot.png")
    plt.close()
    print("Plot for " + title + " saved to: " + outpath + title + "_plot.png")