# Plotting code for Kicflip
import pandas as pd
import numpy as np
from datetime import datetime
from matplotlib import pyplot as plt



def _load_xls(path):
    df =  pd.read_excel(path, header = 9) # note that XLSX files are 1 indexed in the viewer
    # drop row 10
    df = df.drop(df.index[1])
    df = df.drop(df.index[0])
    df = df.replace("---", np.nan)
    df = df.replace("<>", np.nan)
    _strip_df_col_whitespace(df)
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
    plt.savefig(outpath + title + ".png")
    plt.close()
    print("Plot for " + title + " saved to: " + outpath + title + ".png")