import os
import pandas as pd
import pdfplumber

import util

MISSING_SECTORS = {
    "JNJ": "Health Technology",
    "CVX": "Energy",
    "MMM": "Industrial Services",
    "WBD": "Consumer Services",
    "VTI": "Miscellaneous",
}
DIV_YIELDS = {"CVX": 3.47, "JNJ": 2.67, "MMM": 4.62, "T": 7.02}
TYPES = {"VTI": "etp", "VOO": "etp"}


def parse_pdf(filename):
    if not os.path.exists(os.path.abspath(filename)):
        print("File " + filename + " not found for CashApp holdings, ignoring...")
        return None
    with pdfplumber.open(os.path.abspath(filename)) as pdf:
        page = pdf.pages[6]
        text = page.extract_text()

    holdings = []
    columns = None
    appending = False
    for line in text.split("\n"):
        if not appending and line != "HOLDINGS":
            continue
        elif appending and line == "ACTIVITY":
            break
        elif not appending and line == "HOLDINGS":
            appending = True
            continue
        elif appending and line == "Equity":
            continue

        # only reach this block if appending is True and haven't hit continue,
        #   which means we are in the body of the table
        if columns is None:
            columns = line.split()[1:]
            i = 2
            while i < len(columns) - 1:
                spacer = " "
                if i == 10:
                    spacer = ""
                columns[i] = columns[i] + spacer + columns[i + 1]
                columns.pop(i + 1)
                i += 1
            holdings.append(columns)
        else:
            vals = line.split()
            holdings.append(vals[len(vals) - 8 :])

    cashapp = pd.DataFrame(holdings[1:], columns=holdings[0])
    cashapp.set_index("Symbol", drop=True, inplace=True)

    return cashapp


def format_df(df):
    df = df.drop(["Unit Cost", "Total Cost", "A/C Type"], axis=1)

    df_cols = df.columns.tolist()
    df_cols[2] = "Equity"
    df_cols[3] = "Equity Change"
    df.columns = df_cols

    df["Quantity"] = df["Quantity"].astype(float)
    df["Equity"] = df["Equity"].astype(float)
    df.loc[:, "Type"] = df.apply(
        lambda row: "stock" if row.name not in TYPES else TYPES[row.name], axis=1
    )
    df.loc[:, "Dividend Yield"] = df.apply(
        lambda row: 0 if row.name not in DIV_YIELDS else DIV_YIELDS[row.name], axis=1
    )
    df.loc[:, "Dividend Freq"] = df.apply(
        lambda row: "N/A" if row["Dividend Yield"] == 0 else "Quarterly", axis=1
    )
    df.loc[:, "Equity Change"] = df.apply(
        lambda row: float("-" + row["Equity Change"][1:-1])
        if row["Equity Change"][0] == "("
        else float(row["Equity Change"]),
        axis=1,
    )

    df.loc[:, "Annual Dividend Per Share"] = df.apply(
        lambda row: round(float(row["Market Price"]) * row["Dividend Yield"] / 100, 2),
        axis=1,
    )
    df.loc[:, "Dividend Per Period"] = df.apply(
        lambda row: util.get_dividend_per_period(
            row["Annual Dividend Per Share"], row["Dividend Freq"]
        ),
        axis=1,
    )
    df.loc[:, "Annual Dividend Income"] = (
        df.loc[:, "Annual Dividend Per Share"] * df.loc[:, "Quantity"]
    )
    df.loc[:, "Dividend Income Per Period"] = df.apply(
        lambda row: util.get_dividend_per_period(
            row["Annual Dividend Income"], row["Dividend Freq"]
        ),
        axis=1,
    )
    df.loc[:, "Sector"] = df.apply(
        lambda row: MISSING_SECTORS[row.name] if row.name in MISSING_SECTORS else "N/A",
        axis=1,
    )
    df.drop("Market Price", axis=1, inplace=True)

    return df


def run(filename):
    df = parse_pdf(filename)

    # if df is None, that means the file does not exist and can be ignored
    if df is None:
        return None

    return format_df(df)
