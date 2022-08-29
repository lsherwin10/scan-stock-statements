import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

import cashapp
import robinhood
import util


def write_holdings(df, writer):
    df.to_excel(writer, "Holdings")
    workbook = writer.book
    worksheet = writer.sheets["Holdings"]

    center_format = workbook.add_format({"align": "center"})
    percent_format = workbook.add_format({"num_format": "0.0%", "align": "center"})
    money_format = workbook.add_format({"num_format": "$0.00", "align": "center"})

    worksheet.set_column("A:M", None, center_format)
    worksheet.set_column("J:M", None, money_format)

    # set individual columns after ranges to prevent overwriting
    worksheet.set_column("B:B", 26)
    worksheet.set_column("D:D", None, money_format)
    worksheet.set_column("E:E", 13, percent_format)
    worksheet.set_column("F:F", 11.5, money_format)
    worksheet.set_column("G:G", 12.5, percent_format)
    worksheet.set_column("H:H", 19.5)
    worksheet.set_column("I:I", 12)
    worksheet.set_column("J:J", 22)
    worksheet.set_column("K:K", 16)
    worksheet.set_column("L:L", 20)
    worksheet.set_column("M:M", 22.5)


def build_summary(df):
    summary = dict()
    summary["Total Invested"] = (df["Equity"] - df["Equity Change"]).sum()
    summary["Total Value"] = df["Equity"].sum()
    summary["Total P/L"] = df["Equity Change"].sum()
    summary["Total ROI"] = summary["Total P/L"] / summary["Total Invested"]
    summary["Annual Income"] = df["Annual Dividend Income"].sum()
    result = pd.DataFrame.from_dict(summary, orient="index")
    result.columns = [""]
    return result


def write_summary(df, writer):
    df.to_excel(writer, "Summary", header=False)
    workbook = writer.book
    worksheet = writer.sheets["Summary"]

    percent_format = workbook.add_format({"num_format": "0.00%"})
    money_format = workbook.add_format({"num_format": "$0.00"})

    worksheet.set_column("A:A", 12)
    worksheet.set_column("B:B", None, money_format)
    worksheet.write("B4", df.loc["Total ROI", :], percent_format)


def write_image(img, writer):
    workbook = writer.book
    worksheet = workbook.add_worksheet(img[:-4].capitalize())
    worksheet.insert_image("A1", img, {"x_scale": 1.5, "y_scale": 1.5})


def generate_graphs(df):
    equity_pcts = df["Percentage"].values
    sectors = df["Sector"].values
    equity_by_sector = dict()

    for k, v in zip(sectors, equity_pcts):
        if k not in equity_by_sector:
            equity_by_sector[k] = 0
        equity_by_sector[k] += float(v)

    div_incomes_by_sector = dict()
    total_income = 0

    for sector in set(sectors):
        div_incomes_by_sector[sector] = df.loc[
            df["Sector"] == sector, "Annual Dividend Per Share"
        ].sum()
        total_income += div_incomes_by_sector[sector]

    div_incomes_by_sector = {
        k: v / total_income * 100 for k, v in div_incomes_by_sector.items()
    }

    equity_by_sector["ETF"] = equity_by_sector["Miscellaneous"]
    equity_by_sector.pop("Miscellaneous")

    div_incomes_by_sector["ETF"] = div_incomes_by_sector["Miscellaneous"]
    div_incomes_by_sector.pop("Miscellaneous")

    equity_sector_keys, equity_sector_vals = util.remove_zero_vals(equity_by_sector)
    div_sector_keys, div_sector_vals = util.remove_zero_vals(div_incomes_by_sector)

    fig, (ax1, ax2) = plt.subplots(2, 1)
    ax1.pie(
        equity_sector_vals,
        labels=equity_sector_keys,
        textprops={"fontsize": 8},
        autopct="%1.1f%%",
    )
    ax1.set_title("Allocation by Equity")

    ax2.pie(
        div_sector_vals,
        labels=div_sector_keys,
        textprops={"fontsize": 8},
        autopct="%1.1f%%",
    )
    ax2.set_title("Allocation by Dividend Income")

    plt.savefig("Diversification.png")


def run(filename):
    # Begin combination process
    complete_df = robinhood.run()
    cashapp_df = cashapp.run(filename)

    complete_df = complete_df.astype(
        {
            "Quantity": float,
            "Equity": float,
            "Equity Change": float,
            "Percent Change": float,
            "Percentage": float,
            "Dividend Yield": float,
        }
    )

    if cashapp_df is not None:
        missing_cols = list(set(complete_df.columns) - set(cashapp_df.columns))
        missing_symbols = list(set(cashapp_df.index) - set(complete_df.index))
        cashapp_df.loc[:, missing_cols] = np.NaN

        complete_df.loc[
            :,
            [
                "Quantity",
                "Equity",
                "Equity Change",
                "Annual Dividend Per Share",
                "Dividend Per Period",
                "Annual Dividend Income",
                "Dividend Income Per Period",
            ],
        ] = complete_df.apply(
            lambda row: row[
                [
                    "Quantity",
                    "Equity",
                    "Equity Change",
                    "Annual Dividend Per Share",
                    "Dividend Per Period",
                    "Annual Dividend Income",
                    "Dividend Income Per Period",
                ]
            ]
            + cashapp_df.loc[
                row.name,
                [
                    "Quantity",
                    "Equity",
                    "Equity Change",
                    "Annual Dividend Per Share",
                    "Dividend Per Period",
                    "Annual Dividend Income",
                    "Dividend Income Per Period",
                ],
            ]
            if row.name in cashapp_df.index
            else row[
                [
                    "Quantity",
                    "Equity",
                    "Equity Change",
                    "Annual Dividend Per Share",
                    "Dividend Per Period",
                    "Annual Dividend Income",
                    "Dividend Income Per Period",
                ]
            ],
            axis=1,
        )
        complete_df = pd.concat((complete_df, cashapp_df.loc[missing_symbols, :]))
        complete_df = util.compute_percent_equity(complete_df)
        complete_df.loc[:, "Percent Change"] = round(
            complete_df["Equity Change"]
            / (complete_df["Equity"] - complete_df["Equity Change"])
            * 100,
            2,
        )

    generate_graphs(complete_df)

    # WRITE TO EXCEL SHEET
    # https://www.codegrepper.com/code-examples/python/pandas+to+excel+append+to+existing+sheet
    final_df = complete_df.copy()
    final_df.drop(["Type", "Percentage"], axis=1, inplace=True)
    final_df["Quantity"] = final_df["Quantity"].astype(float)
    final_df["Equity"] = final_df["Equity"].astype(float)
    final_df["Percent Change"] = final_df["Percent Change"].astype(float) / 100
    final_df["Equity Change"] = final_df["Equity Change"].astype(float)
    final_df["Dividend Yield"] = final_df["Dividend Yield"].astype(float) / 100

    with pd.ExcelWriter(
        "Investments.xlsx",
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_numbers": True}},
    ) as writer:
        write_summary(build_summary(final_df), writer)
        write_holdings(final_df, writer)
        write_image("Diversification.png", writer)


if __name__ == "__main__":
    run("cashapp.pdf")
