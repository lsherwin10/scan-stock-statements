#%%
import robin_stocks
import robin_stocks.robinhood as r
import pyotp
import pandas as pd
import matplotlib.pyplot as plt

MONTHLY_DIVS = {"PBA", "STAG", "MAIN"}


def get_dividend_per_period(dividend, freq):
    if freq == "Monthly":
        return dividend / 12
    if freq == "Quarterly":
        return dividend / 4
    else:
        return 0


def remove_zero_vals(vals_dict):
    i = 0
    keys = list(vals_dict.keys())
    new_keys, new_vals = [], []
    while i < len(keys):
        key = keys[i]
        if vals_dict[key] > 0:
            new_keys.append(key)
            new_vals.append(vals_dict[key])
        i += 1
    return new_keys, new_vals


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
    worksheet.write("B5", df.loc["Total ROI", :], percent_format)


def write_image(img, writer):
    workbook = writer.book
    worksheet = workbook.add_worksheet(img[:-4].capitalize())
    worksheet.insert_image("A1", img, {"x_scale": 1.5, "y_scale": 1.5})


#%%
# totp = pyotp.TOTP(pyotp.random_base32()).now()
# print("Current OTP:", totp)

# fill in with own credentials: username, password
login = r.login()

my_stocks = r.build_holdings()

#%%
# print(my_stocks)
tickers = list(my_stocks.keys())
fundamentals = r.stocks.get_fundamentals(tickers, "dividend_yield")
sectors = r.stocks.get_fundamentals(tickers, "sector")

#%%
stocks_with_info = my_stocks
for ticker, div, sector in zip(tickers, fundamentals, sectors):
    stocks_with_info[ticker]["dividend_yield"] = float(div) if div is not None else 0.0
    stocks_with_info[ticker]["sector"] = sector
    dividend_freq = "N/A"
    if ticker in MONTHLY_DIVS:
        dividend_freq = "Monthly"
    elif div is not None:
        dividend_freq = "Quarterly"
    stocks_with_info[ticker]["dividend_freq"] = dividend_freq
    stocks_with_info[ticker]["annual_dividend_per_share"] = float(
        stocks_with_info[ticker]["price"]
    ) * (stocks_with_info[ticker]["dividend_yield"] / 100)
    stocks_with_info[ticker]["dividend_per_period"] = get_dividend_per_period(
        stocks_with_info[ticker]["annual_dividend_per_share"], dividend_freq
    )
    stocks_with_info[ticker]["annual_dividend_income"] = stocks_with_info[ticker][
        "annual_dividend_per_share"
    ] * float(stocks_with_info[ticker]["quantity"])
    stocks_with_info[ticker]["dividend_income_per_period"] = get_dividend_per_period(
        stocks_with_info[ticker]["annual_dividend_income"], dividend_freq
    )

# %%
# print(my_stocks)
stocks_df = pd.DataFrame.from_dict(stocks_with_info, orient="index")
# display(stocks_df)

#%%
trimmed_stocks_df = stocks_df.drop(
    ["price", "average_buy_price", "intraday_percent_change", "pe_ratio", "id"], axis=1
)
trimmed_stocks_df.columns = [s.replace("_", " ") for s in trimmed_stocks_df.columns]
trimmed_stocks_df.columns = [
    " ".join([st.capitalize() for st in s.split()]) for s in trimmed_stocks_df.columns
]

cols = trimmed_stocks_df.columns.tolist()
cols.remove("Name")
cols.remove("Type")
cols.insert(0, "Name")
cols.insert(1, "Type")
trimmed_stocks_df = trimmed_stocks_df[cols]
# display(trimmed_stocks_df)


# GRAPHS
# %%
equity_pcts = trimmed_stocks_df["Percentage"].values
sectors = trimmed_stocks_df["Sector"].values
equity_by_sector = dict()

for k, v in zip(sectors, equity_pcts):
    if k not in equity_by_sector:
        equity_by_sector[k] = 0
    equity_by_sector[k] += float(v)

div_incomes_by_sector = dict()
total_income = 0

for sector in set(sectors):
    div_incomes_by_sector[sector] = trimmed_stocks_df.loc[
        trimmed_stocks_df["Sector"] == sector, "Annual Dividend Per Share"
    ].sum()
    total_income += div_incomes_by_sector[sector]

div_incomes_by_sector = {
    k: v / total_income * 100 for k, v in div_incomes_by_sector.items()
}

equity_by_sector["ETF"] = equity_by_sector["Miscellaneous"]
equity_by_sector.pop("Miscellaneous")

div_incomes_by_sector["ETF"] = div_incomes_by_sector["Miscellaneous"]
div_incomes_by_sector.pop("Miscellaneous")

equity_sector_keys, equity_sector_vals = remove_zero_vals(equity_by_sector)
div_sector_keys, div_sector_vals = remove_zero_vals(div_incomes_by_sector)

#%%
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

# WRITE TO EXCEL SHEET
# https://www.codegrepper.com/code-examples/python/pandas+to+excel+append+to+existing+sheet
# %%
final_df = trimmed_stocks_df.copy()
final_df.drop(["Type", "Percentage"], axis=1, inplace=True)
final_df["Quantity"] = final_df["Quantity"].astype(float)
final_df["Equity"] = final_df["Equity"].astype(float)
final_df["Percent Change"] = final_df["Percent Change"].astype(float) / 100
final_df["Equity Change"] = final_df["Equity Change"].astype(float)
final_df["Dividend Yield"] = final_df["Dividend Yield"].astype(float) / 100


with pd.ExcelWriter(
    "investments.xlsx",
    engine="xlsxwriter",
    engine_kwargs={"options": {"strings_to_numbers": True}},
) as writer:
    write_summary(build_summary(final_df), writer)
    write_holdings(final_df, writer)
    write_image("Diversification.png", writer)

# %%
