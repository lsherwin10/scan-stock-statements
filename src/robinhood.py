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

equity_sector_keys = list(equity_by_sector.keys())
equity_sector_vals = [equity_by_sector[key] for key in equity_sector_keys]
div_sector_keys = list(div_incomes_by_sector.keys())
div_sector_vals = [div_incomes_by_sector[key] for key in div_sector_keys]

#%%
fig, (ax1, ax2) = plt.subplots(2, 1)
ax1.pie(equity_sector_vals, labels=equity_sector_keys, textprops={"fontsize": 8})
ax1.set_title("Allocation by Equity")

ax2.pie(div_sector_vals, labels=div_sector_keys, textprops={"fontsize": 8})
ax2.set_title("Allocation by Dividend Income")

plt.show()
