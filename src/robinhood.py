import pandas as pd
import robin_stocks.robinhood as r
import util

MONTHLY_DIVS = {"PBA", "STAG", "MAIN"}


def get_holdings():
    r.login()
    return r.build_holdings()


def get_stocks(stocks):
    tickers = list(stocks.keys())
    fundamentals = r.stocks.get_fundamentals(tickers, "dividend_yield")
    sectors = r.stocks.get_fundamentals(tickers, "sector")

    for ticker, div, sector in zip(tickers, fundamentals, sectors):
        stocks[ticker]["dividend_yield"] = float(div) if div is not None else 0.0
        stocks[ticker]["sector"] = sector
        dividend_freq = "N/A"
        if ticker in MONTHLY_DIVS:
            dividend_freq = "Monthly"
        elif div is not None:
            dividend_freq = "Quarterly"
        stocks[ticker]["dividend_freq"] = dividend_freq
        stocks[ticker]["annual_dividend_per_share"] = float(stocks[ticker]["price"]) * (
            stocks[ticker]["dividend_yield"] / 100
        )
        stocks[ticker]["dividend_per_period"] = util.get_dividend_per_period(
            stocks[ticker]["annual_dividend_per_share"], dividend_freq
        )
        stocks[ticker]["annual_dividend_income"] = stocks[ticker][
            "annual_dividend_per_share"
        ] * float(stocks[ticker]["quantity"])
        stocks[ticker]["dividend_income_per_period"] = util.get_dividend_per_period(
            stocks[ticker]["annual_dividend_income"], dividend_freq
        )

    return stocks


def run():
    stocks = get_holdings()
    stocks = get_stocks(stocks)
    stocks_df = pd.DataFrame.from_dict(stocks, orient="index")
    stocks_df.index.rename("Symbol", inplace=True)

    trimmed_stocks_df = stocks_df.drop(
        ["price", "average_buy_price", "intraday_percent_change", "pe_ratio", "id"],
        axis=1,
    )
    trimmed_stocks_df.columns = [s.replace("_", " ") for s in trimmed_stocks_df.columns]
    trimmed_stocks_df.columns = [
        " ".join([st.capitalize() for st in s.split()])
        for s in trimmed_stocks_df.columns
    ]

    cols = trimmed_stocks_df.columns.tolist()
    cols.remove("Name")
    cols.remove("Type")
    cols.insert(0, "Name")
    cols.insert(1, "Type")
    trimmed_stocks_df = trimmed_stocks_df[cols]
    trimmed_stocks_df.loc["CCL", "Equity Change"] = 0.0

    return trimmed_stocks_df
