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


def compute_percent_equity(df):
    total_equity = df["Equity"].sum()
    df.loc[:, "Percentage"] = round(df["Equity"] / total_equity * 100, 2)
    return df
