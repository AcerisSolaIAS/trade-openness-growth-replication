#!/usr/bin/env python3
"""
Real econometric analysis for trade openness - economic growth paper.
No fabricated numbers. Output is saved to output/ for use in manuscript.
Adds panel unit-root pre-tests (Maddala-Wu Fisher), Engle-Granger residual
cointegration tests, and heterogeneity analysis by development level, time
period, and trade-intensity group.
"""

from pathlib import Path
import numpy as np
import pandas as pd
from scipy import stats
import statsmodels.api as sm
from linearmodels.panel import PooledOLS, PanelOLS
from arch.unitroot import ADF

np.random.seed(42)

REPO_ROOT = Path(__file__).resolve().parent.parent
OUT_DIR = REPO_ROOT / "output"
OUT_DIR.mkdir(exist_ok=True)
DATA_PATH = REPO_ROOT / "data" / "wdi_trade_growth.csv"

# OECD members used as a simple proxy for "developed" economies in the sample.
DEVELOPED_CODES = {
    "AUS", "AUT", "BEL", "CAN", "CHE", "CHL", "CZE", "DEU", "DNK", "ESP",
    "EST", "FIN", "FRA", "GBR", "GRC", "HUN", "IRL", "ISL", "ISR", "ITA",
    "JPN", "KOR", "LTU", "LUX", "LVA", "MEX", "NLD", "NOR", "NZL", "POL",
    "PRT", "SVK", "SVN", "SWE", "TUR", "USA",
}


def load_and_prepare():
    df = pd.read_csv(DATA_PATH)
    df.columns.name = None
    df = df.rename(columns={
        "gdp_per_capita": "gdp",
        "trade_gdp_pct": "trade",
        "investment_gdp_pct": "inv",
        "school_primary": "school",
        "pop_growth": "popg",
    })
    # Drop World Bank regional/income aggregates (common non-country ISO3 codes)
    aggregates = {
        "AFE", "AFW", "ARA", "ARB", "CEB", "EAP", "EAS", "ECA", "ECS", "EUU",
        "HIC", "HPC", "IBD", "IBT", "IDA", "IDB", "IDX", "LAC", "LCN", "LDC",
        "LIC", "LMC", "LMY", "MEA", "MIC", "MNA", "NAC", "OED", "OSS", "PSS",
        "PST", "SAS", "SSA", "SSF", "SST", "TEA", "TEC", "TLA", "TMN", "TSA",
        "TSS", "UMC", "WLD",
    }
    df = df[~df["country_code"].isin(aggregates)]
    # Keep only positive values for variables we log
    df = df[(df["gdp"] > 0) & (df["trade"] > 0) & (df["inv"] > 0)]
    # Logs
    df["lgdp"] = np.log(df["gdp"])
    df["ltrade"] = np.log(df["trade"])
    df["linv"] = np.log(df["inv"])
    df = df.sort_values(["country_code", "year"]).reset_index(drop=True)
    return df


def descriptive_stats(df):
    desc = df.groupby("country_code")[["gdp", "trade", "inv", "school", "popg"]].agg(["mean", "std", "min", "max"])
    desc.columns = [f"{var}_{stat}" for var, stat in desc.columns]
    desc.to_csv(OUT_DIR / "descriptive_stats.csv")
    return desc


def correlations(df):
    corr = df[["lgdp", "ltrade", "linv", "school", "popg"]].corr()
    corr.to_csv(OUT_DIR / "correlation_matrix.csv")
    return corr


def vif_table(df):
    from statsmodels.stats.outliers_influence import variance_inflation_factor
    X = df[["ltrade", "linv", "school", "popg"]].dropna()
    X = sm.add_constant(X)
    vif = pd.DataFrame()
    vif["variable"] = X.columns
    vif["VIF"] = [variance_inflation_factor(X.values, i) for i in range(X.shape[1])]
    vif.to_csv(OUT_DIR / "vif_table.csv", index=False)
    return vif


def adf_per_country(df, var):
    results = []
    for country, g in df.groupby("country_code"):
        y = g[var].dropna()
        if len(y) < 6:
            continue
        try:
            adf = ADF(y, trend="c", lags=1)
            results.append({
                "country_code": country,
                "variable": var,
                "statistic": adf.stat,
                "pvalue": adf.pvalue,
                "lags": adf.lags,
            })
        except Exception as e:
            results.append({
                "country_code": country,
                "variable": var,
                "statistic": np.nan,
                "pvalue": np.nan,
                "lags": np.nan,
                "error": str(e),
            })
    return pd.DataFrame(results)


def _fisher_combined(pvals):
    """Maddala-Wu type Fisher panel unit-root/cointegration test."""
    valid = [p for p in pvals if pd.notna(p) and p > 0]
    if not valid:
        return np.nan, np.nan, 0
    eps = 1e-12
    log_p = np.sum(np.log([max(p, eps) for p in valid]))
    chi2 = -2.0 * log_p
    df2 = 2 * len(valid)
    pvalue = 1.0 - stats.chi2.cdf(chi2, df2)
    return chi2, pvalue, len(valid)


def panel_unit_root_tests(df, vars_to_test=None):
    if vars_to_test is None:
        vars_to_test = ["lgdp", "ltrade", "linv"]
    rows = []
    for var in vars_to_test:
        pvals = []
        reject = 0
        n = 0
        for country, g in df.groupby("country_code"):
            y = g[var].dropna()
            if len(y) < 6:
                continue
            try:
                adf = ADF(y, trend="c", lags=1)
                p = adf.pvalue
                if pd.notna(p):
                    pvals.append(p)
                    n += 1
                    if p < 0.05:
                        reject += 1
            except Exception:
                pass
        chi2, p_fisher, k = _fisher_combined(pvals)
        rows.append({
            "variable": var,
            "n_countries": n,
            "n_reject_5pct": reject,
            "share_reject_5pct": reject / n if n > 0 else np.nan,
            "mean_pvalue": np.mean(pvals) if pvals else np.nan,
            "fisher_chi2": chi2,
            "fisher_pvalue": p_fisher,
            "fisher_k": k,
        })
    rdf = pd.DataFrame(rows)
    rdf.to_csv(OUT_DIR / "panel_unit_root_results.csv", index=False)
    return rdf


def cointegration_test(df, y_var="lgdp", x_vars=None):
    if x_vars is None:
        x_vars = ["ltrade", "linv", "school", "popg"]
    rows = []
    pvals = []
    for country, g in df.groupby("country_code"):
        g = g.dropna(subset=[y_var] + x_vars)
        if len(g) < 10:
            continue
        y = g[y_var]
        X = sm.add_constant(g[x_vars])
        try:
            resid = sm.OLS(y, X).fit().resid
            adf = ADF(resid, trend="n", lags=1)
            p = adf.pvalue
            rows.append({
                "country_code": country,
                "n_obs": len(g),
                "adf_stat": adf.stat,
                "pvalue": p,
            })
            if pd.notna(p):
                pvals.append(p)
        except Exception as e:
            rows.append({
                "country_code": country,
                "n_obs": len(g),
                "error": str(e),
            })
    chi2, p_fisher, k = _fisher_combined(pvals)
    summary = pd.DataFrame([{
        "variable": f"{y_var}_on_" + "_".join(x_vars),
        "n_countries": len(pvals),
        "n_reject_5pct": sum(1 for p in pvals if p < 0.05),
        "mean_pvalue": np.mean(pvals) if pvals else np.nan,
        "fisher_chi2": chi2,
        "fisher_pvalue": p_fisher,
        "fisher_k": k,
    }])
    country_df = pd.DataFrame(rows)
    country_df.to_csv(OUT_DIR / "cointegration_country_results.csv", index=False)
    summary.to_csv(OUT_DIR / "cointegration_summary.csv", index=False)
    return summary, country_df


def _twfe_fit(df_sub, label):
    """Run TWFE and return tidy rows."""
    if df_sub.empty or df_sub["country_code"].nunique() < 3:
        return []
    df_idx = df_sub.set_index(["country_code", "year"])
    exog = sm.add_constant(df_idx[["ltrade", "linv", "school", "popg"]])
    dependent = df_idx["lgdp"]
    try:
        res = PanelOLS(dependent, exog, entity_effects=True, time_effects=True).fit()
    except Exception as e:
        print(f"  {label} TWFE failed: {e}")
        return []
    rows = []
    for var in ["ltrade", "linv", "school", "popg"]:
        rows.append({
            "sample": label,
            "variable": var,
            "coef": res.params[var],
            "se": res.std_errors[var],
            "pvalue": res.pvalues[var],
        })
    rows.append({"sample": label, "variable": "r2", "coef": res.rsquared, "se": np.nan, "pvalue": np.nan})
    rows.append({"sample": label, "variable": "n_obs", "coef": res.nobs, "se": np.nan, "pvalue": np.nan})
    return rows


def heterogeneity_analysis(df):
    """Split sample by development status, time period, and trade intensity."""
    df = df.copy()
    df["developed"] = df["country_code"].isin(DEVELOPED_CODES)

    # Median trade intensity (country-year) to split high vs low openness
    median_trade = df["trade"].median()
    df["high_trade"] = df["trade"] >= median_trade

    rows = []
    samples = {
        "developed": df[df["developed"]],
        "developing": df[~df["developed"]],
        "pre_2008": df[df["year"] <= 2008],
        "post_2008": df[df["year"] >= 2009],
        "high_trade": df[df["high_trade"]],
        "low_trade": df[~df["high_trade"]],
    }
    for label, sub in samples.items():
        print(f"Running TWFE for {label}: {sub['country_code'].nunique()} countries, {len(sub)} obs")
        rows.extend(_twfe_fit(sub, label))
    rdf = pd.DataFrame(rows)
    rdf.to_csv(OUT_DIR / "heterogeneity_results.csv", index=False)
    return rdf


def panel_regressions(df):
    """Pooled OLS, FE, TE, TWFE for log GDP per capita."""
    df_idx = df.set_index(["country_code", "year"])
    exog = sm.add_constant(df_idx[["ltrade", "linv", "school", "popg"]])
    dependent = df_idx["lgdp"]

    models = {}
    models["pooled_ols"] = PooledOLS(dependent, exog).fit()
    models["fe"] = PanelOLS(dependent, exog, entity_effects=True, time_effects=False).fit()
    models["te"] = PanelOLS(dependent, exog, entity_effects=False, time_effects=True).fit()
    models["twfe"] = PanelOLS(dependent, exog, entity_effects=True, time_effects=True).fit()

    rows = []
    for name, res in models.items():
        for var in ["ltrade", "linv", "school", "popg"]:
            try:
                rows.append({
                    "model": name,
                    "variable": var,
                    "coef": res.params[var],
                    "se": res.std_errors[var],
                    "pvalue": res.pvalues[var],
                })
            except KeyError:
                pass
        rows.append({"model": name, "variable": "r2", "coef": res.rsquared, "se": np.nan, "pvalue": np.nan})
        rows.append({"model": name, "variable": "n_obs", "coef": res.nobs, "se": np.nan, "pvalue": np.nan})
    pd.DataFrame(rows).to_csv(OUT_DIR / "panel_regression_results.csv", index=False)
    return models


def dynamic_panel(df):
    """Dynamic panel with lagged log GDP per capita."""
    df_dyn = df.copy()
    df_dyn["lgdp_lag1"] = df_dyn.groupby("country_code")["lgdp"].shift(1)
    df_dyn["ltrade_lag1"] = df_dyn.groupby("country_code")["ltrade"].shift(1)
    df_dyn = df_dyn.dropna()
    df_dyn = df_dyn.set_index(["country_code", "year"])
    exog = sm.add_constant(df_dyn[["lgdp_lag1", "ltrade_lag1", "linv", "school", "popg"]])
    res = PanelOLS(df_dyn["lgdp"], exog, entity_effects=True, time_effects=True).fit()
    rows = []
    for var in ["lgdp_lag1", "ltrade_lag1", "linv", "school", "popg"]:
        rows.append({
            "model": "dynamic_fe",
            "variable": var,
            "coef": res.params[var],
            "se": res.std_errors[var],
            "pvalue": res.pvalues[var],
        })
    rows.append({"model": "dynamic_fe", "variable": "r2", "coef": res.rsquared, "se": np.nan, "pvalue": np.nan})
    rows.append({"model": "dynamic_fe", "variable": "n_obs", "coef": res.nobs, "se": np.nan, "pvalue": np.nan})
    pd.DataFrame(rows).to_csv(OUT_DIR / "dynamic_panel_results.csv", index=False)
    return res


def robustness_checks(df):
    """Exclude very high trade-to-GDP outliers; use lagged trade; first differences."""
    # Exclude countries with mean trade > 150% of GDP (small open economies / tax havens)
    mean_trade = df.groupby("country_code")["trade"].mean()
    normal_trade_countries = mean_trade[mean_trade <= 150].index
    df_normal = df[df["country_code"].isin(normal_trade_countries)].set_index(["country_code", "year"])
    exog_normal = sm.add_constant(df_normal[["ltrade", "linv", "school", "popg"]])
    res_normal = PanelOLS(df_normal["lgdp"], exog_normal, entity_effects=True, time_effects=True).fit()

    # Lagged trade static
    df_lag = df.copy()
    df_lag["ltrade_lag1"] = df_lag.groupby("country_code")["ltrade"].shift(1)
    df_lag = df_lag.dropna().set_index(["country_code", "year"])
    exog_lag = sm.add_constant(df_lag[["ltrade_lag1", "linv", "school", "popg"]])
    res_lag = PanelOLS(df_lag["lgdp"], exog_lag, entity_effects=True, time_effects=True).fit()

    # First differences
    df_diff = df.copy()
    for col in ["lgdp", "ltrade", "linv"]:
        df_diff[f"d_{col}"] = df_diff.groupby("country_code")[col].diff()
    df_diff["d_school"] = df_diff.groupby("country_code")["school"].diff()
    df_diff["d_popg"] = df_diff.groupby("country_code")["popg"].diff()
    df_diff = df_diff.dropna().set_index(["country_code", "year"])
    exog_diff = sm.add_constant(df_diff[["d_ltrade", "d_linv", "d_school", "d_popg"]])
    res_diff = PanelOLS(df_diff["d_lgdp"], exog_diff, entity_effects=True).fit()

    rows = []
    for label, res in [("excl_high_trade", res_normal), ("lag_trade", res_lag), ("first_diff", res_diff)]:
        for var in res.params.index:
            if var != "const":
                rows.append({
                    "model": label,
                    "variable": var,
                    "coef": res.params[var],
                    "se": res.std_errors[var],
                    "pvalue": res.pvalues[var],
                })
        rows.append({"model": label, "variable": "r2", "coef": res.rsquared, "se": np.nan, "pvalue": np.nan})
        rows.append({"model": label, "variable": "n_obs", "coef": res.nobs, "se": np.nan, "pvalue": np.nan})
    pd.DataFrame(rows).to_csv(OUT_DIR / "robustness_results.csv", index=False)
    return res_normal, res_lag, res_diff


def main():
    df = load_and_prepare()
    print("Data shape:", df.shape)
    print(df.head())

    desc = descriptive_stats(df)
    print("\nDescriptive stats:")
    print(desc.head())

    corr = correlations(df)
    print("\nCorrelation matrix:")
    print(corr)

    vif = vif_table(df)
    print("\nVIF table:")
    print(vif)

    print("\nADF tests:")
    for var in ["lgdp", "ltrade", "linv"]:
        adf_df = adf_per_country(df, var)
        print(f"\n{var}: {len(adf_df)} countries")
        print(adf_df.head())
        adf_df.to_csv(OUT_DIR / f"adf_{var}.csv", index=False)

    print("\nPanel unit-root tests (Maddala-Wu Fisher):")
    ur = panel_unit_root_tests(df)
    print(ur)

    print("\nCointegration tests (Engle-Granger residuals):")
    coint_summary, coint_country = cointegration_test(df)
    print(coint_summary)

    print("\nPanel regressions:")
    models = panel_regressions(df)
    for name, res in models.items():
        print(f"\n{name}:")
        print(res.summary.tables[1])

    print("\nDynamic panel:")
    dyn_res = dynamic_panel(df)
    print(dyn_res.summary.tables[1])

    print("\nRobustness checks:")
    robustness_checks(df)

    print("\nHeterogeneity analysis:")
    het = heterogeneity_analysis(df)
    print(het)

    df.to_csv(OUT_DIR / "analysis_data.csv", index=False)
    print("\nAnalysis complete. Outputs saved to", OUT_DIR)


if __name__ == "__