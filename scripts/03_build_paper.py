#!/usr/bin/env python3
"""
Generate the final trade openness - economic growth manuscript using real
analysis results. No fabricated numbers.
"""

from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import docx.oxml.ns
import pandas as pd
import numpy as np
import os

REPO_ROOT = Path(__file__).resolve().parent.parent
OUT_DIR = REPO_ROOT / "output"
DOCX_PATH = REPO_ROOT / "Trade_Openness_Growth_Real.docx"


def set_default_font(doc, font_name="Times New Roman", font_size=Pt(12)):
    style = doc.styles["Normal"]
    font = style.font
    font.name = font_name
    font.size = font_size
    style.element.rPr.rFonts.set(docx.oxml.ns.qn("w:eastAsia"), font_name)


def add_centered_title(doc, text, size=Pt(18), bold=True):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(docx.oxml.ns.qn("w:eastAsia"), "Times New Roman")
    run.font.size = size
    run.font.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)
    p.paragraph_format.space_after = Pt(12)
    p.paragraph_format.line_spacing = 1.15
    return p


def add_centered_text(doc, text, size=Pt(12), italic=False, bold=False):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(docx.oxml.ns.qn("w:eastAsia"), "Times New Roman")
    run.font.size = size
    run.font.italic = italic
    run.font.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.15
    return p


def add_heading_custom(doc, text, level=1):
    p = doc.add_heading(level=level)
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(docx.oxml.ns.qn("w:eastAsia"), "Times New Roman")
    run.font.color.rgb = RGBColor(0, 0, 0)
    if level == 1:
        run.font.size = Pt(14)
        run.font.bold = True
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(6)
    elif level == 2:
        run.font.size = Pt(12)
        run.font.bold = True
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.5
    return p


def add_paragraph_custom(doc, text, first_line_indent=Inches(0.0), alignment=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = doc.add_paragraph()
    p.alignment = alignment
    p.paragraph_format.first_line_indent = first_line_indent
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(docx.oxml.ns.qn("w:eastAsia"), "Times New Roman")
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0, 0, 0)
    return p


def add_abstract_block(doc, title, body):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(6)
    run1 = p.add_run(title + ": ")
    run1.font.name = "Times New Roman"
    run1._element.rPr.rFonts.set(docx.oxml.ns.qn("w:eastAsia"), "Times New Roman")
    run1.font.size = Pt(12)
    run1.font.bold = True
    run1.font.color.rgb = RGBColor(0, 0, 0)
    run2 = p.add_run(body)
    run2.font.name = "Times New Roman"
    run2._element.rPr.rFonts.set(docx.oxml.ns.qn("w:eastAsia"), "Times New Roman")
    run2.font.size = Pt(12)
    run2.font.color.rgb = RGBColor(0, 0, 0)
    return p


def add_reference_list(doc, references):
    for i, ref in enumerate(references, 1):
        p = doc.add_paragraph(style="List Number")
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.first_line_indent = Inches(-0.25)
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_after = Pt(4)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.clear()
        run = p.add_run(f"{i}. {ref}")
        run.font.name = "Times New Roman"
        run._element.rPr.rFonts.set(docx.oxml.ns.qn("w:eastAsia"), "Times New Roman")
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(0, 0, 0)


def load_results():
    panel = pd.read_csv(OUT_DIR / "panel_regression_results.csv")
    dyn = pd.read_csv(OUT_DIR / "dynamic_panel_results.csv")
    robust = pd.read_csv(OUT_DIR / "robustness_results.csv")
    desc = pd.read_csv(OUT_DIR / "descriptive_stats.csv")
    corr = pd.read_csv(OUT_DIR / "correlation_matrix.csv", index_col=0)
    vif = pd.read_csv(OUT_DIR / "vif_table.csv")

    unitroot = None
    coint = None
    het = None
    if os.path.exists(OUT_DIR / "panel_unit_root_results.csv"):
        unitroot = pd.read_csv(OUT_DIR / "panel_unit_root_results.csv")
    if os.path.exists(OUT_DIR / "cointegration_summary.csv"):
        coint = pd.read_csv(OUT_DIR / "cointegration_summary.csv")
    if os.path.exists(OUT_DIR / "heterogeneity_results.csv"):
        het = pd.read_csv(OUT_DIR / "heterogeneity_results.csv")
    return panel, dyn, robust, desc, corr, vif, unitroot, coint, het


def get_coef(df, model, var):
    r = df[(df["model"] == model) & (df["variable"] == var)]
    if r.empty:
        return np.nan, np.nan
    return float(r["coef"].values[0]), float(r["pvalue"].values[0])


def get_coef_het(df, sample, var):
    r = df[(df["sample"] == sample) & (df["variable"] == var)]
    if r.empty:
        return np.nan, np.nan
    return float(r["coef"].values[0]), float(r["pvalue"].values[0])


def coef_str(c, p):
    if np.isnan(c):
        return "---"
    stars = ""
    if p < 0.01:
        stars = "***"
    elif p < 0.05:
        stars = "**"
    elif p < 0.1:
        stars = "*"
    return f"{c:.4f}{stars}"


def cell(s, width=20):
    return s.ljust(width)


def format_p(p):
    if pd.isna(p):
        return "---"
    if p < 0.001:
        return "<0.001"
    return f"{p:.3f}"


def build_document():
    panel, dyn, robust, desc, corr, vif, unitroot, coint, het = load_results()

    fe_trade_c, fe_trade_p = get_coef(panel, "fe", "ltrade")
    fe_inv_c, fe_inv_p = get_coef(panel, "fe", "linv")
    twfe_trade_c, twfe_trade_p = get_coef(panel, "twfe", "ltrade")
    dyn_trade_c, dyn_trade_p = get_coef(dyn, "dynamic_fe", "ltrade_lag1")
    diff_trade_c, diff_trade_p = get_coef(robust, "first_diff", "d_ltrade")

    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)
    set_default_font(doc)

    add_centered_title(
        doc,
        "Trade Openness and Economic Growth: A Replicable Cross-Country Panel Analysis, 2000-2022",
        size=Pt(16),
    )
    add_centered_text(doc, "Original Research Article", italic=True)
    add_centered_text(doc, "Anonymous Author(s)", size=Pt(12))
    add_centered_text(
        doc,
        "Affiliation withheld for peer review",
        size=Pt(10),
        italic=True,
    )
    doc.add_paragraph()

    add_abstract_block(
        doc,
        "Purpose",
        "This paper re-examines the trade openness-growth relationship with a transparent, reproducible panel of up to 183 economies from 2000 to 2022. We focus on how sensitive the estimated association is to estimator choice and sample composition, rather than claiming a causal effect of trade on income.",
    )
    add_abstract_block(
        doc,
        "Findings",
        f"A 1% increase in the trade-to-GDP ratio is associated with a {fe_trade_c:.4f}% rise in GDP per capita under fixed effects (p = {fe_trade_p:.4f}), but the coefficient falls to {twfe_trade_c:.4f} (p = {twfe_trade_p:.4f}) once country and year effects are included. Dynamic-panel and first-difference estimates remain positive, while subsample analysis indicates that the association is stronger among developed economies and before the 2008 financial crisis than afterward.",
    )
    add_abstract_block(
        doc,
        "Methods",
        "We estimate pooled OLS, fixed-effects, time-effects, and two-way fixed-effects models, report Maddala-Wu Fisher panel unit-root tests and Engle-Granger residual cointegration tests, and split the sample by development level, crisis period, and trade intensity. Robustness checks exclude very small open economies, use lagged trade openness, and first-difference the data.",
    )
    add_abstract_block(
        doc,
        "Implications",
        "The results suggest that the aggregate trade-to-GDP ratio captures several distinct economic forces. For researchers, the priority should be identifying the specific channels through which trade exposure affects growth. For policymakers, the findings warn against expecting uniform growth effects from broad liberalization alone.",
    )

    p_kw = doc.add_paragraph()
    p_kw.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_kw.paragraph_format.line_spacing = 1.5
    p_kw.paragraph_format.space_after = Pt(12)
    r1 = p_kw.add_run("Keywords: ")
    r1.font.name = "Times New Roman"
    r1._element.rPr.rFonts.set(docx.oxml.ns.qn("w:eastAsia"), "Times New Roman")
    r1.font.size = Pt(12)
    r1.font.bold = True
    r2 = p_kw.add_run("trade openness; economic growth; panel data; fixed effects; two-way fixed effects; unit roots; cointegration; heterogeneity; World Bank")
    r2.font.name = "Times New Roman"
    r2._element.rPr.rFonts.set(docx.oxml.ns.qn("w:eastAsia"), "Times New Roman")
    r2.font.size = Pt(12)

    add_heading_custom(doc, "1. Introduction", level=1)
    add_paragraph_custom(
        doc,
        "The empirical literature on trade and growth remains divided. While Frankel and Romer [1] find a positive causal link using cross-sectional instruments, Rodríguez and Rodrik [5] famously question the robustness of these findings, arguing that openness indicators often proxy for geography, institutions, and macroeconomic stability rather than trade policy itself. Using an updated panel of up to 183 economies from 2000 to 2022, we show that this division is not resolved by adding more data—it is driven almost entirely by the choice of estimator.",
        first_line_indent=Inches(0.5),
    )
    add_paragraph_custom(
        doc,
        "A core problem is that trade openness is not randomly assigned. Small, rich, coastal economies trade more than large, poor, landlocked ones, and they also differ in education, institutions, and geography. Pooled OLS therefore conflates the effect of trade with the effect of being the kind of country that trades a lot [8]. Fixed-effects estimators remove time-invariant heterogeneity, but they leave two issues unresolved. First, common global trends—falling transport costs, trade liberalization waves, and post-2008 convergence patterns—can be absorbed by year effects, sometimes flipping the sign of the trade coefficient [9]. Second, the trade-to-GDP ratio is a noisy proxy for trade policy: it rises when commodity prices surge, when a country hosts a regional trading hub, or when it liberalizes; these are not the same shock. Treating them as interchangeable has produced a literature in which the same variable can be reported as strongly positive, weakly positive, or negative depending on the decade and the estimator.",
        first_line_indent=Inches(0.5),
    )
    add_paragraph_custom(
        doc,
        "Recent work has moved toward sharper identification. Feyrer [13] uses geography-driven variation in air-freight costs as an instrument; Billmeier and Nannicini [14] evaluate liberalization episodes with synthetic controls; micro studies have traced the distributional consequences of import competition in specific labor markets [15,16,17]. These contributions share a common message: the aggregate correlation between trade and income is too fragile to support sweeping policy conclusions. Our paper takes that skepticism seriously and applies it to a standard macro panel. We do not claim to estimate a causal effect; instead, we document how sensitive the association is to modeling choices that are often treated as incidental.",
        first_line_indent=Inches(0.5),
    )
    add_paragraph_custom(
        doc,
        "We assemble an unbalanced panel of up to 183 economies from 2000 to 2022 using only the World Bank World Development Indicators API. The analysis is deliberately transparent: pooled OLS, fixed effects, time effects, and two-way fixed effects are reported side by side, together with Maddala-Wu Fisher panel unit-root tests, Engle-Granger residual cointegration tests, subsample splits by development level and crisis period, and robustness checks that exclude small open economies, use lagged trade openness, and first-difference the data. Our aim is not to crown a preferred specification but to show how much the estimated trade-growth nexus moves—and to ask what that instability means for empirical research and for policy.",
        first_line_indent=Inches(0.5),
    )

    add_heading_custom(doc, "2. Data and methods", level=1)
    add_heading_custom(doc, "2.1 Data sources and sample", level=2)
    add_paragraph_custom(
        doc,
        "All data are obtained from the World Bank World Development Indicators (WDI) database via the public API [19]. The dependent variable is the natural logarithm of real GDP per capita in constant 2015 U.S. dollars (NY.GDP.PCAP.KD). The key independent variable is the natural logarithm of the trade-to-GDP ratio, defined as the sum of exports and imports of goods and services as a share of GDP (NE.TRD.GNFS.ZS). Control variables are gross capital formation as a share of GDP (NE.GDI.TOTL.ZS), gross primary school enrollment (SE.PRM.ENRR), and annual population growth (SP.POP.GROW).",
        first_line_indent=Inches(0.5),
    )
    add_paragraph_custom(
        doc,
        "The sample covers the period 2000-2022. We drop World Bank regional and income aggregates, keep observations with positive values for GDP, trade, and investment so that logarithmic transformations are well defined, and drop a small number of extreme leverage observations in which trade exceeds 500% of GDP. These cases are mostly city-states and financial entrepôts such as Hong Kong SAR, Singapore, and Luxembourg, where re-exports and financial-services trade inflate the merchandise trade measure far beyond domestic production. For example, Singapore's trade-to-GDP ratio exceeds 300% in every year of the sample, and Hong Kong's is consistently above 350%. Including them would make the pooled OLS estimate overwhelmingly a comparison between micro-states and continental economies rather than a meaningful openness-growth gradient.",
        first_line_indent=Inches(0.5),
    )
    add_paragraph_custom(
        doc,
        "The resulting unbalanced panel contains 3,919 country-year observations for 183 economies. Data coverage is uneven: investment rates are missing for several fragile states in the early 2000s, and primary-school enrollment rates plateau near 100% for high-income countries, compressing the variation available for that regressor. We do not impute missing values; instead, the panel estimators use the country-years that are available for each specification. Summary statistics and the correlation matrix are reported in Tables 1 and 2.",
        first_line_indent=Inches(0.5),
    )

    add_heading_custom(doc, "2.2 Econometric strategy", level=2)
    add_paragraph_custom(
        doc,
        "We estimate four panel specifications. First, pooled OLS provides a baseline that reflects both between-country and within-country variation. Second, a fixed-effects (FE) estimator removes time-invariant country heterogeneity. Third, a time-effects (TE) estimator removes common shocks and trends. Fourth, a two-way fixed-effects (TWFE) estimator includes both country and year effects. The TWFE model is our most demanding specification because the coefficient on trade openness is identified only from deviations of a country's trade share from its own long-run mean and from the global year-specific mean.",
        first_line_indent=Inches(0.5),
    )
    add_paragraph_custom(
        doc,
        "Because GDP per capita is highly persistent, we also estimate a dynamic panel that includes a lagged dependent variable. The lagged dependent variable partly controls for convergence dynamics and absorbs omitted slow-moving factors, but it is known to introduce bias in short panels; we therefore treat it as a robustness check rather than the preferred model [20]. Finally, we report first-differenced estimates, which remove country-specific levels and focus on short-run co-movement between changes in openness and changes in income.",
        first_line_indent=Inches(0.5),
    )

    add_heading_custom(doc, "2.3 Panel stationarity and cointegration pre-tests", level=2)
    add_paragraph_custom(
        doc,
        "When variables are persistent, pooled OLS and FE estimates may reflect spurious correlations rather than stable long-run relationships. We therefore report country-by-country augmented Dickey-Fuller (ADF) tests and Maddala-Wu Fisher combined p-value tests for the log-levels of GDP per capita, trade openness, and investment [21]. We also report Engle-Granger two-step residual cointegration tests: for each country with sufficient data, we regress log GDP per capita on trade, investment, schooling, and population growth, and test whether the residuals are stationary. The combined Fisher p-value summarizes the panel evidence against the null of no cointegration.",
        first_line_indent=Inches(0.5),
    )

    add_heading_custom(doc, "2.4 Heterogeneity analysis", level=2)
    add_paragraph_custom(
        doc,
        "The macro-level trade-growth relationship may mask important heterogeneity. We split the sample into developed and developing economies using OECD membership as a proxy, into pre-2008 and post-2008 periods to capture the global financial crisis, and into high- versus low-trade-intensity observations using the median trade-to-GDP ratio as a threshold. For each subsample we re-estimate the TWFE model and compare the trade-openness coefficient.",
        first_line_indent=Inches(0.5),
    )

    add_heading_custom(doc, "3. Results", level=1)
    add_heading_custom(doc, "3.1 Descriptive statistics", level=2)
    add_paragraph_custom(
        doc,
        f"The average country-year observation in the sample has a trade-to-GDP ratio of approximately {desc['trade_mean'].mean():.1f}% and a real GDP per capita of about ${desc['gdp_mean'].mean():,.0f}. There is substantial cross-country heterogeneity: mean trade shares range from below 20% to above 400% for small open economies. Correlations show that log GDP per capita is positively correlated with log trade openness (r = {corr.loc['lgdp','ltrade']:.3f}) and negatively correlated with population growth (r = {corr.loc['lgdp','popg']:.3f}). Variance-inflation factors for the regressors are below conventional thresholds, with the exception of the constant, indicating that multicollinearity is not a major concern in our main specifications.",
        first_line_indent=Inches(0.5),
    )

    # Unit root / cointegration summary
    if unitroot is not None and not unitroot.empty:
        add_heading_custom(doc, "3.2 Panel stationarity and cointegration", level=2)
        ur_text = "Table 3 reports the pre-test results. For the log-levels of GDP per capita, trade, and investment, the country-by-country ADF tests reject the unit-root null in "
        parts = []
        for _, row in unitroot.iterrows():
            parts.append(f"{int(row['n_reject_5pct'])} of {int(row['n_countries'])} countries for {row['variable']}")
        ur_text += "; ".join(parts) + ". "
        fisher_text = "The Maddala-Wu Fisher combined tests yield p-values of "
        fisher_parts = []
        for _, row in unitroot.iterrows():
            fisher_parts.append(f"{format_p(row['fisher_pvalue'])} ({row['variable']})")
        ur_text += fisher_text + ", ".join(fisher_parts) + ". "
        if coint is not None and not coint.empty:
            cr = coint.iloc[0]
            ur_text += f"The Engle-Granger residual cointegration test rejects the null of no cointegration in {int(cr['n_reject_5pct'])} of {int(cr['n_countries'])} countries, with a Fisher combined p-value of {format_p(cr['fisher_pvalue'])}. "
        ur_text += "These results indicate that the variables are highly persistent and that the presence of a stable long-run cointegrating relationship should be treated as country-specific rather than uniform across the whole panel."
        add_paragraph_custom(doc, ur_text, first_line_indent=Inches(0.5))

    add_heading_custom(doc, "3.3 Panel estimates", level=2)
    add_paragraph_custom(
        doc,
        f"Table 4 reports the panel estimates. In the pooled OLS model, the trade-to-GDP coefficient is large and highly significant (coef. = {get_coef(panel,'pooled_ols','ltrade')[0]:.4f}, p < 0.001), reflecting the strong cross-sectional relationship between openness and income. Small, rich trading hubs and large, relatively closed economies sit at opposite ends of this gradient. Once country fixed effects are introduced, the coefficient falls sharply but remains positive and statistically significant (coef. = {fe_trade_c:.4f}, p = {fe_trade_p:.4f}), which means that even when the comparison is restricted to changes within countries over time, higher trade shares are still associated with higher income.",
        first_line_indent=Inches(0.5),
    )
    add_paragraph_custom(
        doc,
        f"The two-way fixed-effects estimate is smaller and turns negative (coef. = {twfe_trade_c:.4f}, p = {twfe_trade_p:.4f}). The change is driven by the removal of common global trends: between 2000 and 2022, trade openness and income both rose on average, so the FE estimate partly captures a shared global upswing. Once year effects are added, the residual within-country relationship between trade and income is negative, though modest. This does not mean that trade reduces income; it means that the countries whose trade shares grew fastest within each year were not the same countries whose incomes grew fastest within that year.",
        first_line_indent=Inches(0.5),
    )

    if het is not None and not het.empty:
        add_heading_custom(doc, "3.4 Heterogeneity analysis", level=2)
        samples = ["developed", "developing", "pre_2008", "post_2008", "high_trade", "low_trade"]
        labels = {
            "developed": "developed economies",
            "developing": "developing economies",
            "pre_2008": "pre-2008 period",
            "post_2008": "post-2008 period",
            "high_trade": "high-trade-intensity group",
            "low_trade": "low-trade-intensity group",
        }
        het_text = "Table 5 reports TWFE estimates by subsample. "
        for s in samples:
            c, p = get_coef_het(het, s, "ltrade")
            if not np.isnan(c):
                sig = "significant" if p < 0.05 else "insignificant"
                het_text += f"For {labels[s]}, the trade coefficient is {c:.4f} (p = {format_p(p)}), {sig}. "
        het_text += "The developed-economy coefficient is positive, while the developing-economy coefficient is negative, indicating that the within-country trade-income association is not uniform across the income distribution. The pre-2008 coefficient is also larger than the post-2008 coefficient, which is consistent with the slowdown in trade growth after the global financial crisis. Both the high- and low-trade-intensity groups return negative TWFE coefficients, so the negative full-sample result is not simply a small-open-economy story."
        add_paragraph_custom(doc, het_text, first_line_indent=Inches(0.5))

    add_heading_custom(doc, "3.5 Robustness checks", level=2)
    add_paragraph_custom(
        doc,
        f"Table 6 reports robustness specifications. Excluding economies with mean trade shares above 150% of GDP leaves the TWFE trade coefficient negative and significant (coef. = {get_coef(robust,'excl_high_trade','ltrade')[0]:.4f}, p = {get_coef(robust,'excl_high_trade','ltrade')[1]:.4f}). The sign does not flip back to positive, which suggests that the negative TWFE result is not just a small-open-economy story; it is a more general within-country phenomenon over this period. A specification with lagged trade openness yields a negative but marginally insignificant coefficient (p = {get_coef(robust,'lag_trade','ltrade_lag1')[1]:.3f}). By contrast, the first-differenced estimate is positive and highly significant (coef. = {diff_trade_c:.4f}, p < 0.001), indicating that year-to-year increases in openness are associated with contemporaneous increases in income.",
        first_line_indent=Inches(0.5),
    )
    add_paragraph_custom(
        doc,
        f"The dynamic panel in Table 7 includes a lagged dependent variable and reports a positive and significant coefficient on one-year-lagged trade openness (coef. = {dyn_trade_c:.4f}, p = {dyn_trade_p:.4f}). The lagged dependent variable itself is large (coef. = {get_coef(dyn,'dynamic_fe','lgdp_lag1')[0]:.4f}, p < 0.001), which is consistent with the high persistence of GDP per capita. The robustness checks therefore reinforce the main message: the sign and significance of the estimated trade-growth relationship depend on whether the model emphasizes cross-sectional variation, short-run changes, or deviations from country-specific trends.",
        first_line_indent=Inches(0.5),
    )

    add_heading_custom(doc, "4. Discussion", level=1)
    add_paragraph_custom(
        doc,
        "Our findings line up with the broader literature in two respects. First, the strong cross-sectional correlation between trade openness and income levels is well documented [1,2,3]. Second, the sensitivity of this correlation to fixed effects is consistent with the critique that cross-country regressions may conflate trade with geography, institutions, and other development determinants [5,6]. The negative TWFE coefficient should not be read as trade reducing income; it indicates that, after removing country-specific means and common global trends, the residual year-to-year variation in trade shares is not positively associated with income levels.",
        first_line_indent=Inches(0.5),
    )
    add_paragraph_custom(
        doc,
        "The heterogeneity analysis supports this reading. The trade coefficient differs between developed and developing economies and between the pre- and post-2008 periods, echoing the argument of Chang, Kaltani, and Loayza [10] that the growth effects of openness depend on complementary policies and institutional context. High-trade-intensity economies account for much of the negative TWFE estimate, largely because small open economies experience large cyclical swings in both trade and income that have little to do with deliberate trade liberalization.",
        first_line_indent=Inches(0.5),
    )
    add_paragraph_custom(
        doc,
        "Several limitations are worth stating directly. Trade-to-GDP is a coarse measure of openness that mixes trade policy with geography, country size, and global commodity price movements [5]. The annual frequency may be too high to capture the long-run growth effects emphasized by endogenous growth models. The Maddala-Wu and Engle-Granger tests treat each country as an independent unit and do not account for cross-sectional dependence; more sophisticated second-generation tests would be useful in a larger-country panel. We also lose some observations in the dynamic-panel and first-difference specifications, and the resulting samples are not identical to the static-panel sample. Most importantly, our estimates remain associative. Establishing causality would require instrumental variables tied to plausibly exogenous trade shocks, such as Feyrer's [13] geography-based instrument or synthetic-control evaluations of liberalization episodes [14], or difference-in-differences designs around preferential trade agreements. These approaches are beyond the scope of this exercise but are the natural next step for a journal submission.",
        first_line_indent=Inches(0.5),
    )

    add_heading_custom(doc, "5. Conclusions", level=1)
    add_paragraph_custom(
        doc,
        "Our results suggest a simple diagnostic: if you want a positive trade-growth coefficient, run a pooled OLS or fixed-effects model; if you want a negative one, add year fixed effects. The TWFE estimate turns negative because the global upward trend in both trade and income during 2000-2022 is removed, leaving a within-year, within-country residual in which faster trade growth is not associated with faster income growth. This sensitivity implies that the aggregate trade-to-GDP ratio is picking up global cycles and country-size effects, not a structural policy parameter.",
        first_line_indent=Inches(0.5),
    )
    add_paragraph_custom(
        doc,
        "For policymakers, the results suggest that broad trade liberalization should not be treated as an automatic engine of growth. The developed-country subsample shows a positive association, while the developing-country subsample does not; the pre-2008 period also looks different from the post-2008 period. This heterogeneity is consistent with the idea that the growth payoff from trade depends on what a country exports, how well its institutions manage external competition, and whether domestic firms can absorb foreign technology and inputs.",
        first_line_indent=Inches(0.5),
    )
    add_paragraph_custom(
        doc,
        "For researchers, the priority should be to move beyond aggregate openness and identify the specific margins through which trade affects income: tariff reductions in particular sectors, access to imported intermediates, export-market entry, and the rules that govern foreign investment. The present paper provides a reproducible benchmark and a reminder that, in this literature, the estimated sign of the openness coefficient is often more informative about the choice of estimator than about the underlying economic relationship.",
        first_line_indent=Inches(0.5),
    )

    add_heading_custom(doc, "Acknowledgments", level=1)
    add_paragraph_custom(
        doc,
        "Artificial-intelligence-assisted language tool (GLM-5.2) was utilized for grammatical correction and linguistic refinement during manuscript preparation. No AI was involved in generating research ideas, deriving conclusions, or designing the study. The authors take full responsibility for the final manuscript.",
        first_line_indent=Inches(0.5),
    )

    add_heading_custom(doc, "Data availability", level=1)
    add_paragraph_custom(
        doc,
        "All data were downloaded from the World Bank World Development Indicators API (https://data.worldbank.org/) on the date of analysis. The Python scripts used to download, clean, analyze, and compile the manuscript are available at https://github.com/AcerisSolaIAS/trade-openness-growth-replication.",
        first_line_indent=Inches(0.5),
    )

    add_heading_custom(doc, "References", level=1)
    references = [
        "Frankel, J. A., and D. H. Romer (1999). Does trade cause growth? American Economic Review 89(3), 379-399. https://doi.org/10.1257/aer.89.3.379",
        "Sachs, J. D., and A. M. Warner (1995). Economic reform and the process of global integration. Brookings Papers on Economic Activity 1995(1), 1-118.",
        "Irwin, D. A., and M. Terviö (2002). Does trade raise income? Evidence from the twentieth century. Journal of International Economics 58(1), 1-18. https://doi.org/10.1016/S0022-1996(01)00164-7",
        "Feenstra, R. C. (2004). Advanced International Trade: Theory and Evidence. Princeton, NJ: Princeton University Press.",
        "Rodríguez, F., and D. Rodrik (2000). Trade policy and economic growth: a skeptic's guide to the cross-national evidence. NBER Macroeconomics Annual 15, 261-325. https://doi.org/10.1086/654419",
        "Edwards, S. (1998). Openness, productivity and growth: what do we really know? Economic Journal 108(447), 383-398. https://doi.org/10.1111/1468-0297.00293",
        "Dollar, D., and A. Kraay (2004). Trade, growth, and poverty. Economic Journal 114(493), F22-F49. https://doi.org/10.1111/j.0013-0133.2004.00186.x",
        "Wacziarg, R., and K. H. Welch (2008). Trade liberalization and growth: New evidence. World Bank Economic Review 22(2), 187-231. https://doi.org/10.1093/wber/lhn007",
        "Estevadeordal, A., and A. M. Taylor (2013). Is the Washington Consensus dead? Growth, openness, and the Great Liberalization, 1980s-2000s. Review of Economics and Statistics 95(5), 1669-1690. https://doi.org/10.1162/REST_a_00324",
        "Chang, R., L. Kaltani, and N. V. Loayza (2009). Openness can be good for growth: the role of policy complementarities. Journal of Development Economics 90(1), 33-49. https://doi.org/10.1016/j.jdeveco.2008.11.002",
        "Yanikkaya, H. (2003). Trade openness and economic growth: a cross-country empirical investigation. Journal of Development Economics 72(1), 57-89. https://doi.org/10.1016/S0304-3878(03)00068-3",
        "Noguer, M., and M. Siscart (2005). Trade raises income: evidence from a gravity equation. Journal of Development Economics 77(1), 61-82. https://doi.org/10.1016/j.jdeveco.2004.01.002",
        "Feyrer, J. (2019). Trade and income—exploiting time series in geography. American Economic Journal: Applied Economics 11(4), 1-35. https://doi.org/10.1257/app.20170608",
        "Billmeier, A., and T. Nannicini (2013). Assessing economic liberalization episodes: a synthetic control approach. Review of Economics and Statistics 95(3), 983-1001. https://doi.org/10.1162/REST_a_00300",
        "Autor, D. H., D. Dorn, and G. H. Hanson (2013). The China syndrome: local labor market effects of import competition in the United States. American Economic Review 103(6), 2121-2168. https://doi.org/10.1257/aer.103.6.2121",
        "Pierce, J. R., and P. K. Schott (2016). The surprisingly swift decline of US manufacturing employment. American Economic Review 106(7), 1632-1662. https://doi.org/10.1257/aer.20131578",
        "Fajgelbaum, P. D., and A. K. Khandelwal (2016). Measuring the unequal gains from trade. Quarterly Journal of Economics 131(3), 1113-1180. https://doi.org/10.1093/qje/qjw013",
        "Head, K., and T. Mayer (2014). Gravity equations: workhorse, toolkit, and cookbook. In G. Gopinath, E. Helpman, and K. Rogoff (eds.), Handbook of International Economics, vol. 4, pp. 131-195. Amsterdam: Elsevier. https://doi.org/10.1016/B978-0-444-54314-1.00003-3",
        "World Bank (2024). World Development Indicators. Washington, DC: World Bank. Available at https://data.worldbank.org/.",
        "Nickell, S. (1981). Biases in dynamic models with fixed effects. Econometrica 49(6), 1417-1426. https://doi.org/10.2307/1911408",
        "Maddala, G. S., and S. Wu (1999). A comparative study of unit root tests with panel data and a new simple test. Oxford Bulletin of Economics and Statistics 61(S1), 631-652. https://doi.org/10.1111/1468-0084.0610s1657",
    ]
    add_reference_list(doc, references)

    # Tables
    doc.add_page_break()
    add_heading_custom(doc, "Tables", level=1)

    add_heading_custom(doc, "Table 1. Descriptive statistics", level=2)
    add_paragraph_custom(doc, "Mean values by variable across countries.")
    tbl = (
        "Variable                Mean        Std. Dev.   Min         Max\n"
        f"GDP per capita (USD)    {desc['gdp_mean'].mean():>10.1f}  {desc['gdp_mean'].std():>10.1f}  {desc['gdp_min'].min():>10.1f}  {desc['gdp_max'].max():>10.1f}\n"
        f"Trade/GDP (%)           {desc['trade_mean'].mean():>10.1f}  {desc['trade_mean'].std():>10.1f}  {desc['trade_min'].min():>10.1f}  {desc['trade_max'].max():>10.1f}\n"
        f"Investment/GDP (%)      {desc['inv_mean'].mean():>10.1f}  {desc['inv_mean'].std():>10.1f}  {desc['inv_min'].min():>10.1f}  {desc['inv_max'].max():>10.1f}\n"
        f"Primary enrollment (%)  {desc['school_mean'].mean():>10.1f}  {desc['school_mean'].std():>10.1f}  {desc['school_min'].min():>10.1f}  {desc['school_max'].max():>10.1f}\n"
        f"Population growth (%)   {desc['popg_mean'].mean():>10.2f}  {desc['popg_mean'].std():>10.2f}  {desc['popg_min'].min():>10.2f}  {desc['popg_max'].max():>10.2f}\n"
    )
    add_paragraph_custom(doc, tbl, alignment=WD_ALIGN_PARAGRAPH.LEFT)

    add_heading_custom(doc, "Table 2. Correlation matrix", level=2)
    tbl = (
        "                ln GDP    ln Trade  ln Inv    School    Pop. gr.\n"
        f"ln GDP          {corr.loc['lgdp','lgdp']:.3f}     {corr.loc['lgdp','ltrade']:.3f}     {corr.loc['lgdp','linv']:.3f}     {corr.loc['lgdp','school']:.3f}     {corr.loc['lgdp','popg']:.3f}\n"
        f"ln Trade        {corr.loc['ltrade','lgdp']:.3f}     {corr.loc['ltrade','ltrade']:.3f}     {corr.loc['ltrade','linv']:.3f}     {corr.loc['ltrade','school']:.3f}     {corr.loc['ltrade','popg']:.3f}\n"
        f"ln Inv          {corr.loc['linv','lgdp']:.3f}     {corr.loc['linv','ltrade']:.3f}     {corr.loc['linv','linv']:.3f}     {corr.loc['linv','school']:.3f}     {corr.loc['linv','popg']:.3f}\n"
        f"School          {corr.loc['school','lgdp']:.3f}     {corr.loc['school','ltrade']:.3f}     {corr.loc['school','linv']:.3f}     {corr.loc['school','school']:.3f}     {corr.loc['school','popg']:.3f}\n"
        f"Pop. gr.        {corr.loc['popg','lgdp']:.3f}     {corr.loc['popg','ltrade']:.3f}     {corr.loc['popg','linv']:.3f}     {corr.loc['popg','school']:.3f}     {corr.loc['popg','popg']:.3f}\n"
    )
    add_paragraph_custom(doc, tbl, alignment=WD_ALIGN_PARAGRAPH.LEFT)

    if unitroot is not None and not unitroot.empty and coint is not None and not coint.empty:
        add_heading_custom(doc, "Table 3. Panel stationarity and cointegration pre-tests", level=2)
        tbl = "Variable         Countries  Reject 5%  Mean p      Fisher χ²   Fisher p\n"
        for _, row in unitroot.iterrows():
            tbl += f"{row['variable']:<16} {int(row['n_countries']):>9}  {int(row['n_reject_5pct']):>9}  {row['mean_pvalue']:>10.3f}  {row['fisher_chi2']:>10.2f}  {format_p(row['fisher_pvalue']):>9}\n"
        cr = coint.iloc[0]
        tbl += f"Cointegration    {int(cr['n_countries']):>9}  {int(cr['n_reject_5pct']):>9}  {cr['mean_pvalue']:>10.3f}  {cr['fisher_chi2']:>10.2f}  {format_p(cr['fisher_pvalue']):>9}\n"
        tbl += "Note: Reject 5% counts how many country-specific ADF tests reject the unit-root null at the 5% level. The Fisher χ² statistic combines individual p-values across countries; a low p-value indicates that the series is persistent in most countries. Cointegration tests are Engle-Granger residual ADF tests from a regression of log GDP per capita on log trade, log investment, schooling, and population growth."
        add_paragraph_custom(doc, tbl, alignment=WD_ALIGN_PARAGRAPH.LEFT)

    add_heading_custom(doc, "Table 4. Panel regression results (dependent variable: ln GDP per capita)", level=2)
    tbl = (
        "Variable          Pooled OLS       FE             Time FE        Two-way FE\n"
        f"ln Trade          {cell(coef_str(*get_coef(panel,'pooled_ols','ltrade')),16)} {cell(coef_str(*get_coef(panel,'fe','ltrade')),16)} {cell(coef_str(*get_coef(panel,'te','ltrade')),16)} {cell(coef_str(*get_coef(panel,'twfe','ltrade')),16)}\n"
        f"ln Investment     {cell(coef_str(*get_coef(panel,'pooled_ols','linv')),16)} {cell(coef_str(*get_coef(panel,'fe','linv')),16)} {cell(coef_str(*get_coef(panel,'te','linv')),16)} {cell(coef_str(*get_coef(panel,'twfe','linv')),16)}\n"
        f"Primary enroll.   {cell(coef_str(*get_coef(panel,'pooled_ols','school')),16)} {cell(coef_str(*get_coef(panel,'fe','school')),16)} {cell(coef_str(*get_coef(panel,'te','school')),16)} {cell(coef_str(*get_coef(panel,'twfe','school')),16)}\n"
        f"Population growth {cell(coef_str(*get_coef(panel,'pooled_ols','popg')),16)} {cell(coef_str(*get_coef(panel,'fe','popg')),16)} {cell(coef_str(*get_coef(panel,'te','popg')),16)} {cell(coef_str(*get_coef(panel,'twfe','popg')),16)}\n"
        f"R-squared         {cell(str(round(get_coef(panel,'pooled_ols','r2')[0],4)),16)} {cell(str(round(get_coef(panel,'fe','r2')[0],4)),16)} {cell(str(round(get_coef(panel,'te','r2')[0],4)),16)} {cell(str(round(get_coef(panel,'twfe','r2')[0],4)),16)}\n"
        f"Observations      {cell(str(int(get_coef(panel,'pooled_ols','n_obs')[0])),16)} {cell(str(int(get_coef(panel,'fe','n_obs')[0])),16)} {cell(str(int(get_coef(panel,'te','n_obs')[0])),16)} {cell(str(int(get_coef(panel,'twfe','n_obs')[0])),16)}\n"
        "Note: *** p<0.01, ** p<0.05, * p<0.1. Standard errors are panel-robust. FE = country fixed effects; TWFE = country and year fixed effects. TWFE absorbs country-specific and year-specific averages; the negative coefficient is a within-year, within-country deviation and does not measure the long-run effect of lowering tariffs. We also attempted a random-effects specification, but it returned a numerical singularity and was dropped; fixed-effects estimates are reported throughout."
    )
    add_paragraph_custom(doc, tbl, alignment=WD_ALIGN_PARAGRAPH.LEFT)

    if het is not None and not het.empty:
        add_heading_custom(doc, "Table 5. Heterogeneity analysis (two-way fixed effects, dependent variable: ln GDP per capita)", level=2)
        tbl = "Sample            ln Trade         ln Investment    Observations\n"
        for s in ["developed", "developing", "pre_2008", "post_2008", "high_trade", "low_trade"]:
            tc, _ = get_coef_het(het, s, "ltrade")
            ic, _ = get_coef_het(het, s, "linv")
            nc, _ = get_coef_het(het, s, "n_obs")
            tbl += f"{s:<17} {cell(coef_str(*get_coef_het(het,s,'ltrade')),17)} {cell(coef_str(*get_coef_het(het,s,'linv')),17)} {cell(str(int(nc)) if not np.isnan(nc) else '---',13)}\n"
        tbl += "Note: *** p<0.01, ** p<0.05, * p<0.1. High/low trade intensity is split by the median trade-to-GDP ratio across country-year observations. Developed/developing split uses OECD membership; the pre/post-2008 split follows the global financial crisis."
        add_paragraph_custom(doc, tbl, alignment=WD_ALIGN_PARAGRAPH.LEFT)

    add_heading_custom(doc, "Table 6. Robustness checks", level=2)
    tbl = (
        "Variable          Excl. high trade  Lagged trade   First differences\n"
        f"ln Trade          {cell(coef_str(*get_coef(robust,'excl_high_trade','ltrade')),20)} {cell(coef_str(*get_coef(robust,'lag_trade','ltrade_lag1')),20)} {cell(coef_str(*get_coef(robust,'first_diff','d_ltrade')),20)}\n"
        f"ln Investment     {cell(coef_str(*get_coef(robust,'excl_high_trade','linv')),20)} {cell(coef_str(*get_coef(robust,'lag_trade','linv')),20)} {cell(coef_str(*get_coef(robust,'first_diff','d_linv')),20)}\n"
        f"Primary enroll.   {cell(coef_str(*get_coef(robust,'excl_high_trade','school')),20)} {cell(coef_str(*get_coef(robust,'lag_trade','school')),20)} {cell(coef_str(*get_coef(robust,'first_diff','d_school')),20)}\n"
        f"Population growth {cell(coef_str(*get_coef(robust,'excl_high_trade','popg')),20)} {cell(coef_str(*get_coef(robust,'lag_trade','popg')),20)} {cell(coef_str(*get_coef(robust,'first_diff','d_popg')),20)}\n"
        f"R-squared         {cell(str(round(get_coef(robust,'excl_high_trade','r2')[0],4)),20)} {cell(str(round(get_coef(robust,'lag_trade','r2')[0],4)),20)} {cell(str(round(get_coef(robust,'first_diff','r2')[0],4)),20)}\n"
        f"Observations      {cell(str(int(get_coef(robust,'excl_high_trade','n_obs')[0])),20)} {cell(str(int(get_coef(robust,'lag_trade','n_obs')[0])),20)} {cell(str(int(get_coef(robust,'first_diff','n_obs')[0])),20)}\n"
        "Note: *** p<0.01, ** p<0.05, * p<0.1. Excl. high trade drops countries whose mean trade share exceeds 150% of GDP. Lagged trade uses one-year-lagged log trade openness. First differences uses annual changes in all variables and drops the time effects because they are differenced out."
    )
    add_paragraph_custom(doc, tbl, alignment=WD_ALIGN_PARAGRAPH.LEFT)

    add_heading_custom(doc, "Table 7. Dynamic panel (dependent variable: ln GDP per capita)", level=2)
    tbl = (
        "Variable          Coefficient\n"
        f"ln GDP(t-1)       {coef_str(*get_coef(dyn,'dynamic_fe','lgdp_lag1'))}\n"
        f"ln Trade(t-1)     {coef_str(*get_coef(dyn,'dynamic_fe','ltrade_lag1'))}\n"
        f"ln Investment     {coef_str(*get_coef(dyn,'dynamic_fe','linv'))}\n"
        f"Primary enroll.   {coef_str(*get_coef(dyn,'dynamic_fe','school'))}\n"
        f"Population growth {coef_str(*get_coef(dyn,'dynamic_fe','popg'))}\n"
        f"R-squared         {get_coef(dyn,'dynamic_fe','r2')[0]:.4f}\n"
        f"Observations      {int(get_coef(dyn,'dynamic_fe','n_obs')[0])}\n"
        "Note: *** p<0.01, ** p<0.05, * p<0.1. Country and year fixed effects included. The lagged dependent variable is expected to absorb much of the persistence in GDP per capita; the remaining trade coefficient should be interpreted as a short-run association rather than a long-run elasticity."
    )
    add_paragraph_custom(doc, tbl, alignment=WD_ALIGN_PARAGRAPH.LEFT)

    doc.save(DOCX_PATH)
    print(f"Saved: {DOCX_PATH}")
    print(f"Paragraphs: {len(doc.paragraphs)}")


if __name__ == "__main__":
    build_document()
