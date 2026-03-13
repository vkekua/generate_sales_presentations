import pandas as pd
from config import (
    INPUT_FILE,
    TV_METRICS_BY_CHANNEL, TV_METRICS_OTHER, TV_SUMMARY_METRICS,
    OTT_METRICS, SMM_METRICS,
)


# ──────────────────────────────────────────────
# Load & clean master dataframe
# ──────────────────────────────────────────────
def load_data() -> pd.DataFrame:
    sheets = pd.read_excel(INPUT_FILE, sheet_name=["TV", "SMM", "OTT"])
    df = pd.concat(sheets.values(), ignore_index=True)

    cols = ["Date", "Partner", "Country", "Content", "Platform", "Channel", "Metric", "Value", "Rubric"]
    df = df[cols]

    # Drop blank / non-numeric values
    df = df[pd.to_numeric(df["Value"], errors="coerce").notna()]
    df = df[df["Metric"].notna() & (df["Metric"].str.strip() != "")]

    # Cast types
    df["Date"]  = pd.to_datetime(df["Date"], errors="coerce").dt.to_period("M").dt.to_timestamp()
    df["Value"] = pd.to_numeric(df["Value"], errors="coerce").round(0).astype("Int64")

    # Clean string columns
    for col in ["Partner", "Country", "Content", "Platform", "Channel", "Metric", "Rubric"]:
        df[col] = df[col].astype(str).str.strip()

    df["Channel"]  = df["Channel"].str.title()
    df["Metric"]   = df["Metric"].str.lower()
    df["Country"]  = df["Country"].str[:2].str.upper()
    df["Partner"]  = df["Partner"].str.upper()

    return df.reset_index(drop=True)


def load_partners(df: pd.DataFrame) -> pd.DataFrame:
    """Return partners where CreatePPT is True."""
    partners = pd.read_excel(INPUT_FILE, sheet_name="Partners")
    return partners[partners["CreatePPT"] == True].reset_index(drop=True)


# ──────────────────────────────────────────────
# Per-partner data preparation
# ──────────────────────────────────────────────
def prepare_partner_data(df: pd.DataFrame, partner: str, country: str, month) -> dict:
    """
    Build all pivot dataframes needed for one partner/country/month combination.
    """
    month_start = pd.Timestamp(month).to_period("M").to_timestamp()
    month_end   = month_start + pd.offsets.MonthEnd(0)

    # Filter by partner + country + month
    mask = (
        (df["Partner"] == partner) &
        (df["Country"] == country) &
        (df["Date"] >= month_start) &
        (df["Date"] <= month_end)
    )
    df_filtered = df[mask]

    df_tv      = df_filtered.query("Platform == 'TV'")
    df_ott     = df_filtered.query("Platform == 'OTT'")
    df_smm     = df_filtered.query("Platform == 'SMM' and Channel != 'Youtube'")
    df_yt      = df_filtered.query("Platform == 'SMM' and Channel == 'Youtube'")
    df_smm_all = df_filtered.query("Platform == 'SMM'")   # ✅ includes Youtube — for Totals

    # ── TV channel pivot (rows = channels, cols = metrics)
    tv_channel = (
        df_tv[df_tv["Metric"].isin(TV_METRICS_BY_CHANNEL)]
        .pivot_table(index="Channel", columns="Metric", values="Value", aggfunc="sum")
        .reset_index()
    )

    # ── TV other metrics pivot (rows = metrics)
    tv_other = (
        df_tv[df_tv["Metric"].isin(TV_METRICS_OTHER)]
        .pivot_table(index="Metric", values="Value", aggfunc="sum")
        .reset_index()
    )

    # ── TV calculated metrics — filtered by partner + country + month
    tv_spots   = df_tv[df_tv["Metric"].isin(["standard spots",         "live ad spots"])        ]["Value"].sum()
    tv_seconds = df_tv[df_tv["Metric"].isin(["standard spots seconds", "live ad spots seconds"])]["Value"].sum()

    tv_summary_other = (
        df_tv[df_tv["Metric"].isin(TV_SUMMARY_METRICS)]
        .pivot_table(index="Metric", values="Value", aggfunc="sum")
        .reset_index()
    )

    tv_totals = pd.concat([
        pd.DataFrame([
            {"Metric": "TV Spots",   "Value": tv_spots},
            {"Metric": "TV Seconds", "Value": tv_seconds},
        ]),
        tv_summary_other,
    ], ignore_index=True)

    totals_order = ["TV Spots", "TV Seconds"] + TV_SUMMARY_METRICS
    tv_totals["Metric"] = pd.Categorical(tv_totals["Metric"], categories=totals_order, ordered=True)
    tv_totals = tv_totals.sort_values("Metric").reset_index(drop=True)

    # ── OTT pivot (rows = metrics)
    ott = (
        df_ott[df_ott["Metric"].isin(OTT_METRICS)]
        .pivot_table(index="Metric", values="Value", aggfunc="sum")
        .reset_index()
    )

    # ── SMM channel pivot — excludes Youtube (for SMM Summary slide)
    smm_channel = (
        df_smm[df_smm["Metric"].isin(SMM_METRICS)]
        .pivot_table(index="Channel", columns="Metric", values="Value", aggfunc="sum")
        .reset_index()
    )

    # ── SMM totals — includes Youtube (for Totals slide)
    smm_totals = (
        df_smm_all[df_smm_all["Metric"].isin(SMM_METRICS)]
        .pivot_table(index="Metric", values="Value", aggfunc="sum")
        .reset_index()
    )

    # ── YouTube rubric pivot — Youtube only (for YouTube slide)
    yt_rubric = (
        df_yt[df_yt["Metric"].isin(SMM_METRICS)]
        .pivot_table(index="Rubric", columns="Metric", values="Value", aggfunc="sum")
        .reset_index()
    )

    # ── Combined totals (TV + OTT + SMM including Youtube)
    tv_totals_tagged  = tv_totals.copy();  tv_totals_tagged["Platform"]  = "TV"
    ott_tagged        = ott.copy();        ott_tagged["Platform"]        = "OTT"
    smm_totals_tagged = smm_totals.copy(); smm_totals_tagged["Platform"] = "SMM"

    all_totals = pd.concat([tv_totals_tagged, ott_tagged, smm_totals_tagged], ignore_index=True)

    return {
        "tv_channel":   tv_channel,
        "tv_other":     tv_other,
        "tv_totals":    tv_totals,
        "ott":          ott,
        "smm_channel":  smm_channel,
        "smm_totals":   smm_totals,
        "yt_rubric":    yt_rubric,
        "all_totals":   all_totals,
        # ── Platform flags
        "has_tv":       not df_tv.empty,
        "has_ott":      not df_ott.empty,
        "has_smm":      not df_smm.empty,
        "has_yt":       not df_yt.empty,
    }