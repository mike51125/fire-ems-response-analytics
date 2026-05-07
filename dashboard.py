import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(
    page_title="Fire/EMS Operational Performance Dashboard",
    layout="wide"
)

st.title("🚒 Fire/EMS Operational Performance Dashboard")

# =====================================================
# FILE PATHS — LOCAL ONLY
# =====================================================

call_log_file = r"C:\Users\mike5\Desktop\Board Reports\April\April CallLog.csv"
monthly_file = r"C:\Users\mike5\Desktop\Board Reports\April\MonthlyIncidentNumbersApril.csv"
overlap_file = r"C:\Users\mike5\Desktop\Board Reports\April\Overlapping Incidents April 2026.xlsx"

# =====================================================
# LOAD DATA
# =====================================================

calllog = pd.read_csv(call_log_file)
monthly = pd.read_csv(monthly_file)

overlap_df = pd.read_excel(
    overlap_file,
    sheet_name="Fire Incidents",
    header=8
)

calllog.columns = calllog.columns.str.strip()
monthly.columns = monthly.columns.str.strip()
overlap_df.columns = overlap_df.columns.str.strip()

# =====================================================
# CLEAN DATA
# =====================================================

incident_col = "Core incident number"
station_col = "Station"
response_col = "Unit response time"

calllog[incident_col] = calllog[incident_col].astype(str).str.strip()
monthly[incident_col] = monthly[incident_col].astype(str).str.strip()

calllog[station_col] = calllog[station_col].astype(str).str.strip()
monthly[station_col] = monthly[station_col].astype(str).str.strip()

overlap_df["Overlapping"] = pd.to_numeric(
    overlap_df["Overlapping"],
    errors="coerce"
).fillna(0)

# =====================================================
# STATION FILTER
# =====================================================

stations = sorted(monthly[station_col].dropna().unique())

selected_station = st.selectbox(
    "Select Station",
    ["All"] + list(stations)
)

if selected_station != "All":
    calllog_filtered = calllog[calllog[station_col] == selected_station].copy()
    monthly_filtered = monthly[monthly[station_col] == selected_station].copy()
else:
    calllog_filtered = calllog.copy()
    monthly_filtered = monthly.copy()

# =====================================================
# MISSING / INSUFFICIENT DATA
# =====================================================

missing_calls = monthly_filtered[
    ~monthly_filtered[incident_col].isin(calllog_filtered[incident_col])
].copy()

# =====================================================
# RESPONSE TIME / DELAYED CALLS
# =====================================================

calllog_filtered[response_col] = pd.to_numeric(
    calllog_filtered[response_col],
    errors="coerce"
)

delayed_calls = calllog_filtered[
    calllog_filtered[response_col] > 480
].shape[0]

# =====================================================
# SYSTEM STRESS
# Overlapping >= 2 means at least 2 other incidents active
# =====================================================

stress_calls = overlap_df[
    overlap_df["Overlapping"] >= 2
].shape[0]

# =====================================================
# KPI SECTION
# =====================================================

col1, col2, col3, col4, col5 = st.columns(5)

col1.metric(
    "Total Calls",
    monthly_filtered[incident_col].nunique()
)

col2.metric(
    "Analyzed Calls",
    calllog_filtered[incident_col].nunique()
)

col3.metric(
    "Missing / Insufficient",
    missing_calls[incident_col].nunique()
)

col4.metric(
    "System Stress Calls",
    stress_calls
)

col5.metric(
    "Delayed Responses (>480s)",
    delayed_calls
)

st.markdown("---")

# =====================================================
# CALLS BY HOUR
# =====================================================

calllog_filtered["Call DateTime"] = pd.to_datetime(
    calllog_filtered["Date"].astype(str) + " " + calllog_filtered["Time"].astype(str),
    errors="coerce"
)

calllog_filtered["Hour"] = calllog_filtered["Call DateTime"].dt.hour

calls_by_hour = (
    calllog_filtered
    .dropna(subset=["Hour"])
    .groupby("Hour")[incident_col]
    .nunique()
    .reindex(range(24), fill_value=0)
)

st.subheader("📊 Calls by Hour")

fig, ax = plt.subplots(figsize=(11, 4.5))

ax.bar(
    calls_by_hour.index,
    calls_by_hour.values,
    color="#4FA3FF"
)

# Stress window shading
ax.axvspan(14.5, 17.5, color="red", alpha=0.12)
ax.axvspan(19.5, 20.5, color="orange", alpha=0.12)

# Stress window labels
ax.text(
    15.8,
    25.7,
    "Primary Stress Window",
    color="red",
    ha="center",
    fontsize=9
)

ax.text(
    20.2,
    25.7,
    "Secondary Spike",
    color="orange",
    ha="center",
    fontsize=9
)

ax.set_xlabel("Hour of Day", color="white")
ax.set_ylabel("Call Volume", color="white")

ax.set_xticks(range(0, 24, 2))

ax.tick_params(axis="x", colors="white")
ax.tick_params(axis="y", colors="white")

ax.grid(axis="y", alpha=0.3)

fig.patch.set_facecolor("#0E1117")
ax.set_facecolor("#0E1117")

ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
ax.spines["left"].set_color("white")
ax.spines["bottom"].set_color("white")

st.pyplot(fig)

st.markdown("---")

# =====================================================
# STATION CALL COUNTS
# =====================================================

station_counts = (
    monthly
    .groupby(station_col)[incident_col]
    .nunique()
    .reset_index()
)

station_counts.columns = ["Station", "Total Calls"]

missing_by_station = (
    monthly[
        ~monthly[incident_col].isin(calllog[incident_col])
    ]
    .groupby(station_col)[incident_col]
    .nunique()
    .reset_index()
)

missing_by_station.columns = ["Station", "Insufficient Data Calls"]

station_counts = station_counts.merge(
    missing_by_station,
    on="Station",
    how="left"
)

station_counts["Insufficient Data Calls"] = (
    station_counts["Insufficient Data Calls"]
    .fillna(0)
    .astype(int)
)

st.subheader("🏢 Calls by Station")

st.dataframe(
    station_counts,
    use_container_width=True,
    hide_index=True
)

# =====================================================
# KEY FINDINGS
# =====================================================

st.markdown("---")

st.subheader("📊 Key Findings")

st.markdown("""
- Peak system stress occurs between **1500–1700 hours**
- Secondary spike observed around **2000 hours**
- Delayed responses (>480 sec) correlate with **overlapping incidents**
- Missing or insufficient-data calls are included in total call volume but excluded from response-time and overlap metrics
""")