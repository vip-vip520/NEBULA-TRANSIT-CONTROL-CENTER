
# ============================================================
# NEBULA TRANSIT CONTROL CENTER - Streamlit + Plotly + optional PyECharts
# Run:
#   py -3.12 -m streamlit run app_anonymized.py
# Put this file in the same folder as:
#   shock_bus_control_data_anonymized.xlsx
# ============================================================

from pathlib import Path
import math
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components
from html import escape

# -----------------------------
# Page setup
# -----------------------------
st.set_page_config(
    page_title="Nebula Transit Control Center",
    page_icon="🛰️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -----------------------------
# Extreme cyber UI CSS
# -----------------------------
st.markdown("""
<style>
@keyframes pulseGlow {
  0% { box-shadow: 0 0 12px rgba(0,229,255,.25), inset 0 0 12px rgba(255,0,255,.08); }
  50% { box-shadow: 0 0 28px rgba(255,0,255,.35), inset 0 0 18px rgba(0,229,255,.14); }
  100% { box-shadow: 0 0 12px rgba(0,229,255,.25), inset 0 0 12px rgba(255,0,255,.08); }
}
@keyframes scan {
  0% { transform: translateY(-100%); opacity: .2; }
  100% { transform: translateY(100%); opacity: .03; }
}
.stApp {
  background:
    radial-gradient(circle at 10% 10%, rgba(0,229,255,.18), transparent 30%),
    radial-gradient(circle at 90% 15%, rgba(255,0,255,.12), transparent 28%),
    radial-gradient(circle at 50% 95%, rgba(124,255,107,.10), transparent 30%),
    linear-gradient(135deg, #02060A 0%, #06121A 52%, #000000 100%);
  color: #E8FFFF;
}
.stApp:before {
  content: "";
  position: fixed;
  top: -30%;
  left: 0;
  height: 160%;
  width: 100%;
  pointer-events: none;
  background: linear-gradient(to bottom, transparent 0%, rgba(125,249,255,.09) 50%, transparent 100%);
  animation: scan 7s linear infinite;
  z-index: 9999;
}
h1, h2, h3 {
  color: #7DF9FF !important;
  text-shadow: 0 0 12px #00E5FF, 0 0 28px rgba(255,0,255,.55);
}
[data-testid="stSidebar"] {
  background: linear-gradient(180deg, rgba(2,10,16,.92), rgba(8,17,27,.92));
  border-right: 1px solid rgba(125,249,255,.25);
}
.neon-card {
  padding: 18px;
  border-radius: 20px;
  background: linear-gradient(135deg, rgba(0,229,255,.10), rgba(255,0,255,.10));
  border: 1px solid rgba(125,249,255,.35);
  animation: pulseGlow 3.5s ease-in-out infinite;
  min-height: 115px;
}
.neon-label {
  font-size: 12px;
  color: #A7F8FF;
  letter-spacing: 2px;
  text-transform: uppercase;
}
.neon-value {
  font-size: 30px;
  font-weight: 900;
  color: #FFFFFF;
  text-shadow: 0 0 16px #00E5FF;
}
.neon-sub {
  font-size: 12px;
  color: #C8F8FF;
}
hr {
  border: none;
  border-top: 1px solid rgba(125,249,255,.25);
}

/* DataFrame */

[data-testid="stDataFrame"] {
    background-color: rgba(2,10,16,0.95) !important;
    border: 1px solid rgba(125,249,255,.25) !important;
    border-radius: 15px !important;
}

/* Header */

[data-testid="stDataFrame"] th {
    background-color: rgba(0,229,255,.15) !important;
    color: #7DF9FF !important;
    font-weight: bold !important;
}

/* Body */

[data-testid="stDataFrame"] td {
    background-color: rgba(2,10,16,.95) !important;
    color: #E8FFFF !important;
}

/* Hover */

[data-testid="stDataFrame"] tr:hover td {
    background-color: rgba(0,229,255,.12) !important;
}

/* Hide Streamlit top header and remove top white/gray band */
[data-testid="stHeader"] {
  display: none !important;
}

.block-container {
  padding-top: 0rem !important;
  padding-bottom: 1rem !important;
  padding-left: 2rem !important;
  padding-right: 2rem !important;
}

html, body {
  background-color: #02060A !important;
  margin: 0 !important;
  padding: 0 !important;
}

/* Darken dataframe / table containers */
[data-testid="stDataFrame"] {
  background: rgba(2,10,16,0.85) !important;
  border: 1px solid rgba(125,249,255,.25) !important;
  border-radius: 14px !important;
  box-shadow: 0 0 20px rgba(0,229,255,.12);
}

/* Darken expander areas */
.streamlit-expanderHeader {
  background: rgba(2,10,16,0.88) !important;
  color: #7DF9FF !important;
  border-radius: 10px !important;
}

[data-testid="stExpander"] {
  background: rgba(2,10,16,0.72) !important;
  border: 1px solid rgba(125,249,255,.18) !important;
  border-radius: 14px !important;
}

/* Try to blend embedded components such as PyGWalker */
iframe {
  background: #02060A !important;
  border-radius: 14px !important;
}

/* Streamlit Tabs */

button[data-baseweb="tab"] {
    color: #7DF9FF !important;
    font-weight: 700 !important;
    font-size: 18px !important;
    text-shadow: 0 0 10px rgba(0,229,255,.8);
}

/* Selected Tab */

button[data-baseweb="tab"][aria-selected="true"] {
    color: #FFFFFF !important;
    background: rgba(0,229,255,.12) !important;
    border-radius: 10px !important;
    box-shadow: 0 0 15px rgba(0,229,255,.35);
}

/* Tab underline */

[data-testid="stTabs"] {
    background: rgba(2,10,16,.35);
    border-radius: 12px;
}

</style>
""", unsafe_allow_html=True)

# -----------------------------
# Load data
# -----------------------------
@st.cache_data
def load_data():
    # Main anonymized demo data file.
    # Fallback keeps the app usable if you rename the anonymized file back to the old filename.
    possible_paths = [
        Path("shock_bus_control_data_anonymized.xlsx"),
        Path("shock_bus_control_data.xlsx"),
    ]

    path = next((candidate for candidate in possible_paths if candidate.exists()), None)

    if path is None:
        st.error(
            "Cannot find the data file. Put shock_bus_control_data_anonymized.xlsx "
            "in the same folder as this Python file."
        )
        st.stop()

    telemetry = pd.read_excel(path, sheet_name="Telemetry")
    coords = pd.read_excel(path, sheet_name="HubCoordinates")
    flows = pd.read_excel(path, sheet_name="FlowNetwork")
    alerts = pd.read_excel(path, sheet_name="Alerts")
    telemetry["Timestamp"] = pd.to_datetime(telemetry["Timestamp"])
    flows["Timestamp"] = pd.to_datetime(flows["Timestamp"])
    alerts["Timestamp"] = pd.to_datetime(alerts["Timestamp"])
    telemetry["Time Label"] = telemetry["Timestamp"].dt.strftime("%H:%M")
    telemetry["DateTime Label"] = telemetry["Timestamp"].dt.strftime("%Y-%m-%d %H:%M")
    flows["Time Label"] = flows["Timestamp"].dt.strftime("%H:%M")
    return telemetry, coords, flows, alerts

df, coords, flows, alerts = load_data()

# Bright cyberpunk hub colors used across charts.
hub_color_map = {
    "Wildstar Preserve": "#00A6FF",
    "Nova Harbor": "#7DF9FF",
    "Quantum Gardens": "#FF1744",
    "Iron Vista": "#FFB3BA",
    "Crystal Nexus": "#00FFC6",
    "Skyline Circuit": "#A7FF83",
}


# -----------------------------
# Cyberpunk HTML table renderer
# -----------------------------
def render_cyber_table(dataframe, height=520, max_rows=None):
    """Render a dark cyberpunk scroll table with sticky header.
    This avoids Plotly go.Table header/body overlap during scrolling.
    """
    df_view = dataframe.copy()

    if max_rows is not None:
        df_view = df_view.head(max_rows)

    # Format datetime columns cleanly for display.
    for col in df_view.columns:
        if pd.api.types.is_datetime64_any_dtype(df_view[col]):
            df_view[col] = df_view[col].dt.strftime("%Y-%m-%d %H:%M")

    header_html = "".join(
        f"<th>{escape(str(col))}</th>"
        for col in df_view.columns
    )

    body_rows = []
    for _, row in df_view.iterrows():
        cells = []
        for value in row:
            if pd.isna(value):
                display_value = ""
            else:
                display_value = str(value)
            cells.append(f"<td>{escape(display_value)}</td>")
        body_rows.append("<tr>" + "".join(cells) + "</tr>")

    body_html = "".join(body_rows)

    table_html = f"""
    <style>
        .cyber-table-wrap {{
            height: {height}px;
            overflow: auto;
            background: rgba(2, 10, 16, 0.92);
            border: 1px solid rgba(125, 249, 255, 0.35);
            border-radius: 14px;
            box-shadow: 0 0 22px rgba(0, 229, 255, 0.16);
        }}

        .cyber-table {{
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            font-family: "Segoe UI", Arial, sans-serif;
            color: #E8FFFF;
            font-size: 13px;
            table-layout: auto;
        }}

        .cyber-table thead th {{
            position: sticky;
            top: 0;
            z-index: 5;
            background: linear-gradient(180deg, rgba(0, 229, 255, 0.28), rgba(2, 10, 16, 0.98));
            color: #7DF9FF;
            text-shadow: 0 0 8px rgba(0, 229, 255, 0.85);
            font-weight: 800;
            padding: 10px 12px;
            border-bottom: 1px solid rgba(125, 249, 255, 0.55);
            border-right: 1px solid rgba(125, 249, 255, 0.20);
            white-space: nowrap;
        }}

        .cyber-table tbody td {{
            padding: 8px 12px;
            border-bottom: 1px solid rgba(125, 249, 255, 0.15);
            border-right: 1px solid rgba(125, 249, 255, 0.12);
            background: rgba(2, 10, 16, 0.86);
            color: #E8FFFF;
            white-space: nowrap;
            text-align: center;
        }}

        .cyber-table tbody tr:nth-child(even) td {{
            background: rgba(0, 229, 255, 0.055);
        }}

        .cyber-table tbody tr:hover td {{
            background: rgba(0, 229, 255, 0.16);
            color: #FFFFFF;
        }}

        .cyber-table-wrap::-webkit-scrollbar {{
            height: 10px;
            width: 10px;
        }}

        .cyber-table-wrap::-webkit-scrollbar-track {{
            background: rgba(2, 10, 16, 0.95);
        }}

        .cyber-table-wrap::-webkit-scrollbar-thumb {{
            background: rgba(0, 229, 255, 0.45);
            border-radius: 10px;
        }}
    </style>

    <div class="cyber-table-wrap">
        <table class="cyber-table">
            <thead><tr>{header_html}</tr></thead>
            <tbody>{body_html}</tbody>
        </table>
    </div>
    """

    st.markdown(table_html, unsafe_allow_html=True)

# -----------------------------
# Sidebar
# -----------------------------
st.sidebar.title("🎛️ CONTROL PANEL")
hubs = sorted(df["Hub"].unique())
shifts = sorted(df["Shift"].unique())
weather = sorted(df["Weather"].unique())
incident_levels = ["Normal", "Watch", "Alert", "Critical"]

selected_hubs = st.sidebar.multiselect("Hub Signal", hubs, default=hubs)
selected_shifts = st.sidebar.multiselect("Shift Window", shifts, default=shifts)
selected_weather = st.sidebar.multiselect("Weather Mode", weather, default=weather)
selected_incidents = st.sidebar.multiselect("Incident Level", incident_levels, default=incident_levels)

time_min = df["Timestamp"].min().to_pydatetime()
time_max = df["Timestamp"].max().to_pydatetime()
time_range = st.sidebar.slider("Mission Time Range", min_value=time_min, max_value=time_max, value=(time_min, time_max), format="MM/DD HH:mm")

filtered = df[
    df["Hub"].isin(selected_hubs)
    & df["Shift"].isin(selected_shifts)
    & df["Weather"].isin(selected_weather)
    & df["Incident Level"].isin(selected_incidents)
    & (df["Timestamp"] >= pd.Timestamp(time_range[0]))
    & (df["Timestamp"] <= pd.Timestamp(time_range[1]))
].copy()

filtered_flows = flows[
    flows["Source Hub"].isin(selected_hubs)
    & flows["Target Hub"].isin(selected_hubs)
    & (flows["Timestamp"] >= pd.Timestamp(time_range[0]))
    & (flows["Timestamp"] <= pd.Timestamp(time_range[1]))
].copy()

if filtered.empty:
    st.warning("No data after filters.")
    st.stop()

# -----------------------------
# Title
# -----------------------------
st.title("🛰️ NEBULA TRANSIT CONTROL CENTER")
st.caption("Python-powered cyberpunk transit dashboard · Streamlit + Plotly + optional PyECharts")

# -----------------------------
# KPI Cards
# -----------------------------
ridership = int(filtered["Ridership"].sum())
avg_delay = filtered["Avg Delay Min"].mean()
on_time = filtered["On-Time %"].mean()
service = filtered["Service Score"].mean()
critical_count = int((filtered["Incident Level"] == "Critical").sum())

kpi_cols = st.columns(5)
kpis = [
    ("Total Ridership", f"{ridership:,}", "Live passenger energy"),
    ("Avg Delay", f"{avg_delay:.1f} min", "Network pressure"),
    ("On-Time Rate", f"{on_time:.1f}%", "Schedule stability"),
    ("Service Score", f"{service:.1f}", "Overall health"),
    ("Critical Signals", f"{critical_count}", "Attention needed"),
]
for col, (label, value, sub) in zip(kpi_cols, kpis):
    with col:
        st.markdown(f"""
        <div class="neon-card">
          <div class="neon-label">{label}</div>
          <div class="neon-value">{value}</div>
          <div class="neon-sub">{sub}</div>
        </div>
        """, unsafe_allow_html=True)

st.markdown("---")

# -----------------------------
# Tabs
# -----------------------------
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🌌 Command Universe",
    "🕸️ Flow Network",
    "🔥 Heat + Radar",
    "🚨 Alert Room",
    "🧪 Data Explorer"
])

# -----------------------------
# Tab 1: 3D Universe + Animated Bubbles
# -----------------------------
with tab1:
    st.subheader("🌌 3D Traffic Universe")

    fig3d = px.scatter_3d(
        filtered,
        x="X", y="Y", z="Z",
        color="Hub",
        size="Ridership",
        symbol="Incident Level",
        hover_data=["DateTime Label", "Route", "Bus ID", "Driver ID", "Weather", "Load %", "Avg Delay Min", "Service Score"],
        color_discrete_map=hub_color_map,
        opacity=0.92,
        title="Live Transit Particles: Hub Position × Traffic Energy"
    )

    # Strong neon outline so particles stay visible on the dark background.
    fig3d.update_traces(
        marker=dict(
            line=dict(width=1.5, color="rgba(255,255,255,0.80)")
        )
    )
    fig3d.update_layout(
        template="plotly_dark",
        height=720,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(2,10,16,0.92)",
        font=dict(color="#E8FFFF", size=14),
        scene=dict(
            bgcolor="rgba(2,10,16,0.92)",
            xaxis=dict(
                title=dict(text="X Map Position", font=dict(color="#7DF9FF", size=15)),
                tickfont=dict(color="#FFFFFF", size=12),
                gridcolor="rgba(0,229,255,.35)",
                zerolinecolor="#00E5FF"
            ),
            yaxis=dict(
                title=dict(text="Y Map Position", font=dict(color="#FF7CFF", size=15)),
                tickfont=dict(color="#FFFFFF", size=12),
                gridcolor="rgba(255,0,255,.30)",
                zerolinecolor="#FF00FF"
            ),
            zaxis=dict(
                title=dict(text="Traffic Energy", font=dict(color="#A7FF83", size=15)),
                tickfont=dict(color="#FFFFFF", size=12),
                gridcolor="rgba(124,255,107,.25)",
                zerolinecolor="#7CFF6B"
            ),
        ),
        title_font=dict(size=25, color="#7DF9FF"),
        legend=dict(
            title=dict(text="Hub / Incident Level", font=dict(color="#7DF9FF", size=15)),
            bgcolor="rgba(0,0,0,0.55)",
            bordercolor="rgba(125,249,255,.55)",
            borderwidth=1,
            font=dict(color="#FFFFFF", size=14),
            itemclick="toggleothers",
            itemdoubleclick="toggle"
        )
    )
    st.plotly_chart(fig3d, use_container_width=True)

    st.subheader("🎬 Time-Warp Bubble Animation")
    anim = (
        filtered
        .groupby(["Time Label", "Hub"], as_index=False)
        .agg({"Ridership": "sum", "Avg Delay Min": "mean", "Service Score": "mean", "Load %": "mean"})
    )

    # Complete every Time Label x Hub combination.
    # This prevents a Hub from disappearing during animation frames.
    time_labels = sorted(anim["Time Label"].unique())
    hub_labels = sorted(filtered["Hub"].unique())
    full_index = pd.MultiIndex.from_product([time_labels, hub_labels], names=["Time Label", "Hub"])

    anim = (
        anim.set_index(["Time Label", "Hub"])
        .reindex(full_index)
        .reset_index()
    )

    # Fill missing frames with safe baseline values.
    hub_baseline = (
        filtered.groupby("Hub", as_index=True)
        .agg({"Ridership": "median", "Avg Delay Min": "median", "Service Score": "mean", "Load %": "median"})
    )

    for hub in hub_labels:
        mask = anim["Hub"].eq(hub)
        for col in ["Ridership", "Avg Delay Min", "Service Score", "Load %"]:
            anim.loc[mask, col] = anim.loc[mask, col].fillna(hub_baseline.loc[hub, col])

    # Fixed range with extra padding.
    # X axis = Ridership, Y axis = Average Delay Minutes.
    x_min = max(0, anim["Ridership"].min() * 0.78)
    x_max = anim["Ridership"].max() * 1.25
    y_min = max(0, anim["Avg Delay Min"].min() - 3)
    y_max = anim["Avg Delay Min"].max() + 5

    fig_anim = px.scatter(
        anim,
        x="Ridership",
        y="Avg Delay Min",
        size="Load %",
        color="Hub",
        animation_frame="Time Label",
        animation_group="Hub",
        hover_name="Hub",
        hover_data=["Service Score"],
        size_max=72,
        range_x=[x_min, x_max],
        range_y=[y_min, y_max],
        color_discrete_map=hub_color_map,
        title="Ridership vs Delay Through the Day"
    )
    fig_anim.update_layout(
        template="plotly_dark",
        height=620,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(2,10,16,0.88)",
        title_font=dict(color="#7DF9FF"),
        font=dict(color="#C8F8FF"),
        xaxis=dict(
            range=[0, 1400],
            title_font=dict(color="#A7F8FF"),
            tickfont=dict(color="#C8F8FF"),
            gridcolor="rgba(0,229,255,0.18)",
            zerolinecolor="rgba(125,249,255,0.25)"
        ),
        yaxis=dict(
            range=[y_min, y_max],
            title_font=dict(color="#A7F8FF"),
            tickfont=dict(color="#C8F8FF"),
            gridcolor="rgba(255,0,255,0.16)",
            zerolinecolor="rgba(125,249,255,0.25)"
        ),
        legend=dict(
            title=dict(text="Hub", font=dict(color="#7DF9FF", size=15)),
            bgcolor="rgba(0,0,0,0.55)",
            bordercolor="rgba(125,249,255,.55)",
            borderwidth=1,
            font=dict(color="#FFFFFF", size=14)
        )

    )

    fig_anim.update_traces(
        marker=dict(
            opacity=0.86,
            line=dict(width=2.2, color="rgba(255,255,255,0.85)")
        )
    )

    st.plotly_chart(fig_anim, use_container_width=True)

# -----------------------------
# Tab 2: Neon Flow Network
# -----------------------------
with tab2:
    st.subheader("🕸️ Neon Hub Flow Network")

    coord_map = {row["Hub"]: (row["X"], row["Y"]) for _, row in coords.iterrows()}
    flow_agg = (
        filtered_flows
        .groupby(["Source Hub", "Target Hub"], as_index=False)
        .agg({"Passenger Flow": "sum", "Delay Pressure": "mean"})
    )

    fig_net = go.Figure()

    # Add flow lines
    for _, row in flow_agg.iterrows():
        sx, sy = coord_map[row["Source Hub"]]
        tx, ty = coord_map[row["Target Hub"]]
        width = max(1, min(10, row["Passenger Flow"] / 900))
        color = "rgba(0,229,255,.55)" if row["Delay Pressure"] < 8 else "rgba(255,0,255,.65)" if row["Delay Pressure"] < 13 else "rgba(255,70,70,.85)"
        fig_net.add_trace(go.Scatter(
            x=[sx, tx], y=[sy, ty],
            mode="lines",
            line=dict(width=width, color=color),
            hoverinfo="text",
            text=f"{row['Source Hub']} → {row['Target Hub']}<br>Flow: {int(row['Passenger Flow']):,}<br>Delay Pressure: {row['Delay Pressure']:.1f}",
            showlegend=False
        ))

    # Add hub nodes
    hub_stats = filtered.groupby("Hub", as_index=False).agg({"Ridership": "sum", "Avg Delay Min": "mean", "Service Score": "mean"})
    hub_stats["X"] = hub_stats["Hub"].map(lambda h: coord_map[h][0])
    hub_stats["Y"] = hub_stats["Hub"].map(lambda h: coord_map[h][1])
    fig_net.add_trace(go.Scatter(
        x=hub_stats["X"], y=hub_stats["Y"],
        mode="markers+text",
        marker=dict(
            size=np.clip(hub_stats["Ridership"] / 850, 22, 70),
            color=hub_stats["Service Score"],
            colorscale="Turbo",
            line=dict(color="#FFFFFF", width=2),
            showscale=True,
            colorbar=dict(title="Service")
        ),
        text=hub_stats["Hub"],
        textposition="top center",
        textfont=dict(color="#E0FFFF", size=14),
        hovertemplate="<b>%{text}</b><br>Ridership: %{marker.size}<extra></extra>"
    ))

    fig_net.update_layout(
        template="plotly_dark",
        height=720,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(2,10,16,0.88)",
        font=dict(color="#C8F8FF"),
        title="Passenger Flow Between Hubs",
        title_font=dict(color="#7DF9FF", size=24),
        xaxis=dict(visible=False),
        yaxis=dict(visible=False),
    )
    st.plotly_chart(fig_net, use_container_width=True)

    st.subheader("⚡ Route Energy Treemap")
    route_tree = filtered.groupby(["Hub", "Route"], as_index=False).agg({"Ridership": "sum", "Avg Delay Min": "mean"})
    fig_tree = px.treemap(
        route_tree, path=["Hub", "Route"], values="Ridership", color="Avg Delay Min",
        color_continuous_scale="Turbo", title="Ridership Weight + Delay Heat"
    )
    fig_tree.update_layout(template="plotly_dark", height=620, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(2,10,16,0.88)", font=dict(color="#C8F8FF"), title_font=dict(color="#7DF9FF"))
    st.plotly_chart(fig_tree, use_container_width=True)

# -----------------------------
# Tab 3: Heatmaps + Radar
# -----------------------------
with tab3:
    left, right = st.columns([1.05, .95])

    with left:
        st.subheader("🔥 Delay Heat Matrix")
        heat = filtered.pivot_table(index="Hub", columns="Shift", values="Avg Delay Min", aggfunc="mean").fillna(0)
        fig_heat = px.imshow(
            heat, text_auto=".1f", color_continuous_scale="Magma",
            title="Average Delay by Hub and Shift"
        )
        fig_heat.update_layout(
            template="plotly_dark",
            height=520,
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(2,10,16,0.92)",

            font=dict(
                color="#E8FFFF",
                size=16
            ),

            title_font=dict(
                color="#7DF9FF",
                size=24
            ),

            margin=dict(l=80, r=30, t=70, b=80),

            xaxis=dict(
                title=dict(
                    text="Shift",
                    font=dict(color="#7DF9FF", size=22)
                ),
                tickfont=dict(
                    color="#E8FFFF",
                    size=16
                ),
                side="bottom"
            ),

            yaxis=dict(
                title=dict(
                    text="Hub",
                    font=dict(color="#7DF9FF", size=22)
                ),
                tickfont=dict(
                    color="#E8FFFF",
                    size=16
                )
            ),

            coloraxis_colorbar=dict(
                title=dict(
                    text="Delay Min",
                    font=dict(color="#7DF9FF", size=14)
                ),
                tickfont=dict(color="#E8FFFF", size=12)
            )
        )

        fig_heat.update_traces(
            textfont=dict(color="#FFFFFF", size=15)
        )
        st.plotly_chart(fig_heat, use_container_width=True)

    with right:
        st.subheader("📡 System Health Radar")
        radar_df = pd.DataFrame({
            "Metric": ["On-Time", "Service", "Low Delay", "Load Balance", "Fuel Efficiency"],
            "Value": [
                on_time,
                service,
                max(0, 100 - avg_delay * 5),
                max(0, 100 - abs(filtered["Load %"].mean() - 72)),
                max(0, 100 - filtered["Fuel Gal"].mean() * 5),
            ]
        })
        fig_radar = go.Figure()
        fig_radar.add_trace(go.Scatterpolar(
            r=radar_df["Value"],
            theta=radar_df["Metric"],
            fill="toself",
            line=dict(color="#00E5FF", width=4),
            fillcolor="rgba(0,229,255,.25)",
            name="Health"
        ))
        fig_radar.update_layout(
            template="plotly_dark",
            height=520,
            paper_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#C8F8FF"),
            polar=dict(
                bgcolor="rgba(0,0,0,0)",
                radialaxis=dict(visible=True, range=[0,100], gridcolor="rgba(125,249,255,.20)"),
                angularaxis=dict(gridcolor="rgba(255,0,255,.22)")
            ),
            showlegend=False
        )
        st.plotly_chart(fig_radar, use_container_width=True)

    st.subheader("🌊 Ridership Pulse")
    line = filtered.groupby(["Timestamp", "Hub"], as_index=False).agg({"Ridership": "sum", "Avg Delay Min": "mean"})
    fig_line = px.line(
        line, x="Timestamp", y="Ridership", color="Hub",
        line_shape="spline", title="Ridership Pulse Over Time"
    )
    fig_line.update_traces(line=dict(width=4))
    fig_line.update_layout(
        template="plotly_dark",
        height=560,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(2,10,16,0.88)",
        font=dict(color="#C8F8FF"),
        title_font=dict(color="#7DF9FF"),
        xaxis=dict(gridcolor="rgba(0,229,255,0.18)", tickfont=dict(color="#C8F8FF"), title_font=dict(color="#A7F8FF")),
        yaxis=dict(gridcolor="rgba(255,0,255,0.16)", tickfont=dict(color="#C8F8FF"), title_font=dict(color="#A7F8FF")),
        legend=dict(bgcolor="rgba(2,10,16,0.78)", font=dict(color="#C8F8FF"))
    )
    st.plotly_chart(fig_line, use_container_width=True)

# -----------------------------
# Tab 4: Alert Room
# -----------------------------
with tab4:
    st.subheader("🚨 Alert Room")

    alert_view = filtered[filtered["Incident Level"].isin(["Alert", "Critical"])].copy()
    if alert_view.empty:
        st.success("No active alert signals under current filters.")
    else:
        fig_alert = px.scatter(
            alert_view,
            x="Timestamp", y="Hub",
            size="Avg Delay Min", color="Incident Level",
            hover_data=["Route", "Bus ID", "Driver ID", "Weather", "Load %"],
            color_discrete_map={"Alert": "#FFEA00", "Critical": "#FF1744"},
            title="Alert Events Timeline"
        )
        fig_alert.update_layout(
            template="plotly_dark",
            height=520,
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(2,10,16,0.88)",
            font=dict(color="#C8F8FF"),
            margin=dict(
                l=350,
                r=30,
                t=50,
                b=50
            ),
            title_font=dict(color="#7DF9FF"),
            xaxis=dict(gridcolor="rgba(0,229,255,0.18)", tickfont=dict(color="#C8F8FF"), title_font=dict(color="#A7F8FF")),
            yaxis=dict(gridcolor="rgba(255,0,255,0.16)", tickfont=dict(color="#C8F8FF",size=15), title_font=dict(color="#A7F8FF"),ticklabelposition="outside",ticklabelstandoff=18),
            legend=dict(bgcolor="rgba(2,10,16,0.78)", font=dict(color="#C8F8FF"))
        )
        st.plotly_chart(fig_alert, use_container_width=True)

        display_cols = ["Timestamp", "Hub", "Route", "Bus ID", "Driver ID", "Incident Level", "Avg Delay Min", "Load %", "Weather"]
        dark_table_df = alert_view[display_cols].sort_values("Timestamp")

        # Use custom HTML table instead of Plotly Table.
        # This fixes the header/body overlap when scrolling.
        render_cyber_table(dark_table_df, height=430)

    # Optional pyecharts liquid gauge
    st.subheader("💧 Optional PyECharts Neon Liquid Gauge")
    try:
        from pyecharts.charts import Liquid
        from pyecharts import options as opts
        liquid = (
            Liquid(init_opts=opts.InitOpts(width="100%", height="360px", bg_color="#02060A"))
            .add(
                "Service",
                [round(service / 100, 2)],
                color=["#00E5FF", "#FF00FF"],
                background_color="rgba(0,0,0,0)"
            )
            .set_global_opts(
                title_opts=opts.TitleOpts(title="Service Energy Core", title_textstyle_opts=opts.TextStyleOpts(color="#7DF9FF")),
                legend_opts=opts.LegendOpts(is_show=False)
            )
        )
        components.html(liquid.render_embed(), height=390)
    except Exception as e:
        st.markdown(
            f"""
            <div style="
                background: rgba(2,10,16,0.92);
                color: #7DF9FF;
                border: 1px solid rgba(125,249,255,0.35);
                border-radius: 14px;
                padding: 16px;
                box-shadow: 0 0 18px rgba(0,229,255,0.14);
                font-family: Consolas, monospace;
            ">
                PyECharts liquid gauge did not load, but the main dashboard still works.<br><br>
                {str(e)}
            </div>
            """,
            unsafe_allow_html=True
        )

# -----------------------------
# Tab 5: Explorer
# -----------------------------
with tab5:
    st.subheader("🧪 Raw Data + Optional PyGWalker")
    st.markdown(
        """
        <div style="
            background: rgba(2,10,16,0.92);
            color:#7DF9FF;
            border:1px solid rgba(125,249,255,.25);
            border-radius:12px;
            padding:12px;
            margin-bottom:10px;
            font-size:20px;
            font-weight:bold;
        ">
            Filtered Data
        </div>
        """,
        unsafe_allow_html=True
    )
    dark_table_df = filtered.head(80)

    # Use custom HTML table instead of Plotly Table.
    # This keeps the neon style and prevents header/body overlap while scrolling.
    render_cyber_table(dark_table_df, height=560)

    st.markdown("### Drag-and-drop explorer")
    st.markdown(
        """
        <div style="
            background: rgba(2,10,16,0.92);
            color: #7DF9FF;
            border: 1px solid rgba(125,249,255,0.35);
            border-radius: 14px;
            padding: 16px;
            box-shadow: 0 0 18px rgba(0,229,255,0.14);
            font-family: Consolas, monospace;
        ">
            Cloud deployment mode: PyGWalker is disabled to avoid Streamlit Cloud dependency conflicts.<br><br>
            All main dashboard charts, filters, tables, animations, and KPI panels remain active.
        </div>
        """,
        unsafe_allow_html=True
    )
