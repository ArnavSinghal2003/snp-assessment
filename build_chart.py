import pandas as pd
import plotly.graph_objects as go
from pathlib import Path

# PATHS 
BASE_DIR = Path(__file__).resolve().parent.parent
RAW_DATA = BASE_DIR / "data" / "raw" / "Designer Developer Assessment General.xlsx"
OUTPUT_CSV = BASE_DIR / "data" / "processed" / "oil_gas_activity.csv"
OUTPUT_HTML = BASE_DIR / "web" / "plotly_chart.html"

# LOAD EXCEL
df = pd.read_excel(RAW_DATA, header=None)

quarters = [
    "2021 Q1","2021 Q2","2021 Q3","2021 Q4",
    "2022 Q1","2022 Q2","2022 Q3","2022 Q4",
    "2023 Q1","2023 Q2","2023 Q3","2023 Q4",
    "2024 Q1","2024 Q2","2024 Q3","2024 Q4",
    "2025 Q1","2025 Q2","2025 Q3","2025 Q4"
]

# EXTRACT DATA SAFELY 
quarter_ago = pd.to_numeric(df.iloc[8, 1:], errors="coerce").dropna().values
year_ago = pd.to_numeric(df.iloc[9, 1:], errors="coerce").dropna().values

quarter_ago = quarter_ago[:20]
year_ago = year_ago[:20]

# BUILD DATAFRAME 
data = pd.DataFrame({
    "Quarter": quarters,
    "Vs Quarter Ago": quarter_ago,
    "Vs Year Ago": year_ago
})

#  RECENT QUARTERS (LAST 8) 
recent_data = data.tail(8)


x_tick_labels = [q.replace(" ", "<br>") for q in data["Quarter"]]

year_ranges = {
    "All": (0, len(data) - 1),
    "2021–2022": (0, 7),
    "2023–2024": (8, 15),
    "2025": (16, 19)
}



# Save for Tableau
data.to_csv(OUTPUT_CSV, index=False)

# PLOTLY CHART
fig = go.Figure()

fig.add_trace(go.Scatter(
    x=data["Quarter"],
    y=data["Vs Quarter Ago"],
    mode="lines+markers",
    name="Vs. a quarter ago",
    line=dict(color="#1f4e79", width=3)
))

fig.add_trace(go.Scatter(
    x=data["Quarter"],
    y=data["Vs Year Ago"],
    mode="lines+markers",
    name="Vs. a year ago",
    line=dict(color="#7f7f7f", width=3, dash="dash")
))

# Year filter buttons positioned inside chart (top-right)
fig.update_layout(
    updatemenus=[
        dict(
            type="buttons",
            direction="right",
            x=0.98,              # right side of chart
            y=0.98,              # inside plot area
            xanchor="right",
            yanchor="top",
            bgcolor="rgba(255,255,255,0.8)",
            bordercolor="#ccc",
            borderwidth=1,
            buttons=[
                dict(
                    label="All",
                    method="relayout",
                    args=[{"xaxis.range": [-0.5, len(data) - 0.5]}]
                ),
                dict(
                    label="2021–2022",
                    method="relayout",
                    args=[{"xaxis.range": [-0.5, 7.5]}]
                ),
                dict(
                    label="2023–2024",
                    method="relayout",
                    args=[{"xaxis.range": [7.5, 15.5]}]
                ),
                dict(
                    label="2025",
                    method="relayout",
                    args=[{"xaxis.range": [15.5, 19.5]}]
                )
            ]
        )
    ]
)




# Zero baseline
fig.add_hline(y=0, line_dash="dot", line_color="black")

# Editorial annotation
fig.add_annotation(
    x="2025 Q4",
    y=-39,
    text="Business activity saw a sharp contraction in Q4 2025",
    showarrow=True,

    arrowhead=2,
    arrowsize=1,
    arrowwidth=1,
    arrowcolor="#999",

    ax=0,        # ⬅️ no horizontal shift
    ay=-130,      # ⬅️ BIG upward move (this fixes it)

    bgcolor="rgba(255,255,255,0.95)",
    bordercolor="#ddd",
    borderwidth=1,

    font=dict(
        size=12,
        color="#333"
    )
)




fig.update_layout(
    title={
        "text": (
            "10th District Oil & Gas Business Activity<br>"
            "<span style='font-size:14px; color:#555;'>"
            "Index score · Quarterly comparison"
            "</span>"
        ),
        "y": 0.96,
        "x": 0.01,
        "xanchor": "left",
        "yanchor": "top"
    },
    margin=dict(
        t=120,
        l=60,
        r=40,
        b=80
    ),

    xaxis=dict(
        title="Quarter",
        tickmode="array",
        tickvals=data["Quarter"],
        ticktext=x_tick_labels,
        tickfont=dict(size=11)
    ),

    yaxis_title="Index Score",
    hovermode="x unified",
    template="simple_white",

    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="left",
        x=0
    )
    

)

# GROUPED BAR CHART (RECENT QUARTERS) 
bar_fig = go.Figure()

bar_fig.add_trace(go.Bar(
    x=recent_data["Quarter"],
    y=recent_data["Vs Quarter Ago"],
    name="Vs. a quarter ago",
    marker_color="#1f4e79"
))

bar_fig.add_trace(go.Bar(
    x=recent_data["Quarter"],
    y=recent_data["Vs Year Ago"],
    name="Vs. a year ago",
    marker_color="#7f7f7f"
))

bar_fig.update_layout(
    title=dict(
        text="Recent Quarter Comparison<br><span style='font-size:14px;color:#555'>Last 8 quarters</span>",
        x=0,
        xanchor="left",
        pad=dict(b=20)
    ),
    barmode="group",
    xaxis_title="Quarter",
    yaxis_title="Index Score",
    template="simple_white",
    yaxis=dict(
        range=[-60, 20],
        zeroline=True,
        zerolinecolor="black",
    ),
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=0.98,
        xanchor="left",
        x=0.80,
        font=dict(size=12)
    ),
    margin=dict(t=70, l=60, r=40, b=60)
)

# Zero baseline
bar_fig.add_hline(y=0, line_dash="dot", line_color="black")

# Export bar chart
bar_fig.write_html(BASE_DIR / "web" / "plotly_bar_chart.html")




# Export interactive HTML
fig.write_html(OUTPUT_HTML)
