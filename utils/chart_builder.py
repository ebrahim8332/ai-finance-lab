"""
Chart builder for FinanceAI Lab — Module 1: Variance Commentary.

Generates five finance-standard charts from a budget vs. actual dataframe.
All charts use Plotly for interactivity. PNG export (for Word doc) uses kaleido.

Charts produced:
  1. Budget vs Actual — grouped bar chart
  2. Variance Bridge — waterfall chart
  3. Variance % — horizontal bar chart, green/red
  4. Cost Mix — donut chart of actual spend by category
  5. Favorable vs Unfavorable — executive summary donut
"""

import io
import re
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio

# ── Brand colors ──────────────────────────────────────────────────────────
BUDGET_COLOR     = "#2e75b6"
ACTUAL_COLOR     = "#1a2744"
FAVORABLE_COLOR  = "#4CAF50"
UNFAVORABLE_COLOR= "#E53935"
NEUTRAL_COLOR    = "#9E9E9E"
SUBTOTAL_COLOR   = "#78909C"


# ── Column detection ──────────────────────────────────────────────────────

def detect_columns(df: pd.DataFrame) -> dict:
    """
    Finds which columns hold the label, budget, actual, variance, and variance %.
    Uses case-insensitive substring matching so it works with most real-world files.
    Returns a dict: {label, budget, actual, variance, variance_pct}
    Any key not found has value None.
    """
    cols = {c: c.lower() for c in df.columns}
    result = {}

    def find(keywords):
        for col, lower in cols.items():
            if any(k in lower for k in keywords):
                return col
        return None

    result["label"]       = find(["item", "description", "account", "category", "line", "name"])
    result["budget"]      = find(["budget", "plan", "target", "bud"])
    result["actual"]      = find(["actual", "real", "result"])
    result["variance"]    = find(["variance", "var $", "var(", "diff"])
    result["variance_pct"]= find(["%", "percent", "pct", "var %", "variance %"])

    # Fallback: if label not found, use first text column
    if not result["label"]:
        for col in df.columns:
            if df[col].dtype == object:
                result["label"] = col
                break

    # Fallback: if budget/actual not found, use first two numeric columns
    numeric = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    if not result["budget"] and len(numeric) >= 1:
        result["budget"] = numeric[0]
    if not result["actual"] and len(numeric) >= 2:
        result["actual"] = numeric[1]

    return result


def prepare_data(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """
    Detects columns, computes Variance and Variance % if missing.
    Returns (cleaned_df, col_map).
    """
    cols = detect_columns(df)
    df = df.copy()

    # Compute variance if not present
    if not cols["variance"] and cols["budget"] and cols["actual"]:
        df["_Variance"] = pd.to_numeric(df[cols["actual"]], errors="coerce") - \
                          pd.to_numeric(df[cols["budget"]], errors="coerce")
        cols["variance"] = "_Variance"

    # Compute variance % if not present
    if not cols["variance_pct"] and cols["budget"] and cols["variance"]:
        budget_vals = pd.to_numeric(df[cols["budget"]], errors="coerce")
        var_vals    = pd.to_numeric(df[cols["variance"]], errors="coerce")
        df["_VarPct"] = (var_vals / budget_vals.replace(0, float("nan"))) * 100
        cols["variance_pct"] = "_VarPct"

    # Ensure numeric columns are actually numeric
    for key in ["budget", "actual", "variance", "variance_pct"]:
        if cols.get(key):
            df[cols[key]] = pd.to_numeric(df[cols[key]], errors="coerce")

    return df, cols


def is_subtotal_row(label: str) -> bool:
    """Returns True if the row looks like a subtotal or summary line."""
    if not isinstance(label, str):
        return False
    keywords = ["total", "gross profit", "ebit", "net income", "operating income",
                "net ", "gross ", "subtotal", "income before"]
    return any(k in label.lower() for k in keywords)


def is_revenue_row(label: str) -> bool:
    """Returns True if the row looks like a revenue line."""
    if not isinstance(label, str):
        return False
    keywords = ["revenue", "sales", "income", "turnover"]
    return any(k in label.lower() for k in keywords)


# ── Chart 1: Budget vs Actual grouped bar ────────────────────────────────

def chart_budget_vs_actual(df: pd.DataFrame, cols: dict) -> go.Figure:
    """
    Grouped bar chart: Budget (blue) vs Actual (dark) for every line item.
    Subtotal rows shown in a muted colour to distinguish them visually.
    """
    label_col   = cols["label"]
    budget_col  = cols["budget"]
    actual_col  = cols["actual"]

    if not all([label_col, budget_col, actual_col]):
        return None

    clean = df.dropna(subset=[budget_col, actual_col]).copy()
    labels  = clean[label_col].astype(str).tolist()
    budgets = clean[budget_col].tolist()
    actuals = clean[actual_col].tolist()

    budget_colors = [SUBTOTAL_COLOR if is_subtotal_row(l) else BUDGET_COLOR for l in labels]
    actual_colors = [SUBTOTAL_COLOR if is_subtotal_row(l) else ACTUAL_COLOR  for l in labels]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="Budget", x=labels, y=budgets,
        marker_color=budget_colors, opacity=0.85,
        hovertemplate="<b>%{x}</b><br>Budget: %{y:,.0f}<extra></extra>",
    ))
    fig.add_trace(go.Bar(
        name="Actual", x=labels, y=actuals,
        marker_color=actual_colors, opacity=0.95,
        hovertemplate="<b>%{x}</b><br>Actual: %{y:,.0f}<extra></extra>",
    ))

    fig.update_layout(
        title="Budget vs Actual",
        barmode="group",
        xaxis_tickangle=-35,
        yaxis_title="USD ($K)",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(t=60, b=120, l=60, r=20),
        font=dict(family="Calibri, sans-serif", size=12),
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(showgrid=True, gridcolor="#eeeeee")
    return fig


# ── Chart 2: Variance Waterfall ───────────────────────────────────────────

def chart_variance_waterfall(df: pd.DataFrame, cols: dict) -> go.Figure:
    """
    Waterfall chart showing how each line item variance contributes to the total.
    Positive bars = favorable (actual > budget for revenue, or actual < budget for cost).
    Negative bars = unfavorable.
    Subtotal rows shown as totals (connecting bars).
    """
    label_col   = cols["label"]
    variance_col = cols["variance"]

    if not all([label_col, variance_col]):
        return None

    clean = df.dropna(subset=[variance_col]).copy()

    labels   = []
    values   = []
    measures = []
    texts    = []

    for _, row in clean.iterrows():
        lbl = str(row[label_col])
        val = row[variance_col]
        if pd.isna(val):
            continue
        labels.append(lbl)
        values.append(val)
        measures.append("total" if is_subtotal_row(lbl) else "relative")
        texts.append(f"{val:+,.0f}")

    if not labels:
        return None

    colors = []
    for lbl, val, meas in zip(labels, values, measures):
        if meas == "total":
            colors.append(SUBTOTAL_COLOR)
        elif val >= 0:
            colors.append(FAVORABLE_COLOR)
        else:
            colors.append(UNFAVORABLE_COLOR)

    fig = go.Figure(go.Waterfall(
        name="Variance",
        orientation="v",
        measure=measures,
        x=labels,
        y=values,
        text=texts,
        textposition="outside",
        connector={"line": {"color": "#cccccc", "width": 1}},
        increasing={"marker": {"color": FAVORABLE_COLOR}},
        decreasing={"marker": {"color": UNFAVORABLE_COLOR}},
        totals={"marker": {"color": SUBTOTAL_COLOR}},
        hovertemplate="<b>%{x}</b><br>Variance: %{y:+,.0f}<extra></extra>",
    ))

    fig.update_layout(
        title="Variance Bridge (Actual vs Budget)",
        yaxis_title="Variance USD ($K)",
        xaxis_tickangle=-35,
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(t=60, b=120, l=60, r=20),
        font=dict(family="Calibri, sans-serif", size=12),
        showlegend=False,
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(showgrid=True, gridcolor="#eeeeee", zeroline=True, zerolinecolor="#aaaaaa")
    return fig


# ── Chart 3: Variance % horizontal bar ───────────────────────────────────

def chart_variance_pct(df: pd.DataFrame, cols: dict) -> go.Figure:
    """
    Horizontal bar chart of variance % per line item, sorted by absolute magnitude.
    Green = favorable (revenue lines positive, cost lines negative).
    Red = unfavorable.
    """
    label_col   = cols["label"]
    varpct_col  = cols["variance_pct"]

    if not all([label_col, varpct_col]):
        return None

    clean = df.dropna(subset=[varpct_col]).copy()
    clean = clean[~clean[label_col].apply(lambda x: is_subtotal_row(str(x)))]
    clean = clean.dropna(subset=[varpct_col])
    clean = clean.reindex(clean[varpct_col].abs().sort_values().index)

    labels = clean[label_col].astype(str).tolist()
    values = clean[varpct_col].tolist()

    # Determine favorable: revenue lines favor positive variance, expense lines favor negative
    bar_colors = []
    for lbl, val in zip(labels, values):
        if is_revenue_row(lbl):
            bar_colors.append(FAVORABLE_COLOR if val >= 0 else UNFAVORABLE_COLOR)
        else:
            bar_colors.append(FAVORABLE_COLOR if val <= 0 else UNFAVORABLE_COLOR)

    fig = go.Figure(go.Bar(
        x=values,
        y=labels,
        orientation="h",
        marker_color=bar_colors,
        text=[f"{v:+.1f}%" for v in values],
        textposition="outside",
        hovertemplate="<b>%{y}</b><br>Variance: %{x:+.1f}%<extra></extra>",
    ))

    fig.add_vline(x=0, line_width=1.5, line_color="#aaaaaa")

    fig.update_layout(
        title="Variance % by Line Item (sorted by magnitude)",
        xaxis_title="Variance %",
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(t=60, b=40, l=200, r=80),
        font=dict(family="Calibri, sans-serif", size=12),
        showlegend=False,
        height=max(350, len(labels) * 36),
    )
    fig.update_xaxes(showgrid=True, gridcolor="#eeeeee", zeroline=False)
    fig.update_yaxes(showgrid=False)
    return fig


# ── Chart 4: Cost mix donut ───────────────────────────────────────────────

def chart_cost_mix(df: pd.DataFrame, cols: dict) -> go.Figure:
    """
    Donut chart of actual spend broken down by expense category.
    Revenue and subtotal rows are excluded.
    Shows where the money actually went.
    """
    label_col  = cols["label"]
    actual_col = cols["actual"]

    if not all([label_col, actual_col]):
        return None

    clean = df.dropna(subset=[actual_col]).copy()
    # Keep only expense rows (exclude revenue and subtotals)
    clean = clean[~clean[label_col].apply(lambda x: is_subtotal_row(str(x)))]
    clean = clean[~clean[label_col].apply(lambda x: is_revenue_row(str(x)))]
    clean = clean[clean[actual_col] > 0]

    if clean.empty:
        return None

    labels = clean[label_col].astype(str).tolist()
    values = clean[actual_col].abs().tolist()

    fig = go.Figure(go.Pie(
        labels=labels,
        values=values,
        hole=0.45,
        textinfo="label+percent",
        textposition="outside",
        hovertemplate="<b>%{label}</b><br>Actual: %{value:,.0f}<br>Share: %{percent}<extra></extra>",
        marker=dict(
            colors=[
                "#2e75b6", "#1a2744", "#4472C4", "#5B9BD5", "#70AD47",
                "#ED7D31", "#A5A5A5", "#FFC000", "#FF0000", "#9467BD",
            ]
        ),
    ))

    fig.update_layout(
        title="Actual Cost Mix by Category",
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(t=60, b=40, l=20, r=20),
        font=dict(family="Calibri, sans-serif", size=11),
        legend=dict(orientation="v", yanchor="middle", y=0.5),
        showlegend=False,
    )
    return fig


# ── Chart 5: Favorable vs Unfavorable summary donut ──────────────────────

def chart_favorable_summary(df: pd.DataFrame, cols: dict) -> go.Figure:
    """
    Executive summary donut: how many line items beat, missed, or matched budget.
    Excludes subtotal rows. One visual that tells the overall story.
    """
    label_col    = cols["label"]
    variance_col = cols["variance"]

    if not all([label_col, variance_col]):
        return None

    clean = df.dropna(subset=[variance_col]).copy()
    clean = clean[~clean[label_col].apply(lambda x: is_subtotal_row(str(x)))]

    # For revenue lines: positive variance = favorable. For expense: negative = favorable.
    favorable = unfavorable = on_target = 0
    for _, row in clean.iterrows():
        lbl = str(row[label_col])
        val = row[variance_col]
        if pd.isna(val):
            continue
        if abs(val) < 0.5:
            on_target += 1
        elif is_revenue_row(lbl):
            if val > 0: favorable += 1
            else:       unfavorable += 1
        else:
            if val < 0: favorable += 1
            else:       unfavorable += 1

    categories = []
    values     = []
    colors_out = []

    if favorable > 0:
        categories.append(f"Favorable ({favorable})")
        values.append(favorable)
        colors_out.append(FAVORABLE_COLOR)
    if unfavorable > 0:
        categories.append(f"Unfavorable ({unfavorable})")
        values.append(unfavorable)
        colors_out.append(UNFAVORABLE_COLOR)
    if on_target > 0:
        categories.append(f"On Target ({on_target})")
        values.append(on_target)
        colors_out.append(NEUTRAL_COLOR)

    if not categories:
        return None

    total = sum(values)
    fig = go.Figure(go.Pie(
        labels=categories,
        values=values,
        hole=0.55,
        marker=dict(colors=colors_out),
        textinfo="label+percent",
        hovertemplate="<b>%{label}</b><br>%{value} line items<extra></extra>",
    ))

    fig.add_annotation(
        text=f"{total}<br>items",
        x=0.5, y=0.5,
        font=dict(size=16, family="Calibri, sans-serif", color="#1a2744"),
        showarrow=False,
    )

    fig.update_layout(
        title="Variance Summary: Favorable vs Unfavorable",
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(t=60, b=40, l=20, r=20),
        font=dict(family="Calibri, sans-serif", size=12),
        legend=dict(orientation="h", yanchor="bottom", y=-0.15),
    )
    return fig


# ── Build all charts ──────────────────────────────────────────────────────

def build_all_charts(df: pd.DataFrame) -> list[tuple[str, object]]:
    """
    Entry point. Prepares data and builds all five charts.
    Returns list of (title, fig) tuples. Skips any chart that fails gracefully.
    """
    df_clean, cols = prepare_data(df)

    chart_fns = [
        ("Budget vs Actual",             lambda: chart_budget_vs_actual(df_clean, cols)),
        ("Variance Bridge",              lambda: chart_variance_waterfall(df_clean, cols)),
        ("Variance % by Line Item",      lambda: chart_variance_pct(df_clean, cols)),
        ("Actual Cost Mix",              lambda: chart_cost_mix(df_clean, cols)),
        ("Favorable vs Unfavorable",     lambda: chart_favorable_summary(df_clean, cols)),
    ]

    results = []
    for title, fn in chart_fns:
        try:
            fig = fn()
            if fig is not None:
                results.append((title, fig))
        except Exception:
            pass  # skip charts that fail on unusual data

    return results


# ── PNG export for Word doc ───────────────────────────────────────────────

def fig_to_png_bytes(fig, width: int = 900, height: int = 450) -> bytes | None:
    """
    Converts a Plotly figure to PNG bytes using kaleido.
    Returns None if kaleido is not installed or export fails.
    """
    try:
        return pio.to_image(fig, format="png", width=width, height=height, scale=2)
    except Exception:
        return None
