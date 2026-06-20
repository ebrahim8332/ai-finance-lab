"""
Chart builder for FinanceAI Lab — Module 1: Variance Commentary.

Generates three finance-standard charts from a budget vs. actual dataframe.
All charts use Plotly for interactivity. PNG export (for Word doc) uses kaleido.

Charts produced:
  1. Variance Bridge   — waterfall by category (Revenue / COGS / Opex)
  2. Budget vs Actual  — grouped bar chart by category
  3. Top 5 Movers      — horizontal bar of the 5 largest individual variances
"""

import json
import re
import io
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio

# ── Brand colors ──────────────────────────────────────────────────────────
BUDGET_COLOR      = "#2e75b6"
ACTUAL_COLOR      = "#1a2744"
FAVORABLE_COLOR   = "#4CAF50"
UNFAVORABLE_COLOR = "#E53935"
NEUTRAL_COLOR     = "#9E9E9E"
SUBTOTAL_COLOR    = "#78909C"

# Category display order and colors for charts
CATEGORY_ORDER  = ["Revenue", "COGS", "Opex", "Other"]
CATEGORY_COLORS = {
    "Revenue": "#2e75b6",
    "COGS":    "#E53935",
    "Opex":    "#ED7D31",
    "Other":   "#9E9E9E",
}


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

    result["label"]        = find(["item", "description", "account", "category", "line", "name"])
    result["budget"]       = find(["budget", "plan", "target", "bud"])
    result["actual"]       = find(["actual", "real", "result"])
    result["variance"]     = find(["variance", "var $", "var(", "diff"])
    result["variance_pct"] = find(["%", "percent", "pct", "var %", "variance %"])

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
        df["_Variance"] = (
            pd.to_numeric(df[cols["actual"]], errors="coerce") -
            pd.to_numeric(df[cols["budget"]], errors="coerce")
        )
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


# ── AI grouping ───────────────────────────────────────────────────────────

_GROUP_SYSTEM = """You classify finance line items into one of four categories.
Return ONLY a JSON object mapping each line item name to its category.
Categories: Revenue, COGS, Opex, Other

Rules:
- Revenue: income, sales, fees earned, grants, service revenue
- COGS: direct costs tied to delivering revenue (materials, subcontractors, direct labour)
- Opex: indirect operating costs (salaries, rent, marketing, R&D, G&A, depreciation, utilities)
- Other: taxes, interest, financing items, anything that doesn't fit the above

Example output:
{"Revenue": "Revenue", "Cost of Goods Sold": "COGS", "Salaries": "Opex", "Interest Expense": "Other"}"""


def group_line_items_ai(labels: list[str], chain) -> dict[str, str]:
    """
    Uses the AI chain to classify line item names into Revenue / COGS / Opex / Other.
    Sends only the names (no numbers) — typically 50-80 tokens.
    Returns a dict mapping label → category.
    Falls back to keyword classification if AI call fails.
    """
    if not labels:
        return {}

    messages = [
        {"role": "system", "content": _GROUP_SYSTEM},
        {"role": "user",   "content": "Classify these line items:\n" + "\n".join(f"- {l}" for l in labels)},
    ]

    try:
        response, _ = chain.complete(messages, timeout=20)
        clean = re.sub(r"```(?:json)?|```", "", response).strip()
        match = re.search(r"\{.*\}", clean, re.DOTALL)
        if not match:
            raise ValueError("No JSON in response")
        mapping = json.loads(match.group(0))
        # Normalise: keep only valid categories
        valid = {"Revenue", "COGS", "Opex", "Other"}
        return {k: (v if v in valid else "Opex") for k, v in mapping.items()}
    except Exception:
        return _group_fallback(labels)


def _group_fallback(labels: list[str]) -> dict[str, str]:
    """Keyword-based fallback when AI grouping fails."""
    REVENUE_KW = ["revenue", "sales", "turnover", "income", "fees"]
    COGS_KW    = ["cost of goods", "cost of sales", "cogs", "direct cost",
                  "direct labour", "direct labor", "subcontract", "materials"]
    result = {}
    for label in labels:
        low = label.lower()
        if any(k in low for k in REVENUE_KW):
            result[label] = "Revenue"
        elif any(k in low for k in COGS_KW):
            result[label] = "COGS"
        else:
            result[label] = "Opex"
    return result


def _aggregate_groups(df: pd.DataFrame, cols: dict, mapping: dict[str, str]) -> dict:
    """
    Sums budget and actual by category using the AI mapping.
    Skips subtotal rows (they would double-count).
    Returns: {category: {budget: float, actual: float, variance: float}}
    """
    label_col  = cols["label"]
    budget_col = cols["budget"]
    actual_col = cols["actual"]

    groups: dict[str, dict] = {}
    for cat in CATEGORY_ORDER:
        groups[cat] = {"budget": 0.0, "actual": 0.0}

    for _, row in df.iterrows():
        lbl = str(row.get(label_col, ""))
        if is_subtotal_row(lbl):
            continue
        cat = mapping.get(lbl, "Opex")
        b = row.get(budget_col)
        a = row.get(actual_col)
        if pd.notna(b):
            groups[cat]["budget"] += float(b)
        if pd.notna(a):
            groups[cat]["actual"] += float(a)

    # Compute variance per group
    for cat in groups:
        groups[cat]["variance"] = groups[cat]["actual"] - groups[cat]["budget"]

    # Drop empty categories
    groups = {k: v for k, v in groups.items() if v["budget"] != 0 or v["actual"] != 0}
    return groups


# ── Shared number formatter ───────────────────────────────────────────────

def _fmt(value: float, sign: bool = False) -> str:
    """
    Formats a number with M / K suffix for chart labels.
    Examples: 4_185_691 → '+4.2M'   -308_402 → '-308K'   521 → '+521'
    """
    prefix = "+" if sign and value >= 0 else ""
    abs_v  = abs(value)
    if abs_v >= 1_000_000:
        s = f"{abs_v / 1_000_000:.1f}M"
    elif abs_v >= 1_000:
        s = f"{abs_v / 1_000:.0f}K"
    else:
        s = f"{abs_v:.0f}"
    neg = "-" if value < 0 else ""
    return f"{prefix}{neg}{s}"


# ── Chart 1: Variance Bridge waterfall ───────────────────────────────────

def chart_waterfall_bridge(groups: dict) -> go.Figure:
    """
    CFO-standard bridge waterfall anchored on total Revenue budget.

    Start bar: Total Revenue Budget (absolute, always positive and easy to read)
    Middle bars: variance per category — how much each category helped or hurt
    End bar: Total Revenue Actual (absolute total)

    Sign convention:
    - Revenue: positive variance (more revenue) = favorable = green
    - COGS / Opex / Other: negative variance (less cost) = favorable = green
    """
    if not groups:
        return None

    rev       = groups.get("Revenue", {"budget": 0, "actual": 0})
    cost_cats = [groups[c] for c in ["COGS", "Opex", "Other"] if c in groups]

    total_rev_b  = rev["budget"]
    total_rev_a  = rev["actual"]
    total_cost_b = sum(c["budget"] for c in cost_cats)
    total_cost_a = sum(c["actual"] for c in cost_cats)

    # Anchor: Revenue budget. End: Revenue actual.
    # Each cost category bar shows its variance contribution (less cost = positive bar).
    # Revenue bar shows its own variance.
    anchor = total_rev_b
    end    = total_rev_a

    labels   = ["Rev Budget"]
    values   = [anchor]
    measures = ["absolute"]
    texts    = [_fmt(anchor)]
    colors   = [SUBTOTAL_COLOR]

    for cat in CATEGORY_ORDER:
        if cat not in groups:
            continue
        g         = groups[cat]
        raw_delta = g["actual"] - g["budget"]

        if cat == "Revenue":
            delta     = raw_delta       # more revenue = positive = green
            favorable = delta >= 0
        else:
            delta     = -raw_delta      # less cost = positive bar = green
            favorable = delta >= 0

        labels.append(f"{cat} Δ")
        values.append(delta)
        measures.append("relative")
        texts.append(_fmt(raw_delta, sign=True))
        colors.append(FAVORABLE_COLOR if favorable else UNFAVORABLE_COLOR)

    labels.append("Rev Actual")
    values.append(end)
    measures.append("total")
    texts.append(_fmt(end))
    colors.append(SUBTOTAL_COLOR)

    fig = go.Figure(go.Waterfall(
        name="",
        orientation="v",
        measure=measures,
        x=labels,
        y=values,
        text=texts,
        textposition="outside",
        textfont=dict(size=11),
        connector={"line": {"color": "#cccccc", "width": 1}},
        increasing={"marker": {"color": FAVORABLE_COLOR}},
        decreasing={"marker": {"color": UNFAVORABLE_COLOR}},
        totals={"marker":    {"color": SUBTOTAL_COLOR}},
        hovertemplate="<b>%{x}</b><br>%{text}<extra></extra>",
    ))

    fig.update_layout(
        title=dict(
            text="Variance Bridge — Budget to Actual<br>"
                 "<sup style='color:#888;font-size:11px'>"
                 "Green = favorable impact. "
                 "Revenue: above budget = green. "
                 "Costs: below budget = green.</sup>",
            x=0,
        ),
        yaxis_title="Amount ($)",
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(t=80, b=60, l=80, r=40),
        font=dict(family="Calibri, sans-serif", size=12),
        showlegend=False,
    )
    # Format y-axis ticks with M/K suffix
    all_vals = [v for v in values if v is not None]
    max_val  = max(abs(v) for v in all_vals) if all_vals else 1
    if max_val >= 1_000_000:
        tick_div, tick_suffix = 1_000_000, "M"
    elif max_val >= 1_000:
        tick_div, tick_suffix = 1_000, "K"
    else:
        tick_div, tick_suffix = 1, ""

    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(
        showgrid=True, gridcolor="#eeeeee",
        zeroline=True, zerolinecolor="#aaaaaa",
        tickvals=[i * tick_div for i in range(
            int(-(max_val * 1.2) // tick_div),
            int((max_val * 1.2) // tick_div) + 2
        )],
        ticktext=[f"{i}{tick_suffix}" for i in range(
            int(-(max_val * 1.2) // tick_div),
            int((max_val * 1.2) // tick_div) + 2
        )],
    )
    return fig


# ── Chart 2: Budget vs Actual by category ────────────────────────────────

def chart_category_bar(groups: dict) -> go.Figure:
    """
    Grouped bar chart showing Budget vs Actual for each category.
    Far cleaner than showing all 14+ individual line items.
    """
    if not groups:
        return None

    cats    = [c for c in CATEGORY_ORDER if c in groups]
    budgets = [groups[c]["budget"] for c in cats]
    actuals = [groups[c]["actual"] for c in cats]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="Budget",
        x=cats,
        y=budgets,
        marker_color=BUDGET_COLOR,
        opacity=0.85,
        hovertemplate="<b>%{x}</b><br>Budget: %{y:,.0f}<extra></extra>",
    ))
    fig.add_trace(go.Bar(
        name="Actual",
        x=cats,
        y=actuals,
        marker_color=ACTUAL_COLOR,
        opacity=0.95,
        hovertemplate="<b>%{x}</b><br>Actual: %{y:,.0f}<extra></extra>",
    ))

    fig.update_layout(
        title="Budget vs Actual by Category",
        barmode="group",
        yaxis_title="Amount ($)",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(t=60, b=60, l=80, r=20),
        font=dict(family="Calibri, sans-serif", size=12),
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(showgrid=True, gridcolor="#eeeeee")
    return fig


# ── Chart 3: Top 5 movers ─────────────────────────────────────────────────

def chart_top_movers(df: pd.DataFrame, cols: dict, groups: dict = None, n: int = 5) -> go.Figure:
    """
    Horizontal bar chart of the N biggest variance movers by absolute dollar impact.

    Uses CATEGORY-level variances (Revenue, COGS, Opex, Other) as the candidate pool
    because those represent the real dollar story. Individual line items that are
    subtotals are excluded to avoid double-counting, but category totals are included
    since they are the primary movers.
    Green = favorable, red = unfavorable (revenue vs cost sign convention applies).
    """
    label_col    = cols["label"]
    variance_col = cols["variance"]

    if not all([label_col, variance_col]):
        return None

    # Build candidate list from category groups first (the real movers)
    candidates = []
    if groups:
        for cat, g in groups.items():
            candidates.append({"label": cat, "variance": g["variance"], "is_revenue": cat == "Revenue"})

    # Set a noise threshold: individual line items must be at least 10% of the
    # largest category variance to appear. This prevents tiny items ($33, $45)
    # from crowding out the real story when multi-million movers exist.
    max_cat_var = max((abs(g["variance"]) for g in groups.values()), default=0) if groups else 0
    noise_threshold = max_cat_var * 0.10

    # Also include individual non-subtotal lines that clear the threshold
    clean = df.dropna(subset=[variance_col]).copy()
    clean = clean[~clean[label_col].apply(lambda x: is_subtotal_row(str(x)))]
    for _, row in clean.iterrows():
        lbl = str(row[label_col])
        var = row[variance_col]
        if pd.notna(var) and abs(float(var)) >= noise_threshold:
            candidates.append({"label": lbl, "variance": float(var), "is_revenue": is_revenue_row(lbl)})

    if not candidates:
        return None

    # Deduplicate by label, keep highest abs variance per label
    seen = {}
    for c in candidates:
        lbl = c["label"]
        if lbl not in seen or abs(c["variance"]) > abs(seen[lbl]["variance"]):
            seen[lbl] = c

    # Sort by absolute variance, take top N
    sorted_items = sorted(seen.values(), key=lambda x: abs(x["variance"]), reverse=True)[:n]
    sorted_items = sorted(sorted_items, key=lambda x: x["variance"])  # ascending for horizontal bar

    labels = [c["label"] for c in sorted_items]
    values = [c["variance"] for c in sorted_items]
    is_rev = [c["is_revenue"] for c in sorted_items]

    # Convert all variances to profit-impact direction:
    # Revenue: positive variance = more revenue = good = keep as-is
    # Costs:   negative variance = less cost = good = negate so favorable goes right
    profit_impact = []
    for rev_flag, val in zip(is_rev, values):
        profit_impact.append(val if rev_flag else -val)

    bar_colors = [FAVORABLE_COLOR if v >= 0 else UNFAVORABLE_COLOR for v in profit_impact]

    # Truncate long labels so they fit in the margin
    def _short(label, maxlen=26):
        return label if len(label) <= maxlen else label[:maxlen - 1] + "…"

    display_labels = [_short(l) for l in labels]
    bar_texts      = [_fmt(v, sign=True) for v in values]

    # Pad x-axis so outside labels are never clipped
    max_abs = max(abs(v) for v in profit_impact) if profit_impact else 1
    x_pad   = max_abs * 0.30

    fig = go.Figure(go.Bar(
        x=profit_impact,
        y=display_labels,
        orientation="h",
        marker_color=bar_colors,
        text=bar_texts,
        textposition="outside",
        textfont=dict(size=11),
        customdata=values,
        hovertemplate="<b>%{y}</b><br>Variance: %{customdata:+,.0f}<extra></extra>",
    ))

    fig.add_vline(x=0, line_width=1.5, line_color="#aaaaaa")

    fig.update_layout(
        title=f"Top {n} Variance Movers — Profit Impact (Actual vs Budget)",
        xaxis_title="Variance ($)",
        xaxis=dict(
            range=[-(max_abs + x_pad), max_abs + x_pad],
            title="Profit Impact  (← Unfavorable  |  Favorable →)",
        ),
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(t=60, b=40, l=190, r=80),
        font=dict(family="Calibri, sans-serif", size=12),
        showlegend=False,
        height=max(320, n * 75 + 120),
    )
    fig.update_xaxes(showgrid=True, gridcolor="#eeeeee", zeroline=False)
    fig.update_yaxes(showgrid=False, tickfont=dict(size=11))
    return fig


# ── Department charts ─────────────────────────────────────────────────────

def is_total_row(label: str) -> bool:
    """Returns True if the row is a grand total line (excluded from dept charts)."""
    if not isinstance(label, str):
        return False
    return label.strip().lower() in {"total", "grand total", "total expenses", "total budget"}


def chart_dept_bar(df: pd.DataFrame, cols: dict) -> go.Figure:
    """
    Horizontal grouped bar — Budget vs Actual for each department.
    Excludes total rows. Sorted by budget descending so largest depts are at top.
    """
    label_col  = cols["label"]
    budget_col = cols["budget"]
    actual_col = cols["actual"]

    clean = df.dropna(subset=[budget_col, actual_col]).copy()
    clean = clean[~clean[label_col].apply(lambda x: is_total_row(str(x)))]
    clean = clean[~clean[label_col].apply(lambda x: is_subtotal_row(str(x)))]
    clean = clean.sort_values(budget_col, ascending=True)  # ascending = largest at top in horizontal bar

    if clean.empty:
        return None

    def _short(label, maxlen=28):
        return label if len(label) <= maxlen else label[:maxlen - 1] + "…"

    labels  = [_short(str(l)) for l in clean[label_col]]
    budgets = clean[budget_col].tolist()
    actuals = clean[actual_col].tolist()

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="Budget",
        y=labels,
        x=budgets,
        orientation="h",
        marker_color=BUDGET_COLOR,
        opacity=0.85,
        hovertemplate="<b>%{y}</b><br>Budget: %{x:,.0f}<extra></extra>",
    ))
    fig.add_trace(go.Bar(
        name="Actual",
        y=labels,
        x=actuals,
        orientation="h",
        marker_color=ACTUAL_COLOR,
        opacity=0.95,
        hovertemplate="<b>%{y}</b><br>Actual: %{x:,.0f}<extra></extra>",
    ))

    fig.update_layout(
        title="Budget vs Actual by Department",
        barmode="group",
        xaxis_title="Amount ($)",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(t=60, b=60, l=180, r=40),
        font=dict(family="Calibri, sans-serif", size=12),
        height=max(360, len(labels) * 55 + 120),
    )
    fig.update_xaxes(showgrid=True, gridcolor="#eeeeee")
    fig.update_yaxes(showgrid=False, tickfont=dict(size=11))
    return fig


def chart_dept_overspend(df: pd.DataFrame, cols: dict) -> go.Figure:
    """
    Ranked horizontal bar showing only departments that exceeded budget.
    Ordered by overspend amount (largest first = top of chart).
    """
    label_col    = cols["label"]
    budget_col   = cols["budget"]
    actual_col   = cols["actual"]
    variance_col = cols["variance"]

    clean = df.dropna(subset=[budget_col, actual_col]).copy()
    clean = clean[~clean[label_col].apply(lambda x: is_total_row(str(x)))]
    clean = clean[~clean[label_col].apply(lambda x: is_subtotal_row(str(x)))]

    # Keep only over-budget rows (actual > budget = positive variance)
    over = clean[pd.to_numeric(clean[variance_col], errors="coerce") > 0].copy()
    over = over.sort_values(variance_col, ascending=True)  # ascending = largest at top

    if over.empty:
        return None

    def _short(label, maxlen=28):
        return label if len(label) <= maxlen else label[:maxlen - 1] + "…"

    labels   = [_short(str(l)) for l in over[label_col]]
    variances = pd.to_numeric(over[variance_col], errors="coerce").tolist()
    pct_vals  = []
    for _, row in over.iterrows():
        b = pd.to_numeric(row[budget_col], errors="coerce")
        v = pd.to_numeric(row[variance_col], errors="coerce")
        pct_vals.append(round((v / b * 100), 1) if b and b != 0 else 0)

    bar_texts = [f"+{_fmt(v)}  ({p:+.1f}%)" for v, p in zip(variances, pct_vals)]

    max_abs = max(abs(v) for v in variances) if variances else 1
    x_pad   = max_abs * 0.35

    fig = go.Figure(go.Bar(
        x=variances,
        y=labels,
        orientation="h",
        marker_color=UNFAVORABLE_COLOR,
        text=bar_texts,
        textposition="outside",
        textfont=dict(size=11),
        hovertemplate="<b>%{y}</b><br>Overspend: %{x:+,.0f}<extra></extra>",
    ))

    fig.update_layout(
        title="Departments Over Budget — Ranked by Overspend",
        xaxis=dict(
            title="Overspend vs Budget ($)",
            range=[0, max_abs + x_pad],
        ),
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(t=60, b=60, l=180, r=100),
        font=dict(family="Calibri, sans-serif", size=12),
        showlegend=False,
        height=max(300, len(labels) * 60 + 120),
    )
    fig.update_xaxes(showgrid=True, gridcolor="#eeeeee")
    fig.update_yaxes(showgrid=False, tickfont=dict(size=11))
    return fig


def chart_dept_movers(df: pd.DataFrame, cols: dict, n: int = 5) -> go.Figure:
    """
    Top N departments by absolute variance.
    No profit-impact sign convention — over budget is always red, under is green.
    """
    label_col    = cols["label"]
    variance_col = cols["variance"]
    budget_col   = cols["budget"]

    clean = df.dropna(subset=[variance_col]).copy()
    clean = clean[~clean[label_col].apply(lambda x: is_total_row(str(x)))]
    clean = clean[~clean[label_col].apply(lambda x: is_subtotal_row(str(x)))]
    clean["_abs_var"] = pd.to_numeric(clean[variance_col], errors="coerce").abs()
    clean = clean.nlargest(n, "_abs_var")
    clean = clean.sort_values(variance_col, ascending=True)

    if clean.empty:
        return None

    def _short(label, maxlen=28):
        return label if len(label) <= maxlen else label[:maxlen - 1] + "…"

    labels    = [_short(str(l)) for l in clean[label_col]]
    variances = pd.to_numeric(clean[variance_col], errors="coerce").tolist()

    # Over budget (positive variance) = red. Under budget = green.
    bar_colors = [UNFAVORABLE_COLOR if v > 0 else FAVORABLE_COLOR for v in variances]
    bar_texts  = [_fmt(v, sign=True) for v in variances]

    max_abs = max(abs(v) for v in variances) if variances else 1
    x_pad   = max_abs * 0.30

    fig = go.Figure(go.Bar(
        x=variances,
        y=labels,
        orientation="h",
        marker_color=bar_colors,
        text=bar_texts,
        textposition="outside",
        textfont=dict(size=11),
        hovertemplate="<b>%{y}</b><br>Variance: %{x:+,.0f}<extra></extra>",
    ))

    fig.add_vline(x=0, line_width=1.5, line_color="#aaaaaa")

    fig.update_layout(
        title=f"Top {n} Department Movers — Actual vs Budget",
        xaxis=dict(
            title="Variance ($)  (← Under Budget  |  Over Budget →)",
            range=[-(max_abs + x_pad), max_abs + x_pad],
        ),
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(t=60, b=40, l=180, r=80),
        font=dict(family="Calibri, sans-serif", size=12),
        showlegend=False,
        height=max(320, n * 75 + 120),
    )
    fig.update_xaxes(showgrid=True, gridcolor="#eeeeee", zeroline=False)
    fig.update_yaxes(showgrid=False, tickfont=dict(size=11))
    return fig


# ── Build all charts ──────────────────────────────────────────────────────

def build_all_charts(df: pd.DataFrame, chain=None) -> dict:
    """
    Entry point. Prepares data, uses AI to group line items by category,
    auto-detects file type (P&L vs Departmental Budget), then builds the
    appropriate chart set.

    Returns:
        {
            "charts":    list of (title, fig) tuples,
            "file_type": "pl" or "departmental",
        }

    chain: the AI FallbackChain instance used for grouping.
           Pass None to use keyword-only fallback grouping.
    """
    df_clean, cols = prepare_data(df)
    label_col = cols.get("label")

    # Get unique non-subtotal line item names for AI grouping
    if label_col:
        all_labels = df_clean[label_col].dropna().astype(str).tolist()
        labels_to_group = [l for l in all_labels if not is_subtotal_row(l) and not is_total_row(l)]
    else:
        labels_to_group = []

    # AI classification: send names only (no numbers)
    if chain and labels_to_group:
        mapping = group_line_items_ai(labels_to_group, chain)
    else:
        mapping = _group_fallback(labels_to_group)

    # Auto-detect file type.
    # Use COGS count as the signal — a real P&L always has cost-of-goods lines.
    # A departmental budget never does. Revenue count alone is unreliable because
    # "Sales" as a department name can be misclassified as Revenue by the AI.
    cogs_count = sum(1 for v in mapping.values() if v == "COGS")
    file_type = "departmental" if cogs_count == 0 else "pl"

    results = []

    if file_type == "pl":
        groups = _aggregate_groups(df_clean, cols, mapping)
        chart_fns = [
            ("Variance Bridge",       lambda: chart_waterfall_bridge(groups)),
            ("Budget vs Actual",      lambda: chart_category_bar(groups)),
            ("Top 5 Variance Movers", lambda: chart_top_movers(df_clean, cols, groups=groups, n=5)),
        ]
    else:
        chart_fns = [
            ("Budget vs Actual by Department",  lambda: chart_dept_bar(df_clean, cols)),
            ("Departments Over Budget",          lambda: chart_dept_overspend(df_clean, cols)),
            ("Top 5 Department Movers",          lambda: chart_dept_movers(df_clean, cols, n=5)),
        ]

    for title, fn in chart_fns:
        try:
            fig = fn()
            if fig is not None:
                results.append((title, fig))
        except Exception:
            pass

    return {"charts": results, "file_type": file_type}


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
