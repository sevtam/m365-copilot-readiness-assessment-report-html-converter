import io
import base64
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
import streamlit as st

st.set_page_config(page_title="M365 Copilot Readiness Dashboard", page_icon="📊", layout="wide")


CANONICAL_COLUMNS = {
    "service": ["service"],
    "feature": ["feature", "workload", "capability"],
    "status": ["status", "result", "state"],
    "priority": ["priority", "severity", "importance"],
    "observation": ["observation", "finding", "details", "note"],
    "recommendation": ["recommendation", "action", "remediation", "next step"],
    "link_text": ["linktext", "link text", "reference", "doc title"],
    "link_url": ["linkurl", "link url", "url", "link"],
}

EXPORT_THEMES = {
    "Aurora Light": {
        "bg": "#eef4ff",
        "surface": "#ffffff",
        "surface_alt": "#f4f8ff",
        "text": "#0b1a3a",
        "muted": "#3b4f78",
        "accent": "#2a63f6",
        "border": "#d6e4ff",
        "grid": "#d8e4fb",
    },
    "Midnight": {
        "bg": "#070d18",
        "surface": "#0f172a",
        "surface_alt": "#18253f",
        "text": "#e7efff",
        "muted": "#9db0d0",
        "accent": "#57b9ff",
        "border": "#2b3f64",
        "grid": "#324a74",
    },
    "Forest Glass": {
        "bg": "#ecfff8",
        "surface": "#f9fffc",
        "surface_alt": "#ecfbf4",
        "text": "#073329",
        "muted": "#16614f",
        "accent": "#00a884",
        "border": "#bdeedc",
        "grid": "#c6ecd9",
    },
    "Boardroom Contrast": {
        "bg": "#f2f2f2",
        "surface": "#ffffff",
        "surface_alt": "#ececec",
        "text": "#141414",
        "muted": "#4a4a4a",
        "accent": "#5b5b5b",
        "border": "#c4c4c4",
        "grid": "#cdcdcd",
    },
    "Copilot": {
        "bg": "#eef3ff",
        "surface": "#ffffff",
        "surface_alt": "#f4f6ff",
        "text": "#1a1f36",
        "muted": "#4a5678",
        "accent": "#2a63f6",
        "border": "#d8e2ff",
        "grid": "#dce5ff",
    },
}

STATUS_COLOR_MAP = {
    "Success": "#16a34a",
    "Insight": "#0ea5e9",
    "Warning": "#f59e0b",
    "Action Required": "#dc2626",
    "Critical": "#b91c1c",
    "PendingProvisioning": "#f97316",
    "PendingActivation": "#f97316",
    "PendingInput": "#f59e0b",
    "Permission Required": "#ef4444",
    "Disabled": "#6b7280",
    "Missing Prerequisite": "#7c2d12",
}

COPILOT_BRAND = {
    "blue": "#2a63f6",
    "cyan": "#18a7ff",
    "green": "#36d399",
    "purple": "#7c4dff",
    "pink": "#d946ef",
    "orange": "#ff8a4c",
    "yellow": "#f6c945",
}


def summarize_statuses(df: pd.DataFrame) -> dict[str, int]:
    status_counts = df["status"].fillna("Unknown").value_counts() if "status" in df.columns else pd.Series(dtype=int)
    total = len(df)
    success = int(status_counts.get("Success", 0))
    warning = int(status_counts.get("Warning", 0))
    critical = int(status_counts.get("Critical", 0))
    needs_attention = total - success if total else 0
    return {
        "total": total,
        "success": success,
        "warning": warning,
        "critical": critical,
        "needs_attention": max(needs_attention, 0),
    }


def normalize_name(name: str) -> str:
    return " ".join(str(name).strip().lower().replace("_", " ").split())


def map_columns(df: pd.DataFrame) -> pd.DataFrame:
    col_lookup = {normalize_name(col): col for col in df.columns}
    rename_map: dict[str, str] = {}

    for canonical, options in CANONICAL_COLUMNS.items():
        for option in options:
            match = col_lookup.get(normalize_name(option))
            if match:
                rename_map[match] = canonical
                break

    return df.rename(columns=rename_map)


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    cleaned = df.copy()
    if "index" in cleaned.columns:
        as_numeric = pd.to_numeric(cleaned["index"], errors="coerce")
        if as_numeric.notna().all():
            cleaned = cleaned.drop(columns=["index"])
    return cleaned


def read_uploaded_file(uploaded_file, selected_sheet: Optional[str]) -> pd.DataFrame:
    suffix = Path(uploaded_file.name).suffix.lower()
    if suffix == ".csv":
        df = pd.read_csv(uploaded_file)
    elif suffix in {".xlsx", ".xls"}:
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet or 0)
    else:
        raise ValueError("Unsupported file type. Please upload CSV or XLSX.")
    return clean_dataframe(map_columns(df))


def pick_default_sample() -> Optional[Path]:
    cwd = Path.cwd()
    for name in ["m365_recommendations_20260520_134954.csv", "m365_recommendations_20260520_134954.xlsx"]:
        candidate = cwd / name
        if candidate.exists():
            return candidate
    return None


def read_local_sample(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path)
    else:
        df = pd.read_excel(path, sheet_name=0)
    return clean_dataframe(map_columns(df))


def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.header("Filters")
    filtered = df.copy()

    search = st.sidebar.text_input("Search text")

    for col in ["service", "status", "priority"]:
        if col in filtered.columns:
            options = sorted(str(v) for v in filtered[col].dropna().unique())
            selected = st.sidebar.multiselect(f"{col.title()}", options, default=options)
            filtered = filtered[filtered[col].astype(str).isin(selected)]

    if "feature" in filtered.columns:
        finding_options = sorted(str(v) for v in filtered["feature"].dropna().unique())
        selected_findings = st.sidebar.multiselect("Finding", finding_options, default=finding_options)
        filtered = filtered[filtered["feature"].astype(str).isin(selected_findings)]

    if "status" in filtered.columns:
        actionable_only = st.sidebar.checkbox("Show only non-success findings", value=False)
        if actionable_only:
            filtered = filtered[~filtered["status"].fillna("").str.lower().eq("success")]

    if search:
        mask = pd.Series(False, index=filtered.index)
        for col in ["feature", "observation", "recommendation"]:
            if col in filtered.columns:
                mask |= filtered[col].fillna("").astype(str).str.contains(search, case=False, na=False)
        filtered = filtered[mask]

    return filtered


def select_findings_columns(df: pd.DataFrame) -> list[str]:
    default_columns = ["service", "feature", "status", "priority", "observation", "recommendation"]
    available = [c for c in default_columns if c in df.columns]
    if not available:
        return list(df.columns)
    selected = st.sidebar.multiselect("Findings columns", available, default=available)
    return selected if selected else available


def select_report_kpi_cards() -> list[str]:
    options = ["Total checks", "Readiness score", "Success", "Needs attention", "Warnings", "Critical"]
    selected = st.sidebar.multiselect("Report KPI cards", options, default=options)
    return selected if selected else options


def status_score(df: pd.DataFrame) -> float:
    if "status" not in df.columns or len(df) == 0:
        return 0.0
    status = df["status"].fillna("").str.lower()
    good = status.str.contains("success|pass|ok|healthy")
    return round((good.sum() / len(df)) * 100, 1)


def render_kpis(df: pd.DataFrame) -> None:
    summary = summarize_statuses(df)
    score = status_score(df)
    total = summary["total"] if summary["total"] else 1
    cards = [
        ("Total checks", summary["total"], score / 100, "#2563eb"),
        ("Readiness score", f"{score:.1f}%", score / 100, "#2563eb"),
        ("Success", summary["success"], summary["success"] / total, "#16a34a"),
        ("Needs attention", summary["needs_attention"], summary["needs_attention"] / total, "#dc2626"),
        ("Warnings", summary["warning"], summary["warning"] / total, "#f59e0b"),
        ("Critical", summary["critical"], summary["critical"] / total, "#b91c1c"),
    ]

    st.markdown(
        """
        <style>
          .metric-card {
            border: 1px solid #e5e7eb;
            border-radius: 12px;
            padding: 14px 14px 12px;
            background: #ffffff;
            box-shadow: 0 1px 2px rgba(16,24,40,.04);
            min-height: 108px;
          }
          .metric-title { color:#475569;font-size:13px; }
          .metric-value { color:#0f172a;font-size:34px;font-weight:700;line-height:1.1;margin:4px 0 10px; }
          .metric-track { height:8px;background:#e2e8f0;border-radius:999px;overflow:hidden; }
          .metric-fill { height:100%;border-radius:999px; }
        </style>
        """,
        unsafe_allow_html=True,
    )
    cols = st.columns(len(cards))
    for col, (title, value, ratio, color) in zip(cols, cards):
        display_value = f"{value:,}" if isinstance(value, (int, float)) else str(value)
        col.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-title">{title}</div>
                <div class="metric-value">{display_value}</div>
                <div class="metric-track"><div class="metric-fill" style="width:{max(0, min(100, ratio * 100)):.1f}%;background:{color};"></div></div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def create_status_mix_figure(df: pd.DataFrame):
    if "status" not in df.columns or df.empty:
        return None
    status_counts = df["status"].fillna("Unknown").value_counts().reset_index()
    status_counts.columns = ["status", "count"]
    fig = px.pie(
        status_counts,
        names="status",
        values="count",
        hole=0.62,
        title="Status mix",
        color="status",
        color_discrete_map=STATUS_COLOR_MAP,
        template="plotly_white",
    )
    fig.update_layout(legend={"orientation": "v", "x": 1.0, "y": 0.5}, margin={"l": 20, "r": 20, "t": 60, "b": 20})
    return fig


def create_readiness_gauge_figure(df: pd.DataFrame):
    score = status_score(df)
    readiness_label = "Ready" if score >= 80 else "Progressing" if score >= 55 else "Not ready"
    bar_color = "#16a34a" if score >= 80 else "#f59e0b" if score >= 55 else "#ef4444"
    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=score,
            number={"suffix": "%", "font": {"size": 44, "color": "#0f172a"}},
            gauge={
                "axis": {"range": [0, 100], "tickwidth": 0, "showticklabels": False},
                "bar": {"color": bar_color, "thickness": 0.36},
                "bgcolor": "#f1f5f9",
                "steps": [
                    {"range": [0, 45], "color": "#fee2e2"},
                    {"range": [45, 75], "color": "#fef3c7"},
                    {"range": [75, 100], "color": "#dcfce7"},
                ],
                "threshold": {"line": {"color": "#0f172a", "width": 3}, "thickness": 0.8, "value": score},
            },
            title={"text": "Readiness score", "font": {"size": 18}},
        )
    )
    fig.update_layout(
        template="plotly_white",
        margin={"l": 18, "r": 18, "t": 60, "b": 36},
        annotations=[
            {
                "text": readiness_label,
                "xref": "paper",
                "yref": "paper",
                "x": 0.5,
                "y": -0.05,
                "showarrow": False,
                "font": {"size": 14, "color": "#475569"},
            }
        ],
    )
    return fig


def create_service_status_figure(df: pd.DataFrame):
    if "service" not in df.columns or "status" not in df.columns or df.empty:
        return None
    agg = df.groupby(["service", "status"], dropna=False).size().reset_index(name="count")
    service_order = (
        agg.groupby("service", as_index=False)["count"]
        .sum()
        .sort_values("count", ascending=False)
        .head(10)["service"]
        .tolist()
    )
    scoped = agg[agg["service"].isin(service_order)]
    fig = px.bar(
        scoped,
        x="service",
        y="count",
        color="status",
        title="Top services by status",
        barmode="stack",
        color_discrete_map=STATUS_COLOR_MAP,
        template="plotly_white",
    )
    fig.update_layout(legend_title_text="", xaxis_title="", yaxis_title="Checks", margin={"l": 20, "r": 20, "t": 60, "b": 20})
    return fig


def create_priority_balance_figure(df: pd.DataFrame):
    if "priority" not in df.columns or "status" not in df.columns or df.empty:
        return None
    scoped = df.copy()
    scoped["priority"] = scoped["priority"].fillna("Unspecified")
    scoped["bucket"] = scoped["status"].fillna("").str.lower().apply(lambda s: "Success" if s == "success" else "Needs attention")
    agg = scoped.groupby(["priority", "bucket"], dropna=False).size().reset_index(name="count")
    priority_order = ["Critical", "High", "Medium", "Low", "Unspecified"]
    agg["priority"] = pd.Categorical(agg["priority"], categories=priority_order, ordered=True)
    agg = agg.sort_values("priority")
    fig = px.bar(
        agg,
        x="priority",
        y="count",
        color="bucket",
        barmode="group",
        title="By severity",
        color_discrete_map={"Success": "#16a34a", "Needs attention": "#f43f5e"},
        template="plotly_white",
    )
    fig.update_layout(legend_title_text="", xaxis_title="", yaxis_title="Checks", margin={"l": 20, "r": 20, "t": 60, "b": 20})
    return fig


def create_top_recommendations_figure(df: pd.DataFrame, top_n: int):
    if "recommendation" not in df.columns:
        return None
    top = df["recommendation"].dropna().astype(str).value_counts().head(top_n)
    if top.empty:
        return None
    top_df = top.reset_index()
    top_df.columns = ["recommendation", "count"]
    top_df = top_df.sort_values("count", ascending=True)
    fig = px.bar(
        top_df,
        x="count",
        y="recommendation",
        orientation="h",
        title="Top recommendations",
        template="plotly_white",
        color_discrete_sequence=["#2563eb"],
    )
    fig.update_layout(xaxis_title="Occurrences", yaxis_title="", margin={"l": 20, "r": 20, "t": 60, "b": 20})
    fig.update_xaxes(autorange=True, rangemode="tozero")
    fig.update_yaxes(categoryorder="total ascending")
    return fig


def create_high_risk_features_figure(df: pd.DataFrame, top_n: int):
    if "feature" not in df.columns or "status" not in df.columns or df.empty:
        return None
    scoped = df[~df["status"].fillna("").str.lower().eq("success")]
    if scoped.empty:
        return None
    top = scoped["feature"].fillna("Unknown feature").astype(str).value_counts().head(top_n)
    chart_df = top.reset_index()
    chart_df.columns = ["feature", "count"]
    fig = px.line(
        chart_df,
        x="feature",
        y="count",
        markers=True,
        title="By category (needs attention)",
        template="plotly_white",
        color_discrete_sequence=["#f43f5e"],
    )
    fig.update_layout(xaxis_title="", yaxis_title="Findings", margin={"l": 20, "r": 20, "t": 60, "b": 20})
    return fig


def create_licensed_vs_configured_figure(df: pd.DataFrame):
    if "service" not in df.columns or "status" not in df.columns or df.empty:
        return None

    scoped = df.copy()
    status = scoped["status"].fillna("").str.lower()
    scoped["licensed"] = ~status.isin(["missing prerequisite", "disabled"])
    scoped["configured"] = status.isin(["success", "insight"])

    per_service = (
        scoped.groupby("service", dropna=False)
        .agg(licensed=("licensed", "sum"), configured=("configured", "sum"))
        .reset_index()
    )
    if per_service.empty:
        return None

    top_services = per_service.sort_values("licensed", ascending=False).head(10)
    plot_df = top_services.melt(id_vars=["service"], value_vars=["licensed", "configured"], var_name="capability", value_name="count")
    plot_df["capability"] = plot_df["capability"].map({"licensed": "Licensed", "configured": "Configured"})

    fig = px.bar(
        plot_df,
        x="service",
        y="count",
        color="capability",
        barmode="group",
        title="Licensed capability vs configured capability",
        color_discrete_map={"Licensed": "#6366f1", "Configured": "#10b981"},
        template="plotly_white",
    )
    fig.update_layout(legend_title_text="", xaxis_title="", yaxis_title="Capabilities", margin={"l": 20, "r": 20, "t": 60, "b": 20})
    return fig


def render_table(df: pd.DataFrame, selected_columns: list[str]) -> None:
    st.subheader("Detailed findings")
    display_cols = [c for c in selected_columns if c in df.columns]
    st.dataframe(df[display_cols] if display_cols else df, use_container_width=True, hide_index=True)

    csv_bytes = df.to_csv(index=False).encode("utf-8")
    st.download_button("Download filtered data (CSV)", data=csv_bytes, file_name="filtered_m365_copilot_readiness.csv")


def build_logo_data_uri(logo_file) -> Optional[str]:
    if logo_file is None:
        return None
    data = logo_file.getvalue()
    if not data:
        return None
    mime_type = logo_file.type if logo_file.type else "image/png"
    encoded = base64.b64encode(data).decode("ascii")
    return f"data:{mime_type};base64,{encoded}"


def style_figure_for_export(fig, theme: dict) -> None:
    if fig is None:
        return
    fig.update_layout(
        template="plotly_white",
        paper_bgcolor=theme["surface"],
        plot_bgcolor=theme["surface_alt"],
        font={"color": theme["text"], "family": "Segoe UI, Arial, sans-serif"},
        title={"font": {"size": 18, "family": "Segoe UI Semibold, Segoe UI, Arial, sans-serif"}},
        legend={"bgcolor": theme["surface"], "bordercolor": theme["border"], "borderwidth": 1},
        margin={"l": 40, "r": 25, "t": 60, "b": 40},
        hoverlabel={"bgcolor": theme["surface"], "font_size": 12},
        colorway=[
            COPILOT_BRAND["blue"],
            COPILOT_BRAND["purple"],
            COPILOT_BRAND["cyan"],
            COPILOT_BRAND["pink"],
            COPILOT_BRAND["green"],
            COPILOT_BRAND["orange"],
            COPILOT_BRAND["yellow"],
        ],
    )
    for trace in fig.data:
        if hasattr(trace, "marker") and hasattr(trace.marker, "color"):
            color = trace.marker.color
            if isinstance(color, str) and color.startswith("#0000"):
                trace.marker.color = theme["accent"]
        if getattr(trace, "type", "") == "indicator":
            if hasattr(trace, "number") and trace.number is not None:
                trace.number["font"] = {"color": theme["text"], "size": 44}
            if hasattr(trace, "gauge") and trace.gauge is not None:
                trace.gauge["bgcolor"] = theme["surface_alt"]
                if "threshold" in trace.gauge:
                    trace.gauge["threshold"]["line"]["color"] = theme["text"]
            if hasattr(trace, "title") and trace.title is not None:
                trace.title["font"] = {"color": theme["text"], "size": 18}
        if getattr(trace, "type", "") == "pie":
            labels_attr = getattr(trace, "labels", None)
            if labels_attr is not None:
                labels = list(labels_attr)
                if labels:
                    trace.marker.colors = [STATUS_COLOR_MAP.get(str(label), theme["accent"]) for label in labels]
    if fig.layout.annotations:
        for ann in fig.layout.annotations:
            if ann.font is None:
                ann.font = {}
            ann.font["color"] = theme["muted"]
    fig.update_xaxes(gridcolor=theme["grid"], zerolinecolor=theme["grid"])
    fig.update_yaxes(gridcolor=theme["grid"], zerolinecolor=theme["grid"])


def build_html_report(
    filtered_df: pd.DataFrame,
    figures: list,
    source_name: str,
    readiness: float,
    total_rows: int,
    export_title: str,
    theme_name: str,
    logo_data_uri: Optional[str],
    selected_columns: list[str],
    selected_kpi_cards: list[str],
) -> str:
    theme = EXPORT_THEMES[theme_name]
    theme_variant = {
        "Aurora Light": "aurora",
        "Midnight": "midnight",
        "Forest Glass": "forest",
        "Boardroom Contrast": "boardroom",
        "Copilot": "copilot",
    }.get(theme_name, "aurora")
    is_dark = theme_variant == "midnight"
    summary = summarize_statuses(filtered_df)
    total = summary["total"] if summary["total"] else 1
    kpi_all = [
        ("Total checks", f"{summary['total']}", readiness, "#2563eb"),
        ("Readiness score", f"{readiness:.1f}%", readiness, "#2563eb"),
        ("Success", f"{summary['success']}", (summary["success"] / total) * 100, "#16a34a"),
        ("Needs attention", f"{summary['needs_attention']}", (summary["needs_attention"] / total) * 100, "#f43f5e"),
        ("Warnings", f"{summary['warning']}", (summary["warning"] / total) * 100, "#f59e0b"),
        ("Critical", f"{summary['critical']}", (summary["critical"] / total) * 100, "#b91c1c"),
    ]
    selected_set = set(selected_kpi_cards)
    kpi_cards = [item for item in kpi_all if item[0] in selected_set] or kpi_all
    service_values = []
    status_values = []
    priority_values = []
    if "service" in filtered_df.columns:
        service_values = sorted([str(v) for v in filtered_df["service"].dropna().unique()])
    if "status" in filtered_df.columns:
        status_values = sorted([str(v) for v in filtered_df["status"].dropna().unique()])
    if "priority" in filtered_df.columns:
        priority_values = sorted([str(v) for v in filtered_df["priority"].dropna().unique()])
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if theme_variant == "midnight":
        page_bg = "radial-gradient(1200px 620px at -8% -14%, rgba(56,189,248,.18), transparent 58%),radial-gradient(960px 560px at 105% -8%, rgba(99,102,241,.16), transparent 55%),linear-gradient(165deg, #0b1220 0%, #111a2b 100%)"
        hero_bg = "linear-gradient(145deg, rgba(17,26,43,.92), rgba(23,35,58,.86))"
        panel_bg = "linear-gradient(165deg, rgba(17,26,43,.94), rgba(23,35,58,.84))"
        table_bg = "linear-gradient(170deg, rgba(17,26,43,.96), rgba(23,35,58,.9))"
        table_header_bg = "linear-gradient(180deg,#1b2942,#23334f)"
        table_header_text = "#cbd5e1"
        border_rgba = "rgba(148,163,184,.38)"
        soft_border = "rgba(71,85,105,.55)"
        white_sheen = "rgba(255,255,255,.08)"
        track_bg = "#23334f"
        input_bg = "rgba(15,23,42,.55)"
        btn_bg = "linear-gradient(180deg,#1e293b,#0f172a)"
        btn_text = "#e2e8f0"
        row_even_bg = "rgba(30,41,59,.52)"
        link_hover = "#93c5fd"
        hero_glow_1 = "#7c4dff4a"
        hero_glow_2 = "#18a7ff3d"
        card_glow = "#d946ef2e"
        pill_overlay = "linear-gradient(120deg,#2a63f61a,#d946ef1a)"
        button_hover = "linear-gradient(120deg,#18a7ff2a,#7c4dff2a)"
    elif theme_variant == "forest":
        page_bg = "radial-gradient(1100px 600px at -10% -15%, rgba(16,185,129,.2), transparent 56%),radial-gradient(900px 500px at 110% -10%, rgba(34,197,94,.16), transparent 54%),linear-gradient(165deg, #ecfff8 0%, #e6f9f0 100%)"
        hero_bg = "linear-gradient(145deg, rgba(249,255,252,.96), rgba(236,251,244,.86))"
        panel_bg = "linear-gradient(165deg, rgba(249,255,252,.98), rgba(230,247,239,.9))"
        table_bg = "linear-gradient(170deg, rgba(249,255,252,.98), rgba(232,248,240,.95))"
        table_header_bg = "linear-gradient(180deg,#f0fdf8,#ddf7eb)"
        table_header_text = "#1f5e4c"
        border_rgba = "rgba(190,234,216,.92)"
        soft_border = "rgba(155,213,190,.9)"
        white_sheen = "rgba(255,255,255,.92)"
        track_bg = "#d8efe5"
        input_bg = "rgba(255,255,255,.88)"
        btn_bg = "linear-gradient(180deg,#f4fff9,#dcf5ea)"
        btn_text = "#205d4c"
        row_even_bg = "rgba(236,253,245,.85)"
        link_hover = "#047857"
        hero_glow_1 = "#22c55e3a"
        hero_glow_2 = "#14b8a635"
        card_glow = "#34d39933"
        pill_overlay = "linear-gradient(120deg,#22c55e1f,#2dd4bf1f)"
        button_hover = "linear-gradient(120deg,#22c55e24,#14b8a624)"
    elif theme_variant == "boardroom":
        page_bg = "linear-gradient(165deg, #f2f2f2 0%, #e9e9e9 100%)"
        hero_bg = "linear-gradient(145deg, rgba(255,255,255,.97), rgba(244,244,244,.94))"
        panel_bg = "linear-gradient(165deg, rgba(255,255,255,.99), rgba(240,240,240,.94))"
        table_bg = "linear-gradient(170deg, rgba(255,255,255,.99), rgba(242,242,242,.96))"
        table_header_bg = "linear-gradient(180deg,#f3f3f3,#e7e7e7)"
        table_header_text = "#4b4b4b"
        border_rgba = "rgba(196,196,196,.95)"
        soft_border = "rgba(186,186,186,.9)"
        white_sheen = "rgba(255,255,255,.9)"
        track_bg = "#d8d8d8"
        input_bg = "rgba(255,255,255,.9)"
        btn_bg = "linear-gradient(180deg,#f5f5f5,#e2e2e2)"
        btn_text = "#3f3f3f"
        row_even_bg = "rgba(244,244,244,.84)"
        link_hover = "#2f2f2f"
        hero_glow_1 = "#a8a8a844"
        hero_glow_2 = "#bdbdbd30"
        card_glow = "#9ca3af2e"
        pill_overlay = "linear-gradient(120deg,#d4d4d41f,#f0f0f01f)"
        button_hover = "linear-gradient(120deg,#e8e8e8,#d8d8d8)"
    elif theme_variant == "copilot":
        page_bg = (
            "radial-gradient(1100px 620px at -5% -12%, rgba(24,167,255,.24), transparent 58%),"
            "radial-gradient(980px 560px at 110% -12%, rgba(217,70,239,.2), transparent 56%),"
            "radial-gradient(800px 420px at 52% 115%, rgba(255,138,76,.22), transparent 62%),"
            "linear-gradient(165deg, #eef3ff 0%, #f8fbff 100%)"
        )
        hero_bg = (
            "linear-gradient(115deg, rgba(42,99,246,.12), rgba(24,167,255,.1), rgba(124,77,255,.1), rgba(217,70,239,.12), rgba(255,138,76,.12)),"
            "linear-gradient(145deg, rgba(255,255,255,.96), rgba(246,248,255,.92))"
        )
        panel_bg = (
            "linear-gradient(120deg, rgba(42,99,246,.08), rgba(24,167,255,.08), rgba(124,77,255,.08), rgba(217,70,239,.08), rgba(255,138,76,.08)),"
            "linear-gradient(165deg, rgba(255,255,255,.98), rgba(242,246,255,.92))"
        )
        table_bg = (
            "linear-gradient(110deg, rgba(42,99,246,.06), rgba(24,167,255,.06), rgba(124,77,255,.06), rgba(217,70,239,.06), rgba(255,138,76,.06)),"
            "linear-gradient(170deg, rgba(255,255,255,.98), rgba(245,248,255,.95))"
        )
        table_header_bg = "linear-gradient(90deg,#eaf1ff,#e8f5ff,#efe9ff,#ffeef9,#fff2ea)"
        table_header_text = "#3d4c74"
        border_rgba = "rgba(206,220,255,.95)"
        soft_border = "rgba(186,202,240,.9)"
        white_sheen = "rgba(255,255,255,.92)"
        track_bg = "#dce6ff"
        input_bg = "rgba(255,255,255,.92)"
        btn_bg = "linear-gradient(120deg,#e6eeff,#e7f5ff,#f0e8ff,#ffe8f6,#ffefe7)"
        btn_text = "#30456f"
        row_even_bg = "rgba(241,246,255,.88)"
        link_hover = "#7c4dff"
        hero_glow_1 = "#7c4dff4f"
        hero_glow_2 = "#18a7ff45"
        card_glow = "#ff8a4c35"
        pill_overlay = "linear-gradient(120deg,#2a63f62e,#18a7ff25,#7c4dff25,#d946ef24,#ff8a4c24)"
        button_hover = "linear-gradient(120deg,#2a63f638,#18a7ff2e,#7c4dff2e,#d946ef2e,#ff8a4c2a)"
    else:  # aurora
        page_bg = "radial-gradient(1200px 620px at -8% -14%, rgba(59,130,246,.22), transparent 58%),radial-gradient(960px 560px at 105% -8%, rgba(168,85,247,.16), transparent 55%),linear-gradient(165deg, #eef4ff 0%, #f7f9fc 100%)"
        hero_bg = "linear-gradient(145deg, rgba(255,255,255,.96), rgba(245,249,255,.9))"
        panel_bg = "linear-gradient(165deg, rgba(255,255,255,.98), rgba(241,245,255,.9))"
        table_bg = "linear-gradient(170deg, rgba(255,255,255,.98), rgba(245,248,255,.96))"
        table_header_bg = "linear-gradient(180deg,#f5f8ff,#e9efff)"
        table_header_text = "#405072"
        border_rgba = "rgba(214,228,255,.95)"
        soft_border = "rgba(194,213,247,.9)"
        white_sheen = "rgba(255,255,255,.92)"
        track_bg = "#dbe6fb"
        input_bg = "rgba(255,255,255,.9)"
        btn_bg = "linear-gradient(180deg,#f8faff,#e5ecff)"
        btn_text = "#334c7a"
        row_even_bg = "rgba(242,246,255,.84)"
        link_hover = "#1e40af"
        hero_glow_1 = "#7c4dff3f"
        hero_glow_2 = "#18a7ff33"
        card_glow = "#2a63f633"
        pill_overlay = "linear-gradient(120deg,#2a63f61f,#d946ef1f)"
        button_hover = "linear-gradient(120deg,#18a7ff22,#7c4dff20)"

    if theme_variant == "copilot":
        kpi_gradient = {
            "Total checks": "linear-gradient(90deg,#2a63f6,#18a7ff,#7c4dff)",
            "Readiness score": "linear-gradient(90deg,#2a63f6,#18a7ff,#7c4dff)",
            "Success": "linear-gradient(90deg,#36d399,#18a7ff)",
            "Needs attention": "linear-gradient(90deg,#ff8a4c,#d946ef)",
            "Warnings": "linear-gradient(90deg,#f6c945,#ff8a4c)",
            "Critical": "linear-gradient(90deg,#ff8a4c,#d946ef,#7c4dff)",
        }
        kpi_cards = [(title, value, pct, kpi_gradient.get(title, color)) for title, value, pct, color in kpi_cards]

    html_parts: list[str] = [
        "<!doctype html>",
        "<html lang='en'><head><meta charset='utf-8'>",
        "<meta name='viewport' content='width=device-width,initial-scale=1'>",
        "<title>M365 Copilot Readiness Dashboard Export</title>",
        "<style>",
        f":root{{--bg:{theme['bg']};--surface:{theme['surface']};--surface-alt:{theme['surface_alt']};--text:{theme['text']};--muted:{theme['muted']};--accent:{theme['accent']};--border:{theme['border']};--grid:{theme['grid']};"
        f"--cp-blue:{COPILOT_BRAND['blue']};--cp-cyan:{COPILOT_BRAND['cyan']};--cp-green:{COPILOT_BRAND['green']};--cp-purple:{COPILOT_BRAND['purple']};--cp-pink:{COPILOT_BRAND['pink']};--cp-orange:{COPILOT_BRAND['orange']};--cp-yellow:{COPILOT_BRAND['yellow']};}}",
        f"body{{font-family:'Segoe UI',Inter,Arial,sans-serif;margin:0;color:var(--text);background:{page_bg};}}",
        ".container{max-width:1540px;margin:24px auto;padding:0 18px 34px;}",
        ".hero{position:relative;overflow:hidden;background:"
        f"{hero_bg};"
        f"backdrop-filter:blur(6px);border:1px solid {border_rgba};"
        "border-radius:20px;padding:26px 28px;margin-bottom:18px;"
        f"box-shadow:0 18px 44px rgba(15,23,42,.26), inset 0 1px 0 {white_sheen};}}",
        ".hero:before{content:'';position:absolute;inset:auto -12% -55% auto;width:420px;height:420px;border-radius:50%;"
        f"background:radial-gradient(circle, {hero_glow_1}, transparent 68%);pointer-events:none;}}",
        ".hero:after{content:'';position:absolute;left:-130px;top:-150px;width:360px;height:360px;border-radius:50%;"
        f"background:radial-gradient(circle, {hero_glow_2}, transparent 72%);pointer-events:none;}}",
        ".hero-head{display:flex;gap:16px;align-items:center;justify-content:space-between;flex-wrap:wrap;}",
        ".logo-wrap{display:flex;align-items:center;gap:16px;z-index:2;position:relative;}",
        ".logo{max-height:88px;max-width:330px;object-fit:contain;display:block;filter:drop-shadow(0 8px 14px rgba(15,23,42,.18));}",
        ".hero h1{font-size:33px;letter-spacing:-.02em;}",
        f".subtitle{{margin-top:6px;color:var(--muted);font-size:14px;}}",
        ".meta{display:flex;gap:10px;flex-wrap:wrap;margin-top:14px;}",
        f".pill{{background:{pill_overlay},{input_bg};border:1px solid {border_rgba};padding:8px 13px;border-radius:999px;"
        "font-size:12px;font-weight:600;color:var(--muted);box-shadow:0 8px 18px rgba(15,23,42,.12);}",
        ".section{background:transparent;margin-bottom:14px;position:relative;}",
        ".section h2{margin:8px 0 10px;}",
        f".table-wrap{{background:{table_bg};"
        f"border:1px solid {border_rgba};border-radius:16px;padding:14px;overflow:auto;max-height:70vh;"
        f"box-shadow:0 16px 36px rgba(2,6,23,.24), inset 0 1px 0 {white_sheen};}}",
        "table{width:100%;border-collapse:collapse;font-size:12px;line-height:1.45;table-layout:fixed;}",
        f"th,td{{border:1px solid {soft_border};padding:8px 10px;vertical-align:top;text-align:left;"
        "white-space:normal;word-break:break-word;overflow-wrap:anywhere;}",
        f"th{{background:{table_header_bg};position:sticky;top:0;z-index:2;color:{table_header_text};}}",
        f"tr:nth-child(even){{background:{row_even_bg};}}",
        "a{color:var(--accent);text-decoration:none;font-weight:600;}",
        "td a{word-break:break-all;overflow-wrap:anywhere;}",
        f"a:hover{{text-decoration:underline;color:{link_hover};}}",
        ".chart-grid{display:grid;grid-template-columns:repeat(2,minmax(360px,1fr));gap:14px;margin-top:8px;}",
        f".chart-block{{position:relative;overflow:hidden;background:{panel_bg};"
        f"border:1px solid {border_rgba};border-radius:16px;padding:10px;min-height:360px;"
        f"box-shadow:0 18px 36px rgba(2,6,23,.24), inset 0 1px 0 {white_sheen};}}"
        ".chart-block:before{content:'';position:absolute;left:0;right:0;top:0;height:78px;"
        f"background:linear-gradient(180deg, {white_sheen}, rgba(255,255,255,0));pointer-events:none;}}",
        ".cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:11px;margin-top:14px;}",
        f".card{{position:relative;overflow:hidden;background:{panel_bg};"
        f"border:1px solid {border_rgba};border-radius:14px;padding:13px;"
        f"box-shadow:0 14px 28px rgba(2,6,23,.2), inset 0 1px 0 {white_sheen};}}",
        ".card:after{content:'';position:absolute;inset:auto -55px -70px auto;width:130px;height:130px;border-radius:50%;"
        f"background:radial-gradient(circle, {card_glow}, transparent 70%);pointer-events:none;}}",
        ".card-label{font-size:12px;color:var(--muted);font-weight:600;}",
        ".card-value{font-size:29px;font-weight:800;color:var(--text);line-height:1.1;margin:2px 0 8px;}",
        f".track{{height:8px;background:{track_bg};border-radius:999px;overflow:hidden;box-shadow:inset 0 1px 2px rgba(15,23,42,.2);}}",
        ".fill{height:100%;border-radius:999px;box-shadow:0 4px 12px rgba(37,99,235,.35);}",
        f".table-toolbar{{background:{panel_bg};"
        f"border:1px solid {border_rgba};border-radius:15px;padding:10px;display:grid;grid-template-columns:2fr 1fr 1fr 1fr auto auto;"
        "gap:10px;align-items:center;margin-bottom:10px;box-shadow:0 14px 32px rgba(2,6,23,.2);}",
        f".table-toolbar input,.table-toolbar select{{width:100%;border:1px solid {soft_border};border-radius:11px;padding:9px 11px;"
        f"background:{input_bg};color:var(--text);box-shadow:inset 0 1px 2px rgba(15,23,42,.15);}}",
        f".table-toolbar button{{border:1px solid {soft_border};background:{btn_bg};color:{btn_text};"
        "border-radius:11px;padding:9px 13px;cursor:pointer;font-weight:700;}",
        f".table-toolbar button:hover{{background:{button_hover};}}",
        ".table-toolbar .count{color:var(--muted);font-size:12px;justify-self:end;font-weight:600;}",
        "@media (max-width:1100px){.chart-grid{grid-template-columns:1fr;}.cards{grid-template-columns:repeat(2,minmax(140px,1fr));}}",
        "@media (max-width:1100px){.table-toolbar{grid-template-columns:1fr 1fr;}}",
        "</style></head><body>",
        "<div class='container'>",
        "<div class='hero'>",
        "<div class='hero-head'>",
        "<div class='logo-wrap'>",
        f"<img class='logo' src='{logo_data_uri}' alt='Logo'>" if logo_data_uri else "",
        f"<div><h1 style='margin:0'>{export_title}</h1><div class='subtitle'>Interactive M365 Copilot readiness report</div></div>",
        "</div>",
        "</div>",
        f"<div class='meta'><span class='pill'>Generated: {generated_at}</span></div>",
        "<div class='cards'>",
    ]
    for title, value, pct, color in kpi_cards:
        html_parts.append(
            f"<div class='card'><div class='card-label'>{title}</div><div class='card-value'>{value}</div>"
            f"<div class='track'><div class='fill' style='width:{max(0.0, min(100.0, pct)):.1f}%;background:{color};'></div></div></div>"
        )
    html_parts.append("</div>")

    first_chart = True
    html_parts.append("<div class='chart-grid'>")
    for fig in figures:
        if fig is None:
            continue
        style_figure_for_export(fig, theme)
        html_parts.append("<div class='chart-block'>")
        html_parts.append(
            pio.to_html(
                fig,
                full_html=False,
                include_plotlyjs=first_chart,
                config={"displaylogo": False, "displayModeBar": False},
            )
        )
        html_parts.append("</div>")
        first_chart = False
    html_parts.append("</div>")

    export_df = filtered_df.copy()
    cols_to_remove = [c for c in ["link_text", "link_url", "reference_link"] if c in export_df.columns]
    if cols_to_remove:
        export_df = export_df.drop(columns=cols_to_remove)
    if selected_columns:
        cols_for_export = [c for c in selected_columns if c in export_df.columns]
        if cols_for_export:
            export_df = export_df[cols_for_export]

    html_parts.append("<div class='section'><h2>Filtered findings</h2>")
    html_parts.append("<div class='table-toolbar'>")
    html_parts.append("<input id='ff-search' type='text' placeholder='Search findings...'>")
    html_parts.append("<select id='ff-service'><option value=''>All services</option>")
    html_parts.extend([f"<option value=\"{v}\">{v}</option>" for v in service_values])
    html_parts.append("</select>")
    html_parts.append("<select id='ff-status'><option value=''>All statuses</option>")
    html_parts.extend([f"<option value=\"{v}\">{v}</option>" for v in status_values])
    html_parts.append("</select>")
    html_parts.append("<select id='ff-priority'><option value=''>All priorities</option>")
    html_parts.extend([f"<option value=\"{v}\">{v}</option>" for v in priority_values])
    html_parts.append("</select>")
    html_parts.append("<button id='ff-clear' type='button'>Reset</button>")
    html_parts.append("<div class='count' id='ff-count'></div>")
    html_parts.append("</div>")
    html_parts.append("<div class='table-wrap'>")
    html_parts.append(export_df.to_html(index=False, escape=False))
    html_parts.append("</div></div>")
    html_parts.append(
        """
<script>
(function () {
  const table = document.querySelector(".table-wrap table");
  if (!table || !table.tHead || !table.tBodies.length) return;
  const rows = Array.from(table.tBodies[0].rows);
  const headers = Array.from(table.tHead.rows[0].cells).map(c => c.textContent.trim().toLowerCase());
  const idxOf = (name) => headers.indexOf(name);
  const idxService = idxOf("service");
  const idxStatus = idxOf("status");
  const idxPriority = idxOf("priority");
  const elSearch = document.getElementById("ff-search");
  const elService = document.getElementById("ff-service");
  const elStatus = document.getElementById("ff-status");
  const elPriority = document.getElementById("ff-priority");
  const elClear = document.getElementById("ff-clear");
  const elCount = document.getElementById("ff-count");

  const cell = (row, idx) => (idx >= 0 && row.cells[idx]) ? row.cells[idx].innerText.trim() : "";

  function applyFilters() {
    const q = (elSearch?.value || "").toLowerCase().trim();
    const service = elService?.value || "";
    const status = elStatus?.value || "";
    const priority = elPriority?.value || "";
    let visible = 0;

    rows.forEach((row) => {
      const rowText = row.innerText.toLowerCase();
      const matchSearch = !q || rowText.includes(q);
      const matchService = !service || cell(row, idxService) === service;
      const matchStatus = !status || cell(row, idxStatus) === status;
      const matchPriority = !priority || cell(row, idxPriority) === priority;
      const show = matchSearch && matchService && matchStatus && matchPriority;
      row.style.display = show ? "" : "none";
      if (show) visible += 1;
    });
    if (elCount) elCount.textContent = `${visible} shown / ${rows.length}`;
  }

  [elSearch, elService, elStatus, elPriority].forEach((el) => {
    if (!el) return;
    el.addEventListener("input", applyFilters);
    el.addEventListener("change", applyFilters);
  });

  if (elClear) {
    elClear.addEventListener("click", () => {
      if (elSearch) elSearch.value = "";
      if (elService) elService.value = "";
      if (elStatus) elStatus.value = "";
      if (elPriority) elPriority.value = "";
      applyFilters();
    });
  }

  applyFilters();
})();
</script>
        """
    )
    html_parts.append("</div></body></html>")
    return "".join(html_parts)


def main() -> None:
    st.title("M365 Copilot Readiness Dashboard Generator")
    st.caption("Upload a readiness scan report (CSV/XLSX) and instantly generate an interactive dashboard.")

    uploaded_file = st.file_uploader("Upload report file", type=["csv", "xlsx", "xls"])

    selected_sheet = None
    if uploaded_file and Path(uploaded_file.name).suffix.lower() in {".xlsx", ".xls"}:
        workbook = pd.ExcelFile(io.BytesIO(uploaded_file.getvalue()))
        selected_sheet = st.selectbox("Select worksheet", workbook.sheet_names, index=0)

    if uploaded_file:
        df = read_uploaded_file(uploaded_file, selected_sheet)
    else:
        sample = pick_default_sample()
        if sample is None:
            st.warning("Upload a CSV/XLSX report to start.")
            return
        st.info(f"Loaded local sample file: {sample.name}")
        df = read_local_sample(sample)

    if df.empty:
        st.warning("Report contains no rows.")
        return

    top_n = st.sidebar.slider("Top N items in charts", min_value=5, max_value=20, value=10, step=1)
    filtered = apply_filters(df)
    selected_findings_columns = select_findings_columns(df)
    selected_report_kpis = select_report_kpi_cards()
    if filtered.empty:
        st.warning("No rows match the selected filters.")
        return

    readiness = status_score(filtered)
    render_kpis(filtered)

    status_fig = create_status_mix_figure(filtered)
    service_fig = create_service_status_figure(filtered)
    gauge_fig = create_readiness_gauge_figure(filtered)
    severity_fig = create_priority_balance_figure(filtered)
    risk_fig = create_high_risk_features_figure(filtered, top_n=top_n)
    capability_fig = create_licensed_vs_configured_figure(filtered)

    tabs = st.tabs(["Executive Overview", "Risk & Recommendations", "Data Explorer"])
    with tabs[0]:
        c1, c2 = st.columns(2)
        with c1:
            if status_fig is not None:
                st.plotly_chart(status_fig, use_container_width=True)
            else:
                st.info("No status column available for charting.")
        with c2:
            if gauge_fig is not None:
                st.plotly_chart(gauge_fig, use_container_width=True)
            else:
                st.info("No score data available.")

    with tabs[1]:
        c1, c2 = st.columns(2)
        with c1:
            if severity_fig is not None:
                st.plotly_chart(severity_fig, use_container_width=True)
            else:
                st.info("Severity chart needs both priority and status columns.")
        with c2:
            if risk_fig is not None:
                st.plotly_chart(risk_fig, use_container_width=True)
            else:
                st.caption("No non-success findings in current filter.")

        c3, c4 = st.columns(2)
        with c3:
            if service_fig is not None:
                st.plotly_chart(service_fig, use_container_width=True)
            else:
                st.info("Service/status chart needs both columns.")
        with c4:
            if capability_fig is not None:
                st.plotly_chart(capability_fig, use_container_width=True)
                st.caption("Configured is currently estimated as statuses: Success + Insight.")
            else:
                st.info("Capability comparison needs service and status columns.")

    with tabs[2]:
        render_table(filtered, selected_findings_columns)

    st.subheader("Export customization")
    export_col1, export_col2 = st.columns([2, 1])
    with export_col1:
        export_title = st.text_input("Dashboard title in export", value="M365 Copilot Readiness Dashboard")
        theme_name = st.selectbox("Theme", options=list(EXPORT_THEMES.keys()), index=0)
    with export_col2:
        logo_file = st.file_uploader("Optional logo", type=["png", "jpg", "jpeg", "svg"], key="export_logo")
        if logo_file:
            st.image(logo_file, caption="Logo preview", width=170)

    logo_data_uri = build_logo_data_uri(logo_file)
    export_html = build_html_report(
        filtered_df=filtered,
        figures=[status_fig, gauge_fig, severity_fig, risk_fig, service_fig, capability_fig],
        source_name=uploaded_file.name if uploaded_file else "local sample file",
        readiness=readiness,
        total_rows=len(filtered),
        export_title=export_title.strip() or "M365 Copilot Readiness Dashboard",
        theme_name=theme_name,
        logo_data_uri=logo_data_uri,
        selected_columns=selected_findings_columns,
        selected_kpi_cards=selected_report_kpis,
    )
    st.download_button(
        "Export dashboard (single HTML)",
        data=export_html.encode("utf-8"),
        file_name=f"m365_copilot_dashboard_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html",
        mime="text/html",
    )


if __name__ == "__main__":
    main()
