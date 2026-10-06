from __future__ import annotations

import csv
import html
import io
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path

import plotly.graph_objects as go
import streamlit as st

from finance_engine import (
    ANALYSIS_BUCKETS,
    analysis_bucket,
    annual_budget,
    annual_budget_amount,
    annual_income,
    budget_monthly_amounts,
    group_sum,
    historical_snapshots,
    load_config,
    monthly_actual_by_bucket,
    monthly_plan_by_bucket,
    monthly_recurring,
    net_worth,
    norm,
    normalize_transactions,
    parse_date,
    parse_number,
    patrimony_summary,
    uncategorized_transactions,
)

APP_DIR = Path(__file__).resolve().parent
DEFAULT_CONFIG = APP_DIR / "Config_Finanze_Familiari_V2_Aggiornato.xlsx"

IVORY = "#F8F3E8"
BEIGE = "#E9DDC8"
SAND = "#D8C3A5"
OCHRE = "#B88935"
OCHRE_LIGHT = "#D7AE64"
SAGE = "#829A7A"
SAGE_LIGHT = "#B8C6AD"
BRICK = "#A65743"
BRICK_LIGHT = "#CC8C78"
ANTHRACITE = "#3E403F"
POSITIVE = "#2F5D45"
WHITE = "#FFFDF8"
GRID = "#E3D8C5"
PALETTE = [OCHRE, SAND, SAGE, BRICK, OCHRE_LIGHT, SAGE_LIGHT, BRICK_LIGHT, "#A79882"]
TRAVEL_PALETTE = ["#A96832", "#C4873F", "#D7A456", "#E4BD78", "#B97845", "#D4935B", "#E9C994"]
MONTHS_IT = ["Gen", "Feb", "Mar", "Apr", "Mag", "Giu", "Lug", "Ago", "Set", "Ott", "Nov", "Dic"]

APP_VERSION = "Versione 5.0"

st.set_page_config(
    page_title="Finance Hub",
    page_icon="◈",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown(
    f"""
<style>
:root {{ --ivory:{IVORY}; --beige:{BEIGE}; --sand:{SAND}; --ochre:{OCHRE}; --sage:{SAGE}; --brick:{BRICK}; --anthracite:{ANTHRACITE}; }}
.stApp {{ background: {IVORY}; color: {ANTHRACITE}; }}
[data-testid="stHeader"] {{ background: rgba(248,243,232,.92); }}
[data-testid="stSidebar"] {{ background: {BEIGE}; border-right: 1px solid {SAND}; }}
h1, h2, h3, h4 {{ color: {ANTHRACITE} !important; letter-spacing: -0.02em; }}
h1 {{ font-weight: 750 !important; }}
.block-container {{ padding-top: 1.4rem; padding-bottom: 3rem; }}
[data-testid="stSidebar"] .block-container {{ padding-top: 1rem; }}
div[data-testid="stMetric"], .fh-card {{ background: {WHITE}; border: 1px solid {SAND}; border-radius: 14px; padding: 15px 16px; box-shadow: 0 4px 14px rgba(69,58,42,.06); min-height: 112px; }}
.fh-label {{ color: {ANTHRACITE}; font-size: .83rem; line-height: 1.25; min-height: 2.1rem; }}
.fh-value {{ font-size: 1.48rem; font-weight: 760; line-height: 1.18; margin-top: .32rem; }}
.fh-sub {{ color: #776F64; font-size: .75rem; margin-top: .35rem; line-height: 1.25; }}
.fh-positive {{ color: {POSITIVE}; }} .fh-negative {{ color: {BRICK}; }} .fh-neutral {{ color: {ANTHRACITE}; }}
.fh-chip {{ display:inline-block; background:{BEIGE}; border:1px solid {SAND}; border-radius:99px; padding:.22rem .62rem; margin:.1rem .2rem .1rem 0; color:{ANTHRACITE}; font-size:.76rem; }}
.stTabs [data-baseweb="tab-list"] {{ gap:.25rem; border-bottom:1px solid {SAND}; }}
.stTabs [data-baseweb="tab"] {{ height:2.9rem; background:transparent; border-radius:10px 10px 0 0; color:{ANTHRACITE}; padding:0 .85rem; }}
.stTabs [aria-selected="true"] {{ background:{WHITE}; border-bottom:3px solid {OCHRE}; }}
[data-testid="stExpander"] {{ background:rgba(255,253,248,.55); border-color:{SAND}; border-radius:12px; }}
[data-testid="stDataFrame"], [data-testid="stTable"] {{ border:1px solid {SAND}; border-radius:10px; overflow:hidden; }}
hr {{ border-color:{SAND} !important; }}
.fh-note {{ background:{WHITE}; border-left:4px solid {OCHRE}; padding:.75rem 1rem; border-radius:8px; color:{ANTHRACITE}; }}
.fh-table-wrap {{ overflow-x:auto; border:1px solid {SAND}; border-radius:12px; background:{WHITE}; }}
.fh-compare {{ width:100%; border-collapse:separate; border-spacing:0; min-width:1180px; font-size:.78rem; }}
.fh-compare th {{ position:sticky; top:0; z-index:1; background:{BEIGE}; color:{ANTHRACITE}; padding:.55rem .48rem; border-bottom:1px solid {SAND}; text-align:right; white-space:nowrap; }}
.fh-compare th:first-child, .fh-compare td:first-child {{ position:sticky; left:0; z-index:2; text-align:left; min-width:205px; }}
.fh-compare th:first-child {{ z-index:3; }}
.fh-compare td {{ padding:.48rem; border-bottom:1px solid {GRID}; border-right:1px solid {GRID}; text-align:right; vertical-align:top; min-width:92px; }}
.fh-compare td:first-child {{ background:{WHITE}; font-weight:700; color:{ANTHRACITE}; }}
.fh-good {{ background:#E5F1E8; color:{POSITIVE}; }}
.fh-bad {{ background:#F7E5E0; color:{BRICK}; }}
.fh-neutral-cell {{ background:#F4EFE6; color:#776F64; }}
.fh-actual {{ display:block; font-weight:760; }}
.fh-plan {{ display:block; font-size:.68rem; margin-top:.1rem; opacity:.86; }}
.fh-delta {{ display:block; font-size:.68rem; margin-top:.1rem; }}
</style>
""",
    unsafe_allow_html=True,
)


def fmt_number(value: float | None, currency: bool = False) -> str:
    if value is None or abs(float(value)) < 0.0000001:
        return "€ -" if currency else "-"
    value = float(value)
    core = f"{abs(value):,.2f}"
    if currency:
        core = f"€ {core}"
    return f"({core})" if value < 0 else core


def fmt_percent(value: float | None) -> str:
    if value is None:
        return "-"
    return f"{value:,.1f}%"


def metric_card(label: str, value: str, tone: str = "neutral", sub: str = "") -> None:
    st.markdown(
        f'<div class="fh-card"><div class="fh-label">{html.escape(label)}</div>'
        f'<div class="fh-value fh-{tone}">{value}</div>'
        f'<div class="fh-sub">{html.escape(sub)}</div></div>',
        unsafe_allow_html=True,
    )


def chart_style(fig: go.Figure, height: int = 390, legend: bool = True) -> go.Figure:
    fig.update_layout(
        height=height,
        paper_bgcolor=IVORY,
        plot_bgcolor=IVORY,
        font=dict(family="Avenir, Helvetica Neue, Arial", color=ANTHRACITE, size=12),
        margin=dict(l=18, r=18, t=20, b=30),
        colorway=PALETTE,
        hoverlabel=dict(bgcolor=WHITE, font_color=ANTHRACITE),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0) if legend else dict(visible=False),
    )
    fig.update_xaxes(gridcolor=GRID, zerolinecolor=GRID, tickfont=dict(color=ANTHRACITE))
    fig.update_yaxes(gridcolor=GRID, zerolinecolor=GRID, tickfont=dict(color=ANTHRACITE), tickformat=",.2f")
    return fig


def show_plotly(fig: go.Figure) -> None:
    # Da Streamlit 1.51 width è un parametro esplicito; le opzioni Plotly
    # passano solo tramite config, evitando l'avviso sui **kwargs deprecati.
    st.plotly_chart(fig, width="stretch", config={"displayModeBar": False, "responsive": True})


def csv_bytes(rows: list[dict], fieldnames: list[str]) -> bytes:
    buffer = io.StringIO(newline="")
    writer = csv.DictWriter(buffer, fieldnames=fieldnames, extrasaction="ignore", delimiter=";")
    writer.writeheader()
    for row in rows:
        serialized = dict(row)
        for key, value in serialized.items():
            if isinstance(value, date):
                serialized[key] = value.isoformat()
        writer.writerow(serialized)
    return buffer.getvalue().encode("utf-8-sig")


def comparison_table_html(actual: dict[str, list[float]], plan: dict[str, list[float]], actual_through_month: int) -> str:
    parts = ['<div class="fh-table-wrap"><table class="fh-compare"><thead><tr><th>Sezione</th>']
    parts.extend(f"<th>{month}</th>" for month in MONTHS_IT)
    parts.append("<th>Totale</th></tr></thead><tbody>")
    for bucket in ANALYSIS_BUCKETS:
        parts.append(f"<tr><td>{html.escape(bucket)}</td>")
        actual_total = sum(actual[bucket][:actual_through_month])
        plan_total = sum(plan[bucket][:actual_through_month])
        for month_index, (actual_value, plan_value) in enumerate(zip(actual[bucket], plan[bucket]), start=1):
            delta = actual_value - plan_value
            if month_index > actual_through_month or (abs(actual_value) < 0.000001 and abs(plan_value) < 0.000001):
                tone = "fh-neutral-cell"
            else:
                favorable = actual_value >= plan_value if bucket == "Entrate" else actual_value <= plan_value
                tone = "fh-good" if favorable else "fh-bad"
            parts.append(
                f'<td class="{tone}"><span class="fh-actual">{fmt_number(actual_value, True)}</span>'
                f'<span class="fh-plan">Prev. {fmt_number(plan_value, True)}</span>'
                f'<span class="fh-delta">Δ {fmt_number(delta, True)}</span></td>'
            )
        total_favorable = actual_total >= plan_total if bucket == "Entrate" else actual_total <= plan_total
        total_tone = "fh-good" if total_favorable else "fh-bad"
        if actual_through_month == 0 or (abs(actual_total) < 0.000001 and abs(plan_total) < 0.000001):
            total_tone = "fh-neutral-cell"
        parts.append(
            f'<td class="{total_tone}"><span class="fh-actual">{fmt_number(actual_total, True)}</span>'
            f'<span class="fh-plan">Prev. a oggi {fmt_number(plan_total, True)}</span>'
            f'<span class="fh-delta">Δ {fmt_number(actual_total - plan_total, True)}</span></td></tr>'
        )
    parts.append("</tbody></table></div>")
    return "".join(parts)


def transaction_export_rows(rows: list[dict]) -> list[dict]:
    return [
        {
            "Data": row.get("Data"),
            "Importo EUR": round(float(row.get("Importo") or 0.0), 2),
            "Categoria": row.get("Categoria"),
            "Macro-categoria": row.get("Macro-categoria"),
            "Gruppo analisi": analysis_bucket(row),
            "Progetto": row.get("Progetto"),
            "Payee": row.get("Payee"),
            "Conto": row.get("Conto"),
            "Owner": row.get("Owner"),
            "Valuta originale": row.get("Valuta"),
            "Importo originale": row.get("Importo originale"),
            "Note": row.get("Note"),
            "Fonte": row.get("Fonte"),
        }
        for row in rows
    ]


@st.cache_data(show_spinner=False)
def cached_config(config_bytes: bytes):
    return load_config(config_bytes)


@st.cache_data(show_spinner=False)
def cached_transactions(file_payloads, config_bytes):
    config_local = load_config(config_bytes)
    return normalize_transactions(list(file_payloads), config_local)


def end_of_month(d: date) -> date:
    return date(d.year + (1 if d.month == 12 else 0), 1 if d.month == 12 else d.month + 1, 1) - timedelta(days=1)


def budget_rows_for(config, year, start_date, end_date, extraordinary_mode):
    result = []
    for row in config.get("Budget", []):
        if int(parse_number(row.get("Anno")) or year) != year or norm(row.get("Stato")) != "confermato":
            continue
        is_extra = norm(row.get("Classe di spesa")) == "straordinaria"
        if extraordinary_mode == "Solo ordinarie" and is_extra:
            continue
        if extraordinary_mode == "Solo straordinarie" and not is_extra:
            continue
        monthly = budget_monthly_amounts(row, year)
        selected = 0.0
        for month, amount in enumerate(monthly, start=1):
            month_start = date(year, month, 1)
            month_end = end_of_month(month_start)
            overlap_start = max(start_date, month_start)
            overlap_end = min(end_date, month_end)
            if overlap_start <= overlap_end:
                selected += amount * ((overlap_end - overlap_start).days + 1) / ((month_end - month_start).days + 1)
        result.append((row, selected, monthly))
    return result


def filtered_rows(rows, start_date, end_date, owners, accounts, categories, extraordinary_mode):
    output = []
    for tx in rows:
        if tx.get("Trasferimento") or not (start_date <= tx["Data"] <= end_date):
            continue
        if owners and tx["Owner"] not in owners:
            continue
        if accounts and tx["Conto"] not in accounts:
            continue
        if categories and tx["Categoria"] not in categories:
            continue
        is_extra = norm(tx.get("Classe di spesa")) == "straordinaria" or norm(tx.get("Ordinaria/Straordinaria")) == "straordinaria"
        if extraordinary_mode == "Solo ordinarie" and is_extra:
            continue
        if extraordinary_mode == "Solo straordinarie" and not is_extra:
            continue
        output.append(tx)
    return output


def build_quality(config, transactions, ingestion_issues):
    checks = list(ingestion_issues)
    saldi = config.get("Saldi Iniziali", [])
    missing_saldi = [r for r in saldi if norm(r.get("Stato")) != "confermato" or parse_number(r.get("Saldo iniziale (valuta conto)")) is None]
    if missing_saldi:
        checks.append({"Livello": "Avviso", "Ambito": "Saldi iniziali", "Dettaglio": f"{len(missing_saldi)} conti non hanno un saldo iniziale confermato."})
    incomes = config.get("Entrate", [])
    missing_income = [r for r in incomes if norm(r.get("Stato")) != "inattiva" and (parse_number(r.get("Importo")) or 0) <= 0]
    if missing_income:
        checks.append({"Livello": "Avviso", "Ambito": "Entrate", "Dettaglio": f"{len(missing_income)} entrate ricorrenti sono senza importo positivo."})
    seasonal = [r for r in config.get("Budget", []) if norm(r.get("Tipo budget")) == "stagionale" and not (r.get("Mesi stagionali (1-12)") or str(r.get("Mese applicazione") or "").isdigit())]
    if seasonal:
        checks.append({"Livello": "Avviso", "Ambito": "Budget", "Dettaglio": f"{len(seasonal)} budget stagionali non indicano i mesi; sono allocati a dicembre."})
    one_off = [r for r in config.get("Budget", []) if norm(r.get("Tipo budget")) == "una tantum" and not parse_date(r.get("Data una tantum"))]
    if one_off:
        checks.append({"Livello": "Avviso", "Ambito": "Budget", "Dettaglio": f"{len(one_off)} budget una tantum sono senza data."})
    recurrence_candidates = [r for r in config.get("Ricorrenze", []) if norm(r.get("Qualità")) == "da confermare"]
    if recurrence_candidates:
        checks.append({"Livello": "Informazione", "Ambito": "Ricorrenze", "Dettaglio": f"{len(recurrence_candidates)} ricorrenze sono marcate da confermare."})
    mapped = {norm(r.get("Categoria MoneyWiz")) for r in config.get("Mappatura Categorie", [])}
    unmapped = sorted({tx["Categoria"] for tx in transactions if norm(tx["Categoria"]) not in mapped and tx["Categoria"] != "Non classificata"})
    if unmapped:
        checks.append({"Livello": "Avviso", "Ambito": "Mappatura categorie", "Dettaglio": f"{len(unmapped)} categorie caricate non risultano mappate: {', '.join(unmapped[:6])}{'…' if len(unmapped) > 6 else ''}"})
    uncategorized = uncategorized_transactions(transactions)
    if uncategorized:
        checks.append({"Livello": "Avviso", "Ambito": "Categorie", "Dettaglio": f"{len(uncategorized)} transazioni sono senza categoria, dopo l'esclusione dei trasferimenti interni."})
    non_eur_nonzero = []
    for row in saldi:
        amount = parse_number(row.get("Saldo iniziale (valuta conto)")) or 0.0
        currency = str(row.get("Valuta") or "").upper()
        rate = parse_number(row.get("Tasso EUR per unità"))
        if currency != "EUR" and abs(amount) > 0.000001 and rate is None and currency not in {"JPY", "ISK", "EGP"}:
            non_eur_nonzero.append(row)
    if non_eur_nonzero:
        checks.append({"Livello": "Avviso", "Ambito": "Valute", "Dettaglio": f"{len(non_eur_nonzero)} saldi non-EUR sono diversi da zero ma non hanno un tasso EUR disponibile."})
    if not transactions:
        checks.append({"Livello": "Informazione", "Ambito": "Transazioni", "Dettaglio": "Nessun export transazionale caricato: le metriche Actual restano a zero."})
    if not checks:
        checks.append({"Livello": "OK", "Ambito": "Controlli", "Dettaglio": "Nessuna anomalia rilevata."})
    return checks


# ----- Data source -----
with st.sidebar:
    st.markdown("## Finance Hub")
    st.caption(f"{APP_VERSION} · Applicazione per controllo familiare delle finanze")

    with st.expander("Caricamenti", expanded=True):
        config_upload = st.file_uploader(
            "Configurazione (.xlsx)",
            type=["xlsx"],
            help="Carica il file Config_Finanze_Familiari_V2_Aggiornato.xlsx.",
       )

    tx_uploads = st.file_uploader(
        "Transazioni MoneyWiz",
        type=["csv", "xlsx"],
        accept_multiple_files=True,
        help="Carica uno o più export MoneyWiz.",
    )

    st.caption(
        "I file vengono elaborati nella sessione dell'app."
    )

if config_upload is None:
    st.title("Finance Hub")
    st.info("Carica il file Config_Finanze_Familiari_V2_Aggiornato.xlsx")
    st.stop()

config_bytes = config_upload.getvalue()
config = cached_config(config_bytes)

file_payloads = tuple(
(item.name, item.getvalue())
for item in (tx_uploads or [])
)

transactions, ingestion_issues = (
    cached_transactions(
        file_payloads,
        config_bytes,
)
if file_payloads
else ([], [])
)
# ----- Filters -----
today = date.today()
budget_years = {int(parse_number(r.get("Anno")) or today.year) for r in config.get("Budget", [])}
tx_years = {tx["Data"].year for tx in transactions}
years = sorted(budget_years | tx_years | {today.year}, reverse=True)
with st.sidebar:
    with st.expander("Filtri", expanded=False):
        year = st.selectbox("Anno", years, index=0)
        period_mode = st.selectbox("Periodo", ["YTD", "Anno intero", "Personalizzato"])
        if period_mode == "YTD":
            start_date = date(year, 1, 1)
            end_date = min(today, date(year, 12, 31)) if year == today.year else date(year, 12, 31)
        elif period_mode == "Anno intero":
            start_date, end_date = date(year, 1, 1), date(year, 12, 31)
        else:
            chosen = st.date_input("Intervallo", value=(date(year, 1, 1), min(today, date(year, 12, 31))))
            if isinstance(chosen, (tuple, list)) and len(chosen) == 2:
                start_date, end_date = chosen
            else:
                start_date, end_date = date(year, 1, 1), min(today, date(year, 12, 31))
        owner_options = sorted({str(r.get("Owner")) for r in config.get("Conti", []) if r.get("Owner")} | {tx["Owner"] for tx in transactions})
        account_options = sorted({str(r.get("Conto")) for r in config.get("Conti", []) if r.get("Conto")})
        category_options = sorted({str(r.get("Categoria MoneyWiz")) for r in config.get("Mappatura Categorie", []) if r.get("Categoria MoneyWiz")})
        owners = st.multiselect("Owner", owner_options, placeholder="Tutti")
        accounts = st.multiselect("Conti", account_options, placeholder="Tutti")
        categories = st.multiselect("Categorie", category_options, placeholder="Tutte")
        extraordinary_mode = st.selectbox("Straordinari = spese eccezionali o non ordinarie (ad esempio matrimonio, auto nuova, mobili, grandi lavori)", ["Includi tutto", "Solo ordinarie", "Solo straordinarie"])
    st.caption("Conversione EUR attiva: prevale Importo in EUR; fallback su tassi configurati/storici.")

selected = filtered_rows(transactions, start_date, end_date, owners, accounts, categories, extraordinary_mode)
expenses = [tx for tx in selected if tx["Importo"] < 0]
incomes = [tx for tx in selected if tx["Importo"] > 0]
actual_income = sum(tx["Importo"] for tx in incomes)
actual_expenses = -sum(tx["Importo"] for tx in expenses)
actual_result = actual_income - actual_expenses
actual_rate = actual_result / actual_income * 100 if actual_income else None
budget_selected = budget_rows_for(config, year, start_date, end_date, extraordinary_mode)

planned_income = annual_income(config, year)
annual_total_budget = annual_budget(config, year)
fixed_monthly = monthly_recurring(config, "Fissa")
nonfixed_annual = annual_budget(config, year, {"Ricorrente variabile"})
nonfixed_monthly = nonfixed_annual / 12.0
forecast_spending = actual_expenses
forecast_income = actual_income
for month in range(end_date.month + 1, 13):
    forecast_spending += sum(monthly[month - 1] for _, _, monthly in budget_selected)
forecast_income += planned_income / 12.0 * max(0, 12 - end_date.month)
if end_date >= date(year, 12, 31):
    forecast_spending, forecast_income = actual_expenses, actual_income
elif not transactions:
    forecast_spending, forecast_income = annual_total_budget, planned_income
forecast_savings = forecast_income - forecast_spending
forecast_rate = forecast_savings / forecast_income * 100 if forecast_income else None
snapshot = end_date
accounting_nw = net_worth(config, transactions, snapshot, accounting=True)
total_nw = net_worth(config, transactions, snapshot, accounting=False)

# ----- Header -----
st.title("Finance Hub")
st.markdown(
    f'<span class="fh-chip">{start_date.strftime("%d/%m/%Y")} – {end_date.strftime("%d/%m/%Y")}</span>'
    f'<span class="fh-chip">{len(selected):,} transazioni</span>'
    f'<span class="fh-chip">Configurazione: {html.escape(config_upload.name if config_upload else DEFAULT_CONFIG.name)}</span>',
    unsafe_allow_html=True,
)

recap_tab, history_tab, budget_tab, forecast_tab, compare_tab, patrimony_tab, expense_tab, travel_tab, quality_tab = st.tabs(
    ["Recap", "Storico", "Budget vs Actual", "Previsione", "Storico vs Previsione", "Patrimonio", "Dettaglio spese per categoria", "Analisi viaggi", "Data quality"]
)

# ----- Recap -----
with recap_tab:
    st.subheader("Recap")
    row1 = st.columns(4)
    with row1[0]: metric_card("Entrate attuali YTD", fmt_number(actual_income, True), "positive", "Transazioni positive nel periodo")
    with row1[1]: metric_card("Spese attuali YTD", fmt_number(actual_expenses, True), "negative", "Transazioni negative, esclusi trasferimenti")
    with row1[2]: metric_card("Risultato YTD", fmt_number(actual_result, True), "positive" if actual_result >= 0 else "negative", "Entrate meno spese")
    with row1[3]: metric_card("Tasso di risparmio attuale", fmt_percent(actual_rate), "positive" if actual_rate is not None and actual_rate >= 0 else "negative", "Risultato / entrate")
    st.write("")
    row2 = st.columns(4)
    with row2[0]: metric_card("Patrimonio contabile", fmt_number(accounting_nw, True), "positive" if accounting_nw >= 0 else "negative", "Conti inclusi + attività − debiti contabili")
    with row2[1]: metric_card("Patrimonio complessivo", fmt_number(total_nw, True), "positive" if total_nw >= 0 else "negative", "Include beni e finanziamento Mercedes")
    with row2[2]: metric_card("Forecast risultato anno completo", fmt_number(forecast_savings, True), "positive" if forecast_savings >= 0 else "negative", "Risparmio previsto / spesa prevista")
    with row2[3]: metric_card("Tasso di risparmio anno completo", fmt_percent(forecast_rate), "positive" if forecast_rate is not None and forecast_rate >= 0 else "negative", "Forecast risparmio / entrate")

    st.markdown("### Cash flow")
    monthly = {m: {"Entrate": 0.0, "Spese": 0.0, "Risultato": 0.0} for m in range(1, 13)}
    for tx in selected:
        m = tx["Data"].month
        if tx["Importo"] >= 0:
            monthly[m]["Entrate"] += tx["Importo"]
        else:
            monthly[m]["Spese"] += abs(tx["Importo"])
        monthly[m]["Risultato"] += tx["Importo"]
    months = list(range(start_date.month, end_date.month + 1))
    fig = go.Figure()
    fig.add_bar(name="Entrate", x=[MONTHS_IT[m - 1] for m in months], y=[monthly[m]["Entrate"] for m in months], marker_color=SAGE_LIGHT, hovertemplate="%{x}<br>Entrate € %{y:,.2f}<extra></extra>")
    fig.add_bar(name="Spese", x=[MONTHS_IT[m - 1] for m in months], y=[monthly[m]["Spese"] for m in months], marker_color=BRICK_LIGHT, hovertemplate="%{x}<br>Spese € %{y:,.2f}<extra></extra>")
    fig.add_scatter(name="Cash flow netto", x=[MONTHS_IT[m - 1] for m in months], y=[monthly[m]["Risultato"] for m in months], mode="lines+markers", line=dict(color=OCHRE, width=4), marker=dict(color=OCHRE, size=8), hovertemplate="%{x}<br>Cash flow € %{y:,.2f}<extra></extra>")
    fig.update_layout(barmode="group")
    show_plotly(chart_style(fig))
    if not transactions:
        st.info("Carica uno o più export MoneyWiz dalla sezione Caricamenti per popolare le metriche Actual e il cash flow.")

# ----- Budget vs Actual -----
with budget_tab:
    st.subheader("Analisi di spese vs budget")
    budget_macro = defaultdict(float)
    for row, selected_amount, _ in budget_selected:
        category = str(row.get("Categoria MoneyWiz") or "Non classificata")
        macro = category.split(" > ")[0]
        budget_macro[macro] += selected_amount
    actual_macro = group_sum(expenses, "Macro-categoria", absolute=True)
    macros = sorted(set(budget_macro) | set(actual_macro), key=lambda x: actual_macro.get(x, 0) + budget_macro.get(x, 0), reverse=True)
    fig = go.Figure()
    fig.add_bar(name="Budget", x=macros, y=[budget_macro.get(x, 0) for x in macros], marker_color=SAND, hovertemplate="%{x}<br>Budget € %{y:,.2f}<extra></extra>")
    fig.add_bar(name="Actual", x=macros, y=[actual_macro.get(x, 0) for x in macros], marker_color=BRICK, hovertemplate="%{x}<br>Actual € %{y:,.2f}<extra></extra>")
    fig.update_layout(barmode="group")
    fig.update_xaxes(tickangle=-35)
    show_plotly(chart_style(fig, 440))

    st.subheader("Top categorie di spese")
    actual_category = group_sum(expenses, "Categoria", absolute=True)
    top = sorted(actual_category.items(), key=lambda item: item[1], reverse=True)[:12]
    total_top_base = sum(actual_category.values())
    top = list(reversed(top))
    fig = go.Figure(go.Bar(
        x=[v for _, v in top], y=[k for k, _ in top], orientation="h", marker_color=OCHRE_LIGHT,
        text=[f"€ {v:,.2f} · {(v / total_top_base * 100 if total_top_base else 0):.1f}%" for _, v in top], textposition="outside",
        hovertemplate="%{y}<br>€ %{x:,.2f}<extra></extra>",
    ))
    fig.update_xaxes(range=[0, max([v for _, v in top] or [1]) * 1.32])
    show_plotly(chart_style(fig, max(330, 34 * len(top)), legend=False))

    st.subheader("Budget annuale")
    annual_macro = defaultdict(float)
    for row in config.get("Budget", []):
        if int(parse_number(row.get("Anno")) or year) == year and norm(row.get("Stato")) == "confermato":
            annual_macro[str(row.get("Categoria MoneyWiz") or "Non classificata").split(" > ")[0]] += annual_budget_amount(row)
    ordered = sorted(annual_macro.items(), key=lambda item: item[1], reverse=True)
    fig = go.Figure(go.Pie(labels=[k for k, _ in ordered], values=[v for _, v in ordered], hole=.58, marker=dict(colors=PALETTE), textinfo="percent", hovertemplate="%{label}<br>€ %{value:,.2f}<br>%{percent}<extra></extra>"))
    fig.add_annotation(text=f"€ {sum(annual_macro.values()):,.2f}", x=.5, y=.5, showarrow=False, font=dict(size=20, color=ANTHRACITE))
    show_plotly(chart_style(fig, 430))

# ----- Forecast -----
with forecast_tab:
    st.subheader("Previsione")
    cols = st.columns(5)
    with cols[0]: metric_card("Entrate annuali pianificate", fmt_number(planned_income, True), "positive", "Foglio Entrate")
    with cols[1]: metric_card("Spese fisse e ricorrenti mensili", fmt_number(fixed_monthly, True), "negative", "Ricorrenze attive · classe Fissa")
    with cols[2]: metric_card("Spese non fisse ma ricorrenti", fmt_number(nonfixed_monthly, True), "negative", "Budget ricorrente variabile / 12")
    with cols[3]: metric_card("Budget annuo totale", fmt_number(annual_total_budget, True), "neutral", "Tutte le righe confermate")
    predicted_savings = planned_income - annual_total_budget
    with cols[4]: metric_card("Risparmio previsto", fmt_number(predicted_savings, True), "positive" if predicted_savings >= 0 else "negative", "Entrate pianificate − budget")

    class_months = {"Fissa": [0.0] * 12, "Ricorrente variabile": [0.0] * 12, "Straordinaria": [0.0] * 12}
    for row in config.get("Budget", []):
        cls = str(row.get("Classe di spesa") or "Ricorrente variabile")
        if cls not in class_months:
            cls = "Ricorrente variabile"
        amounts = budget_monthly_amounts(row, year)
        class_months[cls] = [a + b for a, b in zip(class_months[cls], amounts)]
    fig = go.Figure()
    forecast_classes = [
        ("Fissa", SAGE),
        ("Ricorrente variabile", OCHRE_LIGHT),
        ("Straordinaria", BRICK),
    ]

    for cls, color in forecast_classes:
        fig.add_bar(
            name=cls,
            x=MONTHS_IT,
            y=class_months[cls],
            marker_color=color,
            customdata=[
                [cls, month_number]
                for month_number in range(1, 13)
    ],
    hovertemplate=(
        "%{x}<br>"
        + cls
        + " € %{y:,.2f}"
        + "<extra></extra>"
    ),
    )
    fig.update_layout(
        barmode="stack",
        clickmode="event+select",
    )

    fig = chart_style(fig, 420)

    forecast_event = st.plotly_chart(
        fig,
        width="stretch",
        key="forecast_interactive_chart",
        on_select="rerun",
        selection_mode="points",
        config={
            "displayModeBar": False,
            "responsive": True,
        },
    )
    forecast_selection = (
        forecast_event.get("selection", {})\
        if forecast_event
        else {}
    )

    forecast_points = forecast_selection.get("points", [])

    if forecast_points:
        selected_point = forecast_points[0]
        selected_data = selected_point.get("customdata", [])

        if len(selected_data) >= 2:
            selected_class = str(selected_data[0])
            selected_month = int(selected_data[1])

            st.markdown(
                f"### Dettaglio {selected_class} · "
                f"{MONTHS_IT[selected_month - 1]} {year}"
            )

            forecast_detail = []

            for row in config.get("Budget", []):
                row_year = int(
                parse_number(row.get("Anno")) or year
            )

            row_status = norm(row.get("Stato"))
            row_class = str(
                row.get("Classe di spesa")
                or "Ricorrente variabile"
            )

            if row_class not in class_months:
                row_class = "Ricorrente variabile"

            if row_year != year:
                continue

            if row_status != "confermato":
                continue

            if norm(row_class) != norm(selected_class):
                continue

            monthly_amounts = budget_monthly_amounts(row, year)
            selected_amount = monthly_amounts[selected_month - 1]

            if abs(selected_amount) < 0.000001:
                continue

            forecast_detail.append(
                {
                    "Categoria": row.get("Categoria MoneyWiz"),
                    "Owner": row.get("Owner"),
                    "Tipo budget": row.get("Tipo budget"),
                    "Classe di spesa": row_class,
                    "Importo previsto": fmt_number(
                        selected_amount,
                        True,
                    ),
                    "Note": (
                        row.get("Note utente")
                        or row.get("Nota tecnica")
                        or ""
                    ),
                }
            )

            if forecast_detail:
                st.dataframe(
                    forecast_detail,
                    width="stretch",
                    hide_index=True,
                )
        else:
            st.info(
                "Nessun dettaglio configurato per la "
                "selezione corrente."
            )
        else:
            st.caption(
                "Seleziona una sezione di una colonna per "
                "visualizzare il dettaglio delle previsioni."
            )
    if planned_income <= 0:
        st.warning("Nel file incluso le righe del foglio Entrate risultano ancora con importo 0; il risparmio previsto resta quindi negativo finché gli importi non vengono compilati.")

# ----- Historical vs Forecast -----
with compare_tab:
    st.subheader("Storico vs Previsione")
    st.caption("Confronto mensile per l'anno selezionato. Gli actual sono espressi come valori positivi sia per le entrate sia per le spese.")
    comparison_rows = filtered_rows(
        transactions, date(year, 1, 1), date(year, 12, 31), owners, accounts, categories, extraordinary_mode
    )
    comparison_actual = monthly_actual_by_bucket(comparison_rows, year)
    comparison_plan = monthly_plan_by_bucket(config, year)
    actual_dates = [row["Data"] for row in comparison_rows if row["Data"].year == year]
    actual_through_month = max((item.month for item in actual_dates), default=0)

    through_label = MONTHS_IT[actual_through_month - 1] if actual_through_month else "nessun mese"
    actual_income_to_date = sum(comparison_actual["Entrate"][:actual_through_month])
    plan_income_to_date = sum(comparison_plan["Entrate"][:actual_through_month])
    expense_buckets = [bucket for bucket in ANALYSIS_BUCKETS if bucket != "Entrate"]
    actual_expense_to_date = sum(sum(comparison_actual[bucket][:actual_through_month]) for bucket in expense_buckets)
    plan_expense_to_date = sum(sum(comparison_plan[bucket][:actual_through_month]) for bucket in expense_buckets)
    cols = st.columns(4)
    with cols[0]: metric_card("Actual entrate", fmt_number(actual_income_to_date, True), "positive", f"Gen–{through_label}")
    with cols[1]: metric_card("Previsione entrate", fmt_number(plan_income_to_date, True), "neutral", f"Gen–{through_label}")
    with cols[2]: metric_card("Actual spese", fmt_number(actual_expense_to_date, True), "negative", f"Gen–{through_label}")
    with cols[3]: metric_card("Previsione spese", fmt_number(plan_expense_to_date, True), "neutral", f"Gen–{through_label}")

    st.markdown(comparison_table_html(comparison_actual, comparison_plan, actual_through_month), unsafe_allow_html=True)
    st.markdown(
        '<div class="fh-note"><b>Logica colori:</b> per le entrate una cella è verde quando Actual ≥ Previsione; '
        'per tutte le spese è verde quando Actual ≤ Previsione. È rossa in caso contrario. I mesi futuri o privi di valori sono neutri. '
        'Le linee Matrimonio, Casa nuova, Luna di miele (Giappone) e Altri viaggi prevalgono sulle classi ordinarie.</div>',
        unsafe_allow_html=True,
    )
    compare_export = []
    for bucket in ANALYSIS_BUCKETS:
        for month in range(1, 13):
            compare_export.append({
                "Anno": year,
                "Mese": month,
                "Mese nome": MONTHS_IT[month - 1],
                "Sezione": bucket,
                "Actual EUR": round(comparison_actual[bucket][month - 1], 2),
                "Previsione EUR": round(comparison_plan[bucket][month - 1], 2),
                "Scostamento EUR": round(comparison_actual[bucket][month - 1] - comparison_plan[bucket][month - 1], 2),
            })
    st.download_button(
        "Esporta confronto CSV",
        data=csv_bytes(compare_export, list(compare_export[0].keys())),
        file_name=f"storico_vs_previsione_{year}.csv",
        mime="text/csv",
    )
    if not transactions:
        st.info("Carica uno o più export MoneyWiz per visualizzare gli actual nel confronto mensile.")

# ----- Historical actuals -----
with history_tab:
    st.subheader("Storico")
    st.caption("Vista basata esclusivamente sulle transazioni actual, nel periodo e con i filtri selezionati.")
    historical_rows = sorted(selected, key=lambda item: item["Data"], reverse=True)
    historical_monthly = defaultdict(lambda: {"Entrate": 0.0, "Spese": 0.0, "Risultato": 0.0})
    for row in historical_rows:
        key = row["Data"].strftime("%Y-%m")
        amount = float(row.get("Importo") or 0.0)
        if amount >= 0:
            historical_monthly[key]["Entrate"] += amount
        else:
            historical_monthly[key]["Spese"] += abs(amount)
        historical_monthly[key]["Risultato"] += amount
    historical_month_numbers = list(
        range(start_date.month, end_date.month + 1)
    )

    historical_keys = [
        f"{year}-{month:02d}"
        for month in historical_month_numbers
    ]

    historical_labels = [
        MONTHS_IT[month - 1]
        for month in historical_month_numbers
    ]

    historical_income_values = [
        historical_monthly[key]["Entrate"]
        for key in historical_keys
    ]

    historical_expense_values = [
        historical_monthly[key]["Spese"]
        for key in historical_keys
    ]

    historical_result_values = [
        historical_monthly[key]["Risultato"]
        for key in historical_keys
    ]

fig = go.Figure()

fig.add_bar(
    name="Entrate actual",
    x=historical_labels,
    y=historical_income_values,
    marker_color=SAGE_LIGHT,
    hovertemplate=(
        "%{x}<br>"
        "Entrate € %{y:,.2f}"
        "<extra></extra>"
    ),
)

fig.add_bar(
    name="Spese actual",
    x=historical_labels,
    y=historical_expense_values,
    marker_color=BRICK_LIGHT,
    hovertemplate=(
        "%{x}<br>"
        "Spese € %{y:,.2f}"
        "<extra></extra>"
    ),
)

fig.add_scatter(
    name="Risultato actual",
    x=historical_labels,
    y=historical_result_values,
    mode="lines+markers",
    line=dict(
        color=OCHRE,
        width=4,
    ),
    marker=dict(
        color=OCHRE,
        size=8,
    ),
    hovertemplate=(
        "%{x}<br>"
        "Risultato € %{y:,.2f}"
        "<extra></extra>"
    ),
)

fig.update_layout(
    barmode="group",
    bargap=0.18,
)

fig.update_xaxes(
    type="category",
    categoryorder="array",
    categoryarray=historical_labels,
    tickmode="array",
    tickvals=historical_labels,
    ticktext=historical_labels,
    tickangle=0,
    automargin=True,
)

show_plotly(
    chart_style(
        fig,
        430,
    )
)

    historical_export = transaction_export_rows(historical_rows)
    historical_display = [
        {
            "Data": row["Data"].strftime("%d/%m/%Y"),
            "Importo EUR": fmt_number(row["Importo"], True),
            "Categoria": row["Categoria"],
            "Gruppo analisi": analysis_bucket(row),
            "Progetto": row["Progetto"],
            "Payee": row["Payee"],
            "Conto": row["Conto"],
            "Owner": row["Owner"],
            "Fonte": row["Fonte"],
        }
        for row in historical_rows
    ]
    st.markdown("### Dettaglio actual")
    st.dataframe(historical_display, width="stretch", hide_index=True)
    if historical_export:
        st.download_button(
            "Esporta storico actual CSV",
            data=csv_bytes(historical_export, list(historical_export[0].keys())),
            file_name=f"storico_actual_{start_date.isoformat()}_{end_date.isoformat()}.csv",
            mime="text/csv",
        )
    if not historical_rows:
        st.info("Nessuna transazione actual nel periodo e nei filtri selezionati.")

# ----- Patrimony -----
with patrimony_tab:
    st.subheader("Patrimonio Netto")
    cols = st.columns(2)
    with cols[0]: metric_card("Patrimonio contabile", fmt_number(accounting_nw, True), "positive" if accounting_nw >= 0 else "negative", "Finanziamento Mercedes escluso")
    with cols[1]: metric_card("Patrimonio complessivo", fmt_number(total_nw, True), "positive" if total_nw >= 0 else "negative", "Finanziamento Mercedes incluso")

    st.markdown("### Evoluzione nel tempo del patrimonio netto")
    history_scope = st.radio(
        "Periodo patrimonio",
        ["Tutto lo storico", "Anno selezionato"],
        horizontal=True,
        label_visibility="collapsed",
        key="patrimony_history_scope",
    )
    candidate_dates = [tx["Data"] for tx in transactions]
    candidate_dates.extend(
        parsed for row in config.get("Saldi Iniziali", [])
        if (parsed := parse_date(row.get("Data saldo iniziale"))) is not None
    )
    candidate_dates.extend(
        parsed for row in config.get("Patrimonio", [])
        for parsed in [parse_date(row.get("Data iniziale")), parse_date(row.get("Data valutazione"))]
        if parsed is not None
    )
    earliest_available = min(candidate_dates, default=date(year, 1, 1))
    history_start = date(year, 1, 1) if history_scope == "Anno selezionato" else earliest_available
    history_end = end_date
    timeline = historical_snapshots(history_start, history_end)
    accounting_series = [net_worth(config, transactions, d, True) for d in timeline]
    total_series = [net_worth(config, transactions, d, False) for d in timeline]
    fig = go.Figure()
    labels = [
        (
            MONTHS_IT[snapshot_date.month - 1]
            if history_scope == "Anno selezionato"
            else (
                f"{MONTHS_IT[snapshot_date.month - 1]} "
                f"{snapshot_date.year}"
            )
        )
        for snapshot_date in timeline
    ]

    hover_labels = [
        snapshot_date.strftime("%d/%m/%Y")
        for snapshot_date in timeline
    ]

    fig.add_scatter(
        name="Patrimonio contabile",
        x=labels,
        y=accounting_series,
        customdata=hover_labels,
        mode="lines+markers",
        line=dict(
            color=SAGE,
            width=4,
        ),
        marker=dict(
            color=SAGE,
            size=8,
        ),
        yaxis="y",
        hovertemplate=(
            "%{customdata}<br>"
            "Patrimonio contabile € %{y:,.2f}"
            "<extra></extra>"
        ),
    )

    fig.add_scatter(
        name="Patrimonio complessivo",
        x=labels,
        y=total_series,
        customdata=hover_labels,
        mode="lines+markers",
        line=dict(
            color=OCHRE,
            width=4,
        ),
        marker=dict(
            color=OCHRE,
            size=8,
        ),
        yaxis="y2",
        hovertemplate=(
            "%{customdata}<br>"
            "Patrimonio complessivo € %{y:,.2f}"
            "<extra></extra>"
        ),
    )

    fig = chart_style(
        fig,
        430,
    )

    fig.update_layout(
        yaxis=dict(
            title="Patrimonio contabile (€)",
            gridcolor=GRID,
            zerolinecolor=GRID,
            tickformat=",.2f",
            side="left",
        ),
        yaxis2=dict(
            title="Patrimonio complessivo (€)",
            overlaying="y",
            side="right",
            showgrid=False,
            zeroline=False,
            tickformat=",.2f",
        ),

        margin=dict(
            l=25,
            r=65,
            t=30,
            b=35,
        ),

    )

    fig.update_xaxes(
        type="category",
        categoryorder="array",
        categoryarray=labels,
        tickmode="array",
        tickvals=labels,
        ticktext=labels,
        tickangle=0,
        automargin=True,
    )
    show_plotly(fig)
    patrimony_history = []
    previous_accounting = previous_total = None
    for snapshot_date, accounting_value, total_value in zip(timeline, accounting_series, total_series):
        patrimony_history.append({
            "Data": snapshot_date,
            "Patrimonio contabile EUR": round(accounting_value, 2),
            "Δ contabile EUR": round(accounting_value - previous_accounting, 2) if previous_accounting is not None else None,
            "Patrimonio complessivo EUR": round(total_value, 2),
            "Δ complessivo EUR": round(total_value - previous_total, 2) if previous_total is not None else None,
        })
        previous_accounting, previous_total = accounting_value, total_value
    st.dataframe(
        [
            {
                "Data": row["Data"].strftime("%d/%m/%Y"),
                "Patrimonio contabile": fmt_number(row["Patrimonio contabile EUR"], True),
                "Δ contabile": fmt_number(row["Δ contabile EUR"], True) if row["Δ contabile EUR"] is not None else "—",
                "Patrimonio complessivo": fmt_number(row["Patrimonio complessivo EUR"], True),
                "Δ complessivo": fmt_number(row["Δ complessivo EUR"], True) if row["Δ complessivo EUR"] is not None else "—",
            }
            for row in reversed(patrimony_history)
        ],
        width="stretch",
        hide_index=True,
    )
    if patrimony_history:
        st.download_button(
            "Esporta patrimonio storico CSV",
            data=csv_bytes(patrimony_history, list(patrimony_history[0].keys())),
            file_name=f"patrimonio_storico_fino_al_{history_end.isoformat()}.csv",
            mime="text/csv",
        )
    st.caption("La serie ricostruisce il patrimonio alle chiusure mensili disponibili e include la data finale del filtro quando il mese è ancora aperto.")

    st.markdown("### Componenti del patrimonio")
    patrimony_metrics = patrimony_summary(config, transactions, snapshot)
    pcols = st.columns(4)
    with pcols[0]: metric_card("Attività finanziarie", fmt_number(patrimony_metrics["Attività finanziarie"], True), "positive", "Conti e investimenti finanziari")
    with pcols[1]: metric_card("Beni fisici", fmt_number(patrimony_metrics["Beni fisici"], True), "neutral", "Immobili, auto e altri beni")
    with pcols[2]: metric_card("Debiti", fmt_number(patrimony_metrics["Debiti"], True), "negative", "Debiti inclusi nel patrimonio complessivo")
    with pcols[3]: metric_card("Patrimonio Netto", fmt_number(patrimony_metrics["Patrimonio Netto"], True), "positive" if patrimony_metrics["Patrimonio Netto"] >= 0 else "negative", "Attività finanziarie + beni fisici − debiti")
    st.write("")
    patrimony_table = []
    for row in config.get("Patrimonio", []):
        value = parse_number(row.get("Valore corrente")) or 0.0
        signed = -abs(value) if norm(row.get("Classe")) == "debito" else value
        patrimony_table.append({
            "Voce": row.get("Nome"), "Classe": row.get("Classe"), "Owner": row.get("Owner"),
            "Valore corrente": fmt_number(signed, True),
            "Patrimonio contabile": row.get("Incluso patrimonio contabile"),
            "Patrimonio complessivo": row.get("Incluso patrimonio complessivo"),
            "Data valutazione": str(row.get("Data valutazione") or "—"),
        })
    st.dataframe(patrimony_table, width="stretch", hide_index=True)
    st.caption("Il finanziamento Mercedes è trattato come debito separato: escluso dal patrimonio contabile e incluso nel patrimonio complessivo.")

# ----- Expenses -----
with expense_tab:
    st.subheader("Dettaglio spese per categoria")
    category_spend = group_sum(expenses, "Categoria", absolute=True)
    total_spend = sum(category_spend.values())
    ranked = sorted(category_spend.items(), key=lambda item: item[1], reverse=True)[:18]
    display = list(reversed(ranked))
    fig = go.Figure(go.Bar(
        x=[v for _, v in display], y=[k for k, _ in display], orientation="h",
        marker=dict(color=[PALETTE[i % len(PALETTE)] for i in range(len(display))]),
        text=[f"€ {v:,.2f} · {(v / total_spend * 100 if total_spend else 0):.1f}%" for _, v in display],
        textposition="outside", cliponaxis=False, hovertemplate="%{y}<br>€ %{x:,.2f}<extra></extra>",
    ))
    fig.update_xaxes(range=[0, max([v for _, v in display] or [1]) * 1.38])
    show_plotly(chart_style(fig, max(390, 34 * len(display)), legend=False))
    spend_table = [
        {"Categoria": cat, "Spesa": fmt_number(value, True), "% sul totale": f"{(value / total_spend * 100 if total_spend else 0):.1f}%"}
        for cat, value in ranked
    ]
    st.dataframe(spend_table, width="stretch", hide_index=True)
    if not expenses:
        st.info("Nessuna spesa nel periodo e nei filtri selezionati.")

# ----- Travel -----
with travel_tab:
    st.subheader("Analisi viaggi")
    all_travel_expenses = [tx for tx in expenses if norm(tx.get("Macro-categoria")) == "viaggi" or norm(tx.get("Categoria")).startswith("viaggi")]
    project_options = sorted({str(tx.get("Progetto") or "Non assegnato") for tx in all_travel_expenses}, key=norm)
    selected_project = st.selectbox("Viaggio / progetto", ["Tutti i viaggi / progetti"] + project_options, key="travel_project")
    travel_expenses = all_travel_expenses if selected_project == "Tutti i viaggi / progetti" else [tx for tx in all_travel_expenses if str(tx.get("Progetto") or "Non assegnato") == selected_project]
    travel_categories = group_sum(travel_expenses, "Categoria", absolute=True)
    travel_total = sum(travel_categories.values())
    travel_ranked = sorted(travel_categories.items(), key=lambda item: item[1], reverse=True)
    travel_display = list(reversed(travel_ranked))
    fig = go.Figure(go.Bar(
        x=[v for _, v in travel_display], y=[k.replace("Viaggi > ", "") for k, _ in travel_display], orientation="h",
        marker=dict(color=[TRAVEL_PALETTE[i % len(TRAVEL_PALETTE)] for i in range(len(travel_display))]),
        text=[f"€ {v:,.2f} · {(v / travel_total * 100 if travel_total else 0):.1f}%" for _, v in travel_display],
        textposition="outside", cliponaxis=False, hovertemplate="%{y}<br>€ %{x:,.2f}<extra></extra>",
    ))
    fig.update_xaxes(range=[0, max([v for _, v in travel_display] or [1]) * 1.38])
    show_plotly(chart_style(fig, max(350, 38 * len(travel_display)), legend=False))
    travel_budget = sum(annual_budget_amount(r) for r in config.get("Budget", []) if str(r.get("Categoria MoneyWiz") or "").startswith("Viaggi") and int(parse_number(r.get("Anno")) or year) == year)
    cols = st.columns(3)
    with cols[0]: metric_card("Spesa viaggi", fmt_number(travel_total, True), "negative", "Actual nel periodo")
    with cols[1]: metric_card("Budget viaggi annuale", fmt_number(travel_budget, True), "neutral", "Configurazione confermata")
    with cols[2]: metric_card("Residuo budget", fmt_number(travel_budget - travel_total, True), "positive" if travel_budget >= travel_total else "negative", "Budget meno Actual")
    st.markdown("### Dettaglio per categoria")
    travel_table = [
        {
            "Categoria": category.replace("Viaggi > ", ""),
            "Spesa": fmt_number(value, True),
            "% sul viaggio/progetto": f"{(value / travel_total * 100 if travel_total else 0):.1f}%",
        }
        for category, value in travel_ranked
    ]
    st.dataframe(travel_table, width="stretch", hide_index=True)
    if not travel_expenses:
        st.info("Nessuna transazione Viaggi nel periodo selezionato.")

# ----- Data quality -----
with quality_tab:
    st.subheader("Data quality")
    checks = build_quality(config, transactions, ingestion_issues)
    uncategorized = uncategorized_transactions(transactions)
    counts = defaultdict(int)
    for item in checks:
        counts[item["Livello"]] += 1
    cols = st.columns(5)
    with cols[0]: metric_card("Errori", str(counts["Errore"]), "negative" if counts["Errore"] else "positive", "File o righe non elaborabili")
    with cols[1]: metric_card("Avvisi", str(counts["Avviso"]), "negative" if counts["Avviso"] else "positive", "Configurazioni da completare")
    with cols[2]: metric_card("Informazioni", str(counts["Informazione"]), "neutral", "Note operative")
    with cols[3]: metric_card("Transazioni valide", f"{len(transactions):,}", "positive", "Dopo deduplica e controlli")
    with cols[4]: metric_card("Senza categoria", f"{len(uncategorized):,}", "negative" if uncategorized else "positive", "Trasferimenti interni esclusi")
    st.dataframe(checks, width="stretch", hide_index=True)
    st.markdown("### Transazioni senza categoria")
    st.caption("La lista usa tutte le transazioni caricate ed esclude i trasferimenti interni riconosciuti.")
    uncategorized_export = transaction_export_rows(uncategorized)
    uncategorized_display = [
        {
            "Data": row["Data"].strftime("%d/%m/%Y"),
            "Importo EUR": fmt_number(row["Importo"], True),
            "Payee": row["Payee"],
            "Conto": row["Conto"],
            "Owner": row["Owner"],
            "Note": row["Note"],
            "Fonte": row["Fonte"],
        }
        for row in uncategorized
    ]
    st.dataframe(uncategorized_display, width="stretch", hide_index=True)
    if uncategorized_export:
        st.download_button(
            "Esporta transazioni senza categoria CSV",
            data=csv_bytes(uncategorized_export, list(uncategorized_export[0].keys())),
            file_name="transazioni_senza_categoria.csv",
            mime="text/csv",
        )
    else:
        st.success("Nessuna transazione senza categoria, dopo l'esclusione dei trasferimenti interni.")
    st.markdown('<div class="fh-note"><b>Conversione EUR attiva.</b> L’app usa prima la colonna Importo in EUR; in alternativa applica i tassi configurati e i tassi storici della prima versione per JPY, ISK ed EGP. Le righe senza un tasso affidabile sono escluse e segnalate.</div>', unsafe_allow_html=True)

st.divider()
st.caption("Finance Hub 4.3 · aggiornamenti 5.2, 5.4, 5.5 e 5.7 · formato numerico 25,000.50 · dati elaborati localmente")
