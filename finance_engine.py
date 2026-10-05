from __future__ import annotations

import csv
import hashlib
import io
import math
import posixpath
import re
import unicodedata
import zipfile
from collections import defaultdict
from datetime import date, datetime, timedelta
from typing import Any, Iterable
from xml.etree import ElementTree as ET

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

# Tassi usati dal workbook storico della prima versione (EUR per unità).
# I valori configurati nel foglio "Saldi Iniziali" hanno sempre precedenza.
HISTORICAL_FX_EUR_PER_UNIT = {
    "EUR": 1.0,
    "JPY": 1.0 / 181.0,
    "ISK": 1.0 / 143.2,
    "EGP": 1.0 / 62.92,
}

ANALYSIS_BUCKETS = [
    "Entrate",
    "Spese fisse inevitabili",
    "Spese fisse evitabili",
    "Spese necessarie",
    "Spese variabili non necessarie",
    "Matrimonio",
    "Casa nuova",
    "Luna di miele (Giappone)",
    "Altri viaggi",
]

UNCATEGORIZED_NAMES = {
    "",
    "non classificata",
    "non classificato",
    "senza categoria",
    "uncategorized",
    "uncategorised",
    "unknown",
}


def clean_text(value: Any) -> str:
    """Ripara il mojibake più comune degli export MoneyWiz/Excel."""
    text = "" if value is None else str(value)
    if "Ã" in text or "Â" in text:
        try:
            text = text.encode("latin-1").decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    return text.replace("\u00a0", " ").strip()


def norm(value: Any) -> str:
    text = clean_text(value)
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def is_yes(value: Any) -> bool:
    return norm(value) in {"si", "yes", "y", "true", "1", "attivo", "attiva", "confermato", "confermata"}


def parse_number(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value) if math.isfinite(float(value)) else None
    raw = str(value).strip()
    if not raw:
        return None
    negative = raw.startswith("(") and raw.endswith(")")
    raw = raw.replace("\u00a0", " ").replace("−", "-")
    raw = re.sub(r"[^0-9,.'+\-]", "", raw).replace("'", "")
    if not raw or raw in {"-", "+"}:
        return None
    if "," in raw and "." in raw:
        if raw.rfind(",") > raw.rfind("."):
            raw = raw.replace(".", "").replace(",", ".")
        else:
            raw = raw.replace(",", "")
    elif "," in raw:
        tail = raw.split(",")[-1]
        raw = raw.replace(",", ".") if len(tail) in {1, 2} else raw.replace(",", "")
    elif raw.count(".") > 1:
        parts = raw.split(".")
        raw = "".join(parts[:-1]) + "." + parts[-1] if len(parts[-1]) in {1, 2} else "".join(parts)
    try:
        number = float(raw)
        return -abs(number) if negative else number
    except ValueError:
        return None


def parse_date(value: Any) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and 1 <= float(value) <= 100000:
        return (datetime(1899, 12, 30) + timedelta(days=float(value))).date()
    text = str(value).strip()
    if not text:
        return None
    text = re.sub(r"\s+.*$", "", text)
    for fmt in (
        "%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y",
        "%m/%d/%Y", "%Y%m%d", "%d/%m/%y", "%m/%d/%y",
    ):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(str(value).replace("Z", "+00:00")).date()
    except Exception:
        return None


def excel_date(serial: float) -> date:
    return (datetime(1899, 12, 30) + timedelta(days=serial)).date()


def _column_index(cell_ref: str) -> int:
    letters = re.match(r"([A-Z]+)", cell_ref.upper())
    if not letters:
        return 0
    index = 0
    for char in letters.group(1):
        index = index * 26 + ord(char) - 64
    return index - 1


def _xlsx_tables(data: bytes) -> dict[str, list[list[Any]]]:
    tables: dict[str, list[list[Any]]] = {}
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        names = set(archive.namelist())
        shared: list[str] = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall(f"{{{NS_MAIN}}}si"):
                shared.append("".join(t.text or "" for t in item.iter(f"{{{NS_MAIN}}}t")))

        date_styles: set[int] = set()
        if "xl/styles.xml" in names:
            root = ET.fromstring(archive.read("xl/styles.xml"))
            custom_formats: dict[int, str] = {}
            num_fmts = root.find(f"{{{NS_MAIN}}}numFmts")
            if num_fmts is not None:
                for item in num_fmts:
                    custom_formats[int(item.attrib.get("numFmtId", 0))] = item.attrib.get("formatCode", "")
            built_in_dates = set(range(14, 23)) | {27, 30, 36, 45, 46, 47, 50, 57}
            cell_xfs = root.find(f"{{{NS_MAIN}}}cellXfs")
            if cell_xfs is not None:
                for idx, xf in enumerate(cell_xfs):
                    num_id = int(xf.attrib.get("numFmtId", 0))
                    fmt = custom_formats.get(num_id, "").lower()
                    fmt_no_literals = re.sub(r'"[^"]*"|\\.', "", fmt)
                    if num_id in built_in_dates or re.search(r"(^|[^a-z])[dmyhs]+([^a-z]|$)", fmt_no_literals):
                        date_styles.add(idx)

        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rels = {r.attrib["Id"]: r.attrib["Target"] for r in rels_root.findall(f"{{{NS_PKG_REL}}}Relationship")}
        for sheet in workbook.find(f"{{{NS_MAIN}}}sheets") or []:
            name = sheet.attrib.get("name", "Sheet")
            rel_id = sheet.attrib.get(f"{{{NS_REL}}}id")
            target = rels.get(rel_id, "")
            if target.startswith("/"):
                path = target.lstrip("/")
            else:
                path = posixpath.normpath(posixpath.join("xl", target))
            if path not in names:
                continue
            root = ET.fromstring(archive.read(path))
            rows: list[list[Any]] = []
            sheet_data = root.find(f"{{{NS_MAIN}}}sheetData")
            if sheet_data is None:
                tables[name] = rows
                continue
            max_col = 0
            temp_rows: list[dict[int, Any]] = []
            for row in sheet_data.findall(f"{{{NS_MAIN}}}row"):
                values: dict[int, Any] = {}
                for cell in row.findall(f"{{{NS_MAIN}}}c"):
                    col = _column_index(cell.attrib.get("r", "A1"))
                    max_col = max(max_col, col + 1)
                    cell_type = cell.attrib.get("t")
                    style = int(cell.attrib.get("s", 0))
                    value_el = cell.find(f"{{{NS_MAIN}}}v")
                    if cell_type == "inlineStr":
                        value = "".join(t.text or "" for t in cell.iter(f"{{{NS_MAIN}}}t"))
                    elif value_el is None:
                        value = None
                    else:
                        raw = value_el.text or ""
                        if cell_type == "s":
                            try:
                                value = shared[int(raw)]
                            except Exception:
                                value = raw
                        elif cell_type == "b":
                            value = raw == "1"
                        elif cell_type in {"str", "e"}:
                            value = raw
                        else:
                            try:
                                num = float(raw)
                                value = excel_date(num) if style in date_styles else (int(num) if num.is_integer() else num)
                            except ValueError:
                                value = raw
                    values[col] = value
                temp_rows.append(values)
            for item in temp_rows:
                rows.append([item.get(i) for i in range(max_col)])
            while rows and all(v in (None, "") for v in rows[-1]):
                rows.pop()
            tables[name] = rows
    return tables


def rows_to_dicts(rows: list[list[Any]]) -> list[dict[str, Any]]:
    if not rows:
        return []
    headers = [str(v).strip() if v not in (None, "") else f"Colonna {i + 1}" for i, v in enumerate(rows[0])]
    output: list[dict[str, Any]] = []
    for row in rows[1:]:
        if not any(v not in (None, "") for v in row):
            continue
        padded = row + [None] * (len(headers) - len(row))
        output.append({headers[i]: padded[i] for i in range(len(headers))})
    return output


def load_config(data: bytes) -> dict[str, Any]:
    grids = _xlsx_tables(data)
    config = {name: rows_to_dicts(rows) for name, rows in grids.items()}
    config["__grids__"] = grids
    return config


def _decode_csv(data: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-16", "cp1252", "latin-1"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            pass
    return data.decode("utf-8", errors="replace")


def _csv_rows(data: bytes) -> list[list[Any]]:
    text = _decode_csv(data)
    sample = text[:8192]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        delimiter = dialect.delimiter
    except csv.Error:
        delimiter = ";" if sample.count(";") > sample.count(",") else ","
    return [list(row) for row in csv.reader(io.StringIO(text), delimiter=delimiter)]


ALIASES = {
    "date": {"data", "date", "transaction date", "data transazione", "data operazione", "giorno"},
    "amount": {"importo", "amount", "value", "valore", "total", "original amount", "importo originale"},
    "amount_eur": {"importo in eur", "importo eur", "amount in eur", "amount eur", "eur amount", "valore eur"},
    "account": {"conto", "account", "account name", "nome conto", "conto moneywiz"},
    "category": {"categoria", "category", "categories", "categoria moneywiz", "category name"},
    "owner": {"owner", "proprietario", "persona", "titolare"},
    "type": {"tipo", "type", "transaction type", "tipo transazione", "income expense", "entrata uscita"},
    "payee": {"payee", "beneficiario", "merchant", "controparte", "esercente", "descrizione", "description", "name"},
    "currency": {"valuta", "currency", "currency code", "codice valuta"},
    "note": {"note", "notes", "memo", "commento", "comment"},
    "transfer": {"trasferimenti", "trasferimento", "transfers", "transfer account", "conto trasferimento"},
    "tags": {"tag", "tags", "etichette", "labels", "progetto", "project", "viaggio"},
}


def _header_map(headers: Iterable[Any]) -> dict[str, int]:
    normalized = [norm(h) for h in headers]
    found: dict[str, int] = {}
    for field, aliases in ALIASES.items():
        for idx, value in enumerate(normalized):
            if value in aliases:
                found[field] = idx
                break
    return found


def _find_table(rows: list[list[Any]]) -> tuple[list[list[Any]], dict[str, int]]:
    best_rows: list[list[Any]] = []
    best_map: dict[str, int] = {}
    for idx, row in enumerate(rows[:30]):
        mapping = _header_map(row)
        has_amount = "amount" in mapping or "amount_eur" in mapping
        best_has_amount = "amount" in best_map or "amount_eur" in best_map
        score = len(mapping) + (3 if "date" in mapping and has_amount else 0)
        best_score = len(best_map) + (3 if "date" in best_map and best_has_amount else 0)
        if score > best_score:
            best_rows, best_map = rows[idx:], mapping
    return best_rows, best_map


def _table_score(sheet_name: str, rows: list[list[Any]], mapping: dict[str, int]) -> int:
    if not rows or "date" not in mapping or not ({"amount", "amount_eur"} & mapping.keys()):
        return -10_000
    sheet = norm(sheet_name)
    score = len(mapping) * 2
    score += 20 if "amount_eur" in mapping else 0
    score += 6 if "category" in mapping else 0
    score += 4 if "account" in mapping else 0
    if sheet in {"storico pulito", "transazioni", "transactions", "movimenti"}:
        score += 100
    elif "storico pulito" in sheet:
        score += 80
    elif sheet in {"estratto", "export", "moneywiz"}:
        score += 15
    if any(word in sheet for word in ("pivot", "budget", "previsione", "patrimonio", "riepilogo")):
        score -= 100
    return score


def _configured_fx(config: dict[str, Any]) -> tuple[dict[str, float], dict[tuple[str, str], float]]:
    by_currency = dict(HISTORICAL_FX_EUR_PER_UNIT)
    by_account: dict[tuple[str, str], float] = {}
    for row in config.get("Saldi Iniziali", []):
        currency = clean_text(row.get("Valuta") or "EUR").upper()
        account = norm(row.get("Conto"))
        rate = parse_number(row.get("Tasso EUR per unità"))
        if rate is None or rate <= 0:
            continue
        by_currency[currency] = rate
        if account:
            by_account[(account, currency)] = rate
    return by_currency, by_account


def _project_from_tags(tags: Any, default: Any = None) -> str:
    excluded = ("needs", "wants", "spese necessarie", "spese fisse", "spese variabili")
    for part in re.split(r"[;|]", clean_text(tags)):
        candidate = part.strip(" ,-\t")
        if candidate and not any(token in norm(candidate) for token in excluded):
            return candidate
    fallback = clean_text(default)
    return fallback or "Non assegnato"


def _value(row: list[Any], mapping: dict[str, int], field: str) -> Any:
    idx = mapping.get(field)
    return row[idx] if idx is not None and idx < len(row) else None


def normalize_transactions(files: list[tuple[str, bytes]], config: dict[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, str]]]:
    account_owner = {norm(r.get("Conto")): r.get("Owner") for r in config.get("Conti", []) if r.get("Conto")}
    mappings = {norm(r.get("Categoria MoneyWiz")): r for r in config.get("Mappatura Categorie", []) if r.get("Categoria MoneyWiz")}
    currency_rates, account_rates = _configured_fx(config)
    transactions: list[dict[str, Any]] = []
    issues: list[dict[str, str]] = []
    seen: set[str] = set()

    for filename, data in files:
        try:
            if filename.lower().endswith(".xlsx"):
                sheets = _xlsx_tables(data)
                candidates = [(f"{filename} · {sheet}", rows) for sheet, rows in sheets.items()]
            else:
                candidates = [(filename, _csv_rows(data))]
        except Exception as exc:
            issues.append({"Livello": "Errore", "Ambito": filename, "Dettaglio": f"File non leggibile: {exc}"})
            continue

        ranked = []
        for source, raw_rows in candidates:
            rows, mapping = _find_table(raw_rows)
            ranked.append((_table_score(source, rows, mapping), source, rows, mapping))
        ranked.sort(key=lambda item: item[0], reverse=True)
        if not ranked or ranked[0][0] < 0:
            issues.append({"Livello": "Errore", "Ambito": filename, "Dettaglio": "Nessun foglio/tabella con colonne Data e Importo riconosciuto."})
            continue

        _, source, rows, mapping = ranked[0]
        converted = 0
        missing_fx: defaultdict[str, int] = defaultdict(int)
        occurrences: defaultdict[str, int] = defaultdict(int)
        for line_no, row in enumerate(rows[1:], start=2):
            if not any(v not in (None, "") for v in row):
                continue
            tx_date = parse_date(_value(row, mapping, "date"))
            raw_amount = parse_number(_value(row, mapping, "amount"))
            eur_amount = parse_number(_value(row, mapping, "amount_eur"))
            if tx_date is None or (raw_amount is None and eur_amount is None):
                issues.append({"Livello": "Avviso", "Ambito": source, "Dettaglio": f"Riga {line_no} esclusa: data o importo non valido."})
                continue
            account = clean_text(_value(row, mapping, "account")) or "Non specificato"
            currency = clean_text(_value(row, mapping, "currency") or "EUR").upper()
            if eur_amount is not None:
                amount = eur_amount
                rate = eur_amount / raw_amount if raw_amount not in (None, 0) else (1.0 if currency == "EUR" else None)
                if currency != "EUR":
                    converted += 1
            elif currency == "EUR":
                amount = float(raw_amount)
                rate = 1.0
            else:
                rate = account_rates.get((norm(account), currency), currency_rates.get(currency))
                if rate is None:
                    missing_fx[currency] += 1
                    continue
                amount = float(raw_amount) * rate
                converted += 1
            tx_type = str(_value(row, mapping, "type") or "")
            type_norm = norm(tx_type)
            if any(word in type_norm for word in ("expense", "spesa", "uscita", "debit", "pagamento")):
                amount = -abs(amount)
            elif any(word in type_norm for word in ("income", "entrata", "credito", "salary", "stipendio")):
                amount = abs(amount)
            category = clean_text(_value(row, mapping, "category")) or "Non classificata"
            owner = clean_text(_value(row, mapping, "owner") or account_owner.get(norm(account)) or "Condiviso")
            payee = clean_text(_value(row, mapping, "payee"))
            note = clean_text(_value(row, mapping, "note"))
            tags = clean_text(_value(row, mapping, "tags"))
            transfer_target = clean_text(_value(row, mapping, "transfer"))
            category_info = mappings.get(norm(category), {})
            macro = clean_text(category_info.get("Macro-categoria") or category.split(" > ")[0] or "Non classificata")
            expense_class = clean_text(category_info.get("Classe di spesa") or "Ricorrente variabile")
            ordinary = clean_text(category_info.get("Ordinaria/Straordinaria") or ("Straordinaria" if norm(expense_class) == "straordinaria" else "Ordinaria"))
            is_transfer = bool(transfer_target) or "trasferiment" in norm(category) or "transfer" in type_norm
            project = _project_from_tags(tags, category_info.get("Progetto default"))
            fingerprint = hashlib.sha1(f"{tx_date}|{amount:.6f}|{account}|{category}|{payee}|{note}".encode("utf-8")).hexdigest()
            occurrences[fingerprint] += 1
            dedupe_key = f"{fingerprint}:{occurrences[fingerprint]}"
            if dedupe_key in seen:
                continue
            seen.add(dedupe_key)
            transactions.append({
                "Data": tx_date,
                "Importo": amount,
                "Conto": account,
                "Categoria": category,
                "Macro-categoria": macro,
                "Owner": owner,
                "Classe di spesa": expense_class,
                "NEED/WANT": clean_text(category_info.get("NEED/WANT")),
                "Fissa/Variabile": clean_text(category_info.get("Fissa/Variabile")),
                "Ordinaria/Straordinaria": ordinary,
                "Payee": payee,
                "Valuta": currency,
                "Importo originale": raw_amount if raw_amount is not None else amount,
                "Tasso EUR": rate,
                "Note": note,
                "Tag": tags,
                "Progetto": project,
                "Trasferimento": is_transfer,
                "Fonte": source,
            })
        if converted:
            issues.append({"Livello": "OK", "Ambito": "Conversione EUR", "Dettaglio": f"{converted} transazioni non-EUR convertite usando la colonna EUR o i tassi configurati/storici."})
        for currency, count in sorted(missing_fx.items()):
            issues.append({"Livello": "Avviso", "Ambito": "Conversione EUR", "Dettaglio": f"{count} transazioni {currency} escluse: tasso EUR non disponibile."})

    transactions.sort(key=lambda item: (item["Data"], item["Conto"], item["Categoria"]))
    return transactions, issues


def is_uncategorized(row: dict[str, Any]) -> bool:
    """Identifica una transazione priva di una categoria utile."""
    return norm(row.get("Categoria")) in UNCATEGORIZED_NAMES


def uncategorized_transactions(transactions: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    """Restituisce le transazioni senza categoria, escludendo i trasferimenti interni."""
    return sorted(
        [row for row in transactions if not row.get("Trasferimento") and is_uncategorized(row)],
        key=lambda item: (item.get("Data") or date.min, str(item.get("Conto") or "")),
        reverse=True,
    )


def analysis_bucket(row: dict[str, Any]) -> str:
    """Assegna una riga a una delle sezioni della vista Storico vs Previsione."""
    amount = parse_number(row.get("Importo"))
    if amount is not None and amount > 0:
        return "Entrate"

    category = norm(row.get("Categoria") or row.get("Categoria MoneyWiz"))
    macro = norm(row.get("Macro-categoria"))
    project = norm(row.get("Progetto") or row.get("Progetto default"))
    tags = norm(row.get("Tag"))
    special_marker = " ".join((category, macro, project, tags))

    # Le linee straordinarie prevalgono sulla classificazione NEED/WANT.
    if "matrimonio" in special_marker:
        return "Matrimonio"
    if any(token in special_marker for token in ("casa nuova", "acquisto casa", "acquisto mobili e arredamenti")):
        return "Casa nuova"
    if any(token in special_marker for token in ("giappone", "luna di miele")):
        return "Luna di miele (Giappone)"
    if macro == "viaggi" or category == "viaggi" or category.startswith("viaggi "):
        return "Altri viaggi"

    # Se presenti, i tag espliciti hanno precedenza sui metadati della mappatura.
    if "spese fisse inevitabili" in tags:
        return "Spese fisse inevitabili"
    if "spese fisse evitabili" in tags:
        return "Spese fisse evitabili"
    if "spese variabili non necessarie" in tags:
        return "Spese variabili non necessarie"
    if "spese necessarie" in tags:
        return "Spese necessarie"

    need_want = norm(row.get("NEED/WANT"))
    fixed_variable = norm(row.get("Fissa/Variabile"))
    expense_class = norm(row.get("Classe di spesa"))
    fixed = fixed_variable == "fissa" or expense_class == "fissa"
    necessary = need_want in {"need", "needs", "necessaria", "necessarie", "necessario", "necessari"}
    if fixed:
        return "Spese fisse inevitabili" if necessary else "Spese fisse evitabili"
    return "Spese necessarie" if necessary else "Spese variabili non necessarie"


def frequency_per_year(frequency: Any) -> float:
    value = norm(frequency)
    return {
        "mensile": 12.0,
        "bimestrale": 6.0,
        "trimestrale": 4.0,
        "semestrale": 2.0,
        "annuale": 1.0,
        "una tantum": 1.0,
        "settimanale": 52.0,
    }.get(value, 12.0)


def income_monthly_amounts(row: dict[str, Any], year: int) -> list[float]:
    """Distribuisce una voce di entrata sui mesi attivi dell'anno."""
    values = [0.0] * 12
    if norm(row.get("Stato")) in {"inattiva", "inattivo", "non attiva", "non attivo"}:
        return values
    amount = max(0.0, parse_number(row.get("Importo")) or 0.0)
    start = parse_date(row.get("Data inizio"))
    end = parse_date(row.get("Data fine"))
    active_months = []
    for month in range(1, 13):
        month_start = date(year, month, 1)
        month_end = date(year + (month == 12), 1 if month == 12 else month + 1, 1) - timedelta(days=1)
        if (not start or month_end >= start) and (not end or month_start <= end):
            active_months.append(month)
    if not active_months or amount <= 0:
        return values

    frequency = norm(row.get("Frequenza"))
    step = {"mensile": 1, "bimestrale": 2, "trimestrale": 3, "semestrale": 6}.get(frequency)
    if step:
        first = start.month if start and start.year == year else active_months[0]
        for month in active_months:
            if (month - first) % step == 0:
                values[month - 1] = amount
    elif frequency in {"annuale", "una tantum"}:
        month = start.month if start and start.year == year else active_months[0]
        if month in active_months:
            values[month - 1] = amount
    elif frequency == "settimanale":
        monthly_equivalent = amount * 52.0 / 12.0
        for month in active_months:
            values[month - 1] = monthly_equivalent
    else:
        for month in active_months:
            values[month - 1] = amount
    return values


def monthly_actual_by_bucket(transactions: Iterable[dict[str, Any]], year: int) -> dict[str, list[float]]:
    """Aggrega gli actual mensili nelle sezioni di analisi, con spese espresse positive."""
    result = {bucket: [0.0] * 12 for bucket in ANALYSIS_BUCKETS}
    for row in transactions:
        tx_date = row.get("Data")
        if not isinstance(tx_date, date) or tx_date.year != year or row.get("Trasferimento"):
            continue
        amount = parse_number(row.get("Importo")) or 0.0
        bucket = analysis_bucket(row)
        result[bucket][tx_date.month - 1] += amount if bucket == "Entrate" else abs(min(0.0, amount))
    return result


def monthly_plan_by_bucket(config: dict[str, Any], year: int) -> dict[str, list[float]]:
    """Aggrega entrate pianificate e budget mensili nelle stesse sezioni degli actual."""
    result = {bucket: [0.0] * 12 for bucket in ANALYSIS_BUCKETS}
    mappings = {
        norm(row.get("Categoria MoneyWiz")): row
        for row in config.get("Mappatura Categorie", [])
        if row.get("Categoria MoneyWiz")
    }
    for row in config.get("Entrate", []):
        monthly = income_monthly_amounts(row, year)
        result["Entrate"] = [a + b for a, b in zip(result["Entrate"], monthly)]

    for row in config.get("Budget", []):
        monthly = budget_monthly_amounts(row, year)
        if not any(monthly):
            continue
        category = clean_text(row.get("Categoria MoneyWiz"))
        mapped = mappings.get(norm(category), {})
        classifier = {
            "Importo": -1.0,
            "Categoria": category,
            "Macro-categoria": clean_text(mapped.get("Macro-categoria") or category.split(" > ")[0]),
            "Progetto": clean_text(row.get("Progetto") or mapped.get("Progetto default")),
            "Progetto default": clean_text(mapped.get("Progetto default")),
            "Tag": clean_text(row.get("Tag")),
            "NEED/WANT": clean_text(mapped.get("NEED/WANT")),
            "Fissa/Variabile": clean_text(mapped.get("Fissa/Variabile")),
            "Classe di spesa": clean_text(row.get("Classe di spesa") or mapped.get("Classe di spesa")),
        }
        bucket = analysis_bucket(classifier)
        result[bucket] = [a + b for a, b in zip(result[bucket], monthly)]
    return result


def historical_snapshots(start: date, end: date) -> list[date]:
    """Genera chiusure mensili complete e, se necessario, l'ultima data parziale."""
    if end < start:
        return []
    snapshots: list[date] = []
    cursor = date(start.year, start.month, 1)
    while cursor <= end:
        month_end = date(cursor.year + (cursor.month == 12), 1 if cursor.month == 12 else cursor.month + 1, 1) - timedelta(days=1)
        if month_end <= end:
            snapshots.append(month_end)
        cursor = date(cursor.year + (cursor.month == 12), 1 if cursor.month == 12 else cursor.month + 1, 1)
    if not snapshots or snapshots[-1] != end:
        snapshots.append(end)
    return snapshots


def annual_income(config: dict[str, Any], year: int) -> float:
    total = 0.0
    for row in config.get("Entrate", []):
        if norm(row.get("Stato")) == "inattiva":
            continue
        amount = parse_number(row.get("Importo")) or 0.0
        start = parse_date(row.get("Data inizio"))
        end = parse_date(row.get("Data fine"))
        if start and start.year > year or end and end.year < year:
            continue
        total += max(0.0, amount) * frequency_per_year(row.get("Frequenza"))
    return total


def annual_budget_amount(row: dict[str, Any]) -> float:
    amount = max(0.0, parse_number(row.get("Importo budget (€)")) or 0.0)
    kind = norm(row.get("Tipo budget"))
    if kind == "mensile":
        return amount * 12.0
    if kind in {"annuale distribuito", "stagionale", "una tantum"}:
        return amount
    if kind == "nessun budget":
        return 0.0
    return amount


def budget_monthly_amounts(row: dict[str, Any], year: int) -> list[float]:
    values = [0.0] * 12
    if int(parse_number(row.get("Anno")) or year) != year or norm(row.get("Stato")) not in {"confermato", "confermata"}:
        return values
    amount = max(0.0, parse_number(row.get("Importo budget (€)")) or 0.0)
    kind = norm(row.get("Tipo budget"))
    if kind == "mensile":
        return [amount] * 12
    if kind == "annuale distribuito":
        return [amount / 12.0] * 12
    if kind == "stagionale":
        raw_months = row.get("Mesi stagionali (1-12)") or row.get("Mese applicazione")
        months = [int(x) for x in re.findall(r"\d+", str(raw_months or "")) if 1 <= int(x) <= 12]
        if not months:
            months = [12]
        share = amount / len(set(months))
        for month in set(months):
            values[month - 1] = share
        return values
    if kind == "una tantum":
        when = parse_date(row.get("Data una tantum"))
        month = when.month if when and when.year == year else 12
        values[month - 1] = amount
    return values


def annual_budget(config: dict[str, Any], year: int, classes: set[str] | None = None) -> float:
    total = 0.0
    for row in config.get("Budget", []):
        if int(parse_number(row.get("Anno")) or year) != year or norm(row.get("Stato")) not in {"confermato", "confermata"}:
            continue
        if classes and norm(row.get("Classe di spesa")) not in {norm(v) for v in classes}:
            continue
        total += annual_budget_amount(row)
    return total


def monthly_recurring(config: dict[str, Any], expense_class: str | None = None) -> float:
    total = 0.0
    for row in config.get("Ricorrenze", []):
        if norm(row.get("Stato")) in {"inattiva", "inattivo", "non attiva", "non attivo"}:
            continue
        if expense_class and norm(row.get("Classe di spesa")) != norm(expense_class):
            continue
        amount = max(0.0, parse_number(row.get("Importo")) or 0.0)
        total += amount * frequency_per_year(row.get("Frequenza")) / 12.0
    return total


def account_balance(config: dict[str, Any], transactions: list[dict[str, Any]], snapshot: date, accounting_only: bool) -> float:
    accounts = {norm(r.get("Conto")): r for r in config.get("Conti", []) if r.get("Conto")}
    currency_rates, account_rates = _configured_fx(config)
    openings: dict[str, tuple[date | None, float, str, str]] = {}
    for row in config.get("Saldi Iniziali", []):
        if norm(row.get("Stato")) != "confermato":
            continue
        account_key = norm(row.get("Conto"))
        currency = clean_text(row.get("Valuta") or "EUR").upper()
        amount_eur = parse_number(row.get("Saldo iniziale (€)"))
        if amount_eur is None:
            raw = parse_number(row.get("Saldo iniziale (valuta conto)")) or 0.0
            rate = account_rates.get((account_key, currency), currency_rates.get(currency))
            amount_eur = raw * rate if rate is not None else 0.0
        openings[account_key] = (
            parse_date(row.get("Data saldo iniziale")), amount_eur,
            str(row.get("Incluso patrimonio contabile") or "No"), currency
        )
    total = 0.0
    for account_key, account_cfg in accounts.items():
        if accounting_only and not is_yes(account_cfg.get("Incluso patrimonio contabile")):
            continue
        opening_date, opening, included, _ = openings.get(account_key, (None, 0.0, "No", "EUR"))
        if opening_date and opening_date > snapshot:
            continue
        balance = opening
        for tx in transactions:
            if norm(tx["Conto"]) != account_key or tx["Data"] > snapshot or tx.get("Trasferimento"):
                continue
            if opening_date and tx["Data"] < opening_date:
                continue
            # normalize_transactions garantisce che Importo sia sempre espresso in EUR.
            balance += tx["Importo"]
        total += balance
    return total


def _months_between(start: date, end: date) -> int:
    return max(0, (end.year - start.year) * 12 + end.month - start.month)


def patrimony_components(config: dict[str, Any], snapshot: date, accounting: bool) -> tuple[float, float]:
    assets = debts = 0.0
    flag = "Incluso patrimonio contabile" if accounting else "Incluso patrimonio complessivo"
    for row in config.get("Patrimonio", []):
        if not is_yes(row.get(flag)) or norm(row.get("Stato")) in {"non attivo", "inattivo"}:
            continue
        current = max(0.0, parse_number(row.get("Valore corrente")) or 0.0)
        current_date = parse_date(row.get("Data valutazione"))
        initial = max(0.0, parse_number(row.get("Valore iniziale")) or current)
        initial_date = parse_date(row.get("Data iniziale"))
        rate = max(0.0, parse_number(row.get("Tasso annuo %")) or 0.0)
        name = norm(row.get("Nome"))
        cls = norm(row.get("Classe"))
        value = current
        if "finanziamento mercedes" in name and initial_date:
            payment = 0.0
            for rec in config.get("Ricorrenze", []):
                if "mercedes financial" in norm(rec.get("Descrizione / Payee")):
                    payment = max(0.0, parse_number(rec.get("Importo")) or 0.0)
                    break
            value = max(0.0, initial - _months_between(initial_date, snapshot) * payment)
        elif current_date and snapshot >= current_date:
            years = (snapshot - current_date).days / 365.25
            value = current * ((1.0 - min(rate, 0.99)) ** years) if rate > 0 else current
        elif initial_date and snapshot >= initial_date:
            years = (snapshot - initial_date).days / 365.25
            value = initial * ((1.0 - min(rate, 0.99)) ** years) if rate > 0 else initial
        if cls == "debito":
            debts += value
        else:
            assets += value
    return assets, debts


def net_worth(config: dict[str, Any], transactions: list[dict[str, Any]], snapshot: date, accounting: bool) -> float:
    cash = account_balance(config, transactions, snapshot, accounting_only=accounting)
    assets, debts = patrimony_components(config, snapshot, accounting)
    return cash + assets - debts


def patrimony_summary(config: dict[str, Any], transactions: list[dict[str, Any]], snapshot: date) -> dict[str, float]:
    """Restituisce le quattro metriche della vista Patrimonio della prima versione."""
    cash = account_balance(config, transactions, snapshot, accounting_only=False)
    financial_assets, _ = patrimony_components(config, snapshot, accounting=True)
    all_assets, all_debts = patrimony_components(config, snapshot, accounting=False)
    activities = cash + financial_assets
    physical = max(0.0, all_assets - financial_assets)
    return {
        "Attività finanziarie": activities,
        "Beni fisici": physical,
        "Debiti": all_debts,
        "Patrimonio Netto": activities + physical - all_debts,
    }


def month_ends(year: int, through_month: int = 12) -> list[date]:
    result: list[date] = []
    for month in range(1, through_month + 1):
        if month == 12:
            result.append(date(year, 12, 31))
        else:
            result.append(date(year, month + 1, 1) - timedelta(days=1))
    return result


def group_sum(rows: Iterable[dict[str, Any]], key: str, amount_key: str = "Importo", absolute: bool = False) -> dict[str, float]:
    grouped: defaultdict[str, float] = defaultdict(float)
    for row in rows:
        value = float(row.get(amount_key) or 0.0)
        grouped[str(row.get(key) or "Non classificata")] += abs(value) if absolute else value
    return dict(grouped)
