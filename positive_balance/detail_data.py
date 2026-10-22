"""Presentation data built from the same audited snapshots as the compact report.

The module does not read Excel, modify history, guess missing values, or copy
figures/comments from reference slides. All amounts are in currency units.

``build_presentation_data(history, period, config)`` returns:

* ``period``, ``previous_period`` (calendar predecessor), ``previous_available``,
  ``date`` (ISO quarter end), ``months`` (six ascending YYYY-MM strings).
* ``countries``: id/name/currency/profile/rate/display_currency, ``monthly``
  (amount, processed_amount, remaining and each category in local currency and
  with _eur suffix; count, processed_count, efficiency_amount/count), ``metrics``,
  ``previous_metrics`` (or None), ``stock_bins`` / ``stock_vintages`` rows
  (label/count/amount/amount_eur/pct_count/pct_amount/source), corresponding
  ``*_totals`` and ``*_reported_totals``, ``global_vintages`` (label, remaining_eur,
  identified_eur), ``categories``, ``comments``, ``issues``, ``show_stock_bins``.
* ``groups`` and ``global``: id/name/entity_ids/countries/monthly/metrics,
  previous_metrics, stock_rows, global_vintages, comments, scope_note. Each stock
  row contains id/name/current_identified_eur/current_remaining_eur,
  previous_identified_eur/previous_remaining_eur/change/vintages and M-5 ratios.

``metrics`` includes flow_eur, processed_eur, count, processed_count,
monthly_average (EUR), flow_local, monthly_average_local, average_balance_local,
average_balance_eur, remaining/identified (local), remaining_eur/identified_eur,
efficiency_amount/count (weighted semester), oldest_amount/count (M-5),
oldest_flow_eur/oldest_processed_eur/oldest_contracts/oldest_processed_contracts,
and stock_change. Group metrics' local fields are also EUR; no currencies are
summed before conversion. Previous-quarter comparisons use that quarter's FX.

KPI_3 is a remaining-stock distribution. KPI_2 identified stock adds unreachable
customers where this category applies; the two measures are never substituted.
"""
from __future__ import annotations

from copy import deepcopy
import math

from .model import complete_sum, parse_period, quarter_end, ratio, rolling_months


CATEGORIES = ("refund", "false_positive", "remaining", "unreachable",
              "authority_refund", "other_without_refund")
AMOUNT_FIELDS = ("amount", "processed_amount", "remaining", "refund",
                 "false_positive", "unreachable", "authority_refund",
                 "other_without_refund")


def previous_period(period: str) -> str:
    """The immediately preceding quarter, including the Q1/year boundary."""
    period = parse_period(period)
    year, quarter = int(period[:4]), int(period[-1])
    return f"{year}Q{quarter - 1}" if quarter > 1 else f"{year - 1}Q4"


def _subtract(a, b):
    return None if a is None or b is None else a - b


def _sum(values):
    values = list(values)
    return complete_sum(values) if values else None


def _change(current, prior):
    return ratio(_subtract(current, prior), prior)


def _fx(snapshot):
    rates = {row["currency"]: row["rate"] for row in snapshot.get("fx", [])}
    rates.setdefault("EUR", 1)
    for code, value in rates.items():
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or value <= 0:
            raise ValueError(f"Invalid presentation FX rate for {code}: {value!r}")
    return rates


def _comments(config, period, key):
    comments = config.get("comments", {}).get(period, {}).get(key, [])
    if isinstance(comments, str):
        return [comments]
    if not isinstance(comments, list) or any(not isinstance(item, str) for item in comments):
        raise ValueError(f"Presentation comments for {period}/{key} must be text or a list of text")
    return list(comments)


def _monthly(entity, rate, months):
    source = {row["month"]: row for row in entity.get("flows", [])}
    active = set(entity["processed_fields"])
    result = []
    for month in months:
        row = source.get(month, {})
        processed = _sum(row.get(field) for field in entity["processed_fields"])
        data = {"month": month, "amount": row.get("amount"),
                "count": row.get("count"), "processed_amount": processed,
                "processed_count": row.get("processed_count"),
                "remaining": _subtract(row.get("amount"), processed),
                "remaining_count": _subtract(row.get("count"), row.get("processed_count")),
                "remaining_source": row.get("remaining_amount_source"),
                "source": row.get("source", "")}
        for field in CATEGORIES:
            if field != "remaining":
                data[field] = row.get(field) if field in active else None
        for field in AMOUNT_FIELDS:
            data[field + "_eur"] = ratio(data[field], rate)
        data["efficiency_amount"] = ratio(processed, data["amount"])
        data["efficiency_count"] = ratio(data["processed_count"], data["count"])
        result.append(data)
    return result


def _distribution(entity, name, rate):
    rows = entity.get(name, [])
    reported = entity.get(name + "_totals", {})
    totals = {"label": "Total", "source": reported.get("source", "")}
    for field in ("count", "amount"):
        detail_sum = _sum(row.get(field) for row in rows)
        # Blank detail rows are unknown, not zero. A reported total is still a
        # usable denominator when the detailed distribution is incomplete.
        totals[field] = detail_sum if detail_sum is not None else reported.get(field)
        totals[field + "_basis"] = "detail_sum" if detail_sum is not None else "reported_total"
    totals["amount_eur"] = ratio(totals["amount"], rate)
    totals["pct_count"] = ratio(totals["count"], totals["count"])
    totals["pct_amount"] = ratio(totals["amount"], totals["amount"])
    label_key = "band" if name == "stock_bins" else "vintage"
    output = []
    for row in rows:
        output.append({"label": row[label_key], label_key: row[label_key],
                       "count": row.get("count"), "amount": row.get("amount"),
                       "amount_eur": ratio(row.get("amount"), rate),
                       "pct_count": ratio(row.get("count"), totals["count"]),
                       "pct_amount": ratio(row.get("amount"), totals["amount"]),
                       "source": row.get("source", "")})
    return output, totals, deepcopy(reported)


def _stock_values(entity, record):
    processed = _sum(record.get(key) for key in entity["processed_fields"])
    remaining = _subtract(record.get("total"), processed)
    unreachable = record.get("unreachable") if "unreachable" in entity["processed_fields"] else None
    identified = _sum([remaining, unreachable]) if "unreachable" in entity["processed_fields"] else remaining
    return remaining, identified


def _global_vintages(entity, rate, period):
    """Rebucket KPI_2 stock without attempting to split older-year buckets."""
    year = int(period[:4])
    labels = [f"<{year - 1}", str(year - 1), str(year)]
    grouped = {label: [] for label in labels}
    for row in entity.get("stock", []):
        vintage = str(row["vintage"]).replace(" ", "")
        if vintage.startswith("<"):
            if int(vintage[1:]) > year - 1:
                raise ValueError(f"{entity['id']}: cannot split stock bucket {vintage} into {labels}")
            label = labels[0]
        else:
            value = int(vintage)
            if value > year:
                raise ValueError(f"{entity['id']}: future stock vintage {vintage}")
            label = labels[0] if value < year - 1 else str(value)
        remaining, identified = _stock_values(entity, row)
        grouped[label].append({"remaining_eur": ratio(remaining, rate),
                               "identified_eur": ratio(identified, rate)})
    return [{"label": label,
             "remaining_eur": _sum(row["remaining_eur"] for row in grouped[label]),
             "identified_eur": _sum(row["identified_eur"] for row in grouped[label])}
            for label in labels]


def _metrics(entity, monthly, rate, period):
    stock = entity.get("stock", [])
    remaining = _sum(_stock_values(entity, row)[0] for row in stock)
    identified = _sum(_stock_values(entity, row)[1] for row in stock)
    total = _sum(row.get("total") for row in stock)
    processed = _subtract(total, remaining)
    unreachable = _sum(row.get("unreachable") for row in stock) if "unreachable" in entity["processed_fields"] else None
    flow = _sum(row["amount"] for row in monthly)
    treated = _sum(row["processed_amount"] for row in monthly)
    count = _sum(row["count"] for row in monthly)
    processed_count = _sum(row["processed_count"] for row in monthly)
    quarter_months = set(rolling_months(period, 3))
    quarter = [row for row in monthly if row["month"] in quarter_months]
    quarter_amount = _sum(row["amount"] for row in quarter)
    quarter_treated = _sum(row["processed_amount"] for row in quarter)
    oldest = monthly[0]
    _, dist_totals, _ = _distribution(entity, "stock_bins", rate)
    metrics = dict(total=total, processed=processed, remaining=remaining,
                   identified=identified, unreachable=unreachable,
                   remaining_eur=ratio(remaining, rate), identified_eur=ratio(identified, rate),
                   contracts=dist_totals["count"], flow_local=flow,
                   flow_eur=ratio(flow, rate), processed_eur=ratio(treated, rate),
                   count=count, processed_count=processed_count,
                   monthly_average=ratio(ratio(flow, rate), len(monthly)),
                   monthly_average_local=ratio(flow, len(monthly)),
                   average_balance_local=ratio(flow, count), average_balance_eur=ratio(ratio(flow, rate), count),
                   efficiency_amount=ratio(treated, flow), efficiency_count=ratio(processed_count, count),
                   oldest_amount=oldest["efficiency_amount"], oldest_count=oldest["efficiency_count"],
                   oldest_flow_eur=oldest["amount_eur"], oldest_processed_eur=oldest["processed_amount_eur"],
                   oldest_contracts=oldest["count"], oldest_processed_contracts=oldest["processed_count"],
                   quarter_flow_eur=ratio(quarter_amount, rate),
                   quarter_processed_eur=ratio(quarter_treated, rate),
                   quarter_efficiency=ratio(quarter_treated, quarter_amount), stock_change=None)
    metrics["cst"] = None if entity.get("requires_cst") else 0
    metrics["including_cst"] = None if entity.get("requires_cst") else identified
    metrics["including_cst_eur"] = ratio(metrics["including_cst"], rate)
    return metrics


def _country(entity, snapshot, period, config):
    rates = _fx(snapshot)
    if entity["currency"] not in rates:
        raise ValueError(f"Missing presentation FX rate for {entity['currency']}")
    rate = rates[entity["currency"]]
    months = sorted(rolling_months(period))
    monthly = _monthly(entity, rate, months)
    options = config.get("country_options", {}).get(entity["id"], {})
    requested_currency = options.get("currency", config.get("country_currency", "local"))
    if requested_currency not in ("local", "EUR"):
        raise ValueError("Presentation country currency must be 'local' or 'EUR'")
    result = {key: entity[key] for key in ("id", "name", "currency", "profile")}
    result.update(rate=rate, report_date=quarter_end(period).isoformat(),
                  display_currency=entity["currency"] if requested_currency == "local" else "EUR",
                  monthly=monthly, metrics=_metrics(entity, monthly, rate, period),
                  categories=[key for key in CATEGORIES if key == "remaining" or key in entity["processed_fields"]],
                  global_vintages=_global_vintages(entity, rate, period),
                  previous_metrics=None, comments=_comments(config, period, entity["id"]),
                  show_stock_bins=options.get("show_stock_bins", True),
                  requires_cst=entity.get("requires_cst", False),
                  issues=deepcopy([row for row in snapshot.get("issues", []) if row.get("entity") == entity["id"]]))
    for name in ("stock_bins", "stock_vintages"):
        result[name], result[name + "_totals"], result[name + "_reported_totals"] = _distribution(entity, name, rate)
    return result


def _aggregate_monthly(countries, months):
    output = []
    for index, month in enumerate(months):
        data = {"month": month}
        rows = [country["monthly"][index] for country in countries]
        for field in AMOUNT_FIELDS:
            # An inapplicable category is excluded, not interpreted as missing.
            applicable = [country["monthly"][index] for country in countries
                          if field in ("amount", "processed_amount", "remaining") or field in country["categories"]]
            data[field] = data[field + "_eur"] = _sum(row[field + "_eur"] for row in applicable)
        for field in ("count", "processed_count", "remaining_count"):
            data[field] = _sum(row[field] for row in rows)
        data["efficiency_amount"] = ratio(data["processed_amount"], data["amount"])
        data["efficiency_count"] = ratio(data["processed_count"], data["count"])
        output.append(data)
    return output


def _aggregate_metrics(countries, previous=False):
    metric_key = "previous_metrics" if previous else "metrics"
    rows = [country.get(metric_key) for country in countries]
    if not rows or any(row is None for row in rows):
        return None
    additive = ("remaining_eur", "identified_eur", "including_cst_eur", "flow_eur", "processed_eur",
                "count", "processed_count", "contracts", "monthly_average",
                "oldest_flow_eur", "oldest_processed_eur", "oldest_contracts", "oldest_processed_contracts",
                "quarter_flow_eur", "quarter_processed_eur")
    result = {key: _sum(row[key] for row in rows) for key in additive}
    result.update(remaining=result["remaining_eur"], identified=result["identified_eur"],
                  flow_local=result["flow_eur"], monthly_average_local=result["monthly_average"],
                  average_balance_local=ratio(result["flow_eur"], result["count"]),
                  average_balance_eur=ratio(result["flow_eur"], result["count"]),
                  efficiency_amount=ratio(result["processed_eur"], result["flow_eur"]),
                  efficiency_count=ratio(result["processed_count"], result["count"]),
                  oldest_amount=ratio(result["oldest_processed_eur"], result["oldest_flow_eur"]),
                  oldest_count=ratio(result["oldest_processed_contracts"], result["oldest_contracts"]),
                  quarter_efficiency=ratio(result["quarter_processed_eur"], result["quarter_flow_eur"]),
                  stock_change=None)
    return result


def _group(identifier, name, countries, months, config, period):
    current = _aggregate_metrics(countries)
    prior = _aggregate_metrics(countries, previous=True)
    if current is not None and prior is not None:
        current["stock_change"] = _change(current["identified_eur"], prior["identified_eur"])
    stock_rows = []
    for country in countries:
        metric, previous = country["metrics"], country["previous_metrics"]
        stock_rows.append({"id": country["id"], "name": country["name"],
                           "current_identified_eur": metric["identified_eur"],
                           "current_remaining_eur": metric["remaining_eur"],
                           "previous_identified_eur": previous["identified_eur"] if previous else None,
                           "previous_remaining_eur": previous["remaining_eur"] if previous else None,
                           "change": metric["stock_change"], "vintages": deepcopy(country["global_vintages"]),
                           "oldest_amount": metric["oldest_amount"], "oldest_count": metric["oldest_count"],
                           "previous_oldest_amount": previous["oldest_amount"] if previous else None,
                           "previous_oldest_count": previous["oldest_count"] if previous else None})
    vintage_totals = []
    year = int(period[:4])
    for index, label in enumerate((f"<{year - 1}", str(year - 1), str(year))):
        vintage_totals.append({"label": label,
                               "identified_eur": _sum(country["global_vintages"][index]["identified_eur"] for country in countries),
                               "remaining_eur": _sum(country["global_vintages"][index]["remaining_eur"] for country in countries)})
    return {"id": identifier, "name": name, "currency": "EUR", "entity_ids": [country["id"] for country in countries],
            "countries": countries, "monthly": _aggregate_monthly(countries, months), "stock_rows": stock_rows,
            "metrics": current, "previous_metrics": prior, "global_vintages": vintage_totals,
            "categories": [key for key in CATEGORIES if any(key in country["categories"] for country in countries)],
            "comments": _comments(config, period, identifier),
            "scope_note": f"Available source perimeter: {len(countries)} countries."
                + (" Nordics stock excludes the separate CST account." if any(c["requires_cst"] for c in countries) else "")}


def build_presentation_data(history: dict, period: str, config: dict | None = None) -> dict:
    """Create a deterministic, non-mutating presentation view of one snapshot."""
    config = config or {"schema_version": 1}
    if config.get("schema_version", 1) != 1:
        raise ValueError("Unsupported presentation configuration schema_version; expected 1")
    if config.get("global_currency", "EUR") != "EUR":
        raise ValueError("Global presentation currency must be EUR")
    period = parse_period(period)
    snapshots = history["snapshots"]
    if period not in snapshots:
        raise ValueError(f"No history snapshot for presentation period {period}")
    snapshot = snapshots[period]
    previous_key = previous_period(period)
    previous = snapshots.get(previous_key)
    previous_entities = {entity["id"]: entity for entity in previous["entities"]} if previous else {}
    countries = [_country(entity, snapshot, period, config) for entity in snapshot["entities"]]
    order = {identifier: index for index, identifier in enumerate(config.get("country_order", []))}
    countries.sort(key=lambda country: (order.get(country["id"], len(order)), country["name"]))
    for country in countries:
        if country["id"] in previous_entities:
            previous_country = _country(previous_entities[country["id"]], previous, previous_key, config)
            country["previous_metrics"] = previous_country["metrics"]
            country["metrics"]["stock_change"] = _change(country["metrics"]["identified_eur"],
                                                          country["previous_metrics"]["identified_eur"])
    months = sorted(rolling_months(period))
    specifications = config.get("groups", [
        {"id": "core", "name": "CORE countries", "entity_ids": [country["id"] for country in countries if country["profile"] != "nordics"]},
        {"id": "noncore", "name": "NON CORE countries", "entity_ids": [country["id"] for country in countries if country["profile"] == "nordics"]},
    ])
    assignments = {}
    group_ids = set()
    for spec in specifications:
        if spec["id"] in group_ids or spec["id"] == "GLOBAL":
            raise ValueError(f"Duplicate or reserved presentation group ID: {spec['id']}")
        group_ids.add(spec["id"])
        for eid in spec.get("entity_ids", []):
            if eid in assignments:
                raise ValueError(f"Presentation entity {eid} is in multiple groups")
            assignments[eid] = spec["id"]
    unassigned = [country["id"] for country in countries if country["id"] not in assignments]
    default_group = config.get("default_group", "core")
    if unassigned and default_group not in group_ids:
        raise ValueError(f"Unassigned presentation entities {unassigned}; configure a valid default_group")
    for eid in unassigned:
        assignments[eid] = default_group
    groups = [_group(spec["id"], spec["name"], [country for country in countries if assignments[country["id"]] == spec["id"]], months, config, period)
              for spec in specifications]
    return {"period": period, "previous_period": previous_key, "previous_available": previous is not None,
            "date": quarter_end(period).isoformat(), "months": months, "countries": countries,
            "groups": groups, "global": _group("GLOBAL", "GLOBAL countries", countries, months, config, period),
            "issues": deepcopy(snapshot.get("issues", [])), "excluded": deepcopy(snapshot.get("excluded", [])),
            "style": deepcopy(config.get("style", {})), "category_labels": deepcopy(config.get("category_labels", {}))}
