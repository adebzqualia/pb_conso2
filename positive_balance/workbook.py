"""Generate a self-contained Excel report, with formulas and Python-calculated caches.

Inputs are snapshots, never workbook links. Excel tables provide filters and stable
key-based summaries. A new quarterly run regenerates all table extents.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name as colname

from .model import complete_sum, ratio, rolling_months

NA = "n.d."
SHEETS = ("Synthese", "Evolution flux", "Consolidation", "Flux", "Stock",
          "Repartition stock", "Taux", "Controles", "Mode emploi")
CATEGORIES = ("refund", "unreachable", "false_positive", "authority_refund", "other_without_refund")


def subtract(a, b):
    return None if a is None or b is None else a - b


def divide(a, b):
    return ratio(a, b)


def entity_metrics(entity: dict, rate: float, period: str) -> dict:
    """Business definitions, independently usable by tests and cache generation."""
    stock = entity["stock"]
    flows = sorted(entity["flows"], key=lambda r: r["month"])
    processed = complete_sum(complete_sum(r[k] for k in entity["processed_fields"]) for r in stock)
    total = complete_sum(r["total"] for r in stock)
    remaining = subtract(total, processed)
    unreachable = complete_sum(r["unreachable"] for r in stock) if "unreachable" in entity["processed_fields"] else None
    identified = complete_sum([remaining, unreachable]) if "unreachable" in entity["processed_fields"] else remaining
    flow_amount = complete_sum(r["amount"] for r in flows)
    flow_processed = complete_sum(complete_sum(r[k] for k in entity["processed_fields"]) for r in flows)
    count = complete_sum(r["count"] for r in flows)
    processed_count = complete_sum(r["processed_count"] for r in flows)
    quarter = [r for r in flows if r["month"] in rolling_months(period, 3)]
    quarter_amount = complete_sum(r["amount"] for r in quarter)
    quarter_processed = complete_sum(complete_sum(r[k] for k in entity["processed_fields"]) for r in quarter)
    oldest = flows[0]
    oldest_processed = complete_sum(oldest[k] for k in entity["processed_fields"])
    return dict(total=total, processed=processed, remaining=remaining, unreachable=unreachable,
                identified=identified, cst=None if entity.get("requires_cst") else 0,
                including_cst=None if entity.get("requires_cst") else identified,
                remaining_eur=divide(remaining, rate), identified_eur=divide(identified, rate),
                including_cst_eur=None if entity.get("requires_cst") else divide(identified, rate),
                contracts=complete_sum(r["count"] for r in entity["stock_bins"]),
                flow_eur=divide(flow_amount, rate), processed_eur=divide(flow_processed, rate),
                count=count, processed_count=processed_count,
                efficiency_amount=ratio(flow_processed, flow_amount), efficiency_count=ratio(processed_count, count),
                monthly_average=divide(divide(flow_amount, rate), len(flows)),
                oldest_amount=ratio(oldest_processed, oldest["amount"]),
                oldest_count=ratio(oldest["processed_count"], oldest["count"]),
                quarter_flow_eur=divide(quarter_amount, rate),
                quarter_processed_eur=divide(quarter_processed, rate),
                quarter_efficiency=ratio(quarter_processed, quarter_amount))


class Report:
    def __init__(self, path, history, period, front_sheets=None):
        self.history, self.period = history, period
        self.snapshots = [history["snapshots"][p] for p in sorted(history["snapshots"])]
        self.periods = [s["period"] for s in self.snapshots]
        self.entities = {}
        for snap in self.snapshots:
            self.entities.update({e["id"]: e["name"] for e in snap["entities"]})
        self.book = xlsxwriter.Workbook(str(path), {"strings_to_formulas": False, "strings_to_urls": False})
        self.book.set_properties({"title": "Positive Balance", "subject": "Consolidation trimestrielle", "author": "Positive Balance"})
        self.book.set_calc_mode("auto")
        sheet_names = SHEETS if front_sheets is None else tuple(front_sheets) + SHEETS[2:]
        self.sheets = {name: self.book.add_worksheet(name) for name in sheet_names}
        base = {"font_name": "Arial", "font_size": 10, "valign": "vcenter"}
        self.formats = {}
        for name, extras in {
            "text": {}, "input": {"font_color": "#1848A0"},
            "number": {"num_format": '#,##0.00;(#,##0.00);"–"', "align": "right"},
            "integer": {"num_format": '#,##0;(#,##0);"–"', "align": "right"},
            "percent": {"num_format": "0.0%", "italic": True, "align": "right"},
            "rate": {"num_format": "0.00000000", "align": "right"},
            "date": {"num_format": "dd/mm/yyyy"},
            "month": {"num_format": "mmm-yy"},
            "title": {"font_size": 16, "bold": True, "font_color": "#16324F"},
            "section": {"bold": True, "font_color": "#16324F", "bottom": 1, "bottom_color": "#9DAAB6"},
            "header": {"bold": True, "bg_color": "#16324F", "font_color": "#FFFFFF", "text_wrap": True, "align": "center"},
            "note": {"font_color": "#566474", "italic": True},
            "selector": {"bg_color": "#FFF0BA", "font_color": "#1848A0", "bold": True},
            "warning": {"bg_color": "#FFF0E3", "font_color": "#993B15"},
            "missing": {"font_color": "#993B15", "align": "right"},
            "total": {"bold": True, "top": 1, "top_color": "#9DAAB6", "num_format": '#,##0.00;(#,##0.00);"–"', "align": "right"},
        }.items():
            self.formats[name] = self.book.add_format({**base, **extras})
        self.last = {}
        self.metrics = {}
        self.records = {}
        for name, sheet in self.sheets.items():
            sheet.hide_gridlines(2)
            sheet.set_default_row(19)
            sheet.set_column("A:AB", 18, self.formats["text"])
            sheet.set_landscape()
            sheet.set_paper(9)
            sheet.fit_to_pages(1, 0)
            sheet.set_margins(.3, .3, .4, .4)
            if name not in {"Synthese", "Mode emploi"}:
                sheet.freeze_panes(5, 2)
                sheet.repeat_rows(0, 4)
        for name in ("Synthese", "Evolution flux"):
            if name in self.sheets:
                self.sheets[name].set_tab_color("#16324F")
        self.sheets["Mode emploi"].set_tab_color("#A4AAB0")

    def write(self, sheet, row, col, value, fmt="text"):
        sh = self.sheets[sheet]
        if value is None:
            sh.write_string(row, col, NA, self.formats["missing"])
        elif isinstance(value, datetime):
            sh.write_datetime(row, col, value, self.formats[fmt])
        else:
            sh.write(row, col, value, self.formats[fmt])

    def formula(self, sheet, row, col, formula, value, fmt="number"):
        self.sheets[sheet].write_formula(row, col, formula, self.formats[fmt], NA if value is None else value)

    def start(self, sheet, title, headers, widths=None, note=None):
        sh = self.sheets[sheet]
        sh.write(1, 0, title, self.formats["title"])
        if note:
            sh.write(2, 0, note, self.formats["note"])
        sh.set_row(4, 42)
        for c, header in enumerate(headers):
            self.write(sheet, 4, c, header, "header")
        if widths:
            for c, width in enumerate(widths):
                sh.set_column(c, c, width, self.formats["text"])
        self.records[sheet] = []

    def table(self, sheet, headers, count, name):
        sh = self.sheets[sheet]
        if count:
            sh.add_table(4, 0, 4 + count, len(headers) - 1,
                         {"name": name, "columns": [{"header": h, "header_format": self.formats["header"]} for h in headers],
                          "style": "Table Style Medium 2"})
            sh.conditional_format(5, 0, 4 + count, len(headers) - 1,
                                  {"type": "text", "criteria": "containing", "value": NA, "format": self.formats["missing"]})
        self.last[sheet] = 5 + count

    def rng(self, sheet, col):
        return f"'{sheet}'!${col}$6:${col}${self.last[sheet]}"

    def sumkey(self, source, column, target_row, extra=None):
        """Keyed sums stay correct after table sorting. Unknowns cannot sum to zero."""
        amount, periods, ids = (self.rng(source, c) for c in (column, "A", "B"))
        terms = f'{periods},$A{target_row},{ids},$B{target_row}'
        if extra:
            for col, criterion in extra:
                terms += f',{self.rng(source, col)},{criterion}'
        return f'IF(OR(COUNTIFS({terms})=0,COUNTIFS({terms},{amount},"{NA}")>0),"{NA}",SUMIFS({amount},{terms}))'

    @staticmethod
    def arithmetic(refs, expression):
        values = ",".join(refs)
        return f'=IF(COUNT({values})={len(refs)},{expression},"{NA}")'

    def rates(self):
        headers = ["Trimestre", "Devise", "Date des taux", "Devise locale pour 1 EUR", "Type de taux", "Source"]
        self.start("Taux", "Taux de conversion", headers, [15, 12, 17, 25, 20, 64],
                   "Montant EUR = montant en devise locale / taux de clôture. Même convention que la synthèse fournie.")
        self.fxrefs, self.fxvalues = {}, {}
        row = 5
        for snap in self.snapshots:
            for rate in snap["fx"]:
                data = [snap["period"], rate["currency"], datetime.fromisoformat(rate["date"]), rate["rate"], "Clôture", rate["source"]]
                for col, val in enumerate(data):
                    self.write("Taux", row, col, val, "date" if col == 2 else "rate" if col == 3 else "text")
                self.fxrefs[(snap["period"], rate["currency"])] = f"'Taux'!D{row+1}"
                self.fxvalues[(snap["period"], rate["currency"])] = rate["rate"]
                row += 1
        self.table("Taux", headers, row - 5, "TauxChange")
        sh = self.sheets["Taux"]
        sh.write(4, 7, "Trimestres disponibles", self.formats["header"])
        sh.set_column(7, 7, 25)
        for index, period in enumerate(self.periods, 5):
            sh.write_string(index, 7, period)
        self.book.define_name("Periodes", f"='Taux'!$H$6:$H${5+len(self.periods)}")

    def rate_formula(self, row):
        return f'=SUMIFS({self.rng("Taux", "D")},{self.rng("Taux", "A")},A{row},{self.rng("Taux", "B")},D{row})'

    def stock(self):
        headers = ["Trimestre", "Code pays", "Pays", "Devise", "Millésime", "Stock initial local",
                   "Remboursé client local", "Injoignables local", "Faux positifs local", "Autorités local", "Autres sans remboursement local",
                   "Traité calculé local", "A traiter local", "Stock identifié local", "Traité déclaré local", "A traiter déclaré local",
                   "Taux local pour 1 EUR", "A traiter EUR", "Stock identifié EUR", "Source"]
        self.start("Stock", "Stock par millésime issu de KPI_2", headers,
                   [15, 12, 22, 12, 16] + [21] * 14 + [58],
                   "Catégories absentes du modèle pays : n.d. Les calculs utilisent uniquement les catégories applicables.")
        row = 5
        fieldcols = {key: 6+i for i, key in enumerate(CATEGORIES)}
        for snap in self.snapshots:
            for entity in snap["entities"]:
                rate = self.fxvalues[(snap["period"], entity["currency"])]
                for record in entity["stock"]:
                    r = row+1
                    data = [snap["period"], entity["id"], entity["name"], entity["currency"], record["vintage"], record["total"]] + [record[k] for k in CATEGORIES]
                    for col, value in enumerate(data):
                        self.write("Stock", row, col, value, "number" if col >= 5 else "text")
                    refs = [f"{colname(fieldcols[k])}{r}" for k in entity["processed_fields"]]
                    processed = complete_sum(record[k] for k in entity["processed_fields"])
                    remaining = subtract(record["total"], processed)
                    addback = "unreachable" in entity["processed_fields"]
                    identified = complete_sum([remaining, record["unreachable"]]) if addback else remaining
                    self.formula("Stock", row, 11, self.arithmetic(refs, f'SUM({",".join(refs)})'), processed)
                    self.formula("Stock", row, 12, self.arithmetic([f"F{r}", f"L{r}"], f"F{r}-L{r}"), remaining)
                    self.formula("Stock", row, 13, self.arithmetic([f"M{r}", f"H{r}"] if addback else [f"M{r}"], f"M{r}+H{r}" if addback else f"M{r}"), identified)
                    self.write("Stock", row, 14, record["processed_source"], "number")
                    self.write("Stock", row, 15, record["remaining_source"], "number")
                    self.formula("Stock", row, 16, self.rate_formula(r), rate, "rate")
                    self.formula("Stock", row, 17, self.arithmetic([f"M{r}", f"Q{r}"], f"M{r}/Q{r}"), divide(remaining, rate))
                    self.formula("Stock", row, 18, self.arithmetic([f"N{r}", f"Q{r}"], f"N{r}/Q{r}"), divide(identified, rate))
                    self.write("Stock", row, 19, record["source"])
                    row += 1
        self.table("Stock", headers, row-5, "StockDetail")

    def flows(self):
        headers = ["Trimestre", "Code pays", "Pays", "Devise", "Mois de cohorte", "Flux local", "Contrats entrants",
                   "Remboursé client local", "Injoignables local", "Faux positifs local", "Autorités local", "Autres sans remboursement local",
                   "Traité calculé local", "Contrats traités", "Reste local", "Contrats restants", "Taux local pour 1 EUR",
                   "Flux EUR", "Traité EUR", "Efficacité montant", "Efficacité nombre", "Traité déclaré local", "Source"]
        self.start("Flux", "Flux par cohorte mensuelle et date de reporting", headers,
                   [15, 12, 22, 12, 20] + [21]*17 + [58],
                   "Une même cohorte peut apparaître dans plusieurs trimestres de reporting. Filtrer le trimestre avant d'agréger.")
        row = 5
        fieldcols = {key: 7+i for i, key in enumerate(CATEGORIES)}
        for snap in self.snapshots:
            for entity in snap["entities"]:
                rate = self.fxvalues[(snap["period"], entity["currency"])]
                for record in entity["flows"]:
                    r = row+1
                    data = [snap["period"], entity["id"], entity["name"], entity["currency"], datetime.strptime(record["month"], "%Y-%m"), record["amount"], record["count"]] + [record[k] for k in CATEGORIES]
                    for col, value in enumerate(data):
                        self.write("Flux", row, col, value, "month" if col == 4 else "integer" if col == 6 else "number" if col >= 5 else "text")
                    refs = [f"{colname(fieldcols[k])}{r}" for k in entity["processed_fields"]]
                    processed = complete_sum(record[k] for k in entity["processed_fields"])
                    self.formula("Flux", row, 12, self.arithmetic(refs, f'SUM({",".join(refs)})'), processed)
                    self.write("Flux", row, 13, record["processed_count"], "integer")
                    self.formula("Flux", row, 14, self.arithmetic([f"F{r}", f"M{r}"], f"F{r}-M{r}"), subtract(record["amount"], processed))
                    self.formula("Flux", row, 15, self.arithmetic([f"G{r}", f"N{r}"], f"G{r}-N{r}"), subtract(record["count"], record["processed_count"]), "integer")
                    self.formula("Flux", row, 16, self.rate_formula(r), rate, "rate")
                    self.formula("Flux", row, 17, self.arithmetic([f"F{r}", f"Q{r}"], f"F{r}/Q{r}"), divide(record["amount"], rate))
                    self.formula("Flux", row, 18, self.arithmetic([f"M{r}", f"Q{r}"], f"M{r}/Q{r}"), divide(processed, rate))
                    self.formula("Flux", row, 19, f'=IF(OR(COUNT(F{r},M{r})<>2,F{r}=0),"{NA}",M{r}/F{r})', ratio(processed, record["amount"]), "percent")
                    self.formula("Flux", row, 20, f'=IF(OR(COUNT(G{r},N{r})<>2,G{r}=0),"{NA}",N{r}/G{r})', ratio(record["processed_count"], record["count"]), "percent")
                    self.write("Flux", row, 21, record["processed_amount"], "number")
                    self.write("Flux", row, 22, record["source"])
                    row += 1
        self.table("Flux", headers, row-5, "FluxDetail")

    def distribution(self):
        headers = ["Trimestre", "Code pays", "Pays", "Devise", "Répartition", "Catégorie source", "Contrats", "Montant local", "Taux local pour 1 EUR", "Montant EUR", "Source"]
        self.start("Repartition stock", "Répartition du stock à traiter issue de KPI_3", headers,
                   [15, 12, 22, 12, 18, 25, 18, 21, 24, 21, 58],
                   "Tranches et millésimes décrivent le même stock. Ne pas additionner les deux répartitions. Libellés de tranche conservés.")
        row = 5
        for snap in self.snapshots:
            for entity in snap["entities"]:
                rate = self.fxvalues[(snap["period"], entity["currency"])]
                for field, kind, label in (("stock_bins", "Tranche", "band"), ("stock_vintages", "Millésime", "vintage")):
                    for record in entity[field]:
                        r = row+1
                        data = [snap["period"], entity["id"], entity["name"], entity["currency"], kind, record[label], record["count"], record["amount"]]
                        for col, value in enumerate(data):
                            self.write("Repartition stock", row, col, value, "integer" if col == 6 else "number" if col == 7 else "text")
                        self.formula("Repartition stock", row, 8, self.rate_formula(r), rate, "rate")
                        self.formula("Repartition stock", row, 9, self.arithmetic([f"H{r}", f"I{r}"], f"H{r}/I{r}"), divide(record["amount"], rate))
                        self.write("Repartition stock", row, 10, record["source"])
                        row += 1
        self.table("Repartition stock", headers, row-5, "StockRepartition")

    def consolidation(self):
        headers = ["Trimestre", "Code pays", "Pays", "Devise", "Taux local pour 1 EUR", "Stock initial local", "Traité local", "A traiter local", "Injoignables local", "Stock identifié local", "CST local", "Stock avec CST local", "A traiter EUR", "Stock identifié EUR", "Stock avec CST EUR", "Contrats KPI3", "Flux 6 mois EUR", "Traité 6 mois EUR", "Contrats entrants 6 mois", "Contrats traités 6 mois", "Efficacité montant 6 mois", "Efficacité nombre 6 mois", "Moyenne mensuelle EUR", "Efficacité montant M-5", "Efficacité nombre M-5", "Flux trimestre EUR", "Traité trimestre EUR", "Efficacité montant trimestre"]
        self.start("Consolidation", "Indicateurs par trimestre de reporting et pays", headers,
                   [15, 12, 22, 12] + [23]*24,
                   "n.d. = non disponible ou non applicable. Le stock identifié exclut les CST Nordics non transmis.")
        row = 5
        for snap in self.snapshots:
            for entity in snap["entities"]:
                period, eid = snap["period"], entity["id"]
                rate = self.fxvalues[(period, entity["currency"])]
                m = entity_metrics(entity, rate, period)
                self.metrics[(period, eid)] = m
                r = row+1
                for col, val in enumerate([period, eid, entity["name"], entity["currency"]]):
                    self.write("Consolidation", row, col, val)
                self.formula("Consolidation", row, 4, self.rate_formula(r), rate, "rate")
                for col, source_col, key in [(5,"F","total"),(6,"L","processed"),(7,"M","remaining"),(9,"N","identified"),(12,"R","remaining_eur"),(13,"S","identified_eur")]:
                    self.formula("Consolidation", row, col, "="+self.sumkey("Stock", source_col, r), m[key])
                if "unreachable" in entity["processed_fields"]:
                    self.formula("Consolidation", row, 8, "="+self.sumkey("Stock", "H", r), m["unreachable"])
                else:
                    self.write("Consolidation", row, 8, None)
                self.write("Consolidation", row, 10, m["cst"], "number")
                self.formula("Consolidation", row, 11, self.arithmetic([f"J{r}",f"K{r}"],f"J{r}+K{r}"), m["including_cst"])
                self.formula("Consolidation", row, 14, self.arithmetic([f"L{r}",f"E{r}"],f"L{r}/E{r}"), m["including_cst_eur"])
                self.formula("Consolidation", row, 15, "="+self.sumkey("Repartition stock","G",r,[("E",'"Tranche"')]), m["contracts"], "integer")
                for col, source_col, key in [(16,"R","flow_eur"),(17,"S","processed_eur"),(18,"G","count"),(19,"N","processed_count")]:
                    self.formula("Consolidation", row, col, "="+self.sumkey("Flux",source_col,r), m[key], "integer" if col>=18 else "number")
                for col, num, den, key in [(20,"R","Q","efficiency_amount"),(21,"T","S","efficiency_count")]:
                    self.formula("Consolidation", row, col, f'=IF(OR(COUNT({num}{r},{den}{r})<>2,{den}{r}=0),"{NA}",{num}{r}/{den}{r})', m[key], "percent")
                n = len(entity["flows"])
                self.formula("Consolidation", row, 22, self.arithmetic([f"Q{r}"], f"Q{r}/{n}"), m["monthly_average"])
                oldest = datetime.strptime(min(x["month"] for x in entity["flows"]), "%Y-%m")
                date_formula = f"DATE({oldest.year},{oldest.month},1)"
                for col, source_col, key in [(23,"T","oldest_amount"),(24,"U","oldest_count")]:
                    self.formula("Consolidation", row, col, "="+self.sumkey("Flux",source_col,r,[("E",date_formula)]), m[key], "percent")
                qyear, qmonth = int(period[:4]), int(period[-1])*3-2
                criterion = f'">="&DATE({qyear},{qmonth},1)'
                for col, source_col, key in [(25,"R","quarter_flow_eur"),(26,"S","quarter_processed_eur")]:
                    self.formula("Consolidation", row, col, "="+self.sumkey("Flux",source_col,r,[("E",criterion)]), m[key])
                self.formula("Consolidation", row, 27, f'=IF(OR(COUNT(Z{r},AA{r})<>2,Z{r}=0),"{NA}",AA{r}/Z{r})', m["quarter_efficiency"], "percent")
                row += 1
        self.table("Consolidation", headers, row-5, "Indicateurs")

    def lookup(self, metric_col, period_expr, id_expr):
        amount = self.rng("Consolidation", metric_col)
        criteria = f'{self.rng("Consolidation","A")},{period_expr},{self.rng("Consolidation","B")},{id_expr}'
        return f'=IF(OR(COUNTIFS({criteria})<>1,COUNTIFS({criteria},{amount},"{NA}")>0),"{NA}",SUMIFS({amount},{criteria}))'

    def header(self, sheet, row, headers, first=1):
        self.sheets[sheet].set_row(row, 44)
        for i, value in enumerate(headers, first):
            self.write(sheet, row, i, value, "header")

    def summary(self):
        sh = self.sheets["Synthese"]
        sh.set_column("A:A", 3)
        sh.set_column("B:B", 24)
        sh.set_column("C:J", 20)
        sh.write(1,1,"Positive Balance",self.formats["title"])
        sh.write(2,1,"Périmètre configuré uniquement. Les pays absents de la livraison ne sont pas reconstitués.",self.formats["note"])
        sh.write(3,1,"Trimestre",self.formats["section"])
        sh.write_string(3,2,self.period,self.formats["selector"])
        sh.data_validation(3,2,3,2,{"validate":"list","source":"=Periodes","error_type":"stop","error_title":"Trimestre inconnu","error_message":"Choisir un trimestre disponible."})
        count = len(self.history["snapshots"][self.period]["entities"])
        self.formula("Synthese",3,4,f'=COUNTIFS({self.rng("Consolidation","A")},$C$4)',count,"integer")
        sh.write(3,5,"entités dans ce trimestre",self.formats["note"])
        sh.write(5,1,"Stock à la clôture",self.formats["section"])
        self.header("Synthese",6,["Pays","A traiter EUR","Injoignables EUR","Stock identifié EUR","CST local","Devise CST","Stock avec CST EUR","Contrats KPI3"])
        ids = list(self.entities)
        first = 7
        for index,eid in enumerate(ids,first):
            self.write("Synthese",index,1,self.entities[eid])
            m=self.metrics.get((self.period,eid),{})
            for col,metric,key,fmt in [(2,"M","remaining_eur","number"),(4,"N","identified_eur","number"),(5,"K","cst","number"),(7,"O","including_cst_eur","number"),(8,"P","contracts","integer")]:
                self.formula("Synthese",index,col,self.lookup(metric,"$C$4",f'"{eid}"'),m.get(key),fmt)
            # Convert unreachable through the same period-specific rate, without mixing currencies.
            unreach=self.lookup("I","$C$4",f'"{eid}"')[1:]
            rate=self.lookup("E","$C$4",f'"{eid}"')[1:]
            self.formula("Synthese",index,3,f'=IF(ISNUMBER({unreach}),({unreach})/({rate}),"{NA}")',divide(m.get("unreachable"),self.fxvalues.get((self.period,next((e["currency"] for e in self.history["snapshots"][self.period]["entities"] if e["id"]==eid),"")),1)))
            currency=next((e["currency"] for s in reversed(self.snapshots) for e in s["entities"] if e["id"]==eid),"")
            self.write("Synthese",index,6,currency)
        end=first+len(ids)
        self.write("Synthese",end,1,"Total du périmètre","section")
        for col,key in [(2,"remaining_eur"),(3,"unreachable_eur"),(4,"identified_eur"),(7,"including_cst_eur"),(8,"contracts")]:
            refs=f"{colname(col)}{first+1}:{colname(col)}{end}"
            vals=[]
            for eid in ids:
                m=self.metrics.get((self.period,eid),{})
                if key=="unreachable_eur":
                    e=next((e for e in self.history["snapshots"][self.period]["entities"] if e["id"]==eid),None)
                    vals.append(divide(m.get("unreachable"),self.fxvalues[(self.period,e["currency"])]) if e else None)
                else:
                    vals.append(m.get(key))
            self.formula("Synthese",end,col,f'=IF(COUNT({refs})={len(ids)},SUM({refs}),"{NA}")',complete_sum(vals),"integer" if col==8 else "total")
        sh.write(end+2,1,"Stock identifié = à traiter + injoignables lorsque cette catégorie existe. CST Nordics non transmis : total avec CST indisponible.",self.formats["note"])
        sh.write(end+3,1,"n.d. = donnée non disponible ou catégorie non applicable. Les totaux restent indisponibles si une valeur nécessaire manque.",self.formats["note"])
        start=end+6
        sh.write(start,1,"Flux du semestre glissant à la date de reporting",self.formats["section"])
        self.header("Synthese",start+1,["Pays","Flux 6 mois EUR","Moyenne mensuelle EUR","Traité 6 mois EUR","Efficacité montant","Efficacité nombre","Cohorte M-5 montant","Cohorte M-5 nombre"])
        for row,eid in enumerate(ids,start+2):
            self.write("Synthese",row,1,self.entities[eid])
            m=self.metrics.get((self.period,eid),{})
            for col,metric,key in [(2,"Q","flow_eur"),(3,"W","monthly_average"),(4,"R","processed_eur"),(5,"U","efficiency_amount"),(6,"V","efficiency_count"),(7,"X","oldest_amount"),(8,"Y","oldest_count")]:
                self.formula("Synthese",row,col,self.lookup(metric,"$C$4",f'"{eid}"'),m.get(key),"percent" if col>=5 else "number")
        totalrow = start+2+len(ids)
        self.write("Synthese",totalrow,1,"Total du périmètre","section")
        for col,key in [(2,"flow_eur"),(3,"monthly_average"),(4,"processed_eur")]:
            refs=f"{colname(col)}{start+3}:{colname(col)}{totalrow}"
            self.formula("Synthese",totalrow,col,f'=IF(COUNT({refs})={len(ids)},SUM({refs}),"{NA}")',
                         complete_sum(self.metrics.get((self.period,eid),{}).get(key) for eid in ids),"total")
        rr=totalrow+1
        total_flow=complete_sum(self.metrics.get((self.period,eid),{}).get("flow_eur") for eid in ids)
        total_processed=complete_sum(self.metrics.get((self.period,eid),{}).get("processed_eur") for eid in ids)
        self.formula("Synthese",totalrow,5,f'=IF(OR(COUNT(C{rr},E{rr})<>2,C{rr}=0),"{NA}",E{rr}/C{rr})',ratio(total_processed,total_flow),"percent")
        terms=f'{self.rng("Consolidation","A")},$C$4'
        entrants=self.rng("Consolidation","S")
        treated=self.rng("Consolidation","T")
        self.formula("Synthese",totalrow,6,
                     f'=IF(OR(COUNTIFS({terms})<>{len(ids)},COUNTIFS({terms},{entrants},"{NA}")>0,COUNTIFS({terms},{treated},"{NA}")>0,SUMIFS({entrants},{terms})=0),"{NA}",SUMIFS({treated},{terms})/SUMIFS({entrants},{terms}))',
                     ratio(complete_sum(self.metrics.get((self.period,eid),{}).get("processed_count") for eid in ids),complete_sum(self.metrics.get((self.period,eid),{}).get("count") for eid in ids)),"percent")
        sh.write(start+len(ids)+4,1,"Efficacité pondérée = total traité / total entrant. Cohorte M-5 = mois le plus ancien du semestre, mesuré à la clôture.",self.formats["note"])
        row=start+len(ids)+7
        for title,metric,key in [("Evolution du stock à traiter (EUR)","M","remaining_eur"),("Evolution du stock identifié hors CST (EUR)","N","identified_eur")]:
            sh.write(row,1,title,self.formats["section"])
            self.header("Synthese",row+1,["Pays"]+self.periods)
            for offset,eid in enumerate(ids,row+2):
                self.write("Synthese",offset,1,self.entities[eid])
                for col,period in enumerate(self.periods,2):
                    self.formula("Synthese",offset,col,self.lookup(metric,f'{colname(col)}${row+2}',f'"{eid}"'),self.metrics.get((period,eid),{}).get(key))
            row+=len(ids)+5
        sh.freeze_panes(7,2)
        sh.print_area(0,1,row,max(8,1+len(self.periods)))

    def evolution(self):
        sh=self.sheets["Evolution flux"]
        sh.set_column("A:A",3)
        sh.set_column("B:B",24)
        sh.set_column(2,max(8,len(self.periods)+1),20)
        sh.write(1,1,"Evolution des flux et efficacités",self.formats["title"])
        sh.write(2,1,"Trimestres de reporting de gauche à droite. Les semestres se chevauchent : ne pas les additionner.",self.formats["note"])
        row=4
        for title,metric,key,fmt in [
            ("Flux des cohortes du semestre glissant (EUR)","Q","flow_eur","number"),
            ("Moyenne mensuelle du semestre (EUR)","W","monthly_average","number"),
            ("Efficacité pondérée du semestre en montant","U","efficiency_amount","percent"),
            ("Efficacité pondérée du semestre en nombre","V","efficiency_count","percent"),
            ("Efficacité de la cohorte M-5 en montant","X","oldest_amount","percent"),
            ("Efficacité de la cohorte M-5 en nombre","Y","oldest_count","percent"),
            ("Flux des trois cohortes du trimestre (EUR)","Z","quarter_flow_eur","number"),
            ("Efficacité des trois cohortes du trimestre en montant","AB","quarter_efficiency","percent")]:
            sh.write(row,1,title,self.formats["section"])
            self.header("Evolution flux",row+1,["Pays"]+self.periods)
            for rr,eid in enumerate(self.entities,row+2):
                self.write("Evolution flux",rr,1,self.entities[eid])
                for col,period in enumerate(self.periods,2):
                    self.formula("Evolution flux",rr,col,self.lookup(metric,f'{colname(col)}${row+2}',f'"{eid}"'),self.metrics.get((period,eid),{}).get(key),fmt)
            row+=len(self.entities)+5
        sh.freeze_panes(6,2)
        sh.print_area(0,1,row,max(8,1+len(self.periods)))

    def controls(self):
        headers=["Trimestre","Niveau","Code","Pays","Observation","Source"]
        self.start("Controles","Contrôles et limites de la livraison",headers,[15,14,26,12,100,64],
                   "Observations détectées lors de la génération. Relancer Python après toute correction des fichiers source.")
        row=5
        observation_format=self.book.add_format({"font_name":"Arial","font_size":10,"text_wrap":True,"valign":"vcenter"})
        for snap in self.snapshots:
            for issue in snap.get("issues",[]):
                data=[snap["period"],issue["severity"],issue["code"],issue.get("entity",""),issue["message"],issue.get("source","")]
                for col,value in enumerate(data):
                    if col == 4:
                        self.sheets["Controles"].write_string(row,col,value,observation_format)
                    else:
                        self.write("Controles",row,col,value)
                self.sheets["Controles"].set_row(row, 32 if len(issue["message"]) > 100 else 19)
                row+=1
        if row==5:
            for col,value in enumerate([self.period,"info","NO_OBSERVATIONS","","Aucune observation détectée sur les contrôles exécutés.",""]):
                self.write("Controles",row,col,value)
            row+=1
        self.table("Controles",headers,row-5,"Observations")
        sh=self.sheets["Controles"]
        sh.conditional_format(5,1,row-1,1,{"type":"text","criteria":"containing","value":"warning","format":self.formats["warning"]})
        row+=3
        sh.write(row,0,"Traçabilité des fichiers importés",self.formats["section"])
        row+=1
        for snap in self.snapshots:
            for source in snap.get("sources",[]) + snap.get("configurations",[]):
                sh.write_row(row,0,[snap["period"],source["file"],source["sha256"]])
                row+=1

    def guide(self):
        sh=self.sheets["Mode emploi"]
        sh.set_column("A:A",3)
        sh.set_column("B:B",35)
        sh.set_column("C:C",112)
        sh.write(1,1,"Utilisation du reporting",self.formats["title"])
        lines=[
            ("Période", "Choisir un trimestre en Synthese!C4. Les périodes disponibles proviennent de l'historique enregistré."),
            ("Périmètre", "Ce classeur consolide uniquement les entités configurées et fournies. Il ne représente pas tous les pays de l'ancienne synthèse."),
            ("Trimestre suivant", "Déposer les fichiers pays et Taux.xlsx dans QX_20XX puis relancer le programme avec le même fichier d'historique."),
            ("Nouvelle exécution", "Le trimestre est remplacé, sans doublon. Les autres trimestres conservent leurs valeurs source et leurs propres taux."),
            ("Tableaux Excel", "Filtrer les tables par trimestre et pays. Leur taille est reconstruite à chaque exécution. Ce ne sont pas des tableaux croisés dynamiques."),
            ("Stock à traiter", "KPI_2 : stock initial moins les catégories traitées applicables au pays. KPI_3 fournit la répartition de ce stock."),
            ("Stock identifié", "Stock à traiter plus les clients injoignables pour le modèle standard. Pour Nordics, stock à traiter hors CST non transmis."),
            ("CST Nordics", "Le compte CST reçu séparément par courriel n'est pas disponible ici. Le stock total avec CST reste n.d. Aucun montant n'est inventé."),
            ("Flux", "Six cohortes mensuelles observées à la date du reporting. Une cohorte peut être réévaluée au trimestre suivant."),
            ("Efficacité", "Le ratio pondéré est total traité / total entrant. La cohorte M-5 correspond à l'indicateur trimestriel repris dans l'ancienne synthèse."),
            ("Taux de change", "Montants locaux divisés par le taux de clôture (devise locale pour 1 EUR), pour les stocks comme pour les flux."),
            ("Données absentes", "n.d. signifie non disponible ou non applicable. Les zéros numériques sont conservés. Une absence nécessaire rend le résultat indisponible."),
            ("Source des calculs", "Les données proviennent des fichiers pays. Les fichiers source restent inchangés. Les calculs du classeur n'ont aucun lien externe."),
            ("Formules source", "Le programme lit les résultats enregistrés par Excel. Il refuse les formules sans résultat en cache, mais ne peut pas détecter un cache ancien."),
            ("Corrections", "Corriger les sources ou la configuration puis relancer Python. Les changements faits directement dans la sortie sont écrasés au prochain run."),
            ("Contrôles", "Consulter Controles. Les observations d'import sont un constat à la génération, pas un contrôle automatique des modifications manuelles du classeur."),
            ("Historique antérieur", "Il commence avec les trimestres importés. Aucune migration des anciens onglets de synthèse n'a été effectuée."),
            ("Exceptions différées", "Roumanie, Brésil, correction Germany, chiffres UK et saisie du CST Nordics nécessitent des règles métier spécifiques ultérieures."),
        ]
        wrap=self.book.add_format({"font_name":"Arial","font_size":10,"text_wrap":True,"valign":"top"})
        for row,(label,text) in enumerate(lines,4):
            sh.write(row,1,label,self.formats["section"])
            sh.write(row,2,text,wrap)
            sh.set_row(row,34)
        sh.print_area(0,1,4+len(lines),2)

    def build(self):
        try:
            self.rates()
            self.stock()
            self.flows()
            self.distribution()
            self.consolidation()
            self.summary()
            self.evolution()
            self.controls()
            self.guide()
            self.sheets["Synthese"].activate()
        finally:
            self.book.close()


def write_workbook(path: Path, history: dict, period: str) -> None:
    Report(path, history, period).build()
