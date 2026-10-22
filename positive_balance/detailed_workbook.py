"""Slide-oriented Excel pages with native, editable combination charts.

The compact writer remains the calculation owner. These pages link to its
normalized inputs/calculations and add only presentation-specific groupings.
"""
from __future__ import annotations

from datetime import datetime
import math
from pathlib import Path

from xlsxwriter.utility import xl_col_to_name as colname

from .detail_data import build_presentation_data
from .model import complete_sum, ratio
from .workbook import Report, NA


class DetailedReport(Report):
    def __init__(self, path, history, period, config):
        self.data = build_presentation_data(history, period, config)
        self.config = config
        self.country_sheets = {c["id"]: f"{c['id']} {c['name']}"[:31] for c in self.data["countries"]}
        self.group_sheets = {}
        front = ["Global"]
        for group in self.data["groups"]:
            if not group["countries"]:
                continue
            label = "CORE" if group["id"] == "core" else "NON CORE" if group["id"] == "noncore" else group["id"].upper()[:18]
            stock_name = f"{label} Stock"
            flow_name = f"{label} Flows" if group["id"] != "noncore" else None
            self.group_sheets[group["id"]] = (stock_name, flow_name)
            front.append(stock_name)
            if flow_name:
                front.append(flow_name)
        front += list(self.country_sheets.values())
        if len({n.casefold() for n in front}) != len(front):
            raise ValueError("Duplicate presentation sheet names; use distinct country names/IDs")
        super().__init__(path, history, period, front_sheets=front)
        self.sheets["GraphData"] = self.book.add_worksheet("GraphData")
        self.sheets["Presentation inputs"] = self.book.add_worksheet("Presentation inputs")
        self.colors = {"country_header":"00627A","global_header":"77B900","total_fill":"B7DDE8",
                       "refund":"92C43D","false_positive":"006633","remaining":"A6A6A6","unreachable":"D9A72C",
                       "authority_refund":"6CAED6","other_without_refund":"B4A7D6","efficiency_amount":"005B63",
                       "efficiency_count":"7E245A","vintage_old":"0095B6","vintage_prior":"D9A72C","vintage_current":"77B900"}
        self.colors.update(self.data["style"])
        self.colors.pop("font", None)
        self.font = self.data["style"].get("font","Arial")
        self.labels = {"refund":"With client refund","false_positive":"False positive balance","remaining":"To be processed",
                       "unreachable":"Customer not reachable","authority_refund":"Public authority refund","other_without_refund":"Other processing without refund"}
        self.labels.update(self.data["category_labels"])
        self.pf = {}
        base = {"font_name":self.font,"font_size":10,"valign":"vcenter"}
        for key, extra in {
            "body":{"border":1,"border_color":"8B969B"},
            "label":{"border":1,"border_color":"8B969B","text_wrap":True},
            "num":{"border":1,"border_color":"8B969B","num_format":'#,##0;(#,##0);"–"',"align":"right"},
            "pct":{"border":1,"border_color":"8B969B","num_format":"0.0%","align":"right"},
            "total":{"border":1,"border_color":"607D8B","bg_color":self.color("total_fill"),"bold":True,"num_format":'#,##0;(#,##0);"–"',"align":"right"},
            "totalpct":{"border":1,"border_color":"607D8B","bg_color":self.color("total_fill"),"bold":True,"num_format":"0.0%","align":"right"},
            "header":{"border":1,"border_color":"607D8B","bg_color":"EAF3F4","bold":True,"text_wrap":True,"align":"center"},
            "date":{"border":1,"border_color":"607D8B","bg_color":"EAF3F4","bold":True,"num_format":"dd/mm/yyyy","align":"center","font_color":self.color("country_header")},
            "country_title":{"bold":True,"font_size":24,"font_color":"FFFFFF","bg_color":self.color("country_header")},
            "global_title":{"font_size":23,"font_color":"173331","bg_color":self.color("global_header")},
            "note":{"font_size":9,"font_color":"526269","text_wrap":True},
            "insight":{"font_size":12,"text_wrap":True,"valign":"top","font_color":"263A42"},
            "footer":{"font_size":8,"font_color":"526269"},
            "box":{"border":1,"border_color":"6C9EA8"},
            "helper":{"num_format":"#,##0.00;(#,##0.00);\"–\""},
            "helper_header":{"bg_color":"18334F","font_color":"FFFFFF","bold":True,"text_wrap":True},
        }.items():
            self.pf[key]=self.book.add_format({**base,**extra})
        self.helper_row=4
        self.helper_meta={}
        self.input_refs={}
        self.chart_count=0
        self.sheets["GraphData"].hide_gridlines(2)
        self.sheets["GraphData"].set_column("A:A",26)
        self.sheets["GraphData"].set_column("B:Q",20)
        self.sheets["GraphData"].write("A1","Formula-linked chart data",self.formats["title"])

    def color(self,key):
        return "#"+self.colors[key].lstrip("#")

    def merged(self,sheet,r,c,r2,c2,value,fmt="body",formula=None):
        sh=self.sheets[sheet]
        if (r,c)!=(r2,c2):
            sh.merge_range(r,c,r2,c2,"",self.pf[fmt])
        if formula is not None:
            sh.write_formula(r,c,formula,self.pf[fmt],NA if value is None else value)
        elif value is None:
            sh.write_string(r,c,NA,self.pf[fmt])
        elif isinstance(value,datetime):
            sh.write_datetime(r,c,value,self.pf[fmt])
        else:
            sh.write(r,c,value,self.pf[fmt])
        return f"'{sheet}'!{colname(c)}{r+1}"

    def page(self,sheet,title,offset=0,country=False,height=42):
        sh=self.sheets[sheet]
        sh.freeze_panes(0,0)
        sh.set_default_row(13.5)
        sh.set_column("A:A",2.3)
        sh.set_column("B:AF",5.3)
        sh.set_zoom(80)
        sh.set_landscape()
        sh.set_paper(9)
        sh.fit_to_pages(1,0)
        sh.set_margins(.15,.15,.2,.2)
        sh.repeat_rows(0,0)
        sh.set_tab_color(self.color("country_header" if country else "global_header"))
        strip=self.book.add_format({"bg_color":self.color("country_header" if country else "global_header")})
        for r in range(offset+1,offset+height):
            sh.write_blank(r,0,None,strip)
        self.merged(sheet,offset+1,1,offset+3,12 if country else 31,title,"country_title" if country else "global_title")
        self.merged(sheet,offset+height-1,1,offset+height-1,31,f"Positive Balance     {self.period}     Source: country workbooks and Taux.xlsx     Internal","footer")
        sh.print_area(0,0,offset+height-1,31)

    def query(self,sheet,column,period,eid,extra=None):
        vals=self.rng(sheet,column)
        if extra:
            # Direct equality keeps literal labels like '<2019' and '> 1 000 €'
            # distinct from spreadsheet criteria operators and wildcards.
            filters=[f'--({self.rng(sheet,"A")}="{period}")',f'--({self.rng(sheet,"B")}="{eid}")']
            filters += [f'--({self.rng(sheet,col)}={criterion})' for col,criterion in extra]
            conditions=','.join(filters)
            return f'IF(OR(SUMPRODUCT({conditions})=0,SUMPRODUCT({conditions},--({vals}="{NA}"))>0),"{NA}",SUMPRODUCT({conditions},{vals}))'
        terms=f'{self.rng(sheet,"A")},"{period}",{self.rng(sheet,"B")},"{eid}"'
        return f'IF(OR(COUNTIFS({terms})=0,COUNTIFS({terms},{vals},"{NA}")>0),"{NA}",SUMIFS({vals},{terms}))'

    def metric(self,period,eid,column):
        return self.lookup(column,f'"{period}"',f'"{eid}"')

    def helper(self,title,headers,rows,formulas=None):
        sh=self.sheets["GraphData"]
        start=self.helper_row
        sh.write(start,0,title,self.formats["section"])
        sh.write_row(start+1,0,headers,self.pf["helper_header"])
        sh.set_row(start+1,32)
        first=start+2
        for index,row in enumerate(rows):
            for col,value in enumerate(row):
                formula=(formulas[index][col] if formulas else None)
                if value is None:
                    # A real blank produces a chart gap; it must not become 0%.
                    sh.write_blank(first+index,col,None)
                elif formula:
                    sh.write_formula(first+index,col,formula,self.pf["helper"],value)
                else:
                    sh.write(first+index,col,value)
        self.helper_row=first+len(rows)+3
        result={"sheet":"GraphData","first":first,"last":first+len(rows)-1,"headers":headers,"rows":rows}
        self.helper_meta[title]=result
        return result

    @staticmethod
    def series_range(block,col):
        return [block["sheet"],block["first"],col,block["last"],col]

    @staticmethod
    def href(block,index,col):
        return f"'GraphData'!{colname(col)}{block['first']+index+1}"

    def chart_style(self,chart,title,width,height,percent=False):
        chart.set_title({"name":title,"name_font":{"name":self.font,"size":11,"bold":False}})
        chart.set_size({"width":width,"height":height})
        chart.set_chartarea({"border":{"color":"CFD8DC"},"fill":{"color":"FFFFFF"}})
        chart.set_plotarea({"border":{"none":True},"fill":{"color":"FFFFFF"}})
        chart.set_x_axis({"num_font":{"name":self.font,"size":8},"line":{"color":"BAC5CA"},"major_tick_mark":"none"})
        chart.set_y_axis({"min":0,"num_format":"0%" if percent else "#,##0","num_font":{"name":self.font,"size":8},"major_gridlines":{"visible":True,"line":{"color":"DDE3E5","width":.5}},"line":{"none":True}})
        chart.set_legend({"position":"bottom","font":{"name":self.font,"size":8}})
        chart.show_blanks_as("gap")

    def insert(self,sheet,row,col,chart):
        self.sheets[sheet].insert_chart(row,col,chart,{"x_offset":2,"y_offset":2,"object_position":1})
        self.chart_count+=1

    def presentation_inputs(self):
        sh=self.sheets["Presentation inputs"]
        sh.hide_gridlines(2)
        sh.set_column("A:C",20)
        sh.set_column("D:E",22)
        sh.set_column("F:F",65)
        sh.write("A2","Source totals and commentary",self.formats["title"])
        sh.write_row(4,0,["Period","Country","Distribution","Reported count","Reported amount local","Source"],self.formats["header"])
        row=5
        for country in self.data["countries"]:
            for kind in ("stock_bins","stock_vintages"):
                total=country[kind+"_reported_totals"]
                sh.write_row(row,0,[self.period,country["id"],kind])
                for col,field in ((3,"count"),(4,"amount")):
                    if total.get(field) is None:
                        sh.write_string(row,col,NA)
                    else:
                        sh.write_number(row,col,total[field],self.formats["number"])
                    self.input_refs[(country["id"],kind,field)]=f"'Presentation inputs'!{colname(col)}{row+1}"
                sh.write(row,5,total.get("source",""))
                row+=1
        sh.add_table(4,0,row-1,5,{"name":"PresentationSourceTotals","columns":[{"header":v} for v in ["Period","Country","Distribution","Reported count","Reported amount local","Source"]],"style":"Table Style Medium 2"})
        self.comment_row=row+3

    def country_chart(self,country,sheet):
        eid=country["id"]
        categories=country["categories"]
        is_eur=country["display_currency"]=="EUR"
        sourcecols={"refund":"H","false_positive":"J","remaining":"O","unreachable":"I","authority_refund":"K","other_without_refund":"L"}
        rows,forms=[],[]
        for record in country["monthly"]:
            month=datetime.strptime(record["month"],"%Y-%m")
            criteria=[("E",f"DATE({month.year},{month.month},1)")]
            values=[month.strftime("%b-%y")]
            expressions=[None]
            for cat in categories:
                val=record[cat+"_eur"] if is_eur else record[cat]
                query=self.query("Flux",sourcecols[cat],self.period,eid,criteria)
                formula="="+query
                if is_eur:
                    rate=self.metric(self.period,eid,"E")[1:]
                    formula=f'=IF(ISNUMBER({query}),({query})/({rate}),"{NA}")'
                values.append(val)
                expressions.append(formula)
            for key,col in (("efficiency_amount","T"),("efficiency_count","U")):
                values.append(record[key])
                expressions.append("="+self.query("Flux",col,self.period,eid,criteria))
            rows.append(values)
            forms.append(expressions)
        block=self.helper(eid+" monthly chart",["Month"]+[self.labels[c] for c in categories]+["Efficiency in amount","Efficiency in number"],rows,forms)
        columns=self.book.add_chart({"type":"column","subtype":"stacked"})
        for index,cat in enumerate(categories,1):
            columns.add_series({"name":self.labels[cat],"categories":self.series_range(block,0),"values":self.series_range(block,index),"fill":{"color":self.color(cat)},"border":{"none":True},"gap":70})
        lines=self.book.add_chart({"type":"line"})
        for index,key,label,pos in [(len(categories)+1,"efficiency_amount","Efficiency in amount","above"),(len(categories)+2,"efficiency_count","Efficiency in number","below")]:
            lines.add_series({"name":label,"categories":self.series_range(block,0),"values":self.series_range(block,index),"y2_axis":True,"smooth":False,"line":{"color":self.color(key),"width":2.25},"marker":{"type":"none"},"data_labels":{"value":True,"position":pos,"num_format":"0%","font":{"name":self.font,"size":8,"color":self.color(key)}}})
        columns.combine(lines)
        self.chart_style(columns,f"{country['name']}: Monthly inflow in {country['display_currency']} and efficiency",640,365)
        effs=[r[k] for r in country["monthly"] for k in ("efficiency_amount","efficiency_count") if r[k] is not None]
        maximum=max(1,max(effs,default=1))
        lines.set_y2_axis({"min":0,"max":math.ceil((maximum-1e-9)*5)/5,"major_unit":.2,"num_format":"0%","num_font":{"size":8},"major_gridlines":{"visible":False},"line":{"none":True}})
        self.insert(sheet,6,1,columns)
        return block

    def stock_distribution(self,country,sheet,kind,start):
        rows=country[kind]
        total=country[kind+"_totals"]
        is_eur=country["display_currency"]=="EUR"
        spans=[(20,22),(23,24),(25,27),(28,29),(30,31)]
        title="By step of amount" if kind=="stock_bins" else "By date of positive balance"
        self.merged(sheet,start,17,start+len(rows)-1,19,title,"header")
        totalrow=start+len(rows)
        for i,record in enumerate(rows):
            r=start+i
            label=record["label"]
            self.merged(sheet,r,spans[0][0],r,spans[0][1],label,"label")
            for colpair,field,sourcecol,fmt in [(spans[1],"count","G","num"),(spans[2],"amount_eur" if is_eur else "amount","J" if is_eur else "H","num")]:
                query=self.query("Repartition stock",sourcecol,self.period,country["id"],[("E",'"Tranche"' if kind=="stock_bins" else '"Millésime"'),("F",f'"{label.replace(chr(34),chr(34)*2)}"')])
                self.merged(sheet,r,colpair[0],r,colpair[1],record[field],fmt,formula="="+query)
            for cols,field,numcol,dencol in [(spans[3],"pct_count","X","X"),(spans[4],"pct_amount","Z","Z")]:
                n=f"{numcol}{r+1}"; d=f"{dencol}{totalrow+1}"
                self.merged(sheet,r,cols[0],r,cols[1],record[field],"pct",formula=f'=IF(OR(COUNT({n},{d})<>2,{d}=0),"{NA}",{n}/{d})')
        self.merged(sheet,totalrow,17,totalrow,22,"Total:","total")
        for columns,field in [(spans[1],"count"),(spans[2],"amount")]:
            targetcol=colname(columns[0])
            rng=f"{targetcol}{start+1}:{targetcol}{totalrow}"
            if total[field+"_basis"]=="detail_sum":
                expr=f'=IF(COUNT({rng})={len(rows)},SUM({rng}),"{NA}")'
            else:
                ref=self.input_refs[(country["id"],kind,field)]
                expr=f'={ref}'
                if field=="amount" and is_eur:
                    expr=f'=IF(ISNUMBER({ref}),{ref}/({self.metric(self.period,country["id"],"E")[1:]}),"{NA}")'
            value=total["amount_eur"] if field=="amount" and is_eur else total[field]
            self.merged(sheet,totalrow,columns[0],totalrow,columns[1],value,"total",formula=expr)
        for columns,field,den in [(spans[3],"pct_count","X"),(spans[4],"pct_amount","Z")]:
            ref=f"{den}{totalrow+1}"
            self.merged(sheet,totalrow,columns[0],totalrow,columns[1],total[field],"totalpct",formula=f'=IF(OR(NOT(ISNUMBER({ref})),{ref}=0),"{NA}",1)')
        return totalrow+2

    def country_page(self,country):
        sheet=self.country_sheets[country["id"]]
        end=6+(len(country["stock_bins"])+2 if country["show_stock_bins"] else 0)+len(country["stock_vintages"])+2
        notes_start=max(29,end+4)
        self.page(sheet,country["name"]+(f" in {country['display_currency']}" if country["display_currency"]!="EUR" else ""),country=True,height=notes_start+13)
        self.merged(sheet,4,1,5,15,f"Reporting at {datetime.fromisoformat(self.data['date']).strftime('%d %B %Y')}. Six monthly cohorts.","note")
        self.country_chart(country,sheet)
        self.merged(sheet,1,17,4,22,"STOCK OF POSITIVE BALANCES\nto be processed (KPI_3)","header")
        self.merged(sheet,1,23,1,31,datetime.fromisoformat(self.data["date"]),"date")
        for c,c2,label in [(23,24,"Number of\ncontracts"),(25,27,f"Total amount\n{country['display_currency']}"),(28,29,"% in\nnumber"),(30,31,"% in\namount")]:
            self.merged(sheet,2,c,4,c2,label,"header")
        row=6
        if country["show_stock_bins"]:
            row=self.stock_distribution(country,sheet,"stock_bins",row)
        row=self.stock_distribution(country,sheet,"stock_vintages",row)
        warnings=[]
        if any(r["amount"] is None or r["count"] is None for r in country["stock_vintages"]):
            warnings.append("Some vintage details are missing. The reported total is retained.")
        if country["id"]=="ZA" and not country["show_stock_bins"]:
            warnings.append("Amount bands omitted: source distribution is not informative.")
        if country["display_currency"]!=country["currency"] and country["show_stock_bins"]:
            warnings.append(f"Amount-band thresholds remain in source currency ({country['currency']}); amounts are converted to EUR.")
        if warnings:
            self.merged(sheet,row,17,row+2,31," ".join(warnings),"note")
        self.country_insights(country,sheet,notes_start)

    def country_insights(self,country,sheet,notes_start=29):
        eid=country["id"]; m=country["metrics"]
        eur=country["display_currency"]=="EUR"
        currency=country["display_currency"]
        avg=m["monthly_average"] if eur else m["monthly_average_local"]
        balance=m["average_balance_eur"] if eur else m["average_balance_local"]
        values=[avg,balance,m["oldest_amount"],m["oldest_count"],m["remaining_eur"] if eur else m["remaining"],m["identified_eur"] if eur else m["identified"],m["stock_change"]]
        rate=self.metric(self.period,eid,"E")[1:]
        avg_expr=self.metric(self.period,eid,"W")[1:]
        if not eur:
            avg_expr=f"({avg_expr})*({rate})"
        flow=self.metric(self.period,eid,"Q")[1:]
        count=self.metric(self.period,eid,"S")[1:]
        bal_expr=f'IF(OR(NOT(ISNUMBER({flow})),NOT(ISNUMBER({count})),({count})=0),"{NA}",({flow})/({count})'+(')' if eur else f'*({rate}))')
        forms=["="+avg_expr,"="+bal_expr,self.metric(self.period,eid,"X"),self.metric(self.period,eid,"Y"),self.metric(self.period,eid,"M" if eur else "H"),self.metric(self.period,eid,"N" if eur else "J"),None]
        helper=self.helper(eid+" insight metrics",["Monthly average","Average contract","M-5 amount","M-5 count","Remaining","Identified","Change"],[values],[forms])
        refs=[self.href(helper,0,i) for i in range(len(values))]
        def txt(ref,fmt="#,##0"):
            return f'IF(ISNUMBER({ref}),TEXT({ref},"{fmt}"),"n.d.")'
        def num(value,pct=False):
            return "n.d." if value is None else f"{value:.1%}" if pct else f"{value:,.0f}"
        captions=[
            (f"Average monthly new flow: {num(avg)} {currency}. Average balance per contract: {num(balance)} {currency}.",f'="Average monthly new flow: "&{txt(refs[0])}&" {currency}. Average balance per contract: "&{txt(refs[1])}&" {currency}."'),
            (f"Oldest cohort (M-5): {num(values[2],True)} treated in amount and {num(values[3],True)} in number.",f'="Oldest cohort (M-5): "&{txt(refs[2],"0.0%")}&" treated in amount and "&{txt(refs[3],"0.0%")}&" in number."'),
            (f"Stock to process: {num(values[4])} {currency}. Identified stock: {num(values[5])} {currency}.",f'="Stock to process: "&{txt(refs[4])}&" {currency}. Identified stock: "&{txt(refs[5])}&" {currency}."'),
        ]
        if country["requires_cst"]:
            captions.append(("Separate CST account not supplied. Identified stock excludes CST.",None))
        elif country["previous_metrics"] is None:
            captions.append((f"Prior-quarter comparison unavailable: no {self.data['previous_period']} source snapshot has been imported.",None))
        else:
            captions.append((f"Identified stock change vs {self.data['previous_period']}: {num(m['stock_change'],True)}.",None))
        self.comment_box(sheet,notes_start,1,notes_start+10,31,captions,country["comments"])

    def comment_box(self,sheet,r,c,r2,c2,lines,manual=()):
        sh=self.sheets[sheet]
        sh.conditional_format(r,c,r2,c2,{"type":"no_blanks","format":self.pf["box"]})
        # Merged lines are individually editable/copyable, like the reference notes box.
        available=r2-r+1
        contents=list(lines)+[(str(v),None) for v in manual]
        if len(contents)*2>available:
            # Store overflow in a separate input section, never on a later page.
            capacity=max(1,available//2-1)
            overflow=contents[capacity:]
            contents=contents[:capacity]+[("Additional commentary is available in Presentation inputs.",None)]
            start=self.comment_row
            self.merged("Presentation inputs",start,0,start,5,sheet+" — additional commentary","header")
            for index,(text,formula) in enumerate(overflow):
                self.merged("Presentation inputs",start+1+index*2,0,start+2+index*2,5,text,"insight",formula)
            self.comment_row=start+len(overflow)*2+3
        for i,(text,formula) in enumerate(contents):
            self.merged(sheet,r+i*2,c,r+i*2+1,c2,text,"insight",formula)

    def stock_matrix(self,group):
        rows=[];forms=[]
        for record in group["stock_rows"]:
            eid=record["id"]
            current=record["current_identified_eur"]
            prior=record["previous_identified_eur"]
            vals=[record["name"],prior,current,record["change"],record["current_remaining_eur"]]
            expressions=[None,self.metric(self.data["previous_period"],eid,"N"),self.metric(self.period,eid,"N"),None,self.metric(self.period,eid,"M")]
            for metric in ("identified_eur","remaining_eur"):
                sourcecol="S" if metric=="identified_eur" else "R"
                for vintage in record["vintages"]:
                    label=vintage["label"]
                    if label.startswith("<"):
                        # Source old bucket + individual earlier years, never split a bucket.
                        source=next(e for e in self.history["snapshots"][self.period]["entities"] if e["id"]==eid)
                        labels=[x["vintage"] for x in source["stock"] if x["vintage"].startswith("<") or int(x["vintage"])<int(self.period[:4])-1]
                    else:
                        labels=[label]
                    queries=[self.query("Stock",sourcecol,self.period,eid,[("E",f'"{v}"')]) for v in labels]
                    expr=f'=IF(COUNT({",".join(queries)})={len(queries)},SUM({",".join(queries)}),"{NA}")' if queries else None
                    vals.append(vintage[metric]);expressions.append(expr)
            rows.append(vals);forms.append(expressions)
        block=self.helper(group["id"]+" stock matrix",["Country","Prior identified EUR","Current identified EUR","Change","Remaining EUR"]+["Identified "+v["label"] for v in group["global_vintages"]]+["Remaining "+v["label"] for v in group["global_vintages"]],rows,forms)
        for i,record in enumerate(group["stock_rows"]):
            if record["change"] is not None:
                r=block["first"]+i
                self.sheets["GraphData"].write_formula(r,3,f'=IF(OR(COUNT(B{r+1},C{r+1})<>2,B{r+1}=0),"{NA}",C{r+1}/B{r+1}-1)',self.pf["helper"],record["change"])
        return block

    def vintage_chart(self,block,title,firstcol,width=595,height=250):
        chart=self.book.add_chart({"type":"column","subtype":"stacked"})
        year=int(self.period[:4])
        for offset,label,color in [(0,f"<{year-1}","vintage_old"),(1,str(year-1),"vintage_prior"),(2,str(year),"vintage_current")]:
            chart.add_series({"name":label,"categories":self.series_range(block,0),"values":self.series_range(block,firstcol+offset),"fill":{"color":self.color(color)},"border":{"none":True},"gap":85})
        self.chart_style(chart,title,width,height)
        return chart

    def stock_page(self,group,sheet,offset=0,noncore=False):
        self.page(sheet,group["name"]+": Stock analysis",offset)
        self.merged(sheet,offset+4,1,offset+5,31,group["scope_note"],"note")
        block=self.stock_matrix(group)
        spans=[(1,6),(7,9),(10,12),(13,15)]
        for span,label in zip(spans,["Countries",self.data["previous_period"],self.period,"Evolution\ncurrent / prior"]):
            self.merged(sheet,offset+7,span[0],offset+9,span[1],label,"header")
        for i,record in enumerate(group["stock_rows"]):
            for j,span in enumerate(spans):
                value=block["rows"][i][j]
                ref=self.href(block,i,j)
                self.merged(sheet,offset+10+i,span[0],offset+10+i,span[1],value,"body" if j==0 else "pct" if j==3 else "num",None if j==0 else f'=IF(ISNUMBER({ref}),{ref},"{NA}")')
        totalrow=offset+10+len(group["countries"])
        self.merged(sheet,totalrow,1,totalrow,6,"TOTAL available perimeter","total")
        for span,key,m in [(spans[1],"identified_eur",group["previous_metrics"]),(spans[2],"identified_eur",group["metrics"])]:
            letter=colname(span[0]);rng=f"{letter}{offset+11}:{letter}{totalrow}"
            self.merged(sheet,totalrow,span[0],totalrow,span[1],m[key] if m else None,"total",f'=IF(COUNT({rng})={len(group["countries"])},SUM({rng}),"{NA}")')
        self.merged(sheet,totalrow,13,totalrow,15,group["metrics"]["stock_change"],"totalpct",f'=IF(OR(COUNT(H{totalrow+1},K{totalrow+1})<>2,H{totalrow+1}=0),"{NA}",K{totalrow+1}/H{totalrow+1}-1)')
        notes=[("Identified stock = stock to be processed + unreachable customers where this category applies.",None),
               ("The vintage table excludes the separate CST accounts." if noncore else "The lower chart isolates stock still to be processed. Vintage buckets are calculated from KPI_2.",None)]
        if not self.data["previous_available"]:
            notes.append((f"{self.data['previous_period']} has not been imported. The comparison is left unavailable.",None))
        self.comment_box(sheet,max(totalrow+3,offset+23),1,offset+38,15,notes,group["comments"])
        if noncore:
            self.noncore_tables(group,sheet,offset,block)
        else:
            self.insert(sheet,offset+7,17,self.vintage_chart(block,f"Total identified stock in EUR: {self.period}",5))
            self.insert(sheet,offset+23,17,self.vintage_chart(block,f"Stock to be processed in EUR: {self.period}",8))

    def noncore_tables(self,group,sheet,offset,block):
        spans=[(17,21),(22,24),(25,26),(27,28),(29,31)]
        headers=["Stock in EUR","Total"]+[v["label"] for v in group["global_vintages"]]
        for span,label in zip(spans,headers):
            self.merged(sheet,offset+7,span[0],offset+9,span[1],label,"header")
        for i,record in enumerate(group["stock_rows"]):
            vals=[record["name"],record["current_identified_eur"]]+[v["identified_eur"] for v in record["vintages"]]
            for j,(span,value) in enumerate(zip(spans,vals)):
                col=2 if j==1 else j+3
                ref=self.href(block,i,col)
                self.merged(sheet,offset+10+i,span[0],offset+10+i,span[1],value,"body" if j==0 else "num",None if j==0 else f'=IF(ISNUMBER({ref}),{ref},"{NA}")')
        totalrow=offset+10+len(group["countries"])
        self.merged(sheet,totalrow,17,totalrow,21,"TOTAL available","total")
        vals=[group["metrics"]["identified_eur"]]+[v["identified_eur"] for v in group["global_vintages"]]
        for span,value in zip(spans[1:],vals):
            letter=colname(span[0]);rng=f"{letter}{offset+11}:{letter}{totalrow}"
            self.merged(sheet,totalrow,span[0],totalrow,span[1],value,"total",f'=IF(COUNT({rng})={len(group["countries"])},SUM({rng}),"{NA}")')
        self.merged(sheet,offset+20,17,offset+21,31,"Waiting accounts for Nordics (CST)","header")
        cstspans=[(17,21),(22,25),(26,29),(30,31)]
        for span,label in zip(cstspans,["Country",self.data["previous_period"]+" EUR",self.period+" EUR","Change"]):
            self.merged(sheet,offset+22,span[0],offset+24,span[1],label,"header")
        for i,country in enumerate(group["countries"]):
            self.merged(sheet,offset+25+i,17,offset+25+i,21,country["name"],"body")
            for j,span in enumerate(cstspans[1:]):
                self.merged(sheet,offset+25+i,span[0],offset+25+i,span[1],None,"num" if j<2 else "pct")
        self.merged(sheet,offset+30,17,offset+36,31,"CST balances are received separately and are not present in the source workbooks.\n\nRomania and Brazil are not included in this report.","note")

    def monthly_group(self,group):
        rows=[];forms=[]
        for i,monthtext in enumerate(self.data["months"]):
            month=datetime.strptime(monthtext,"%Y-%m")
            rows.append([month.strftime("%b-%y")]+[c["monthly"][i]["amount_eur"] for c in group["countries"]])
            forms.append([None]+["="+self.query("Flux","R",self.period,c["id"],[("E",f"DATE({month.year},{month.month},1)")]) for c in group["countries"]])
        return self.helper(group["id"]+" monthly by country",["Month"]+[c["name"] for c in group["countries"]],rows,forms)

    def efficiency_group(self,group):
        rows=[];forms=[]
        for country in group["countries"]:
            previous=country["previous_metrics"]
            rows.append([country["name"],previous["oldest_count"] if previous else None,country["metrics"]["oldest_count"],previous["oldest_amount"] if previous else None,country["metrics"]["oldest_amount"]])
            forms.append([None,self.metric(self.data["previous_period"],country["id"],"Y"),self.metric(self.period,country["id"],"Y"),self.metric(self.data["previous_period"],country["id"],"X"),self.metric(self.period,country["id"],"X")])
        return self.helper(group["id"]+" oldest cohort comparison",["Country","Prior number","Current number","Prior amount","Current amount"],rows,forms)

    def flow_page(self,group,sheet,offset=0):
        tablesrow=max(10+len(group["countries"])+3,20)
        height=max(42,tablesrow+len(group["countries"])+10)
        self.page(sheet,group["name"]+": Monthly inflow and treatment efficiency",offset,height=height)
        self.merged(sheet,offset+4,1,offset+5,31,"EUR amounts. M-5 efficiency uses the oldest cohort in each reporting snapshot. "+group["scope_note"],"note")
        monthly=self.monthly_group(group)
        flowchart=self.book.add_chart({"type":"column","subtype":"stacked"})
        palette=["00627A","B75132","4AA3B2","7E245A","549742","D9A72C","344D67","CD5858"]
        for i,country in enumerate(group["countries"],1):
            flowchart.add_series({"name":country["name"],"categories":self.series_range(monthly,0),"values":self.series_range(monthly,i),"fill":{"color":"#"+palette[(i-1)%len(palette)]},"border":{"none":True}})
        self.chart_style(flowchart,"Monthly flow in EUR",520,260)
        self.insert(sheet,offset+7,1,flowchart)
        eff=self.efficiency_group(group)
        bars=self.book.add_chart({"type":"column"})
        prior_available=any(c["previous_metrics"] is not None for c in group["countries"])
        for col,name,color in [(1,f"Number {self.data['previous_period']}","C5E3E6"),(2,f"Number {self.period}","0095B6")]:
            if col==1 and not prior_available:
                continue
            bars.add_series({"name":name,"categories":self.series_range(eff,0),"values":self.series_range(eff,col),"fill":{"color":"#"+color},"border":{"none":True}})
        lines=self.book.add_chart({"type":"line"})
        for col,name,dash in [(3,f"Amount {self.data['previous_period']}","dash"),(4,f"Amount {self.period}","solid")]:
            if col==3 and not prior_available:
                continue
            lines.add_series({"name":name,"categories":self.series_range(eff,0),"values":self.series_range(eff,col),"smooth":False,"line":{"color":"#C04040","width":2,"dash_type":dash},"marker":{"type":"none"}})
        bars.combine(lines)
        self.chart_style(bars,"Efficiency after 6 months (M-5 cohort)",520,260,True)
        bars.set_y_axis({"min":0,"max":1.2,"major_unit":.2,"num_format":"0%","num_font":{"size":8},"major_gridlines":{"visible":True,"line":{"color":"DDE3E5"}}})
        self.insert(sheet,offset+23,1,bars)
        # Reference layout: newest month first in the table, chronological charts.
        spans=[(14,18),(19,20)]+[(i,i+1) for i in range(21,32,2)]
        # Six compact month columns fit in K..P-like pairs by using single columns.
        spans=[(14,18),(19,20),(21,22),(23,24),(25,26),(27,28),(29,30),(31,31)]
        headers=["Monthly flow in EUR","Monthly average"]+list(reversed(monthly["rows"]))
        for j,span in enumerate(spans):
            label=headers[j] if j<2 else headers[j][0]
            self.merged(sheet,offset+7,span[0],offset+9,span[1],label,"header")
        # Last month has one narrow column: expand page right edge locally via widths.
        self.sheets[sheet].set_column(31,31,10.6)
        totalrow=offset+10
        self.merged(sheet,totalrow,14,totalrow,18,"Total available","total")
        avgrefs=[f'T{totalrow+2+i}' for i in range(len(group["countries"]))]
        self.merged(sheet,totalrow,19,totalrow,20,group["metrics"]["monthly_average"],"total",self.arithmetic(avgrefs,f'SUM({",".join(avgrefs)})'))
        for j,idx in enumerate(reversed(range(6)),2):
            refs=[self.href(monthly,idx,k+1) for k in range(len(group["countries"]))]
            val=group["monthly"][idx]["amount_eur"]
            self.merged(sheet,totalrow,spans[j][0],totalrow,spans[j][1],val,"total",self.arithmetic(refs,f'SUM({",".join(refs)})'))
        for i,country in enumerate(group["countries"]):
            r=totalrow+1+i
            self.merged(sheet,r,14,r,18,country["name"],"body")
            self.merged(sheet,r,19,r,20,country["metrics"]["monthly_average"],"num",self.metric(self.period,country["id"],"W"))
            for j,idx in enumerate(reversed(range(6)),2):
                ref=self.href(monthly,idx,i+1)
                self.merged(sheet,r,spans[j][0],r,spans[j][1],country["monthly"][idx]["amount_eur"],"num",f'=IF(ISNUMBER({ref}),{ref},"{NA}")')
        tablesrow=max(totalrow+len(group["countries"])+3,offset+20)
        for start,title,indices in [(14,"Efficiency in amount",(3,4)),(23,"Efficiency in number",(1,2))]:
            self.merged(sheet,tablesrow,start,tablesrow,start+8,title,"header")
            localspans=[(start,start+3),(start+4,start+5),(start+6,start+8)]
            for span,label in zip(localspans,["Country",self.data["previous_period"],self.period]):
                self.merged(sheet,tablesrow+1,span[0],tablesrow+2,span[1],label,"header")
            for i,country in enumerate(group["countries"]):
                r=tablesrow+3+i
                self.merged(sheet,r,start,r,start+3,country["name"],"body")
                for span,index in zip(localspans[1:],indices):
                    val=eff["rows"][i][index]
                    self.merged(sheet,r,span[0],r,span[1],val,"pct",None if val is None else "="+self.href(eff,i,index))
        notesrow=tablesrow+len(group["countries"])+4
        if notesrow<=offset+height-4:
            text="Previous-quarter series are omitted until that snapshot is imported." if not prior_available else "Compare the same M-5 cohort maturity across reporting dates."
            self.merged(sheet,notesrow,14,min(notesrow+2,offset+height-3),31,text,"note")

    def build(self):
        try:
            self.rates();self.stock();self.flows();self.distribution();self.consolidation();self.controls();self.guide()
            self.presentation_inputs()
            self.stock_page(self.data["global"],"Global")
            self.flow_page(self.data["global"],"Global",44)
            self.sheets["Global"].set_h_pagebreaks([44])
            for group in self.data["groups"]:
                if group["id"] not in self.group_sheets:
                    continue
                stock,flows=self.group_sheets[group["id"]]
                self.stock_page(group,stock,noncore=group["id"]=="noncore")
                if flows:
                    self.flow_page(group,flows)
            for country in self.data["countries"]:
                self.country_page(country)
            guide=self.sheets["Mode emploi"]
            guide.write("C5","Les pages détaillées présentent le trimestre généré. Relancer Python avec --report detailed ou both pour le trimestre suivant.",self.formats["text"])
            guide.write("C6","Les vues Global/Core/Non Core couvrent uniquement les pays configurés et importés.",self.formats["text"])
            self.sheets["Global"].activate()
        finally:
            self.book.close()


def write_detailed_workbook(path: Path, history: dict, period: str, presentation_config: dict | None = None) -> None:
    DetailedReport(path,history,period,presentation_config or {"schema_version":1}).build()
