"""Rebuild the Excel dashboard from a deterministic synthetic portfolio dataset."""
from pathlib import Path
import random
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment

OUT = Path("Community_Engagement_Dashboard.xlsx")
DEPARTMENTS = ["Information Science", "Environmental Sciences", "Business", "Health Sciences", "Education"]
TYPES = ["Health Initiative", "Community Research", "Service Learning", "Technology Program", "Workforce Development"]
IMPACTS = ["Technology Access", "Health and Wellness", "Education", "Economic Development", "Environment"]

def distribute(rows, key, target):
    rows[-1][key] += target - sum(row[key] for row in rows)

def build_rows():
    rng = random.Random(2027)
    rows=[]
    for i in range(100):
        rows.append({"id":f"P{i+1:03d}","department":DEPARTMENTS[i%5],"type":TYPES[(i*3)%5],"impact":IMPACTS[(i*2)%5],"participants":rng.randint(100,399),"volunteers":rng.randint(10,69),"funding":rng.randint(50000,199999),"satisfaction":round(rng.uniform(3.5,4.5),1)})
    distribute(rows,"participants",25363); distribute(rows,"volunteers",3736); distribute(rows,"funding",12478642)
    rows[-1]["satisfaction"] = round(rows[-1]["satisfaction"] + 406 - sum(x["satisfaction"] for x in rows),2)
    return rows

def main():
    rows=build_rows(); wb=Workbook(); dash=wb.active; dash.title="Dashboard"; src=wb.create_sheet("Partnerships")
    headers=["Partnership ID","Department","Partnership Type","Impact Area","Participants","Student Volunteers","Funding (USD)","Satisfaction (1-5)"]
    src.append(headers)
    for x in rows: src.append([x["id"],x["department"],x["type"],x["impact"],x["participants"],x["volunteers"],x["funding"],x["satisfaction"]])
    for cell in src[1]: cell.fill=PatternFill("solid",fgColor="17365D"); cell.font=Font(color="FFFFFF",bold=True); cell.alignment=Alignment(horizontal="center")
    src.freeze_panes="A2"; src.auto_filter.ref="A1:H101"
    widths=[15,24,24,22,14,19,16,18]
    for i,w in enumerate(widths,1): src.column_dimensions[chr(64+i)].width=w
    for c in src["G"][1:]: c.number_format='$#,##0';
    for c in src["H"][1:]: c.number_format='0.00'
    dash["A2"]="Community Engagement and University Partnerships Dashboard"; dash["A2"].font=Font(size=16,bold=True,color="17365D")
    kpis=[("A4","Partnerships","=COUNTA(Partnerships!A2:A101)"),("D4","Participants","=SUM(Partnerships!E2:E101)"),("G4","Student volunteers","=SUM(Partnerships!F2:F101)"),("A7","Total funding","=SUM(Partnerships!G2:G101)"),("D7","Average satisfaction","=AVERAGE(Partnerships!H2:H101)")]
    for cell,label,formula in kpis: dash[cell]=label; dash[cell].font=Font(bold=True,color="17365D"); value=dash.cell(dash[cell].row+1,dash[cell].column); value.value=formula; value.font=Font(size=15,bold=True,color="17365D")
    dash["A8"].number_format='$#,##0'; dash["D8"].number_format='0.00'
    dash.append([])
    dash["A11"]="Department"; dash["B11"]="Funding (USD)"
    for i,d in enumerate(DEPARTMENTS,12): dash[f"A{i}"]=d; dash[f"B{i}"]=f'=SUMIF(Partnerships!$B$2:$B$101,A{i},Partnerships!$G$2:$G$101)'; dash[f"B{i}"].number_format='$#,##0'
    dash["D11"]="Impact area"; dash["E11"]="Participants"
    for i,v in enumerate(IMPACTS,12): dash[f"D{i}"]=v; dash[f"E{i}"]=f'=SUMIF(Partnerships!$D$2:$D$101,D{i},Partnerships!$E$2:$E$101)'
    for rng in ("A11:B11","D11:E11"):
        for row in dash[rng]:
            for cell in row: cell.fill=PatternFill("solid",fgColor="17365D"); cell.font=Font(color="FFFFFF",bold=True)
    c1=BarChart(); c1.title="Funding by department (USD)"; c1.add_data(Reference(dash,min_col=2,min_row=11,max_row=16),titles_from_data=True); c1.set_categories(Reference(dash,min_col=1,min_row=12,max_row=16)); c1.height=7; c1.width=11; dash.add_chart(c1,"A18")
    c2=BarChart(); c2.title="Participants by impact area"; c2.add_data(Reference(dash,min_col=5,min_row=11,max_row=16),titles_from_data=True); c2.set_categories(Reference(dash,min_col=4,min_row=12,max_row=16)); c2.height=7; c2.width=11; dash.add_chart(c2,"E18")
    for col in "ABCDEFGH": dash.column_dimensions[col].width=16
    dash.sheet_view.showGridLines=False; src.sheet_view.showGridLines=False
    wb.save(OUT); print(f"Wrote {OUT}")

if __name__ == "__main__": main()

