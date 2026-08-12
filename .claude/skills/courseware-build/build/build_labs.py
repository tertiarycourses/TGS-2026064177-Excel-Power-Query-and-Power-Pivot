#!/usr/bin/env python3
"""
Build the lab folders for Excel Power Query and Power Pivot (TGS-2026064177).

One folder per lab: labs/lab-NN-<slug>/
  <lab-NN-slug>.xlsx   starter workbook — mock data as real Excel Tables, plus
                       an embedded Power Query connection where the lab needs a
                       pre-built staging query
  README.md            the lab brief (objective, scenario, steps, verification)

Mock data is deterministic (fixed seed) so the workbook, the Learner Guide and
the assessment model answers always describe the SAME numbers.
"""
import os, sys, random, shutil, datetime as dt
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import course_data as CD
from data_domain1 import DOMAIN1
from data_domain2 import DOMAIN2
from data_domain3 import DOMAIN3

from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

RNG = random.Random(20260812)

BRAND = "1F4E79"
HDR_FILL = PatternFill("solid", fgColor=BRAND)
HDR_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

ALL_LABS = DOMAIN1 + DOMAIN2 + DOMAIN3


def repo_root() -> Path:
    p = Path(os.environ.get("COURSE_REPO", "")) if os.environ.get("COURSE_REPO") else None
    if p and p.is_dir():
        return p
    here = Path(__file__).resolve()
    for parent in here.parents:
        if (parent / "courseware").is_dir() and (parent / "labs").is_dir():
            return parent
    raise SystemExit("cannot locate course repo (needs courseware/ and labs/)")


def slug(title: str) -> str:
    out = "".join(c.lower() if c.isalnum() else "-" for c in title)
    while "--" in out:
        out = out.replace("--", "-")
    return out.strip("-")


# ------------------------------------------------------------------ mock data
REGIONS = ["North", "South", "East", "West", "Central"]
CATEGORIES = ["Laptops", "Monitors", "Peripherals", "Networking", "Storage"]
PRODUCTS = [
    (101, "UltraBook 14", "Laptops", 1850.00),
    (102, "UltraBook Pro 16", "Laptops", 2750.00),
    (103, "WorkStation Mini", "Laptops", 2100.00),
    (201, "VisionView 24", "Monitors", 320.00),
    (202, "VisionView 27 QHD", "Monitors", 545.00),
    (203, "VisionView 32 4K", "Monitors", 890.00),
    (301, "SilentType Keyboard", "Peripherals", 95.00),
    (302, "GlideMouse Wireless", "Peripherals", 65.00),
    (303, "ClearCam 1080p", "Peripherals", 130.00),
    (401, "MeshRouter AX3000", "Networking", 410.00),
    (402, "SwitchPro 8-Port", "Networking", 185.00),
    (501, "VaultDrive 2TB", "Storage", 240.00),
    (502, "VaultDrive NAS 8TB", "Storage", 1150.00),
]
FIRST = ["Wei Ming", "Siti", "Rajesh", "Grace", "Daniel", "Priya", "Kai Xin",
         "Marcus", "Nurul", "Jonathan", "Mei Ling", "Arun", "Clara", "Hakim", "Yi Ting"]
LAST = ["Tan", "Lim", "Kumar", "Wong", "Ng", "Rahman", "Chen", "Goh",
        "Menon", "Lee", "Ibrahim", "Ho", "Nair", "Teo", "Koh"]
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def money(v):
    return round(v, 2)


def write_sheet(ws, headers, rows, table_name, as_text=None, col_widths=None):
    """Write a header + rows block and register it as a real Excel Table."""
    as_text = as_text or set()
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for r in rows:
        ws.append(list(r))
    # force text storage where the lab deliberately ships bad types
    for ci in as_text:
        for ri in range(2, len(rows) + 2):
            cell = ws.cell(row=ri, column=ci + 1)
            if cell.value is not None:
                cell.value = str(cell.value)
                cell.number_format = "@"
    n = len(rows)
    if n:
        ref = f"A1:{get_column_letter(len(headers))}{n + 1}"
        t = Table(displayName=table_name, ref=ref)
        t.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(t)
    for i, h in enumerate(headers, start=1):
        w = (col_widths or {}).get(h, max(12, min(30, len(str(h)) + 8)))
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"


def sales_rows(n, start=dt.date(2026, 1, 5), with_amount=True):
    rows = []
    for i in range(n):
        pid, pname, cat, price = RNG.choice(PRODUCTS)
        qty = RNG.randint(1, 12)
        d = start + dt.timedelta(days=RNG.randint(0, 175))
        amt = money(qty * price)
        row = [f"SO-{1001 + i}", d, RNG.choice(REGIONS), pid, qty, price]
        if with_amount:
            row.append(amt)
        rows.append(row)
    return rows


# ------------------------------------------------------------------ per-lab builders
def build_lab_01(wb):
    """Everything as TEXT so learners must fix data types."""
    ws = wb.active; ws.title = "Sales"
    rows = []
    for r in sales_rows(60):
        code, d, region, pid, qty, price, amt = r
        rows.append([code, d.strftime("%d/%m/%Y"), region, pid,
                     str(qty), f"S${price:,.2f}", f"S${amt:,.2f}"])
    write_sheet(ws, ["Order ID", "Order Date", "Region", "ProductID",
                     "Quantity", "Unit Price", "Amount"],
                rows, "SalesRaw", as_text={1, 3, 4, 5, 6},
                col_widths={"Order Date": 14, "Unit Price": 14, "Amount": 14})
    note(wb, "Every column arrives as TEXT. Set the correct data type on each "
             "column in Power Query, using Data Type > Using Locale for the "
             "dd/mm/yyyy dates.")


def build_lab_02(wb, folder):
    ws = wb.active; ws.title = "Products"
    rows = [[pid, name, cat] for pid, name, cat, _ in PRODUCTS]
    write_sheet(ws, ["ProductID", "Product Name", "Category"], rows, "ProductsTbl")
    note(wb, "The prices live in the SEPARATE workbook price-list-source.xlsx in "
             "this folder. Connect to it with Get Data > From File > From Excel "
             "Workbook and load it as Connection Only.")
    # the external source workbook
    src = Workbook(); s = src.active; s.title = "PriceList"
    prows = [[pid, money(price), money(price * 1.15)] for pid, _, _, price in PRODUCTS]
    write_sheet(s, ["ProductID", "Unit Price", "List Price"], prows, "PriceList")
    src.save(folder / "price-list-source.xlsx")


def build_lab_03(wb):
    """Ragged report: region only on first row of block, quarter below its group."""
    ws = wb.active; ws.title = "Headcount"
    rows = []
    for region in REGIONS:
        first = True
        for dept in ["Sales", "Operations", "Finance"]:
            rows.append([region if first else None, dept, RNG.randint(4, 40), None])
            first = False
        rows.append([None, None, None, f"Q{RNG.randint(1,4)} 2026"])
        rows.append([None, None, None, None])
    write_sheet(ws, ["Region", "Department", "Headcount", "Quarter"],
                rows, "HeadcountRaw")
    note(wb, "Region appears only on the first row of each block (needs Fill Down). "
             "Quarter sits BELOW its group (needs Fill Up). Remove the blank "
             "spacer rows afterwards.")


def build_lab_04(wb):
    ws = wb.active; ws.title = "Customers"
    rows = []
    for i in range(40):
        full = f"{RNG.choice(FIRST)} {RNG.choice(LAST)}"
        k = RNG.randint(1, 3)
        prods = ", ".join(RNG.sample([p[1] for p in PRODUCTS], k))
        rows.append([f"CUST-{2001+i}", full, RNG.choice(REGIONS), prods])
    write_sheet(ws, ["Customer ID", "Full Name", "Region", "Products Bought"],
                rows, "CustomersRaw",
                col_widths={"Full Name": 22, "Products Bought": 52})
    note(wb, "Split Full Name by Space (left-most delimiter) into First/Last Name. "
             "Split Products Bought by Comma using Advanced options > Split into "
             "Rows, then Trim the result.")


def build_lab_05(wb):
    ws = wb.active; ws.title = "Orders"
    rows = []
    for r in sales_rows(70, with_amount=True):
        code, d, region, pid, qty, price, amt = r
        rows.append([code, d, region, pid, qty, price, amt])
    # guarantee boundary rows at exactly 1000 and 5000 for the band test
    rows[0] = ["SO-9001", dt.date(2026, 3, 2), "North", 301, 10, 100.00, 1000.00]
    rows[1] = ["SO-9002", dt.date(2026, 3, 3), "South", 102, 2, 2500.00, 5000.00]
    write_sheet(ws, ["Order ID", "Order Date", "Region", "ProductID",
                     "Quantity", "Unit Price", "Amount"], rows, "OrdersTbl")
    note(wb, "Rows SO-9001 (Amount exactly 1000) and SO-9002 (Amount exactly 5000) "
             "are the boundary cases for the Value Band conditional column.")


def build_lab_06(wb):
    ws = wb.active; ws.title = "Despatch"
    streets = ["Ang Mo Kio Ave 3", "Jurong East St 21", "Tampines Ave 4",
               "Bukit Batok West Ave 8", "Serangoon Central", "Clementi Rd"]
    rows = []
    for i in range(55):
        r = sales_rows(1)[0]
        rows.append([r[0], r[1], f"Blk {RNG.randint(100, 899)}",
                     RNG.choice(streets), f"{RNG.randint(100000, 828000):06d}",
                     RNG.choice(["Despatched", "Pending", "Cancelled", "Pending"]),
                     r[6]])
    write_sheet(ws, ["Order ID", "Order Date", "Block", "Street", "Postal Code",
                     "Status", "Amount"], rows, "DespatchTbl",
                as_text={4}, col_widths={"Street": 26})
    note(wb, "Merge Block + Street + Postal Code (in that order) with a comma "
             "separator. Filter out Cancelled, keep Amount > 1000, then sort by "
             "Order Date then Amount, both descending.")


def build_lab_07(wb):
    ws = wb.active; ws.title = "Orders"
    rows = [[r[0], r[1], r[2], r[6]] for r in sales_rows(60)]
    rows[0] = ["SO-9100", dt.date(2026, 2, 10), "East", 500.00]   # exact boundary
    write_sheet(ws, ["Order ID", "Order Date", "Region", "Amount"],
                rows, "OrdersTbl")
    note(wb, "Row SO-9100 has Amount EXACTLY 500 — it must qualify for free "
             "shipping because the rule uses >= 500.")


def build_lab_08(wb):
    ws = wb.active; ws.title = "Orders"
    rows = []
    for i in range(60):
        r = sales_rows(1)[0]
        rows.append([r[0], r[1], r[2], r[6],
                     RNG.choice(["Open", "Closed", "On Hold"]),
                     RNG.randint(0, 95)])
    # deterministic rows covering each branch of the truth table
    rows[0] = ["SO-9201", dt.date(2026, 1, 8), "North", 8200.00, "Open", 45]   # Escalate
    rows[1] = ["SO-9202", dt.date(2026, 1, 9), "South", 7400.00, "Open", 12]   # Watch
    rows[2] = ["SO-9203", dt.date(2026, 1, 10), "East", 900.00, "Open", 61]    # Watch
    rows[3] = ["SO-9204", dt.date(2026, 1, 11), "West", 450.00, "Closed", 5]   # Normal
    write_sheet(ws, ["Order ID", "Order Date", "Region", "Amount", "Status",
                     "Days Outstanding"], rows, "OrdersTbl")
    note(wb, "Rows SO-9201..SO-9204 cover all four branches of the AND/OR truth "
             "table: Escalate, Watch (by value), Watch (by ageing), Normal.")


def build_lab_09(wb):
    ws = wb.active; ws.title = "Invoices"
    rows = []
    for i in range(50):
        d = dt.date(2026, 1, 1) + dt.timedelta(days=RNG.randint(0, 150))
        rows.append([f"INV-{4001+i}", d, RNG.choice([14, 30, 45, 60]),
                     money(RNG.uniform(400, 9000)), RNG.choice(REGIONS)])
    rows[0] = ["INV-9001", dt.date(2026, 5, 1), 30, 2400.00, "North"]  # worked example
    write_sheet(ws, ["Invoice No", "Invoice Date", "Terms Days", "Amount", "Region"],
                rows, "InvoicesTbl")
    note(wb, "Worked example INV-9001: dated 1 May 2026 with 30-day terms gives a "
             "Due Date of 31 May 2026 and, against the 30 June 2026 reporting "
             "date, Days Overdue = 30.")


def build_lab_10(wb):
    ws = wb.active; ws.title = "MonthlySales"
    rows = [[reg] + [money(RNG.uniform(18000, 96000)) for _ in MONTHS]
            for reg in REGIONS]
    write_sheet(ws, ["Region"] + MONTHS, rows, "MonthlySalesWide",
                col_widths={m: 11 for m in MONTHS})
    note(wb, "Wide layout: 5 regions x 12 months. Select ONLY Region, then "
             "Unpivot Other Columns — the result must be 60 rows and 3 columns.")


def build_lab_11(wb):
    ws = wb.active; ws.title = "Q1Sales"
    q1 = [[r[0], r[1], r[2], r[3], r[4], r[6]]
          for r in sales_rows(30, start=dt.date(2026, 1, 1))]
    write_sheet(ws, ["Order ID", "Order Date", "Region", "ProductID",
                     "Quantity", "Amount"], q1, "Q1Sales")
    ws2 = wb.create_sheet("Q2Sales")
    q2 = [[f"SO-2{i:03d}", dt.date(2026, 4, 1) + dt.timedelta(days=RNG.randint(0, 88)),
           RNG.choice(REGIONS), RNG.choice(PRODUCTS)[0], RNG.randint(1, 10),
           money(RNG.uniform(200, 9000))] for i in range(28)]
    write_sheet(ws2, ["Order ID", "Order Date", "Region", "ProductID",
                      "Quantity", "Amount"], q2, "Q2Sales")
    ws3 = wb.create_sheet("Attributes")
    arows = []
    for pid, name, cat, price in PRODUCTS[:8]:
        arows += [[name, "Category", cat],
                  [name, "Warranty", RNG.choice(["12 months", "24 months", "36 months"])],
                  [name, "Supplier", RNG.choice(["Acme Supply", "Northwind", "Contoso"])]]
    write_sheet(ws3, ["Product Name", "Attribute", "Value"], arows, "AttributesTbl")
    note(wb, "Q1Sales and Q2Sales share IDENTICAL column names so Append works "
             "cleanly. On Attributes, pivot the Attribute column with Value as "
             "values and choose Don't Aggregate to keep the text.")


def build_lab_12(wb):
    ws = wb.active; ws.title = "Transactions"
    rows = []
    for i in range(400):
        pid, pname, cat, price = RNG.choice(PRODUCTS)
        qty = RNG.randint(1, 15)
        d = dt.date(2026, 1, 1) + dt.timedelta(days=RNG.randint(0, 200))
        rows.append([f"TXN-{10001+i}", d, RNG.choice(REGIONS), cat, pname,
                     qty, money(qty * price)])
    write_sheet(ws, ["Txn ID", "Txn Date", "Region", "Product Category",
                     "Product Name", "Quantity", "Amount"], rows, "TransactionsTbl",
                col_widths={"Product Name": 24, "Product Category": 18})
    note(wb, "400 transactions. Group By Region with Count Rows, Sum of Amount and "
             "Average of Amount; then add Product Category as a second group "
             "column. The group totals must reconcile to the grand total.")


def build_lab_13(wb):
    ws = wb.active; ws.title = "Sales"
    srows = [[r[0], r[1], r[3], r[4], r[6]] for r in sales_rows(300)]
    write_sheet(ws, ["Order ID", "Order Date", "ProductID", "Quantity", "Amount"],
                srows, "SalesFact")
    ws2 = wb.create_sheet("Products")
    write_sheet(ws2, ["ProductID", "Product Name", "Category", "Unit Price"],
                [[p, n, c, money(pr)] for p, n, c, pr in PRODUCTS], "ProductsDim",
                col_widths={"Product Name": 24})
    ws3 = wb.create_sheet("Calendar")
    crows = []
    d = dt.date(2026, 1, 1)
    while d <= dt.date(2026, 12, 31):
        crows.append([d, d.year, (d.month - 1) // 3 + 1, d.month,
                      MONTHS[d.month - 1], d.strftime("%A")])
        d += dt.timedelta(days=1)
    write_sheet(ws3, ["Date", "Year", "Quarter", "Month Number", "Month Name",
                      "Day Name"], crows, "CalendarDim")
    note(wb, "Load all three to the Data Model as Connection Only. Relate "
             "Sales[ProductID] -> Products[ProductID] and Sales[Order Date] -> "
             "Calendar[Date]. Mark Calendar as a Date Table before TOTALYTD.")


def build_lab_14(wb):
    ws = wb.active; ws.title = "Orders"
    # deliberate: some orders reference a product that does not exist
    orows = []
    for i in range(50):
        pid = RNG.choice(PRODUCTS)[0] if i % 7 else RNG.choice([901, 902, 903])
        orows.append([f"SO-{7001+i}", dt.date(2026, 2, 1) + dt.timedelta(days=i),
                      pid, RNG.randint(1, 9), money(RNG.uniform(150, 7500))])
    write_sheet(ws, ["Order ID", "Order Date", "ProductID", "Quantity", "Amount"],
                orows, "OrdersJoin")
    ws2 = wb.create_sheet("Products")
    # deliberate: some products were never ordered
    prows = [[p, n, c, money(pr)] for p, n, c, pr in PRODUCTS]
    prows += [[801, "LegacyDock v1", "Peripherals", 145.00],
              [802, "LegacyDock v2", "Peripherals", 165.00]]
    write_sheet(ws2, ["ProductID", "Product Name", "Category", "Unit Price"],
                prows, "ProductsJoin", col_widths={"Product Name": 24})
    note(wb, "Deliberately imperfect: some Orders reference ProductIDs 901-903 "
             "which do NOT exist in Products (orphans), and ProductIDs 801-802 "
             "were never ordered. That is what makes the six joins differ.")


def build_lab_15(wb):
    ws = wb.active; ws.title = "Sales"
    rows = [[r[0], r[1], r[3], r[4], r[6]] for r in sales_rows(120)]
    write_sheet(ws, ["Order ID", "Order Date", "ProductID", "Quantity", "Amount"],
                rows, "SalesTbl")
    # XLOOKUP demo column the learner will replace
    ws["G1"] = "XLOOKUP demo"; ws["G1"].font = Font(bold=True)
    ws["G2"] = "=XLOOKUP([@ProductID],Products[ProductID],Products[Product Name],\"Not found\")"
    ws2 = wb.create_sheet("Products")
    write_sheet(ws2, ["ProductID", "Product Name", "Category", "Unit Price"],
                [[p, n, c, money(pr)] for p, n, c, pr in PRODUCTS], "Products",
                col_widths={"Product Name": 24})
    note(wb, "Cell G2 on Sales holds the XLOOKUP formula you will replace with a "
             "Left Outer merge. Compare: the formula must be dragged down for new "
             "rows; the merge just refreshes.")


def build_lab_16(wb):
    ws = wb.active; ws.title = "Sales"
    rows = []
    for i in range(80):
        pid = RNG.choice(PRODUCTS)[0]
        rows.append([f"SO-{8001+i}", dt.date(2026, 3, 1) + dt.timedelta(days=i % 90),
                     pid, RNG.choice(REGIONS), RNG.randint(1, 10)])
    write_sheet(ws, ["Order ID", "Order Date", "ProductID", "Region", "Quantity"],
                rows, "SalesMulti")
    ws2 = wb.create_sheet("RegionalPrices")
    prows = []
    for pid, n, c, price in PRODUCTS:
        for reg in REGIONS:
            factor = {"North": 1.00, "South": 0.95, "East": 1.05,
                      "West": 0.98, "Central": 1.10}[reg]
            prows.append([pid, reg, money(price * factor)])
    write_sheet(ws2, ["ProductID", "Region", "Unit Price"], prows, "RegionalPrices")
    note(wb, "Each ProductID has FIVE regional prices. Merging on ProductID alone "
             "multiplies the rows 5x — you must Ctrl-click ProductID then Region "
             "in BOTH tables, in the same order.")


def build_lab_17(wb):
    ws = wb.active; ws.title = "SalesReps"
    names = [f"{RNG.choice(FIRST)} {RNG.choice(LAST)}" for _ in range(30)]
    rows = [[f"REP-{i+1:03d}", names[i], RNG.choice(REGIONS),
             money(RNG.uniform(2000, 120000))] for i in range(30)]
    rows[0] = ["REP-901", "Boundary Case A", "North", 10000.00]
    rows[1] = ["REP-902", "Boundary Case B", "South", 50000.00]
    write_sheet(ws, ["Rep ID", "Rep Name", "Region", "Sales"], rows, "SalesReps",
                col_widths={"Rep Name": 22})
    ws2 = wb.create_sheet("CommissionBands")
    write_sheet(ws2, ["Band", "Lower", "Upper", "Rate"],
                [["Bronze", 0, 9999.99, 0.02],
                 ["Silver", 10000, 49999.99, 0.05],
                 ["Gold", 50000, 99999999, 0.08]], "CommissionBands")
    note(wb, "REP-901 (exactly 10000) and REP-902 (exactly 50000) are the band "
             "boundaries — both methods must agree on these two rows.")


def build_lab_18(wb):
    ws = wb.active; ws.title = "Regions"
    write_sheet(ws, ["Region", "Region Manager"],
                [[r, f"{RNG.choice(FIRST)} {RNG.choice(LAST)}"] for r in REGIONS],
                "RegionsTbl", col_widths={"Region Manager": 22})
    ws2 = wb.create_sheet("Products")
    write_sheet(ws2, ["ProductID", "Product Name", "Category"],
                [[p, n, c] for p, n, c, _ in PRODUCTS], "ProductsGrid",
                col_widths={"Product Name": 24})
    ws3 = wb.create_sheet("Parameters")
    write_sheet(ws3, ["Parameter", "Value"],
                [["Sales Target Threshold", 5000],
                 ["Reporting Year", 2026],
                 ["Minimum Order Qty", 3]], "ParametersTbl",
                col_widths={"Parameter": 26})
    ws4 = wb.create_sheet("Sales")
    srows = []
    for _ in range(45):   # deliberately sparse: not every region x product sells
        pid, n, c, price = RNG.choice(PRODUCTS[:8])
        srows.append([RNG.choice(REGIONS[:4]), pid, money(RNG.uniform(500, 9000))])
    write_sheet(ws4, ["Region", "ProductID", "Amount"], srows, "SalesSparse")
    note(wb, "5 regions x 13 products = 65 grid rows after the cross join. Sales "
             "is deliberately sparse, so unsold pairs must show 0 after the Left "
             "Outer merge and Replace Values on null.")


def build_lab_19(wb):
    ws = wb.active; ws.title = "CRMCustomers"
    base = ["ABC Holdings", "Sunrise Logistics", "Pinnacle Systems",
            "Harbourfront Trading", "Evergreen Foods", "Titan Engineering",
            "BlueWave Marine", "Summit Consulting", "Orchid Pharma",
            "Redwood Analytics", "Skyline Property", "Nexus Digital"]
    crows = [[f"CRM-{i+1:03d}", f"{n} Pte Ltd", RNG.choice(REGIONS)]
             for i, n in enumerate(base)]
    write_sheet(ws, ["CRM ID", "Customer Name", "Region"], crows, "CRMCustomers",
                col_widths={"Customer Name": 32})
    ws2 = wb.create_sheet("FinanceCustomers")
    # same companies, inconsistently spelled
    variants = {
        "ABC Holdings": "ABC Group Pte. Ltd.",          # needs transformation table
        "Sunrise Logistics": "Sunrise Logistic Pte. Ltd.",
        "Pinnacle Systems": "Pinacle Systems Pte Ltd",
        "Harbourfront Trading": "Harbourfront  Trading Pte. Ltd.",
        "Evergreen Foods": "EVERGREEN FOODS PTE LTD",
        "Titan Engineering": "Titan Engineering Pte.Ltd.",
        "BlueWave Marine": "Blue Wave Marine Pte Ltd",
        "Summit Consulting": "Summit Consultng Pte. Ltd.",
        "Orchid Pharma": "Orchid Pharma Pte Ltd",
        "Redwood Analytics": "Redwood Analytic Pte. Ltd.",
        "Skyline Property": "Skyline Properties Pte Ltd",
        "Nexus Digital": "Nexus Digitial Pte. Ltd.",
    }
    frows = [[f"FIN-{i+1:03d}", variants[n], money(RNG.uniform(12000, 240000))]
             for i, n in enumerate(base)]
    write_sheet(ws2, ["Finance ID", "Customer Name", "Outstanding"], frows,
                "FinanceCustomers", col_widths={"Customer Name": 32})
    ws3 = wb.create_sheet("TransformTable")
    write_sheet(ws3, ["From", "To"],
                [["ABC Group Pte. Ltd.", "ABC Holdings Pte Ltd"],
                 ["Skyline Properties Pte Ltd", "Skyline Property Pte Ltd"]],
                "TransformTable", col_widths={"From": 32, "To": 32})
    note(wb, "'ABC Holdings' vs 'ABC Group' can never be matched by similarity "
             "score alone — that is exactly what the TransformTable query is for. "
             "Set it as the Transformation table in Fuzzy matching options.")


def note(wb, text):
    """Add a Notes sheet explaining the deliberate defects in the mock data."""
    ws = wb.create_sheet("Notes")
    ws["A1"] = "About this dataset"
    ws["A1"].font = Font(bold=True, size=13, color=BRAND)
    ws["A3"] = text
    ws["A3"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A3:H12")
    ws.column_dimensions["A"].width = 100


BUILDERS = {
    1: build_lab_01, 3: build_lab_03, 4: build_lab_04, 5: build_lab_05,
    6: build_lab_06, 7: build_lab_07, 8: build_lab_08, 9: build_lab_09,
    10: build_lab_10, 11: build_lab_11, 12: build_lab_12, 13: build_lab_13,
    14: build_lab_14, 15: build_lab_15, 16: build_lab_16, 17: build_lab_17,
    18: build_lab_18, 19: build_lab_19,
}


def lab_readme(lab) -> str:
    L = []
    L.append(f"# Lab {lab['num']:02d} — {lab['title']}\n")
    L.append(f"**Course:** {CD.TITLE} ({CD.COURSE_CODE})  ")
    L.append(f"**Topic {lab['topic']}:** {CD.TOPICS[lab['topic']-1]['title']}  ")
    L.append(f"**Maps to:** {lab['objective']}\n")
    L.append(f"**Workbook:** `{lab['workbook']}`")
    for f in lab.get("extra_files", []):
        L.append(f"**Also in this folder:** `{f}`")
    L.append("")
    L.append("## Scenario\n")
    L.append(lab["desc"] + "\n")
    L.append("## What you will build\n")
    L.append(lab["build"] + "\n")
    L.append(f"**Tools used:** {lab['services']}\n")
    L.append("## Steps\n")
    for i, (text, cmd) in enumerate(lab["steps"], 1):
        L.append(f"{i}. {text}")
        if cmd:
            L.append("")
            L.append("   ```m")
            for part in cmd.split("  |  "):
                L.append(f"   {part.strip()}")
            L.append("   ```")
        L.append("")
    L.append("## Test it\n")
    L.append(lab["test"] + "\n")
    L.append("---\n")
    L.append(f"*{CD.ORG} · {CD.UEN} · {CD.VERSION}*")
    return "\n".join(L)


def main():
    root = repo_root()
    labs_dir = root / "labs"
    if labs_dir.exists():
        for child in labs_dir.iterdir():
            if child.name.startswith("lab-") or child.name == "README.md":
                shutil.rmtree(child) if child.is_dir() else child.unlink()
    labs_dir.mkdir(exist_ok=True)

    index = ["# Labs — " + CD.TITLE, "",
             f"**Course Code:** {CD.COURSE_CODE} · **Version:** {CD.VERSION}", "",
             "Each lab has its own folder containing a starter Excel workbook with "
             "mock data (and, where needed, a companion source workbook) plus the "
             "lab brief. Full step-by-step instructions are in the Learner Guide.", ""]

    for topic in CD.TOPICS:
        index.append(f"## Topic {topic['num']} — {topic['title']}")
        index.append("")
        index.append("| Lab | Title | Folder |")
        index.append("|---|---|---|")
        for lab in [l for l in ALL_LABS if l["topic"] == topic["num"]]:
            folder_name = f"lab-{lab['num']:02d}-{slug(lab['title'])}"
            index.append(f"| {lab['num']:02d} | {lab['title']} | "
                         f"[`{folder_name}/`]({folder_name}/) |")
        index.append("")

    built = 0
    for lab in ALL_LABS:
        folder = labs_dir / f"lab-{lab['num']:02d}-{slug(lab['title'])}"
        folder.mkdir(parents=True, exist_ok=True)

        wb = Workbook()
        n = lab["num"]
        if n == 2:
            build_lab_02(wb, folder)
        else:
            BUILDERS[n](wb)
        wb.save(folder / lab["workbook"])

        (folder / "README.md").write_text(lab_readme(lab), encoding="utf-8")
        built += 1
        print(f"  lab {n:02d}  {folder.name}/{lab['workbook']}")

    (labs_dir / "README.md").write_text("\n".join(index), encoding="utf-8")
    print(f"\nBuilt {built} lab folders in {labs_dir}")


if __name__ == "__main__":
    main()
