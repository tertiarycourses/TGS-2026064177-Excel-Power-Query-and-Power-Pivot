#!/usr/bin/env python3
"""
Inject real Power Query (M) queries into the lab starter workbooks.

openpyxl cannot write Power Query, so this post-processes the saved .xlsx
(a zip of OOXML parts) and adds the Data Mashup structures Excel reads:

  customXml/item1.xml          the DataMashup payload (base64 zip holding
                               Formulas/Section1.m — the actual M code)
  customXml/itemProps1.xml     the Microsoft.DataMashup schema ref
  xl/connections.xml           one WORKBOOK QUERY connection per query
  xl/queryTables/...           not needed for connection-only queries

Every lab ships with at least a `Source` staging query already written, so the
learner opens Data > Queries & Connections and sees a real, editable query. The
labs that build on a prior query (02, 11, 13, 14, 15, 16, 17, 18, 19) get their
supporting queries pre-built too.
"""
import base64, io, os, re, shutil, sys, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import course_data as CD
from data_domain1 import DOMAIN1
from data_domain2 import DOMAIN2
from data_domain3 import DOMAIN3

ALL_LABS = DOMAIN1 + DOMAIN2 + DOMAIN3


def slug(title: str) -> str:
    out = "".join(c.lower() if c.isalnum() else "-" for c in title)
    while "--" in out:
        out = out.replace("--", "-")
    return out.strip("-")


def repo_root() -> Path:
    if os.environ.get("COURSE_REPO"):
        return Path(os.environ["COURSE_REPO"])
    here = Path(__file__).resolve()
    for parent in here.parents:
        if (parent / "courseware").is_dir() and (parent / "labs").is_dir():
            return parent
    raise SystemExit("cannot locate course repo")


# ------------------------------------------------------------------ M builders
def m_from_table(tbl: str) -> str:
    """A plain 'load this worksheet table' staging query."""
    return (
        f'let\n'
        f'    Source = Excel.CurrentWorkbook(){{[Name="{tbl}"]}}[Content]\n'
        f'in\n'
        f'    Source'
    )


def m_typed(tbl: str, types: list) -> str:
    tl = ", ".join(f'{{"{c}", {t}}}' for c, t in types)
    return (
        f'let\n'
        f'    Source = Excel.CurrentWorkbook(){{[Name="{tbl}"]}}[Content],\n'
        f'    #"Changed Type" = Table.TransformColumnTypes(Source,{{{tl}}})\n'
        f'in\n'
        f'    #"Changed Type"'
    )


T_TXT, T_INT, T_NUM, T_DATE = "type text", "Int64.Type", "type number", "type date"

# queries pre-built into each lab workbook: {lab_num: [(query_name, m_code)]}
QUERIES = {
    1: [("Sales", m_from_table("SalesRaw"))],
    2: [("Products", m_typed("ProductsTbl",
                             [("ProductID", T_INT), ("Product Name", T_TXT),
                              ("Category", T_TXT)]))],
    3: [("Headcount", m_from_table("HeadcountRaw"))],
    4: [("Customers", m_from_table("CustomersRaw"))],
    5: [("Orders", m_typed("OrdersTbl",
                           [("Order ID", T_TXT), ("Order Date", T_DATE),
                            ("Region", T_TXT), ("ProductID", T_INT),
                            ("Quantity", T_INT), ("Unit Price", T_NUM),
                            ("Amount", T_NUM)]))],
    6: [("Despatch", m_from_table("DespatchTbl"))],
    7: [("Orders", m_typed("OrdersTbl",
                           [("Order ID", T_TXT), ("Order Date", T_DATE),
                            ("Region", T_TXT), ("Amount", T_NUM)]))],
    8: [("Orders", m_typed("OrdersTbl",
                           [("Order ID", T_TXT), ("Order Date", T_DATE),
                            ("Region", T_TXT), ("Amount", T_NUM),
                            ("Status", T_TXT), ("Days Outstanding", T_INT)]))],
    9: [("Invoices", m_typed("InvoicesTbl",
                             [("Invoice No", T_TXT), ("Invoice Date", T_DATE),
                              ("Terms Days", T_INT), ("Amount", T_NUM),
                              ("Region", T_TXT)]))],
    10: [("MonthlySales", m_from_table("MonthlySalesWide"))],
    11: [("Q1Sales", m_from_table("Q1Sales")),
         ("Q2Sales", m_from_table("Q2Sales")),
         ("Attributes", m_from_table("AttributesTbl"))],
    12: [("Transactions", m_typed("TransactionsTbl",
                                  [("Txn ID", T_TXT), ("Txn Date", T_DATE),
                                   ("Region", T_TXT), ("Product Category", T_TXT),
                                   ("Product Name", T_TXT), ("Quantity", T_INT),
                                   ("Amount", T_NUM)]))],
    13: [("Sales", m_typed("SalesFact",
                           [("Order ID", T_TXT), ("Order Date", T_DATE),
                            ("ProductID", T_INT), ("Quantity", T_INT),
                            ("Amount", T_NUM)])),
         ("Products", m_typed("ProductsDim",
                              [("ProductID", T_INT), ("Product Name", T_TXT),
                               ("Category", T_TXT), ("Unit Price", T_NUM)])),
         ("Calendar", m_typed("CalendarDim",
                              [("Date", T_DATE), ("Year", T_INT),
                               ("Quarter", T_INT), ("Month Number", T_INT),
                               ("Month Name", T_TXT), ("Day Name", T_TXT)]))],
    14: [("Orders", m_from_table("OrdersJoin")),
         ("Products", m_from_table("ProductsJoin"))],
    15: [("Sales", m_from_table("SalesTbl")),
         ("Products", m_from_table("Products"))],
    16: [("Sales", m_from_table("SalesMulti")),
         ("RegionalPrices", m_from_table("RegionalPrices"))],
    17: [("SalesReps", m_from_table("SalesReps")),
         ("CommissionBands", m_from_table("CommissionBands"))],
    18: [("Regions", m_from_table("RegionsTbl")),
         ("Products", m_from_table("ProductsGrid")),
         ("Parameters", m_from_table("ParametersTbl")),
         ("Sales", m_from_table("SalesSparse"))],
    19: [("CRMCustomers", m_from_table("CRMCustomers")),
         ("FinanceCustomers", m_from_table("FinanceCustomers")),
         ("TransformTable", m_from_table("TransformTable"))],
}


def build_section1_m(queries) -> str:
    """The Section1.m document that holds every query in the workbook."""
    parts = ['section Section1;', '']
    for name, code in queries:
        safe = f'#"{name}"' if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", name) else name
        parts.append(f'shared {safe} = {code};')
        parts.append('')
    return "\n".join(parts)


def build_datamashup(queries) -> bytes:
    """
    The DataMashup part: a binary header + an embedded zip package whose
    Formulas/Section1.m carries the M code, then the permission/metadata
    sections Excel expects.
    """
    # 1. the inner package zip
    pkg = io.BytesIO()
    with zipfile.ZipFile(pkg, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0" encoding="utf-8"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="m" ContentType="application/x-ms-m"/>'
                   '<Override PartName="/Config/Package.xml" ContentType="application/xml"/>'
                   '</Types>')
        z.writestr("Config/Package.xml",
                   '<?xml version="1.0" encoding="utf-8"?>'
                   '<Package xmlns="http://schemas.microsoft.com/DataMashup">'
                   '<Version>2.0</Version><MinVersion>1.5</MinVersion>'
                   '<Culture>en-SG</Culture></Package>')
        z.writestr("Formulas/Section1.m",
                   build_section1_m(queries).encode("utf-8"))
    pkg_bytes = pkg.getvalue()

    permissions = (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<PermissionList xmlns:xsd="http://www.w3.org/2001/XMLSchema" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<CanEvaluateFuturePackages>false</CanEvaluateFuturePackages>'
        '<FirewallEnabled>true</FirewallEnabled>'
        '<WorkbookGroupType xsi:nil="true"/></PermissionList>'
    ).encode("utf-8")

    metadata = (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<LocalPackageMetadataFile xmlns:xsd="http://www.w3.org/2001/XMLSchema" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<Items><Item><ItemLocation>'
        '<ItemType>AllFormulas</ItemType><ItemPath/><ItemName/>'
        '</ItemLocation><StableEntries/></Item></Items>'
        '</LocalPackageMetadataFile>'
    ).encode("utf-8")

    # metadata block = 4-byte version + metadata xml + 4-byte content length
    meta_block = (len(metadata)).to_bytes(4, "little") + metadata + (0).to_bytes(4, "little")

    out = bytearray()
    out += (0).to_bytes(4, "little")                     # version
    out += len(pkg_bytes).to_bytes(4, "little")          # package length
    out += pkg_bytes
    out += len(permissions).to_bytes(4, "little")
    out += permissions
    out += len(meta_block).to_bytes(4, "little")
    out += meta_block
    out += (0).to_bytes(4, "little")                     # permission bindings
    return bytes(out)


CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def inject(xlsx: Path, queries) -> None:
    tmp = xlsx.with_suffix(".tmp.xlsx")
    payload = build_datamashup(queries)
    b64 = base64.b64encode(payload).decode("ascii")

    item1 = ('<?xml version="1.0" encoding="utf-8"?>'
             f'<DataMashup xmlns="http://schemas.microsoft.com/DataMashup">{b64}</DataMashup>')
    itemprops1 = ('<?xml version="1.0" encoding="utf-8"?>'
                  '<ds:datastoreItem ds:itemID="{F4A0C4E1-9B2D-4A7E-8C31-2E5D7A9B0C11}" '
                  'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">'
                  '<ds:schemaRefs><ds:schemaRef ds:uri="http://schemas.microsoft.com/DataMashup"/>'
                  '</ds:schemaRefs></ds:datastoreItem>')

    # one workbook-query connection per query, so Excel lists them
    conns = []
    for i, (name, _) in enumerate(queries, start=1):
        conns.append(
            f'<connection id="{i}" xr16:uid="{{B1E6F5A{i}-0000-0000-0000-00000000000{i}}}" '
            f'name="Query - {name}" description="Connection to the \'{name}\' query in the workbook." '
            f'type="5" refreshedVersion="8" background="1" saveData="1">'
            f'<dbPr connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;'
            f'Location={name};Extended Properties=&quot;&quot;" '
            f'command="SELECT * FROM [{name}]"/></connection>')
    connections = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:xr16="http://schemas.microsoft.com/office/spreadsheetml/2017/revision16" '
        'mc:Ignorable="xr16">' + "".join(conns) + '</connections>')

    with zipfile.ZipFile(xlsx) as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        names = set(zin.namelist())
        for item in zin.infolist():
            data = zin.read(item.filename)

            if item.filename == "[Content_Types].xml":
                x = data.decode("utf-8")
                add = ""
                if "customXml/item1.xml" not in x:
                    add += ('<Override PartName="/customXml/item1.xml" '
                            'ContentType="application/xml"/>'
                            '<Override PartName="/customXml/itemProps1.xml" '
                            'ContentType="application/vnd.openxmlformats-officedocument.'
                            'customXmlProperties+xml"/>')
                if "/xl/connections.xml" not in x:
                    add += ('<Override PartName="/xl/connections.xml" '
                            'ContentType="application/vnd.openxmlformats-officedocument.'
                            'spreadsheetml.connections+xml"/>')
                x = x.replace("</Types>", add + "</Types>")
                data = x.encode("utf-8")

            elif item.filename == "xl/_rels/workbook.xml.rels":
                x = data.decode("utf-8")
                if "connections.xml" not in x:
                    x = x.replace("</Relationships>",
                                  '<Relationship Id="rIdPQconn" Type="http://schemas.openxmlformats.org/'
                                  'officeDocument/2006/relationships/connections" Target="connections.xml"/>'
                                  '</Relationships>')
                data = x.encode("utf-8")

            elif item.filename == "_rels/.rels":
                x = data.decode("utf-8")
                if "customXml/item1.xml" not in x:
                    x = x.replace("</Relationships>",
                                  '<Relationship Id="rIdPQcustx" Type="http://schemas.openxmlformats.org/'
                                  'officeDocument/2006/relationships/customXml" Target="customXml/item1.xml"/>'
                                  '</Relationships>')
                data = x.encode("utf-8")

            zout.writestr(item, data)

        zout.writestr("customXml/item1.xml", item1)
        zout.writestr("customXml/itemProps1.xml", itemprops1)
        zout.writestr("customXml/_rels/item1.xml.rels",
                      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                      f'<Relationships xmlns="{REL_NS}">'
                      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
                      'officeDocument/2006/relationships/customXmlProps" '
                      'Target="itemProps1.xml"/></Relationships>')
        zout.writestr("xl/connections.xml", connections)

    shutil.move(str(tmp), str(xlsx))


def main():
    root = repo_root()
    n = 0
    for lab in ALL_LABS:
        qs = QUERIES.get(lab["num"])
        if not qs:
            continue
        folder = root / "labs" / f"lab-{lab['num']:02d}-{slug(lab['title'])}"
        xlsx = folder / lab["workbook"]
        if not xlsx.exists():
            print(f"  !! missing {xlsx}")
            continue
        inject(xlsx, qs)
        n += 1
        print(f"  lab {lab['num']:02d}  +{len(qs)} query: {', '.join(q[0] for q in qs)}")
    print(f"\nPower Query injected into {n} workbooks")


if __name__ == "__main__":
    main()
