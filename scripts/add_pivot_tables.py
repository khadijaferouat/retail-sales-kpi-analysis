"""
Ajoute de vrais tableaux croisés dynamiques (TCD) dans l'onglet TCD,
via LibreOffice (API UNO), puis recalcule et enregistre au format .xlsx.

Prérequis : LibreOffice installé (soffice dans le PATH) et le module python `uno`.
Usage : python scripts/add_pivot_tables.py
"""
import os, subprocess, time, tempfile
from pathlib import Path
import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.table import CellAddress, CellRangeAddress
from com.sun.star.sheet.DataPilotFieldOrientation import ROW, COLUMN, DATA
from com.sun.star.sheet.GeneralFunction import SUM
from com.sun.star.sheet import DataPilotFieldReference
from com.sun.star.sheet.DataPilotFieldReferenceType import ROW_PERCENTAGE

ROOT = Path(__file__).resolve().parents[1]
XLSX = ROOT / "excel" / "analyse_ventes_retail_2025.xlsx"

def prop(n, v):
    p = PropertyValue(); p.Name, p.Value = n, v; return p

for lock in XLSX.parent.glob(".~lock*"):
    lock.unlink()
profile = Path(tempfile.mkdtemp()).as_uri()
proc = subprocess.Popen(["soffice", "--headless", "--invisible", "--norestore",
    f"-env:UserInstallation={profile}",
    "--accept=socket,host=localhost,port=2002;urp;"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
for _ in range(60):
    try:
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"); break
    except Exception:
        time.sleep(0.5)
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
doc = desktop.loadComponentFromURL(XLSX.as_uri(), "_blank", 0, (prop("Hidden", True),))
if doc is None:
    proc.kill()
    raise SystemExit("Impossible d'ouvrir le classeur dans LibreOffice")

sheets = doc.Sheets
src = sheets.getByName("Données")
cur = src.createCursor(); cur.gotoEndOfUsedArea(False)
last_row = cur.RangeAddress.EndRow
src_idx = [i for i in range(sheets.Count) if sheets.getByIndex(i).Name == "Données"][0]
tcd = sheets.getByName("TCD")
tcd_idx = [i for i in range(sheets.Count) if sheets.getByIndex(i).Name == "TCD"][0]
area = CellRangeAddress(src_idx, 0, 0, 9, last_row)  # A1:J<n>

def add_pivot(name, row, value_field, label, title, pct=False):
    tcd.getCellByPosition(0, row - 2).String = title
    dps = tcd.DataPilotTables
    d = dps.createDataPilotDescriptor()
    d.SourceRange = area
    fields = d.DataPilotFields
    def field(n):
        for i in range(fields.Count):
            f = fields.getByIndex(i)
            if f.Name == n:
                return f
    field("Mois").Orientation = ROW
    field("Magasin").Orientation = COLUMN
    f = field(value_field)
    f.Orientation = DATA
    f.Function = SUM
    dps.insertNewByName(name, CellAddress(tcd_idx, 0, row), d)
    t = dps.getByName(name)
    if pct:
        ref = DataPilotFieldReference()
        ref.ReferenceType = ROW_PERCENTAGE
        dfs = t.getDataPilotFields()
        for i in range(dfs.Count):
            fld = dfs.getByIndex(i)
            if fld.Orientation == DATA:
                fld.setPropertyValue("Reference", ref)
    fmt = "0.0%" if pct else "# ##0"
    keys = doc.NumberFormats
    loc = uno.createUnoStruct("com.sun.star.lang.Locale")
    key = keys.queryKey(fmt, loc, False)
    if key == -1:
        key = keys.addNew(fmt, loc)
    out = t.OutputRange
    tcd.getCellRangeByPosition(1, out.StartRow + 2, out.EndColumn, out.EndRow).NumberFormat = key
    for prop_name, val in (("GrandTotalName", "Total"), ("ShowFilterButton", False)):
        try:
            t.setPropertyValue(prop_name, val)
        except Exception:
            pass  # propriété absente selon la version de LibreOffice

add_pivot("TCD_CA", 3, "CA_HT", "Somme de CA_HT", "TCD 1 : somme du CA HT par mois et par magasin")
add_pivot("TCD_Transactions", 22, "Nb_transactions", "Somme de Nb_transactions", "TCD 2 : somme des transactions par mois et par magasin")
add_pivot("TCD_PartCA", 41, "CA_HT", "Part du CA (% ligne)", "TCD 3 : part de chaque magasin dans le CA du mois (% de la ligne). Lyon Part-Dieu décroche en 2025-10.", pct=True)

doc.calculateAll()
doc.storeToURL(XLSX.as_uri(), (prop("FilterName", "Calc MS Excel 2007 XML"),))
doc.close(True)
try:
    desktop.terminate()
except Exception:
    pass
proc.wait(timeout=30)

# --- Post-traitement OOXML : libellés français et affichage "% du total de la ligne"
# (LibreOffice calcule le % mais n'écrit pas l'attribut showDataAs pour Excel)
import zipfile, shutil
tmp = XLSX.with_suffix(".tmp.xlsx")
with zipfile.ZipFile(XLSX) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename.startswith("xl/pivotTables/pivotTable") and item.filename.endswith(".xml"):
            x = data.decode("utf-8").replace('dataCaption="Values"', 'dataCaption="Valeurs"')
            if 'name="TCD_PartCA"' in x:
                x = x.replace('name="Sum - CA_HT" fld="7" subtotal="sum" numFmtId="',
                              'name="Part du CA (% ligne)" fld="7" subtotal="sum" showDataAs="percentOfRow" baseField="0" baseItem="0" numFmtId="')
            x = x.replace('name="Sum - CA_HT"', 'name="Somme de CA_HT"')
            x = x.replace('name="Sum - Nb_transactions"', 'name="Somme de Nb_transactions"')
            data = x.encode("utf-8")
        if item.filename.startswith("xl/worksheets/") or item.filename == "xl/sharedStrings.xml":
            x = data.decode("utf-8")
            x = x.replace(">Sum - CA_HT<", ">Somme de CA_HT<").replace(">Sum - Nb_transactions<", ">Somme de Nb_transactions<")
            data = x.encode("utf-8")
        zout.writestr(item, data)
shutil.move(tmp, XLSX)
print("TCD ajoutés et classeur enregistré :", XLSX)
