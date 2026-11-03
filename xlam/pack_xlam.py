"""Inject the custom ribbon (customUI14.xml + images) into an .xlam saved by Excel.  usage: pack_xlam.py file.xlam out_dir"""
import sys, zipfile, shutil
from pathlib import Path

xlam, out = Path(sys.argv[1]), Path(sys.argv[2])
zin = zipfile.ZipFile(xlam)
items = {n: zin.read(n) for n in zin.namelist()}
zin.close()

rels = items["_rels/.rels"].decode("utf-8")
rel = '<Relationship Id="rIdPro100xUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'
if "customUI14" not in rels:
    rels = rels.replace("</Relationships>", rel + "</Relationships>")
items["_rels/.rels"] = rels.encode("utf-8")

ct = items["[Content_Types].xml"].decode("utf-8")
if 'Extension="png"' not in ct:
    ct = ct.replace("<Override", '<Default Extension="png" ContentType="image/png"/><Override', 1)
items["[Content_Types].xml"] = ct.encode("utf-8")

items["customUI/customUI14.xml"] = (out / "ribbon.xml").read_bytes()
pngs = sorted((out / "images").glob("*_32.png"))
r = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
for p in pngs:
    key = p.name[:-7]
    items[f"customUI/images/{p.name}"] = p.read_bytes()
    r.append(f'<Relationship Id="{key}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="images/{p.name}"/>')
r.append("</Relationships>")
items["customUI/_rels/customUI14.xml.rels"] = "".join(r).encode("utf-8")

tmp = xlam.with_suffix(".tmp")
with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", items.pop("[Content_Types].xml"))
    for n, b in items.items():
        z.writestr(n, b)
shutil.move(tmp, xlam)
print("packed", xlam.name, len(pngs), "icons")
