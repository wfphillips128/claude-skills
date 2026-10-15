"""
inject_ribbon.py - put a customUI14.xml ribbon (and optional custom icons) into
an .xlsm / .xlam WITHOUT the Office RibbonX Editor.

This is the scripted alternative. The RibbonX Editor is the default route
because you can see and validate each step; use this when you want the step
repeatable (a build script, a test fixture).

An .xlsm is a zip. A ribbon needs three things in it, and missing any one of
them gives you NO ribbon and NO error:
  customUI/customUI14.xml                  the XML itself
  _rels/.rels                              a relationship of type
      http://schemas.microsoft.com/office/2007/relationships/ui/extensibility
      pointing at it (yes, "2007" for the 2010+ customUI14 part)
  customUI/_rels/customUI14.xml.rels       only for custom icons: one image
      relationship per icon, whose Id is the name used in image="..."

The input file is never modified: the result is written to --out. Close the
workbook in Excel first anyway - Excel holds an open file and would not see
the change until reopened.

Usage:
  python inject_ribbon.py BOOK.xlsm ribbon.xml --out BOOK-ribbon.xlsm
  python inject_ribbon.py BOOK.xlsm ribbon.xml --out BOOK-ribbon.xlsm --image myIcon=icons/my.png

Exit code: 0 = written, 1 = refused (bad XML, missing file), 2 = usage.
"""
import argparse
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

NS14 = "http://schemas.microsoft.com/office/2009/07/customui"
REL_UI14 = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
PKG_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
PART = "customUI/customUI14.xml"
PART_RELS = "customUI/_rels/customUI14.xml.rels"
IMAGE_TYPES = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
               ".gif": "image/gif", ".bmp": "image/bmp", ".ico": "image/x-icon"}


def fail(msg):
    print("REFUSED:", msg)
    sys.exit(1)


def main():
    ap = argparse.ArgumentParser(description="Insert customUI14.xml into an .xlsm/.xlam.")
    ap.add_argument("workbook")
    ap.add_argument("ribbon_xml")
    ap.add_argument("--out", required=True, help="where to write the result (must differ from the input)")
    ap.add_argument("--image", action="append", default=[], metavar="ID=FILE",
                    help="embed a custom icon; reference it in the XML as image=\"ID\"")
    a = ap.parse_args()

    if os.path.abspath(a.out) == os.path.abspath(a.workbook):
        fail("--out must be a different file; the input is never overwritten")
    xml_bytes = open(a.ribbon_xml, "rb").read()
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError as e:
        fail(f"{a.ribbon_xml} is not well-formed XML: {e}")
    if root.tag != f"{{{NS14}}}customUI":
        fail(f"root must be <customUI xmlns=\"{NS14}\"> for customUI14.xml")

    images = {}
    for spec in a.image:
        if "=" not in spec:
            fail(f"--image {spec}: use ID=FILE")
        iid, path = spec.split("=", 1)
        ext = os.path.splitext(path)[1].lower()
        if not re.fullmatch(r"[A-Za-z_][\w]*", iid):
            fail(f"--image {iid}: the id must be a plain name (letters, digits, underscore)")
        if ext not in IMAGE_TYPES or not os.path.isfile(path):
            fail(f"--image {path}: needs an existing .png/.jpg/.gif/.bmp/.ico file")
        images[iid] = (path, ext)

    used = {el.get("image") for el in root.iter() if el.get("image")}
    zin = zipfile.ZipFile(a.workbook)
    names = zin.namelist()

    # images already in the file stay, unless replaced by --image
    old_imgs = {}
    if PART_RELS in names:
        for r in ET.fromstring(zin.read(PART_RELS)):
            if r.get("Type") == REL_IMAGE:
                old_imgs[r.get("Id")] = "customUI/" + r.get("Target").lstrip("/")
    missing = sorted(u for u in used if u not in images and u not in old_imgs)
    if missing and root.get("loadImage") is None:
        fail(f"the XML uses image=\"{', '.join(missing)}\" but no such icon is embedded or given with --image")

    # _rels/.rels: drop any old ribbon relationship, add ours
    ET.register_namespace("", PKG_RELS)
    rels = ET.fromstring(zin.read("_rels/.rels"))
    for r in list(rels):
        if r.get("Type", "").endswith("/ui/extensibility") and r.get("Target", "").lstrip("/").lower() == PART.lower():
            rels.remove(r)
    taken = {r.get("Id") for r in rels}
    rid = "rCustomUI14"
    while rid in taken:
        rid += "x"
    ET.SubElement(rels, f"{{{PKG_RELS}}}Relationship", Id=rid, Type=REL_UI14, Target=PART)
    rels_bytes = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + ET.tostring(rels)

    # customUI14.xml.rels
    img_rel = ET.Element(f"{{{PKG_RELS}}}Relationships")
    keep = {k: v for k, v in old_imgs.items() if k not in images}
    for iid, target in keep.items():
        ET.SubElement(img_rel, f"{{{PKG_RELS}}}Relationship", Id=iid, Type=REL_IMAGE,
                      Target=target[len("customUI/"):])
    for iid, (_, ext) in images.items():
        ET.SubElement(img_rel, f"{{{PKG_RELS}}}Relationship", Id=iid, Type=REL_IMAGE,
                      Target=f"images/{iid}{ext}")
    img_rel_bytes = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + ET.tostring(img_rel)

    # [Content_Types].xml: image extensions need a Default entry
    ct = zin.read("[Content_Types].xml").decode("utf-8")
    for _, ext in images.values():
        if not re.search(rf'Extension="{ext[1:]}"', ct, re.I):
            ct = ct.replace("<Default ", f'<Default Extension="{ext[1:]}" ContentType="{IMAGE_TYPES[ext]}"/><Default ', 1)

    replaced = {"_rels/.rels", "[Content_Types].xml", PART, PART_RELS} | \
               {f"customUI/images/{i}{e}" for i, (_, e) in images.items()}
    with zipfile.ZipFile(a.out, "w", zipfile.ZIP_DEFLATED) as zout:
        zout.writestr("[Content_Types].xml", ct)     # conventionally first
        for item in zin.infolist():
            if item.filename not in replaced:
                zout.writestr(item, zin.read(item.filename))
        zout.writestr("_rels/.rels", rels_bytes)
        zout.writestr(PART, xml_bytes)
        if keep or images:
            zout.writestr(PART_RELS, img_rel_bytes)
        for iid, (path, ext) in images.items():
            zout.write(path, f"customUI/images/{iid}{ext}")
    zin.close()
    print(f"Wrote {a.out}: ribbon from {os.path.basename(a.ribbon_xml)}"
          + (f", {len(images)} icon(s) embedded" if images else ""))
    print("Next: python review_addin.py \"%s\"" % a.out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
