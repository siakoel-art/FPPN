# -*- coding: utf-8 -*-
"""Recopie les commentaires modernes (fil de discussion PowerPoint) d'une version relue vers un rapport régénéré.

Le commentaire est rattaché à la diapositive de même rang (le sommaire, diapositive 2) ; ses repères de diapositive
(identifiant et identifiant de création) sont recalés sur le nouveau fichier.

Usage : python3 injecter_commentaire_fppn.py <version relue.pptx> <rapport régénéré.pptx> <sortie.pptx>
"""
import re
import shutil
import sys
import zipfile

SRC, CIBLE, SORTIE = sys.argv[1:4]
# v14 : correspondance facultative « diapositive de la version relue » -> « diapositive du rapport régénéré » (fichier JSON)
import json  # noqa: E402
CORR = json.load(open(sys.argv[4], encoding="utf-8")) if len(sys.argv) > 4 else {}
zs = zipfile.ZipFile(SRC)
parts = {}
for n in zs.namelist():
    if n.startswith("ppt/comments/") or n == "ppt/authors.xml":
        parts[n] = zs.read(n)
# diapositive de chaque commentaire dans la version relue
lien = {}
for n in zs.namelist():
    if n.startswith("ppt/slides/_rels/"):
        for cible in re.findall(r'Target="\.\./comments/([^"]+)"', zs.read(n).decode()):
            s_ = n.split("/")[-1].replace(".rels", "")
            lien["ppt/comments/" + cible] = CORR.get(s_, s_)

zc = zipfile.ZipFile(CIBLE)
pres = zc.read("ppt/presentation.xml").decode()
rels_p = zc.read("ppt/_rels/presentation.xml.rels").decode()
rid_par_slide = {re.search(r'Target="slides/(slide\d+\.xml)"', m).group(1): re.search(r'Id="(rId\d+)"', m).group(1)
                 for m in re.findall(r"<Relationship [^>]*slides/slide\d+\.xml[^>]*>", rels_p)}
sldid_par_rid = dict((r_, i_) for i_, r_ in re.findall(r'<p:sldId id="(\d+)" r:id="(rId\d+)"', pres))
ajouts, modifs = dict(parts), {}
ct = zc.read("[Content_Types].xml").decode()
for n in parts:
    nom = "/" + n
    typ = "application/vnd.ms-powerpoint.authors+xml" if n.endswith("authors.xml") else "application/vnd.ms-powerpoint.comments+xml"
    if nom not in ct:
        ct = ct.replace("</Types>", f'<Override PartName="{nom}" ContentType="{typ}"/></Types>')
modifs["[Content_Types].xml"] = ct.encode()
if "relationships/authors" not in rels_p:
    rels_p = rels_p.replace("</Relationships>", '<Relationship Id="rIdAuteurs1" Type="http://schemas.microsoft.com/office/2018/10/relationships/authors" Target="authors.xml"/></Relationships>')
modifs["ppt/_rels/presentation.xml.rels"] = rels_p.encode()
for n, slide in lien.items():
    rels_s = "ppt/slides/_rels/" + slide + ".rels"
    r_ = zc.read(rels_s).decode()
    if n.split("/")[-1] not in r_:
        r_ = r_.replace("</Relationships>", f'<Relationship Id="rIdCom1" Type="http://schemas.microsoft.com/office/2018/10/relationships/comments" Target="../comments/{n.split("/")[-1]}"/></Relationships>')
    modifs[rels_s] = r_.encode()
    xml_s = zc.read("ppt/slides/" + slide).decode()
    cid = re.search(r'p14:creationId [^>]*val="(\d+)"', xml_s)
    sid = sldid_par_rid[rid_par_slide[slide]]
    c_ = parts[n].decode()
    if cid:
        c_ = re.sub(r'cId="\d+"', f'cId="{cid.group(1)}"', c_)
    c_ = re.sub(r'sldId="\d+"', f'sldId="{sid}"', c_)
    ajouts[n] = c_.encode()
shutil.copy(CIBLE, SORTIE + ".tmp")
with zipfile.ZipFile(SORTIE + ".tmp") as zi, zipfile.ZipFile(SORTIE, "w", zipfile.ZIP_DEFLATED) as zo:
    for it in zi.infolist():
        zo.writestr(it, modifs.get(it.filename, zi.read(it.filename)))
    for n, d in ajouts.items():
        if n not in zi.namelist():
            zo.writestr(n, d)
import os  # noqa: E402
os.remove(SORTIE + ".tmp")
print(SORTIE, "commentaires recopiés :", len(lien))
