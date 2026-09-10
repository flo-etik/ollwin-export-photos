#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Ajoute une colonne 'Vignette' en tete de la feuille avec les photos Cloudinary
incrustees (images ancrees a la cellule -> compatibles Excel ET LibreOffice Calc).
Le reste du fichier d'origine est laisse tel quel.

Usage : python embed_photos.py "export.xlsx" [NomFeuille] [EnteteColonnePhoto]
"""
import os
import re
import sys
import urllib.request

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage

SRC = sys.argv[1] if len(sys.argv) > 1 else r"C:\Users\Florent\Downloads\Ville de Marseille_diag_Parcs et jardins - Parc de la Ravelle_1789052912458.xlsx"
OUT = os.path.splitext(SRC)[0] + "_vignettes.xlsx"
SHEET = sys.argv[2] if len(sys.argv) > 2 else "Préconisations"
PHOTO_HEADER = sys.argv[3] if len(sys.argv) > 3 else "Photo"

BOX = 140          # cote max d'une vignette en pixels
GAP = 6            # marge en pixels
EMU = 9525         # EMU par pixel
CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "imgcache")
os.makedirs(CACHE, exist_ok=True)
URL_RE = re.compile(r"https?://[^\s;,]+(?:,[a-z]_[^\s;,]+)*", re.I)  # tolere les virgules internes Cloudinary


def fetch(url):
    key = re.sub(r"[^0-9A-Za-z]", "_", url)[-120:] + ".img"
    path = os.path.join(CACHE, key)
    if not os.path.exists(path):
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=30) as r:
            data = r.read()
        with open(path, "wb") as f:
            f.write(data)
    return path


def find_urls(value):
    if not isinstance(value, str):
        return []
    return re.findall(r"https?://[^\s]+", value)


wb = openpyxl.load_workbook(SRC)
if SHEET not in wb.sheetnames:
    sys.exit(f"Feuille '{SHEET}' introuvable. Feuilles : {wb.sheetnames}")
ws = wb[SHEET]

# reperer la colonne des URLs photo
photo_col = None
for c in ws[1]:
    if str(c.value or "").strip().lower() == PHOTO_HEADER.lower():
        photo_col = c.column
        break
if photo_col is None:
    sys.exit(f"Colonne '{PHOTO_HEADER}' introuvable dans l'en-tete")

# collecter les URLs AVANT insertion de colonne
rows_urls = {r: find_urls(ws.cell(row=r, column=photo_col).value)
             for r in range(2, ws.max_row + 1)}

# inserer la colonne A "Vignette"
ws.insert_cols(1)
hdr = ws.cell(row=1, column=1, value="Vignette")
src_hdr = ws.cell(row=1, column=photo_col + 1)  # photo_col a decale de +1
try:
    hdr.font = src_hdr.font
    hdr.fill = src_hdr.fill
    hdr.alignment = src_hdr.alignment
except Exception:
    pass
ws.column_dimensions["A"].width = (BOX + 2 * GAP) / 7.0

placed = 0
for r in range(2, ws.max_row + 1):
    urls = rows_urls.get(r, [])
    if not urls:
        continue
    y = GAP
    for url in urls:
        try:
            path = fetch(url)
            with PILImage.open(path) as im:
                w, h = im.size
        except Exception as e:
            print(f"  !! ligne {r}: {e}")
            continue
        scale = min(BOX / w, BOX / h, 1.0)
        dw, dh = int(w * scale), int(h * scale)
        pic = XLImage(path)
        pic.width, pic.height = dw, dh
        marker = AnchorMarker(col=0, colOff=GAP * EMU, row=r - 1, rowOff=int(y * EMU))
        pic.anchor = OneCellAnchor(_from=marker,
                                   ext=XDRPositiveSize2D(cx=dw * EMU, cy=dh * EMU))
        ws.add_image(pic)
        placed += 1
        y += dh + GAP
    need_h = y * 0.75  # px -> points
    cur = ws.row_dimensions[r].height or 15
    if need_h > cur:
        ws.row_dimensions[r].height = need_h

wb.save(OUT)
print(f"OK -> {OUT}  ({placed} vignettes)")
