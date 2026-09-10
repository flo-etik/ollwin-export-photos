# -*- coding: utf-8 -*-
"""
Mini-appli web : l'utilisateur televerse un export .xlsx, l'appli ajoute une
colonne 'Vignette' avec les photos Cloudinary incrustees, et renvoie le fichier.

Lancer en local :   streamlit run app.py
"""
import io
import os
import re
import urllib.request
from urllib.parse import urlparse

import openpyxl
import streamlit as st
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from PIL import Image as PILImage

EMU = 9525
URL_RE = re.compile(r"https?://[^\s]+", re.I)
TIMEOUT = 30
MAX_BYTES = 15 * 1024 * 1024          # taille max par image
ALLOWED_HOSTS = ("res.cloudinary.com", "cloudinary.com")  # l'appli ne va chercher que la

st.set_page_config(page_title="Photos dans l'export", page_icon="🖼️")
st.title("🖼️ Ajouter les photos dans l'export Excel")
st.caption(
    "Televersez l'export .xlsx : les URLs Cloudinary de la colonne photo "
    "sont remplacees par les images, dans une nouvelle colonne « Vignette » "
    "en tete de feuille. Le reste du fichier n'est pas modifie."
)

with st.sidebar:
    st.header("Options")
    sheet_name = st.text_input("Nom de la feuille", value="Préconisations")
    photo_header = st.text_input("En-tete de la colonne des URLs", value="Photo")
    box = st.slider("Taille max des vignettes (px)", 80, 320, 140, 10)
    gap = 6

uploaded = st.file_uploader("Fichier .xlsx", type=["xlsx"])


def fetch_image(url: str) -> bytes:
    p = urlparse(url)
    host = (p.hostname or "").lower()
    if p.scheme != "https" or not (host in ALLOWED_HOSTS or host.endswith(".cloudinary.com")):
        raise ValueError("URL non autorisée (seul cloudinary.com est accepté)")
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
        data = r.read(MAX_BYTES + 1)
    if len(data) > MAX_BYTES:
        raise ValueError("image trop volumineuse")
    return data


def process(data: bytes) -> tuple[bytes, int, list[str]]:
    wb = openpyxl.load_workbook(io.BytesIO(data))
    if sheet_name not in wb.sheetnames:
        raise ValueError(
            f"Feuille « {sheet_name} » absente. Feuilles trouvees : {', '.join(wb.sheetnames)}"
        )
    ws = wb[sheet_name]

    photo_col = None
    for c in ws[1]:
        if str(c.value or "").strip().lower() == photo_header.strip().lower():
            photo_col = c.column
            break
    if photo_col is None:
        raise ValueError(f"Colonne « {photo_header} » introuvable dans la 1re ligne.")

    rows_urls = {
        r: URL_RE.findall(str(ws.cell(row=r, column=photo_col).value or ""))
        for r in range(2, ws.max_row + 1)
    }

    ws.insert_cols(1)
    ws.cell(row=1, column=1, value="Vignette")
    try:
        src = ws.cell(row=1, column=photo_col + 1)
        ws.cell(row=1, column=1).font = src.font
        ws.cell(row=1, column=1).fill = src.fill
        ws.cell(row=1, column=1).alignment = src.alignment
    except Exception:
        pass
    ws.column_dimensions["A"].width = (box + 2 * gap) / 7.0

    placed, warnings = 0, []
    for r in range(2, ws.max_row + 1):
        urls = rows_urls.get(r, [])
        y = gap
        for url in urls:
            try:
                raw = fetch_image(url)
                im = PILImage.open(io.BytesIO(raw))
                w, h = im.size
            except Exception as e:
                warnings.append(f"Ligne {r} : {url[:60]}… ({e})")
                continue
            scale = min(box / w, box / h, 1.0)
            dw, dh = int(w * scale), int(h * scale)
            pic = XLImage(io.BytesIO(raw))
            pic.width, pic.height = dw, dh
            pic.anchor = OneCellAnchor(
                _from=AnchorMarker(col=0, colOff=gap * EMU, row=r - 1, rowOff=int(y * EMU)),
                ext=XDRPositiveSize2D(cx=dw * EMU, cy=dh * EMU),
            )
            ws.add_image(pic)
            placed += 1
            y += dh + gap
        need_h = y * 0.75
        cur = ws.row_dimensions[r].height or 15
        if need_h > cur:
            ws.row_dimensions[r].height = need_h

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), placed, warnings


if uploaded is not None:
    if st.button("Générer le fichier avec les photos", type="primary"):
        with st.spinner("Téléchargement des photos et génération…"):
            try:
                result, n, warnings = process(uploaded.getvalue())
            except Exception as e:
                st.error(str(e))
                st.stop()
        st.success(f"{n} photo(s) incrustée(s).")
        for w in warnings:
            st.warning(w)
        base = os.path.splitext(uploaded.name)[0]
        st.download_button(
            "⬇️ Télécharger l'export avec les photos",
            data=result,
            file_name=f"{base}_photos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
