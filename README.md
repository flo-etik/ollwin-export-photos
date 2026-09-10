# Photos dans l'export Excel

Ajoute une colonne « Vignette » en tête de la feuille avec les photos Cloudinary
incrustées (images ancrées à la cellule → compatibles Excel **et** LibreOffice Calc).

## En ligne (utilisateurs)
Appli web : téléverser l'export .xlsx → récupérer le fichier avec les photos.
Aucune installation côté utilisateur.

## Lancer en local
```
pip install -r requirements.txt
streamlit run app.py
```

## Script en ligne de commande (sans web)
```
pip install openpyxl pillow
python embed_photos.py "export.xlsx" "Préconisations" "Photo"
```
Produit `export_vignettes.xlsx`.
