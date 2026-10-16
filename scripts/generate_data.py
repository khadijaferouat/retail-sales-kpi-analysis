"""
Génère un jeu de données de ventes retail fictif (année 2025).

4 magasins x 5 catégories x 365 jours. Saisonnalité mensuelle, effet jour de
semaine et bruit aléatoire. Une anomalie est volontairement injectée :
le magasin de Lyon perd environ 15 % de fréquentation en octobre 2025
(simulation d'une panne prolongée du système d'encaissement).

Usage : python scripts/generate_data.py
"""
import numpy as np
import pandas as pd
from pathlib import Path

SEED = 8
rng = np.random.default_rng(SEED)

MAGASINS = {  # magasin: (région, transactions moyennes / jour / catégorie)
    "Paris Rivoli": ("Île-de-France", 95),
    "Lyon Part-Dieu": ("Auvergne-Rhône-Alpes", 70),
    "Lille Centre": ("Hauts-de-France", 55),
    "Bordeaux Lac": ("Nouvelle-Aquitaine", 60),
}
CATEGORIES = {  # catégorie: (panier moyen €, taux de marge, articles / ticket)
    "Alimentaire": (24.0, 0.28, 6.5),
    "Maison": (38.0, 0.42, 2.1),
    "Textile": (46.0, 0.55, 1.8),
    "Électronique": (112.0, 0.22, 1.1),
    "Beauté": (29.0, 0.48, 2.4),
}
SAISON = {1: 0.88, 2: 0.86, 3: 0.95, 4: 0.97, 5: 1.00, 6: 1.02,
          7: 0.98, 8: 0.90, 9: 0.99, 10: 1.03, 11: 1.12, 12: 1.30}
JOUR = {0: 0.85, 1: 0.88, 2: 0.95, 3: 0.97, 4: 1.08, 5: 1.35, 6: 0.92}

ANOMALIE = {"magasin": "Lyon Part-Dieu", "mois": 10, "facteur": 0.85}

rows = []
for date in pd.date_range("2025-01-01", "2025-12-31", freq="D"):
    for mag, (region, base_tx) in MAGASINS.items():
        for cat, (panier, marge, art) in CATEGORIES.items():
            lam = base_tx * SAISON[date.month] * JOUR[date.dayofweek]
            if mag == ANOMALIE["magasin"] and date.month == ANOMALIE["mois"]:
                lam *= ANOMALIE["facteur"]
            tx = int(rng.poisson(lam))
            pm = panier * rng.normal(1.0, 0.04)
            ca = round(tx * pm, 2)
            qte = int(round(tx * art * rng.normal(1.0, 0.05)))
            cout = round(ca * (1 - marge * rng.normal(1.0, 0.03)), 2)
            rows.append({
                "Date": date.date(),
                "Mois": date.strftime("%Y-%m"),
                "Magasin": mag,
                "Région": region,
                "Catégorie": cat,
                "Nb_transactions": tx,
                "Quantité": qte,
                "CA_HT": ca,
                "Coût_achat": cout,
            })

df = pd.DataFrame(rows)
out = Path(__file__).resolve().parents[1] / "data" / "ventes_retail_2025.csv"
df.to_csv(out, index=False, sep=";", decimal=",", encoding="utf-8-sig")
print(f"{len(df)} lignes écrites dans {out}")
