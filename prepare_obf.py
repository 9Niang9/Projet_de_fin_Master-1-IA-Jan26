#!/usr/bin/env python3
"""
Collecte et nettoyage d'Open Beauty Facts pour le projet de chatbot
d'aide au conseil cosmétique en pharmacie.

Usage :
    pip install pandas requests tqdm
    python prepare_obf.py                      # télécharge puis nettoie
    python prepare_obf.py --input obf.csv      # utilise un fichier déjà téléchargé
    python prepare_obf.py --country france     # filtre sur un pays de vente

Sorties (dossier data_Open Beauty Facts/) :
    products_clean.csv    catalogue nettoyé (un produit par ligne)
    products_clean.jsonl  même contenu, prêt pour l'indexation RAG
    ingredients_freq.csv  fréquence des ingrédients (aide à construire le référentiel)
    report.txt            statistiques de qualité des données

Source : Open Beauty Facts (https://world.openbeautyfacts.org), licence ODbL.
Pensez à citer la source et à garder l'attribution dans votre rapport.
"""

import argparse
import re
import unicodedata
from collections import Counter
from pathlib import Path

import pandas as pd
import requests

OBF_URL = "https://static.openbeautyfacts.org/data/en.openbeautyfacts.org.products.csv.gz"
DATA_DIR = Path("data_Open Beauty Facts")

KEEP_COLUMNS = [
    "code",
    "product_name",
    "brands",
    "brands_en",
    "categories_tags",
    "countries_tags",
    "ingredients_text",
    "quantity",
    "url",
    "last_modified_t",
]

# 26 allergènes à déclaration obligatoire historiques (annexe III, règlement CE 1223/2009).
# Le règlement UE 2023/1545 en ajoute d'autres : à compléter à partir de CosIng.
EU_ALLERGENS = {
    "amyl cinnamal", "amylcinnamyl alcohol", "anise alcohol", "benzyl alcohol",
    "benzyl benzoate", "benzyl cinnamate", "benzyl salicylate", "butylphenyl methylpropional",
    "cinnamal", "cinnamyl alcohol", "citral", "citronellol", "coumarin", "eugenol",
    "evernia furfuracea extract", "evernia prunastri extract", "farnesol", "geraniol",
    "hexyl cinnamal", "hydroxycitronellal", "hydroxyisohexyl 3-cyclohexene carboxaldehyde",
    "isoeugenol", "limonene", "linalool", "methyl 2-octynoate", "alpha-isomethyl ionone",
    "parfum", "fragrance",
}

# Marqueurs INDICATIFS à faire valider par un pharmacien (ne pas utiliser tels quels
# comme règles médicales). Servent de point de départ pour le moteur de règles.
SENSITIVE_FLAGS = {
    "retinoid": ["retinol", "retinal", "retinaldehyde", "retinyl palmitate", "retinyl acetate",
                 "tretinoin", "adapalene", "tazarotene"],
    "salicylic_acid": ["salicylic acid", "betaine salicylate"],
    "aha_bha_peel": ["glycolic acid", "lactic acid", "mandelic acid", "citric acid"],
    "benzoyl_peroxide": ["benzoyl peroxide"],
    "essential_oils": ["lavandula angustifolia", "melaleuca alternifolia", "mentha piperita",
                       "eucalyptus globulus", "rosmarinus officinalis", "citrus"],
    "alcohol_denat": ["alcohol denat", "alcohol denat."],
    "uv_filters": ["octocrylene", "avobenzone", "butyl methoxydibenzoylmethane",
                   "homosalate", "oxybenzone", "benzophenone-3", "ethylhexyl methoxycinnamate"],
}


def download(url: str, dest: Path) -> Path:
    """Télécharge le fichier OBF (plusieurs centaines de Mo, patience)."""
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists():
        print(f"[info] {dest} existe déjà, téléchargement ignoré")
        return dest
    print(f"[info] téléchargement de {url}")
    headers = {"User-Agent": "OBF-Data-Prep/1.0 (Open Beauty Facts bulk data download)"}
    with requests.get(url, headers=headers, stream=True, timeout=60) as r:
        r.raise_for_status()
        with open(dest, "wb") as f:
            for chunk in r.iter_content(chunk_size=1 << 20):
                f.write(chunk)
    return dest


def strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def normalize_ingredient(raw: str) -> str:
    s = strip_accents(raw).lower().strip()
    s = re.sub(r"[_*]", "", s)                 # marqueurs de gras OBF
    s = re.sub(r"\(.*?\)", "", s)              # parenthèses (précisions, % ...)
    s = re.sub(r"\d+([.,]\d+)?\s*%", "", s)    # pourcentages
    s = re.sub(r"\s+", " ", s).strip(" .;:-")
    return s


def split_ingredients(text: str) -> list[str]:
    if not isinstance(text, str) or not text.strip():
        return []
    parts = re.split(r"[,;]", text)
    cleaned = [normalize_ingredient(p) for p in parts]
    return [p for p in cleaned if 2 <= len(p) <= 80]


def flag_product(ingredients: list[str]) -> dict:
    joined = " | ".join(ingredients)
    allergens = sorted({i for i in ingredients if i in EU_ALLERGENS})
    flags = sorted(
        name for name, keys in SENSITIVE_FLAGS.items() if any(k in joined for k in keys)
    )
    return {"allergens": allergens, "sensitive_flags": flags}


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", type=Path, help="CSV OBF déjà téléchargé")
    ap.add_argument("--country", default=None, help="ex. france, senegal (filtre countries_tags)")
    ap.add_argument("--limit", type=int, default=None, help="limiter le nombre de lignes lues (test)")
    args = ap.parse_args()

    src = args.input or download(OBF_URL, DATA_DIR / "obf_raw.csv.gz")

    print("[info] lecture du CSV…")
    df = pd.read_csv(
        src, sep="\t", usecols=lambda c: c in KEEP_COLUMNS, nrows=args.limit,
        dtype=str, on_bad_lines="skip", low_memory=False,
    )
    if "brands" not in df.columns:
        df["brands"] = df["brands_en"] if "brands_en" in df.columns else ""
    df = df.drop(columns=["brands_en"], errors="ignore")
    n_raw = len(df)

    text_columns = df.columns
    replacement_chars = sum(
        int(df[col].str.count("\ufffd").sum()) for col in text_columns
    )
    for col in text_columns:
        df[col] = df[col].str.replace("\ufffd", " ", regex=False)
        df[col] = df[col].str.replace(r"\s+", " ", regex=True).str.strip()

    # Nettoyage de base
    df = df.dropna(subset=["product_name", "ingredients_text"])
    df = df[df["product_name"].str.len() >= 3]
    df = df.drop_duplicates(subset=["code"])
    df = df.drop_duplicates(subset=["product_name", "brands"])

    if args.country:
        mask = df["countries_tags"].fillna("").str.contains(args.country, case=False)
        df = df[mask]

    # Ingrédients normalisés + marqueurs
    df["ingredients_list"] = df["ingredients_text"].apply(split_ingredients)
    df = df[df["ingredients_list"].str.len() >= 3]  # liste INCI exploitable

    flags = df["ingredients_list"].apply(flag_product).apply(pd.Series)
    df = pd.concat([df, flags], axis=1)
    df["n_ingredients"] = df["ingredients_list"].str.len()

    # Exports
    DATA_DIR.mkdir(exist_ok=True)
    out = df.copy()
    for col in ["ingredients_list", "allergens", "sensitive_flags"]:
        out[col + "_str"] = out[col].apply(lambda x: "; ".join(x))
    out.drop(columns=["ingredients_list", "allergens", "sensitive_flags"]).to_csv(
        DATA_DIR / "products_clean.csv", index=False
    )
    df.to_json(DATA_DIR / "products_clean.jsonl", orient="records", lines=True, force_ascii=False)

    freq = Counter(i for lst in df["ingredients_list"] for i in lst)
    pd.DataFrame(freq.most_common(), columns=["ingredient", "count"]).to_csv(
        DATA_DIR / "ingredients_freq.csv", index=False
    )

    report = [
        f"Lignes brutes lues        : {n_raw}",
        f"Caractères illisibles corrigés : {replacement_chars}",
        f"Produits après nettoyage  : {len(df)}",
        f"Ingrédients distincts     : {len(freq)}",
        f"Produits avec allergènes  : {(df['allergens'].str.len() > 0).sum()}",
        f"Produits avec marqueurs   : {(df['sensitive_flags'].str.len() > 0).sum()}",
        "",
        "Source : Open Beauty Facts (ODbL) - attribution obligatoire.",
        "Les marqueurs sensibles sont INDICATIFS et doivent être validés par un pharmacien.",
    ]
    (DATA_DIR / "report.txt").write_text("\n".join(report), encoding="utf-8")
    print("\n".join(report))


if __name__ == "__main__":
    main()
