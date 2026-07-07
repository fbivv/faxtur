# -*- coding: utf-8 -*-
# ============================================================================
# Faxtur
# Copyright © 2026 Frédéric Brouard
#
# This Source Code Form is subject to the terms of the
# Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file,
# You can obtain one at https://mozilla.org/MPL/2.0/
# ============================================================================
"""
Convertisseur PDF -> Factur-X, avec fiche société et validation d'un modèle de facture.

Objectif :
- l'utilisateur renseigne une société émettrice dans un fichier JSON ;
- il choisit une facture type PDF ;
- le programme analyse la facture, vérifie que l'émetteur correspond à la société configurée ;
- l'utilisateur valide le modèle ;
- le modèle validé est stocké dans le JSON ;
- la conversion PDF -> Factur-X réutilise ces informations.

Dépendances :
    pip install pypdf

Optionnel pour l'Excel du journal :
    pip install pandas openpyxl

Lancement GUI :
    py convertisseur_facturx_assistant_modele.py --gui --config societe.json

Conversion dossier :
    py convertisseur_facturx_assistant_modele.py --config societe.json --input factures --output facturx --todo a_traiter
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import shutil
import sys
import traceback
import subprocess
import tempfile
import uuid
from dataclasses import dataclass, asdict
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from xml.sax.saxutils import escape
from engine.paths import VERAPDF, GHOSTSCRIPT, JAVA, ICC, SETTINGS_JSON


def _run_subprocess_hidden(cmd, **kwargs):
    """Lance Ghostscript sans fenêtre console sous Windows."""
    if sys.platform.startswith("win"):
        kwargs.setdefault("creationflags", getattr(subprocess, "CREATE_NO_WINDOW", 0))
        try:
            si = subprocess.STARTUPINFO()
            si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            kwargs.setdefault("startupinfo", si)
        except Exception:
            pass
    return subprocess.run(cmd, **kwargs)

try:
    from pypdf import PdfReader, PdfWriter
    from pypdf.generic import NameObject, TextStringObject, NumberObject, ArrayObject, ByteStringObject, DecodedStreamObject, DictionaryObject
except Exception as exc:  # pragma: no cover
    PdfReader = None
    PdfWriter = None
    NameObject = TextStringObject = NumberObject = ArrayObject = ByteStringObject = DecodedStreamObject = DictionaryObject = None
    PYPDF_IMPORT_ERROR = exc
else:
    PYPDF_IMPORT_ERROR = None

from engine.pdf_generation import (
    normalize_pdfa_with_ghostscript,
    embed_xml_in_pdf,
)

APP_NAME = "Faxtur"
VERSION = "1.0.0"

# ---------------------------------------------------------------------------
# Utilitaires texte / montants / dates
# ---------------------------------------------------------------------------

def normalize_text(s: str) -> str:
    s = s or ""
    s = s.replace("\u00a0", " ")
    s = s.replace("\u202f", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n[ \t]+", "\n", s)
    return s.strip()


def compact_digits(s: str) -> str:
    return re.sub(r"\D", "", s or "")


def compact_alnum(s: str) -> str:
    return re.sub(r"[^0-9A-Za-z]", "", (s or "")).upper()


def money_to_decimal(raw: str) -> Optional[Decimal]:
    if not raw:
        return None
    s = raw.strip()
    s = s.replace("€", "").replace("EUR", "").replace("euro", "")
    s = s.replace("\u00a0", " ").replace("\u202f", " ")
    s = s.replace(" ", "")
    # formats possibles : 1 234,56 / 1234.56 / 1.234,56
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    # garde signe et chiffres
    s = re.sub(r"[^0-9.\-]", "", s)
    if not s or s in {"-", "."}:
        return None
    try:
        return Decimal(s).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except Exception:
        return None


def dec_to_str(d: Optional[Decimal]) -> str:
    if d is None:
        return ""
    return f"{d:.2f}"


def dec_xml(d: Optional[Decimal]) -> str:
    return dec_to_str(d) if d is not None else "0.00"


MONTHS_FR = {
    "janvier": "01", "fevrier": "02", "février": "02", "mars": "03", "avril": "04", "avr": "04", "avr.": "04",
    "mai": "05", "juin": "06", "juillet": "07", "aout": "08", "août": "08",
    "septembre": "09", "octobre": "10", "novembre": "11", "decembre": "12", "décembre": "12",
}


def normalize_date(raw: str) -> Optional[str]:
    if not raw:
        return None
    s = raw.strip().lower()
    s = re.sub(r"\b(lundi|mardi|mercredi|jeudi|vendredi|samedi|dimanche)\b", "", s).strip()
    # 11/06/2026, 11.06.2026, 12-05-26
    m = re.search(r"(\d{1,2})[./\-](\d{1,2})[./\-](\d{2,4})", s)
    if m:
        d, mo, y = m.groups()
        if len(y) == 2:
            y = "20" + y
        try:
            return datetime(int(y), int(mo), int(d)).strftime("%Y%m%d")
        except ValueError:
            return None
    # 2026-06-11
    m = re.search(r"(20\d{2})[./\-](\d{1,2})[./\-](\d{1,2})", s)
    if m:
        y, mo, d = m.groups()
        try:
            return datetime(int(y), int(mo), int(d)).strftime("%Y%m%d")
        except ValueError:
            return None
    # 12 Mai 2026
    m = re.search(r"(\d{1,2})\s+([a-zéû.]+)\s+(20\d{2})", s)
    if m:
        d, mo_name, y = m.groups()
        mo = MONTHS_FR.get(mo_name) or MONTHS_FR.get(mo_name.rstrip("."))
        if mo:
            try:
                return datetime(int(y), int(mo), int(d)).strftime("%Y%m%d")
            except ValueError:
                return None
    return None


def date_display(yyyymmdd: Optional[str]) -> str:
    if not yyyymmdd:
        return ""
    try:
        return datetime.strptime(yyyymmdd, "%Y%m%d").strftime("%d/%m/%Y")
    except Exception:
        return yyyymmdd

# ---------------------------------------------------------------------------
# Config société
# ---------------------------------------------------------------------------

DEFAULT_COMPANY: Dict[str, Any] = {
    "name": "",
    "legal_form": "SAS",
    "siren": "",
    "siret": "",
    "vat": "",
    "ape": "",
    "address1": "",
    "address2": "",
    "postcode": "",
    "city": "",
    "country": "FR",
    "email": "",
    "phone": "",
    "iban": "",
    "bic": "",
    "invoice_template": {
        "validated": False,
        "validated_at": "",
        "sample_file": "",
        "score": 0,
        "fields": {},
        "notes": "",
        "issuer_match": {},
        "model_hints": {},
    },
}


def load_config(path: str | Path) -> Dict[str, Any]:
    p = Path(path)
    if not p.exists():
        cfg = json.loads(json.dumps(DEFAULT_COMPANY))
        save_config(p, cfg)
        return cfg
    with p.open("r", encoding="utf-8") as f:
        cfg = json.load(f)
    # merge shallow defaults
    merged = json.loads(json.dumps(DEFAULT_COMPANY))
    for k, v in cfg.items():
        if k == "invoice_template" and isinstance(v, dict):
            merged[k].update(v)
        else:
            merged[k] = v
    return merged


def save_config(path: str | Path, cfg: Dict[str, Any]) -> None:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    with p.open("w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

# ---------------------------------------------------------------------------
# Extraction PDF
# ---------------------------------------------------------------------------

def extract_pdf_text(pdf_path: str | Path) -> str:
    if PdfReader is None:
        raise RuntimeError(
            "Le module pypdf est absent. Installez-le avec : python -m pip install pypdf"
        ) from PYPDF_IMPORT_ERROR
    reader = PdfReader(str(pdf_path))
    pages = []
    for page in reader.pages:
        try:
            pages.append(page.extract_text() or "")
        except Exception:
            pages.append("")
    return normalize_text("\n".join(pages))

# ---------------------------------------------------------------------------
# Analyse facture type
# ---------------------------------------------------------------------------

@dataclass
class Candidate:
    value: Any = None
    confidence: int = 0
    source: str = ""


def find_invoice_number(text: str) -> Candidate:
    patterns = [
        (r"Num[ée]ro\s+([A-Z]-[0-9]{4}-[0-9]{2}-[0-9]+)", 98),
        (r"Num[ée]ro\s*[:\n ]+([A-Z0-9][A-Z0-9/_\-]{4,})", 96),
        (r"Facture\s+(FAC/[0-9]{4}/[0-9]+)", 96),
        (r"Facture\s+([A-Z]{2,4}[0-9]{4,}[-/]?[0-9A-Z]*)", 94),
        (r"FACTURE\s*N[°o]?\s*([A-Z0-9][A-Z0-9/\-]+)", 94),
        (r"Facture\s*N[°o]?\s*([A-Z0-9][A-Z0-9/\-]+)", 94),
        (r"Num[eé]ro\s+de\s+la\s+facture\s*[:\-]?\s*([A-Z0-9][A-Z0-9/\-]+)", 92),
        (r"Facture\s*n[°o]?\s*([A-Z0-9][A-Z0-9/\-]+)", 92),
        (r"Invoice\s+number\s*Date\s*\n?\s*(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\s+([A-Z0-9][A-Z0-9/\-]+)", 90),
        (r"Invoice\s+number\s*[:\-]?\s*([A-Z0-9][A-Z0-9/\-]+)", 90),
        (r"\b(INV-[A-Z]{3}-[0-9\-]+)\b", 88),
        (r"\b(FR[0-9A-Z]{8,})\b", 84),
        (r"\b(FA\d{4}-\d{4})\b", 96),
        (r"\b(FAC-\d{8}-\d+)\b", 96),
    ]
    for pat, score in patterns:
        m = re.search(pat, text, re.I)
        if m:
            val = m.group(m.lastindex or 1)
            # cas invoice number/date où le numéro est groupe 2
            if "Invoice" in pat and (m.lastindex or 1) >= 2:
                val = m.group(2)
            val = val.strip()
            if val.upper() in {"NUM", "NUMERO", "N"}:
                continue
            return Candidate(val, score, m.group(0)[:120])
    # fallback : nom de fichier non disponible ici, donc rien
    return Candidate(None, 0, "")


def find_invoice_date(text: str) -> Candidate:
    patterns = [
        (r"Date\s+d[’\']?[ée]mission\s*[:\n ]+([0-9]{1,2}[./\-][0-9]{1,2}[./\-][0-9]{2,4})", 98),
        (r"Date\s+d[’\']?[ée]mission\s*[:\n ]+([0-9]{1,2}\s+(?:janvier|février|fevrier|mars|avril|avr\.?|mai|juin|juillet|août|aout|septembre|octobre|novembre|décembre|decembre)\s+20[0-9]{2})", 98),
        (r"Date\s+Facture\s*[:\n ]+([0-9]{1,2}[./\-][0-9]{1,2}[./\-][0-9]{2,4})", 96),
        (r"Date\s+facturation\s*[:\n ]+([0-9]{1,2}[./\-][0-9]{1,2}[./\-][0-9]{2,4})", 96),
        (r"Date\s+de\s+la\s+facture[^0-9]{0,40}([0-9]{1,2}[./\-][0-9]{1,2}[./\-][0-9]{2,4})", 96),
        (r"En\s+date\s+du\s*[:\n ]+([0-9]{1,2}[./\-][0-9]{1,2}[./\-][0-9]{2,4})", 96),
        (r"Facture\s+N[°o]?.{0,80}?du\s+([0-9]{1,2}[./\-][0-9]{1,2}[./\-][0-9]{2,4})", 94),
        (r"FACTURE\s+N[°o]?.{0,80}?du\s+([0-9]{1,2}[./\-][0-9]{1,2}[./\-][0-9]{2,4})", 94),
        (r"Emise\s+le\s+([0-9]{1,2}[./\-][0-9]{1,2}[./\-][0-9]{2,4})", 92),
        (r"Invoice\s+number\s*Date\s*\n?\s*([0-9]{1,2}[./\-][0-9]{1,2}[./\-][0-9]{2,4})", 92),
        (r"\b([0-9]{1,2}\s+(?:janvier|février|fevrier|mars|avril|avr\.?|mai|juin|juillet|août|aout|septembre|octobre|novembre|décembre|decembre)\s+20[0-9]{2})\b", 82),
        (r"\b([0-9]{1,2}[./\-][0-9]{1,2}[./\-][0-9]{2,4})\b", 60),
    ]
    for pat, score in patterns:
        m = re.search(pat, text, re.I | re.S)
        if m:
            norm = normalize_date(m.group(1))
            if norm:
                return Candidate(norm, score, m.group(0)[:160])
    return Candidate(None, 0, "")


def find_due_date(text: str) -> Candidate:
    patterns = [
        (r"Date\s+d[’\']?échéance\s*[:\n ]+([0-9]{1,2}[./\-][0-9]{1,2}[./\-][0-9]{2,4})", 92),
        (r"Date\s+d[’\']?échéance\s*[:\n ]+([0-9]{1,2}\s+(?:janvier|février|fevrier|mars|avril|avr\.?|mai|juin|juillet|août|aout|septembre|octobre|novembre|décembre|decembre)\s+20[0-9]{2})", 92),
        (r"Date\s+limite\s+de\s+r[èe]glement\s*[:\n ]+([0-9]{1,2}[./\-][0-9]{1,2}[./\-][0-9]{2,4})", 92),
        (r"Due\s+date\s*\n?\s*([0-9]{1,2}[./\-][0-9]{1,2}[./\-][0-9]{2,4})", 90),
        (r"Conditions\s+de\s+r[èe]glement\s*:\s*le\s+([0-9]{1,2}[./\-][0-9]{1,2}[./\-][0-9]{2,4})", 85),
        (r"Conditions\s+de\s+r[èe]glement\s*:\s*le\s+([0-9]{1,2}[./\-][0-9]{2})", 75),
    ]
    for pat, score in patterns:
        m = re.search(pat, text, re.I | re.S)
        if m:
            norm = normalize_date(m.group(1))
            if norm:
                return Candidate(norm, score, m.group(0)[:160])
    return Candidate(None, 0, "")



def find_luciole_recap_totals(text: str) -> Tuple[Candidate, Candidate, Candidate]:
    """Cas des factures Luciole Energies : l'extraction texte remonte parfois
    les trois montants du récapitulatif en début de page, séparés des libellés
    Total HT / Total TVA / Total TTC. On les reprend si le bloc récapitulatif
    est bien présent.
    """
    if not re.search(r"Récapitulatif\s+Total\s+HT\s+Total\s+TVA\s+Total\s+TTC", text, re.I | re.S):
        return Candidate(None, 0, ""), Candidate(None, 0, ""), Candidate(None, 0, "")
    # Les premières lignes contiennent généralement : HT, TVA, TTC.
    head = "\n".join(text.splitlines()[:8])
    amounts = re.findall(r"\b([0-9][0-9 .\u00a0\u202f]*,[0-9]{2})\s*€?\b", head)
    if len(amounts) >= 3:
        return (
            Candidate(money_to_decimal(amounts[0]), 98, "Récapitulatif Luciole : Total HT"),
            Candidate(money_to_decimal(amounts[1]), 98, "Récapitulatif Luciole : Total TVA"),
            Candidate(money_to_decimal(amounts[2]), 98, "Récapitulatif Luciole : Total TTC"),
        )
    return Candidate(None, 0, ""), Candidate(None, 0, ""), Candidate(None, 0, "")


def find_totals(text: str) -> Tuple[Candidate, Candidate, Candidate]:
    ht, vat, ttc = find_luciole_recap_totals(text)

    # HT : on prend le dernier total fiable, car certains tableaux contiennent des sous-totaux avant le total final.
    if ht.value is None:
        for pat, score in [
        (r"Total\s*€\s*HT\s*([0-9][0-9 .\u00a0\u202f]*,[0-9]{2})", 99),
        (r"Total\s+net\s+HT\s*([0-9][0-9 .\u00a0\u202f]*,[0-9]{2})", 96),
        (r"Total\s*(?:€\s*)?HT\s*([0-9][0-9 .\u00a0\u202f]*,[0-9]{2})", 94),
        (r"Montant\s+HT\s*([0-9][0-9 .\u00a0\u202f]*,[0-9]{2})", 90),
        (r"Amount\s*\n?\s*([0-9][0-9 .\u00a0\u202f]*,[0-9]{2})", 70),
    ]:
            matches = list(re.finditer(pat, text, re.I))
            if matches:
                m = matches[-1]
                ht = Candidate(money_to_decimal(m.group(1)), score, m.group(0)[:120])
                break

    # TVA
    if vat.value is None:
        for pat, score in [
        (r"Total\s*(?:€\s*)?TVA\s*([0-9][0-9 .\u00a0\u202f]*,[0-9]{2})", 96),
        (r"TVA\s*20\s*%\s*([0-9][0-9 .\u00a0\u202f]*,[0-9]{2})", 94),
        (r"20\s*%\s*([0-9][0-9 .\u00a0\u202f]*,[0-9]{2})\s*€?\s+[0-9][0-9 .\u00a0\u202f]*,[0-9]{2}", 92),
        (r"Total\s+TVA\s*20[,.]0\s*%\s*([0-9][0-9 .\u00a0\u202f]*,[0-9]{2})", 94),
        (r"Montant\s+TVA\s*([0-9][0-9 .\u00a0\u202f]*,[0-9]{2})", 90),
        (r"Total\s+TVA\s*([0-9][0-9 .\u00a0\u202f]*,[0-9]{2})", 90),
    ]:
            matches = list(re.finditer(pat, text, re.I))
            if matches:
                m = matches[-1]
                vat = Candidate(money_to_decimal(m.group(1)), score, m.group(0)[:120])
                break

    # TTC / total à payer
    if ttc.value is None:
        for pat, score in [
        (r"Total\s*(?:€\s*)?TTC\s*([0-9][0-9 .\u00a0\u202f]*,[0-9]{2})", 98),
        (r"Montant\s+total\s+TTC\s*([0-9][0-9 .\u00a0\u202f]*,[0-9]{2})", 98),
        (r"TOTAL\s+TTC\s*([0-9][0-9 .\u00a0\u202f]*,[0-9]{2})", 98),
        (r"NET\s+A\s+PAYER\s+TTC\s*([0-9][0-9 .\u00a0\u202f]*,[0-9]{2})", 96),
        (r"Net\s+à\s+payer\s*\n?.{0,80}?([0-9][0-9 .\u00a0\u202f]*,[0-9]{2})", 88),
        (r"Total\s+à\s+r[ée]gler\s*([0-9][0-9 .\u00a0\u202f]*,[0-9]{2})", 96),
        (r"Total\s+à\s+payer\s*([0-9][0-9 .\u00a0\u202f]*,[0-9]{2})", 96),
        (r"Prix\s+total\s*([0-9][0-9 .\u00a0\u202f]*,[0-9]{2})", 92),
        (r"Total\s+Due.*?([0-9][0-9 .\u00a0\u202f]*,[0-9]{2})", 88),
        (r"Total\s+([0-9][0-9 .\u00a0\u202f]*,[0-9]{2})\s*€", 70),
    ]:
            matches = list(re.finditer(pat, text, re.I | re.S))
            if matches:
                m = matches[-1]
                ttc = Candidate(money_to_decimal(m.group(1)), score, m.group(0)[:140])
                break

    # Si TTC absent et HT+TVA présents
    if ttc.value is None and ht.value is not None and vat.value is not None:
        ttc = Candidate((ht.value + vat.value).quantize(Decimal("0.01")), 75, "calcul HT + TVA")
    # Si HT absent et TTC+TVA présents
    if ht.value is None and ttc.value is not None and vat.value is not None:
        ht = Candidate((ttc.value - vat.value).quantize(Decimal("0.01")), 70, "calcul TTC - TVA")
    # Si TVA absent et HT+TTC présents
    if vat.value is None and ht.value is not None and ttc.value is not None:
        vat = Candidate((ttc.value - ht.value).quantize(Decimal("0.01")), 65, "calcul TTC - HT")

    return ht, vat, ttc


def find_customer_block(text: str, company: Dict[str, Any]) -> Candidate:
    # Très simple : on extrait un bloc autour d'une adresse client évidente.
    patterns = [
        r"Client\s+ou\s+Cliente\s*\n(.{0,240}?)(?:Période|Produits|Détails|Récapitulatif)",
        r"Adress[ée] à\s*\n(.{0,240}?)(?:Num[eé]ro TVA|Désignation|Catégorie|Commande|Montants exprimés)",
        r"Invoicing address\s+BuyerRef.*?\n(.{0,240}?)(?:Invoice issued by|Agent|Terms)",
        r"Votre N[°o]\s+de\s+TVA.*?\n(.{0,240}?)(?:Représentant|Commande|Code Article)",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.I | re.S)
        if m:
            block = normalize_text(m.group(1))
            return Candidate(block[:300], 70, block[:120])
    # fallback : si la société émettrice est en haut, client souvent dans première moitié. On ne force pas.
    return Candidate("Client non extrait", 20, "fallback")


def find_vat_rate(text: str, ht: Optional[Decimal], vat: Optional[Decimal]) -> Decimal:
    if ht and vat and ht != Decimal("0.00"):
        rate = (vat / ht * Decimal("100")).quantize(Decimal("0.01"))
        if Decimal("19.00") <= rate <= Decimal("21.00"):
            return Decimal("20.00")
        if Decimal("0.00") <= rate <= Decimal("0.50"):
            return Decimal("0.00")
        return rate
    m = re.search(r"TVA\s*(?:à)?\s*(20(?:,00|\.00)?|0(?:,00|\.00)?)\s*%", text, re.I)
    if m:
        return Decimal(m.group(1).replace(",", ".")).quantize(Decimal("0.01"))
    m = re.search(r"(20(?:,00|\.00)?|0(?:,00|\.00)?)\s*%", text)
    if m:
        return Decimal(m.group(1).replace(",", ".")).quantize(Decimal("0.01"))
    return Decimal("20.00") if vat and vat != 0 else Decimal("0.00")


def issuer_match_score(text: str, company: Dict[str, Any]) -> Dict[str, Any]:
    clean = compact_alnum(text)
    score = 0
    checks = []

    def add(label: str, ok: bool, pts: int):
        nonlocal score
        if ok:
            score += pts
        checks.append({"label": label, "ok": ok, "points": pts if ok else 0})

    vat = compact_alnum(company.get("vat", ""))
    siret = compact_digits(company.get("siret", ""))
    siren = compact_digits(company.get("siren", ""))
    name = compact_alnum(company.get("name", ""))
    email = (company.get("email") or "").lower().strip()
    iban = compact_alnum(company.get("iban", ""))

    add("TVA émetteur présente", bool(vat and vat in clean), 35)
    add("SIRET émetteur présent", bool(siret and siret in compact_digits(text)), 25)
    add("SIREN émetteur présent", bool(siren and siren in compact_digits(text)), 15)
    add("Nom société présent", bool(name and name in clean), 15)
    add("Email présent", bool(email and email in text.lower()), 5)
    add("IBAN présent", bool(iban and iban in clean), 5)

    # Ne pas pénaliser trop si SIRET absent sur facture mais TVA/nom/IBAN présents.
    return {"score": min(score, 100), "checks": checks}


def analyze_invoice_model(pdf_path: str | Path, company: Dict[str, Any]) -> Dict[str, Any]:
    text = extract_pdf_text(pdf_path)

    invoice_number = find_invoice_number(text)
    invoice_date = find_invoice_date(text)
    due_date = find_due_date(text)
    ht, vat, ttc = find_totals(text)
    vat_rate = find_vat_rate(text, ht.value, vat.value)
    customer = find_customer_block(text, company)
    issuer = issuer_match_score(text, company)

    score = 0
    controls = []

    def ctl(label: str, ok: bool, pts: int, detail: str = ""):
        nonlocal score
        if ok:
            score += pts
        controls.append({"label": label, "ok": ok, "points": pts if ok else 0, "detail": detail})

    ctl("émetteur reconnu", issuer["score"] >= 45, 25, f"score émetteur {issuer['score']}/100")
    ctl("numéro facture trouvé", invoice_number.value is not None, 15, invoice_number.source)
    ctl("date facture trouvée", invoice_date.value is not None, 15, invoice_date.source)
    ctl("total HT trouvé", ht.value is not None, 10, ht.source)
    ctl("TVA trouvée", vat.value is not None, 10, vat.source)
    ctl("total TTC trouvé", ttc.value is not None, 15, ttc.source)

    # Si une facture présente un total HT avant escompte/remise mais une TVA calculée sur une base nette,
    # la règle EN16931 impose d'utiliser la base taxable nette dans BG-23.
    coherence = False
    if ht.value is not None and vat.value is not None and ttc.value is not None:
        coherence = abs((ht.value + vat.value) - ttc.value) <= Decimal("0.02")
        if not coherence:
            derived_ht = (ttc.value - vat.value).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            if derived_ht > Decimal("0.00") and abs((derived_ht + vat.value) - ttc.value) <= Decimal("0.02"):
                ht = Candidate(derived_ht, 86, "base taxable recalculée : TTC - TVA, utile en cas d'escompte/remise")
                coherence = True
    ctl("cohérence HT + TVA = TTC", coherence, 10, f"{dec_to_str(ht.value)} + {dec_to_str(vat.value)} = {dec_to_str(ttc.value)}")

    score = min(score, 100)
    decision = "OK" if score >= 75 else "A_CONTROLER"

    fields = {
        "invoice_number": invoice_number.value or "",
        "invoice_date": invoice_date.value or "",
        "invoice_date_display": date_display(invoice_date.value),
        "due_date": due_date.value or "",
        "due_date_display": date_display(due_date.value),
        "total_ht": dec_to_str(ht.value),
        "total_tva": dec_to_str(vat.value),
        "total_ttc": dec_to_str(ttc.value),
        "vat_rate": dec_to_str(vat_rate),
        "customer_block": customer.value or "",
        "buyer_vat": find_buyer_vat(text),
        "description": guess_description(text),
    }

    return {
        "version": VERSION,
        "pdf": str(pdf_path),
        "score": score,
        "decision": decision,
        "issuer_match": issuer,
        "controls": controls,
        "fields": fields,
        "text_preview": text[:2500],
    }


def guess_description(text: str) -> str:
    # Description globale par défaut. On évite d'essayer de reconstruire toutes les lignes ici.
    for pat in [
        r"Objet\s*:\s*(.{10,160})",
        r"Commande\s*:\s*(.{10,160})",
        r"Désignation\s+Quantité\s+Prix\s+Montant\s+HT\s*(.{10,160})",
        r"Code Article\s+Désignation.*?\n(.{10,160})",
    ]:
        m = re.search(pat, text, re.I | re.S)
        if m:
            return normalize_text(m.group(1)).replace("\n", " ")[:180]
    return "Facture selon document PDF original"


def format_analysis_report(res: Dict[str, Any]) -> str:
    f = res.get("fields", {})
    lines = []
    lines.append("RÉSULTAT DE L'ANALYSE DU MODÈLE")
    lines.append("========================================")
    lines.append(f"Score : {res.get('score', 0)}/100")
    lines.append("")
    lines.append("Champs extraits :")
    lines.append(f"- Numéro facture : {f.get('invoice_number', '')}")
    lines.append(f"- Date facture   : {f.get('invoice_date_display', '')}")
    lines.append(f"- Date échéance  : {f.get('due_date_display', '')}")
    lines.append(f"- Total HT       : {f.get('total_ht', '')}")
    lines.append(f"- TVA            : {f.get('total_tva', '')}")
    lines.append(f"- Total TTC      : {f.get('total_ttc', '')}")
    lines.append(f"- Taux TVA       : {f.get('vat_rate', '')}%")
    lines.append(f"- Description    : {f.get('description', '')}")
    lines.append("")
    lines.append("Contrôles :")
    for c in res.get("controls", []):
        mark = "OK" if c.get("ok") else "NON"
        lines.append(f"- {mark:3s} {c.get('label')} ({c.get('points',0)} pts) {c.get('detail','')}")
    lines.append("")
    lines.append("Contrôle émetteur :")
    im = res.get("issuer_match", {})
    lines.append(f"- score émetteur : {im.get('score',0)}/100")
    for c in im.get("checks", []):
        mark = "OK" if c.get("ok") else "NON"
        lines.append(f"  - {mark:3s} {c.get('label')} ({c.get('points',0)} pts)")
    lines.append("")
    if res.get("decision") == "OK":
        lines.append("Décision proposée : OK. Le modèle paraît exploitable pour une conversion automatique.")
    else:
        lines.append("Décision proposée : À CONTRÔLER. Le modèle n'est pas assez fiable pour une conversion automatique.")
    return "\n".join(lines)



def french_vat_to_siren(vat: str) -> str:
    vat = compact_alnum(vat)
    m = re.match(r"FR[A-Z0-9]{2}(\d{9})$", vat)
    return m.group(1) if m else ""

def find_buyer_vat(text: str) -> str:
    patterns = [
        r"Votre\s+N[°o]\s+de\s+TVA\s*[:\s]+(FR[A-Z0-9 ]{11,16})",
        r"(FR[A-Z0-9 ]{11,16})\s*Votre\s+N[°o]\s+de\s+TVA",
        r"Buyer's\s+VAT-number\s*[:\s]+(FR[A-Z0-9 ]{11,16})",
        r"Num\s+Tva\s+Intra\s*[:\s]+(FR[A-Z0-9 ]{11,16})",
        r"TVA\s+Intracommunautaire\s*[:\s]+(FR[A-Z0-9 ]{11,16})",
        r"Num[eé]ro\s+TVA\s*[:\s]+(FR[A-Z0-9 ]{11,16})",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.I)
        if m:
            return compact_alnum(m.group(1))
    return ""

def ensure_buyer_electronic_address(fields: Dict[str, Any]) -> str:
    buyer_vat = compact_alnum(fields.get("buyer_vat", ""))
    buyer_siren = french_vat_to_siren(buyer_vat)
    if buyer_siren:
        return buyer_siren
    return "000000000"

def default_french_notes() -> List[Tuple[str, str]]:
    return [
        ("AAB", "Escompte pour paiement anticipé : néant."),
        ("PMD", "En cas de retard de paiement, des pénalités de retard sont exigibles selon les conditions légales ou contractuelles."),
        ("PMT", "Indemnité forfaitaire pour frais de recouvrement : 40 euros."),
    ]

# ---------------------------------------------------------------------------
# XML Factur-X simple
# ---------------------------------------------------------------------------

def build_facturx_xml(company: Dict[str, Any], fields: Dict[str, Any]) -> bytes:
    invoice_id = fields.get("invoice_number") or "SANS-NUMERO"
    issue_date = fields.get("invoice_date") or datetime.now().strftime("%Y%m%d")
    due_date = fields.get("due_date") or issue_date
    ht = money_to_decimal(str(fields.get("total_ht", ""))) or Decimal("0.00")
    tva = money_to_decimal(str(fields.get("total_tva", ""))) or Decimal("0.00")
    ttc = money_to_decimal(str(fields.get("total_ttc", ""))) or (ht + tva)
    rate = money_to_decimal(str(fields.get("vat_rate", ""))) or (Decimal("20.00") if tva else Decimal("0.00"))
    tax_cat = "S" if rate > 0 else "Z"
    description = fields.get("description") or "Facture selon document PDF original"
    customer_block = fields.get("customer_block") or "Client selon PDF original"

    seller_name = company.get("name") or "Société émettrice"
    seller_addr = company.get("address1") or ""
    seller_postcode = company.get("postcode") or ""
    seller_city = company.get("city") or ""
    seller_country = company.get("country") or "FR"
    seller_vat = compact_alnum(company.get("vat", ""))
    seller_siret = compact_digits(company.get("siret", ""))
    seller_siren = compact_digits(company.get("siren", "")) or (seller_siret[:9] if len(seller_siret) >= 9 else "")
    iban = compact_alnum(company.get("iban", ""))
    bic = compact_alnum(company.get("bic", ""))
    buyer_vat = compact_alnum(fields.get("buyer_vat", ""))
    # BT-48 ne doit être présent que si l'identifiant TVA acheteur est réellement exploitable.
    # Un élément vide ou un identifiant sans préfixe ISO déclenche BR-CO-09.
    buyer_vat_xml = ""
    if re.match(r"^[A-Z]{2}[A-Z0-9]{2,}$", buyer_vat):
        buyer_vat_xml = f'<ram:SpecifiedTaxRegistration><ram:ID schemeID="VA">{escape(buyer_vat)}</ram:ID></ram:SpecifiedTaxRegistration>'
    buyer_electronic = ensure_buyer_electronic_address(fields)
    notes_xml = "\n".join([f"      <ram:IncludedNote><ram:Content>{escape(txt)}</ram:Content><ram:SubjectCode>{code}</ram:SubjectCode></ram:IncludedNote>" for code, txt in default_french_notes()])

    # Profil BASIC : autorise une ligne globale unique.
    xml = f'''<?xml version="1.0" encoding="UTF-8"?>
<rsm:CrossIndustryInvoice
 xmlns:rsm="urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100"
 xmlns:ram="urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100"
 xmlns:udt="urn:un:unece:uncefact:data:standard:UnqualifiedDataType:100">
  <rsm:ExchangedDocumentContext>
    <ram:GuidelineSpecifiedDocumentContextParameter>
      <ram:ID>urn:factur-x.eu:1p0:basic</ram:ID>
    </ram:GuidelineSpecifiedDocumentContextParameter>
  </rsm:ExchangedDocumentContext>
  <rsm:ExchangedDocument>
    <ram:ID>{escape(invoice_id)}</ram:ID>
    <ram:TypeCode>380</ram:TypeCode>
    <ram:IssueDateTime><udt:DateTimeString format="102">{escape(issue_date)}</udt:DateTimeString></ram:IssueDateTime>
{notes_xml}
  </rsm:ExchangedDocument>
  <rsm:SupplyChainTradeTransaction>
    <ram:IncludedSupplyChainTradeLineItem>
      <ram:AssociatedDocumentLineDocument><ram:LineID>1</ram:LineID></ram:AssociatedDocumentLineDocument>
      <ram:SpecifiedTradeProduct><ram:Name>{escape(description)}</ram:Name></ram:SpecifiedTradeProduct>
      <ram:SpecifiedLineTradeAgreement>
        <ram:NetPriceProductTradePrice><ram:ChargeAmount>{dec_xml(ht)}</ram:ChargeAmount></ram:NetPriceProductTradePrice>
      </ram:SpecifiedLineTradeAgreement>
      <ram:SpecifiedLineTradeDelivery><ram:BilledQuantity unitCode="C62">1</ram:BilledQuantity></ram:SpecifiedLineTradeDelivery>
      <ram:SpecifiedLineTradeSettlement>
        <ram:ApplicableTradeTax><ram:TypeCode>VAT</ram:TypeCode><ram:CategoryCode>{tax_cat}</ram:CategoryCode><ram:RateApplicablePercent>{dec_xml(rate)}</ram:RateApplicablePercent></ram:ApplicableTradeTax>
        <ram:SpecifiedTradeSettlementLineMonetarySummation><ram:LineTotalAmount>{dec_xml(ht)}</ram:LineTotalAmount></ram:SpecifiedTradeSettlementLineMonetarySummation>
      </ram:SpecifiedLineTradeSettlement>
    </ram:IncludedSupplyChainTradeLineItem>
    <ram:ApplicableHeaderTradeAgreement>
      <ram:SellerTradeParty>
        <ram:Name>{escape(seller_name)}</ram:Name>
        <ram:URIUniversalCommunication><ram:URIID schemeID="0002">{escape(seller_siren)}</ram:URIID></ram:URIUniversalCommunication>
        <ram:SpecifiedLegalOrganization><ram:ID schemeID="0002">{escape(seller_siren)}</ram:ID></ram:SpecifiedLegalOrganization>
        <ram:PostalTradeAddress>
          <ram:PostcodeCode>{escape(seller_postcode)}</ram:PostcodeCode>
          <ram:LineOne>{escape(seller_addr)}</ram:LineOne>
          <ram:CityName>{escape(seller_city)}</ram:CityName>
          <ram:CountryID>{escape(seller_country)}</ram:CountryID>
        </ram:PostalTradeAddress>
        <ram:SpecifiedTaxRegistration><ram:ID schemeID="VA">{escape(seller_vat)}</ram:ID></ram:SpecifiedTaxRegistration>
      </ram:SellerTradeParty>
      <ram:BuyerTradeParty>
        <ram:Name>{escape(customer_block.splitlines()[0] if customer_block else 'Client')}</ram:Name>
        <ram:URIUniversalCommunication><ram:URIID schemeID="0002">{escape(buyer_electronic)}</ram:URIID></ram:URIUniversalCommunication>
        <ram:PostalTradeAddress>
          <ram:LineOne>{escape(customer_block.replace(chr(10), ' ')[:180])}</ram:LineOne>
          <ram:CountryID>FR</ram:CountryID>
        </ram:PostalTradeAddress>
        {buyer_vat_xml}
      </ram:BuyerTradeParty>
    </ram:ApplicableHeaderTradeAgreement>
    <ram:ApplicableHeaderTradeDelivery/>
    <ram:ApplicableHeaderTradeSettlement>
      <ram:InvoiceCurrencyCode>EUR</ram:InvoiceCurrencyCode>
      <ram:SpecifiedTradeSettlementPaymentMeans>
        <ram:TypeCode>30</ram:TypeCode>
        <ram:PayeePartyCreditorFinancialAccount><ram:IBANID>{escape(iban)}</ram:IBANID></ram:PayeePartyCreditorFinancialAccount>
        <ram:PayeeSpecifiedCreditorFinancialInstitution><ram:BICID>{escape(bic)}</ram:BICID></ram:PayeeSpecifiedCreditorFinancialInstitution>
      </ram:SpecifiedTradeSettlementPaymentMeans>
      <ram:ApplicableTradeTax>
        <ram:CalculatedAmount>{dec_xml(tva)}</ram:CalculatedAmount>
        <ram:TypeCode>VAT</ram:TypeCode>
        <ram:BasisAmount>{dec_xml(ht)}</ram:BasisAmount>
        <ram:CategoryCode>{tax_cat}</ram:CategoryCode>
        <ram:RateApplicablePercent>{dec_xml(rate)}</ram:RateApplicablePercent>
      </ram:ApplicableTradeTax>
      <ram:SpecifiedTradePaymentTerms>
        <ram:DueDateDateTime><udt:DateTimeString format="102">{escape(due_date)}</udt:DateTimeString></ram:DueDateDateTime>
      </ram:SpecifiedTradePaymentTerms>
      <ram:SpecifiedTradeSettlementHeaderMonetarySummation>
        <ram:LineTotalAmount>{dec_xml(ht)}</ram:LineTotalAmount>
        <ram:TaxBasisTotalAmount>{dec_xml(ht)}</ram:TaxBasisTotalAmount>
        <ram:TaxTotalAmount currencyID="EUR">{dec_xml(tva)}</ram:TaxTotalAmount>
        <ram:GrandTotalAmount>{dec_xml(ttc)}</ram:GrandTotalAmount>
        <ram:DuePayableAmount>{dec_xml(ttc)}</ram:DuePayableAmount>
      </ram:SpecifiedTradeSettlementHeaderMonetarySummation>
    </ram:ApplicableHeaderTradeSettlement>
  </rsm:SupplyChainTradeTransaction>
</rsm:CrossIndustryInvoice>
'''
    return xml.encode("utf-8")


def pdf_date_now() -> TextStringObject:
    """Date PDF au format D:YYYYMMDDHHmmSSZ."""
    return TextStringObject(datetime.utcnow().strftime("D:%Y%m%d%H%M%SZ"))


# ---------------------------------------------------------------------------
# Conversion dossier
# ---------------------------------------------------------------------------

def convert_one(pdf_path: Path, cfg: Dict[str, Any], output_dir: Path, todo_dir: Path) -> Dict[str, Any]:
    row = {"fichier": pdf_path.name, "statut": "", "score": "", "numero": "", "date": "", "ht": "", "tva": "", "ttc": "", "message": ""}
    try:
        # On analyse chaque facture réelle. Si un modèle est validé, on accepte plus facilement.
        res = analyze_invoice_model(pdf_path, cfg)
        fields = res["fields"]
        row.update({
            "score": str(res["score"]),
            "numero": fields.get("invoice_number", ""),
            "date": fields.get("invoice_date_display", ""),
            "ht": fields.get("total_ht", ""),
            "tva": fields.get("total_tva", ""),
            "ttc": fields.get("total_ttc", ""),
        })
        template_valid = bool(cfg.get("invoice_template", {}).get("validated"))
        if res["score"] < (60 if template_valid else 75):
            todo_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(pdf_path, todo_dir / pdf_path.name)
            row["statut"] = "A_TRAITER"
            row["message"] = f"score insuffisant {res['score']}/100"
            return row
        xml = build_facturx_xml(cfg, fields)
        out = output_dir / f"{pdf_path.stem}_facturx.pdf"
        embed_xml_in_pdf(pdf_path, out, xml)
        row["statut"] = "OK"
        row["message"] = str(out)
        return row
    except Exception as exc:
        todo_dir.mkdir(parents=True, exist_ok=True)
        try:
            shutil.copy2(pdf_path, todo_dir / pdf_path.name)
        except Exception:
            pass
        row["statut"] = "ERREUR"
        row["message"] = f"{type(exc).__name__}: {exc}"
        return row


def convert_folder(config: str | Path, input_dir: str | Path, output_dir: str | Path, todo_dir: str | Path) -> Path:
    cfg = load_config(config)
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)
    todo_dir = Path(todo_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    todo_dir.mkdir(parents=True, exist_ok=True)
    rows = []
    for pdf in sorted(input_dir.glob("*.pdf")):
        rows.append(convert_one(pdf, cfg, output_dir, todo_dir))
    journal = output_dir / "journal_facturx.csv"
    with journal.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["fichier", "statut", "score", "numero", "date", "ht", "tva", "ttc", "message"], delimiter=";")
        writer.writeheader()
        writer.writerows(rows)
    return journal

# ---------------------------------------------------------------------------
# GUI Tkinter
# ---------------------------------------------------------------------------

def run_gui(config_path: str | Path) -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    cfg = load_config(config_path)
    last_analysis: Dict[str, Any] = {}

    root = tk.Tk()
    root.title(APP_NAME)
    root.geometry("1120x780")

    nb = ttk.Notebook(root)
    nb.pack(fill="both", expand=True, padx=8, pady=8)

    # Onglet société
    tab_company = ttk.Frame(nb)
    nb.add(tab_company, text="1. Société émettrice")

    fields_meta = [
        ("name", "Nom"), ("legal_form", "Forme juridique"), ("siren", "SIREN"), ("siret", "SIRET"),
        ("vat", "N° TVA"), ("ape", "APE/NAF"), ("address1", "Adresse 1"), ("address2", "Adresse 2"),
        ("postcode", "Code postal"), ("city", "Ville"), ("country", "Pays"), ("email", "Email"),
        ("phone", "Téléphone"), ("iban", "IBAN"), ("bic", "BIC"),
    ]
    vars_: Dict[str, tk.StringVar] = {}
    for r, (key, label) in enumerate(fields_meta):
        ttk.Label(tab_company, text=label).grid(row=r, column=0, sticky="w", padx=8, pady=4)
        v = tk.StringVar(value=str(cfg.get(key, "")))
        vars_[key] = v
        ttk.Entry(tab_company, textvariable=v, width=70).grid(row=r, column=1, sticky="ew", padx=8, pady=4)
    tab_company.columnconfigure(1, weight=1)

    def update_cfg_from_form():
        for key, _label in fields_meta:
            cfg[key] = vars_[key].get().strip()

    def save_company():
        update_cfg_from_form()
        save_config(config_path, cfg)
        messagebox.showinfo(APP_NAME, f"Fiche société enregistrée :\n{config_path}")

    ttk.Button(tab_company, text="Enregistrer la société", command=save_company).grid(row=len(fields_meta), column=1, sticky="e", padx=8, pady=12)

    # Onglet modèle
    tab_model = ttk.Frame(nb)
    nb.add(tab_model, text="2. Facture type / modèle")

    pdf_var = tk.StringVar(value=cfg.get("invoice_template", {}).get("sample_file", ""))
    ttk.Label(tab_model, text="Facture type PDF").grid(row=0, column=0, padx=8, pady=8, sticky="w")
    ttk.Entry(tab_model, textvariable=pdf_var).grid(row=0, column=1, padx=8, pady=8, sticky="ew")
    tab_model.columnconfigure(1, weight=1)

    def choose_pdf():
        p = filedialog.askopenfilename(title="Choisir une facture type", filetypes=[("PDF", "*.pdf")])
        if p:
            pdf_var.set(p)

    ttk.Button(tab_model, text="Choisir...", command=choose_pdf).grid(row=0, column=2, padx=8, pady=8)

    report = tk.Text(tab_model, height=26, wrap="word", font=("Consolas", 10))
    report.grid(row=4, column=0, columnspan=3, sticky="nsew", padx=8, pady=8)
    tab_model.rowconfigure(4, weight=1)

    info = ttk.Label(tab_model, text="Principe : choisissez une facture réellement émise par la société configurée. Le logiciel vérifie l'émetteur, extrait les champs clés, puis vous validez le modèle.", wraplength=1000)
    info.grid(row=2, column=0, columnspan=3, sticky="w", padx=8, pady=8)

    def analyze_model_button():
        nonlocal last_analysis
        update_cfg_from_form()
        p = Path(pdf_var.get())
        if not p.exists():
            messagebox.showerror(APP_NAME, "Le fichier PDF indiqué n'existe pas.")
            return
        try:
            res = analyze_invoice_model(p, cfg)
            last_analysis = res
            report.delete("1.0", "end")
            report.insert("1.0", format_analysis_report(res))
        except Exception as exc:
            report.delete("1.0", "end")
            report.insert("1.0", traceback.format_exc())
            messagebox.showerror(APP_NAME, f"Erreur d'analyse : {exc}")

    def validate_model_button():
        nonlocal last_analysis
        if not last_analysis:
            messagebox.showwarning(APP_NAME, "Analysez d'abord une facture type.")
            return
        if last_analysis.get("score", 0) < 50:
            if not messagebox.askyesno(APP_NAME, "Le score est faible. Valider quand même ce modèle ?"):
                return
        update_cfg_from_form()
        cfg["invoice_template"] = {
            "validated": True,
            "validated_at": datetime.now().isoformat(timespec="seconds"),
            "sample_file": pdf_var.get(),
            "score": last_analysis.get("score", 0),
            "fields": last_analysis.get("fields", {}),
            "notes": "Modèle validé manuellement par l'utilisateur.",
            "issuer_match": last_analysis.get("issuer_match", {}),
            "model_hints": {
                "invoice_number_source": find_invoice_number_source(last_analysis),
                "date_source": find_date_source(last_analysis),
            },
        }
        save_config(config_path, cfg)
        messagebox.showinfo(APP_NAME, "Modèle validé et stocké dans le fichier société.")

    ttk.Button(tab_model, text="Analyser la facture type", command=analyze_model_button).grid(row=1, column=1, sticky="w", padx=8, pady=4)
    ttk.Button(tab_model, text="Valider et stocker ce modèle", command=validate_model_button).grid(row=1, column=1, sticky="e", padx=8, pady=4)

    # Onglet conversion
    tab_conv = ttk.Frame(nb)
    nb.add(tab_conv, text="3. Conversion dossier")
    input_var = tk.StringVar(value="factures")
    output_var = tk.StringVar(value="facturx")
    todo_var = tk.StringVar(value="a_traiter")

    for r, (label, var) in enumerate([("Dossier factures PDF", input_var), ("Dossier sortie Factur-X", output_var), ("Dossier à traiter", todo_var)]):
        ttk.Label(tab_conv, text=label).grid(row=r, column=0, sticky="w", padx=8, pady=8)
        ttk.Entry(tab_conv, textvariable=var).grid(row=r, column=1, sticky="ew", padx=8, pady=8)
        def browse(v=var):
            p = filedialog.askdirectory()
            if p:
                v.set(p)
        ttk.Button(tab_conv, text="Choisir...", command=browse).grid(row=r, column=2, padx=8, pady=8)
    tab_conv.columnconfigure(1, weight=1)
    conv_report = tk.Text(tab_conv, height=20, wrap="word", font=("Consolas", 10))
    conv_report.grid(row=5, column=0, columnspan=3, sticky="nsew", padx=8, pady=8)
    tab_conv.rowconfigure(5, weight=1)

    def convert_button():
        update_cfg_from_form()
        save_config(config_path, cfg)
        try:
            journal = convert_folder(config_path, input_var.get(), output_var.get(), todo_var.get())
            conv_report.delete("1.0", "end")
            conv_report.insert("1.0", f"Conversion terminée.\nJournal : {journal}\n")
            messagebox.showinfo(APP_NAME, f"Conversion terminée.\nJournal : {journal}")
        except Exception as exc:
            conv_report.delete("1.0", "end")
            conv_report.insert("1.0", traceback.format_exc())
            messagebox.showerror(APP_NAME, f"Erreur de conversion : {exc}")

    ttk.Button(tab_conv, text="Convertir le dossier", command=convert_button).grid(row=3, column=1, sticky="e", padx=8, pady=12)

    root.mainloop()


def find_invoice_number_source(res: Dict[str, Any]) -> str:
    for c in res.get("controls", []):
        if "numéro" in c.get("label", ""):
            return c.get("detail", "")
    return ""


def find_date_source(res: Dict[str, Any]) -> str:
    for c in res.get("controls", []):
        if "date facture" in c.get("label", ""):
            return c.get("detail", "")
    return ""

# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=APP_NAME)
    parser.add_argument("--config", default="societe.json", help="Fichier JSON société")
    parser.add_argument("--gui", action="store_true", help="Lancer l'interface graphique")
    parser.add_argument("--analyze", help="Analyser une facture type PDF")
    parser.add_argument("--input", help="Dossier de PDF à convertir")
    parser.add_argument("--output", default="facturx", help="Dossier de sortie")
    parser.add_argument("--todo", default="a_traiter", help="Dossier des PDF à traiter manuellement")
    args = parser.parse_args(argv)

    if args.gui:
        run_gui(args.config)
        return 0

    cfg = load_config(args.config)

    if args.analyze:
        res = analyze_invoice_model(args.analyze, cfg)
        print(format_analysis_report(res))
        return 0

    if args.input:
        journal = convert_folder(args.config, args.input, args.output, args.todo)
        print(f"Conversion terminée. Journal : {journal}")
        return 0

    parser.print_help()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
