# ============================================================================
# Faxtur
# Copyright © 2026 Frédéric Brouard
#
# This Source Code Form is subject to the terms of the
# Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file,
# You can obtain one at https://mozilla.org/MPL/2.0/
# ============================================================================
from __future__ import annotations

import csv
import sys
import traceback
import subprocess
import json
import xml.etree.ElementTree as ET
import os
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List

# Moteur figé.
from engine import engine_v1 as engine

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from company_manager import CompanyManager
from settings import SettingsStore, resolve_verapdf_path, resolve_ghostscript_path, resolve_icc_profile, resolve_java_path, configure_runtime_environment
from engine.paths import VERAPDF, GHOSTSCRIPT, JAVA, ICC, SETTINGS_JSON

def get_resource_path(root_dir: Path, relative: str) -> Path:
    """Retourne le chemin d'une ressource, compatible source et PyInstaller."""
    base = Path(getattr(sys, "_MEIPASS", root_dir))
    candidates = [base / relative, Path(root_dir) / relative]
    for c in candidates:
        if c.exists():
            return c
    return candidates[0]


def run_hidden_subprocess(cmd, **kwargs):
    """Lance un outil externe sans ouvrir de fenêtre noire sous Windows."""
    if sys.platform.startswith("win"):
        kwargs.setdefault("creationflags", getattr(subprocess, "CREATE_NO_WINDOW", 0))
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        kwargs.setdefault("startupinfo", si)
    return subprocess.run(cmd, **kwargs)

FIELDS = [
    ("name", "Nom"),
    ("legal_form", "Forme juridique"),
    ("siren", "SIREN"),
    ("siret", "SIRET"),
    ("vat", "N° TVA"),
    ("ape", "APE/NAF"),
    ("address1", "Adresse 1"),
    ("address2", "Adresse 2"),
    ("postcode", "Code postal"),
    ("city", "Ville"),
    ("country", "Pays"),
    ("email", "Email"),
    ("phone", "Téléphone"),
    ("iban", "IBAN"),
    ("bic", "BIC"),
]

class Faxtur(tk.Tk):
    def __init__(self, root_dir: Path):
        super().__init__()
        self.root_dir = Path(root_dir)
        self.settings = SettingsStore(self.root_dir)
        self.company_manager = CompanyManager(self.root_dir)
        self.current_company_file = ""
        self.current_company: Dict[str, Any] = {}
        self.last_analysis: Dict[str, Any] = {}
        self.title("Faxtur 1.1.0")
        self._apply_app_icon()
        self.geometry("1180x780")
        self.minsize(980, 650)
        self._configure_styles()
        self._build_ui()
        self._load_initial_company()

    def _apply_app_icon(self):
        """Applique l'icône Faxtur à la fenêtre principale."""
        self._app_icon_image = None
        ico = get_resource_path(self.root_dir, "resources/Faxtur.ico")
        png = get_resource_path(self.root_dir, "resources/Faxtur_icon.png")
        try:
            if sys.platform.startswith("win") and ico.exists():
                self.iconbitmap(str(ico))
        except Exception:
            pass
        try:
            if png.exists():
                self._app_icon_image = tk.PhotoImage(file=str(png))
                self.iconphoto(True, self._app_icon_image)
        except Exception:
            pass

    def _configure_styles(self):
        self.accent = "#0b63ce"
        self.ok_color = "#15803d"
        self.warn_color = "#b45309"
        self.err_color = "#b91c1c"
        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except Exception:
            pass
        self.style.configure("Header.TFrame", background="#f7fbff")
        self.style.configure("HeaderTitle.TLabel", background="#f7fbff", foreground="#0f172a", font=("Segoe UI", 18, "bold"))
        self.style.configure("HeaderStatus.TLabel", background="#f7fbff", foreground=self.ok_color, font=("Segoe UI", 11, "bold"))
        self.style.configure("Primary.TButton", font=("Segoe UI", 13, "bold"), padding=(18, 12))
        self.style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=(10, 6))
        self.style.configure("Status.TLabel", relief="sunken", anchor="w", padding=(8, 3))
        self.style.configure("Card.TLabelframe", padding=(8, 6))
        self.style.configure("Card.TLabelframe.Label", font=("Segoe UI", 10, "bold"))
        self.style.configure("Ok.TLabel", foreground=self.ok_color, font=("Segoe UI", 10, "bold"))
        self.style.configure("Warn.TLabel", foreground=self.warn_color, font=("Segoe UI", 10, "bold"))
        self.style.configure("Err.TLabel", foreground=self.err_color, font=("Segoe UI", 10, "bold"))

    def short_path(self, value: str) -> str:
        try:
            p = Path(value)
            desktop = Path.home() / "Desktop"
            if desktop.exists():
                try:
                    return str(p).replace(str(desktop), "Bureau")
                except Exception:
                    pass
            parts = p.parts
            if len(parts) > 3:
                return "…" + str(Path(*parts[-3:]))
            return str(p)
        except Exception:
            return value or ""

    def set_status(self, message: str):
        if hasattr(self, "status_var"):
            self.status_var.set(message)

    def refresh_header(self):
        company = self.current_company.get("name") or self.current_company_file or "Aucune société"
        tmpl = self.current_company.get("invoice_template", {}) if self.current_company else {}
        model = "modèle validé" if tmpl.get("validated") else "modèle à valider"
        if hasattr(self, "header_status_var"):
            self.header_status_var.set(f"🟢 Prêt · {company} · {model}")
        if hasattr(self, "dashboard_company_var"):
            self.dashboard_company_var.set(f"Société : {company}")
        if hasattr(self, "dashboard_model_var"):
            self.dashboard_model_var.set(f"Modèle : {model}")
        if hasattr(self, "status_var"):
            self.status_var.set(f"Faxtur 1.1.0   |   Société : {company}   |   {model}")

    def _build_ui(self):
        self.configure(background="#eef3f8")

        # En-tête compact : le logo ne doit jamais prendre toute la fenêtre.
        header = ttk.Frame(self, style="Header.TFrame")
        header.pack(fill="x")
        if getattr(self, "_app_icon_image", None):
            try:
                self._header_icon_image = self._app_icon_image.subsample(6, 6)
            except Exception:
                self._header_icon_image = self._app_icon_image
            ttk.Label(header, image=self._header_icon_image, background="#f7fbff").pack(side="left", padx=(14, 10), pady=8)
        title_box = ttk.Frame(header, style="Header.TFrame")
        title_box.pack(side="left", fill="x", expand=True, pady=8)
        ttk.Label(title_box, text="Faxtur", style="HeaderTitle.TLabel").pack(anchor="w")
        ttk.Label(title_box, text="Facturation électronique", background="#f7fbff", foreground="#334155", font=("Segoe UI", 10)).pack(anchor="w")
        self.header_status_var = tk.StringVar(value="🟢 Prêt")
        ttk.Label(title_box, textvariable=self.header_status_var, style="HeaderStatus.TLabel").pack(anchor="w")
        ttk.Button(header, text="Quitter", command=self.destroy).pack(side="right", padx=12, pady=10)

        # Sélection société.
        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=(8, 4))
        ttk.Label(top, text="Société actuelle :", font=("Segoe UI", 10, "bold")).pack(side="left")
        self.company_combo = ttk.Combobox(top, state="readonly", width=42)
        self.company_combo.pack(side="left", padx=6)
        self.company_combo.bind("<<ComboboxSelected>>", lambda _e: self.select_company(self.company_combo.get()))
        ttk.Button(top, text="Nouvelle", command=self.new_company).pack(side="left", padx=2)
        ttk.Button(top, text="Dupliquer", command=self.duplicate_company).pack(side="left", padx=2)
        ttk.Button(top, text="Supprimer", command=self.delete_company).pack(side="left", padx=2)

        # Corps : menu vertical à gauche, page active à droite.
        body = ttk.Frame(self)
        body.pack(fill="both", expand=True, padx=10, pady=6)
        sidebar = tk.Frame(body, background="#0f315b", width=210)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)
        content = ttk.Frame(body)
        content.pack(side="left", fill="both", expand=True, padx=(10, 0))
        self.content = content
        self.pages: Dict[str, ttk.Frame] = {}
        self.nav_buttons: Dict[str, tk.Button] = {}

        def make_page(key: str) -> ttk.Frame:
            frame = ttk.Frame(content)
            frame.grid(row=0, column=0, sticky="nsew")
            self.pages[key] = frame
            return frame
        content.rowconfigure(0, weight=1)
        content.columnconfigure(0, weight=1)

        self.tab_dashboard = make_page("dashboard")
        self.tab_company = make_page("company")
        self.tab_model = make_page("model")
        self.tab_convert = make_page("convert")
        self.tab_validation = make_page("validation")
        self.tab_settings = make_page("settings")

        nav_items = [
            ("dashboard", "🏠  Tableau de bord"),
            ("company", "🏢  Société"),
            ("model", "📄  Modèle"),
            ("convert", "🔄  Conversion"),
            ("validation", "✔  Validation"),
            ("settings", "⚙  Paramètres"),
        ]
        tk.Label(sidebar, text="Faxtur", bg="#0f315b", fg="white", font=("Segoe UI", 18, "bold")).pack(anchor="w", padx=18, pady=(18, 2))
        tk.Label(sidebar, text="1.1.0", bg="#0f315b", fg="#bfdbfe", font=("Segoe UI", 9)).pack(anchor="w", padx=20, pady=(0, 18))
        for key, label in nav_items:
            btn = tk.Button(sidebar, text=label, anchor="w", relief="flat", bd=0, padx=18, pady=10,
                            bg="#0f315b", fg="white", activebackground="#1d4f91", activeforeground="white",
                            font=("Segoe UI", 10), command=lambda k=key: self.show_page(k))
            btn.pack(fill="x", padx=8, pady=1)
            self.nav_buttons[key] = btn
        tk.Label(sidebar, text="© 2026 Frédéric Brouard\nMPL 2.0", bg="#0f315b", fg="#cbd5e1", font=("Segoe UI", 8), justify="left").pack(side="bottom", anchor="w", padx=18, pady=18)

        self._build_dashboard_tab()
        self._build_company_tab()
        self._build_model_tab()
        self._build_convert_tab()
        self._build_validation_tab()
        self._build_settings_tab()

        self.status_var = tk.StringVar(value="Faxtur 1.1.0 prêt.")
        ttk.Label(self, textvariable=self.status_var, style="Status.TLabel").pack(fill="x", side="bottom")
        self.show_page("dashboard")

    def show_page(self, key: str):
        if key in self.pages:
            self.pages[key].tkraise()
        for k, btn in getattr(self, "nav_buttons", {}).items():
            if k == key:
                btn.configure(bg="#1d4f91")
            else:
                btn.configure(bg="#0f315b")

    def _build_dashboard_tab(self):
        title = ttk.Label(self.tab_dashboard, text="Tableau de bord", font=("Segoe UI", 20, "bold"))
        title.pack(anchor="w", padx=16, pady=(16, 8))
        self.dashboard_status = ttk.Label(self.tab_dashboard, text="Faxtur est prêt.", style="Ok.TLabel")
        self.dashboard_status.pack(anchor="w", padx=16, pady=(0, 14))
        cards = ttk.Frame(self.tab_dashboard)
        cards.pack(fill="x", padx=16, pady=8)
        self.dashboard_company_var = tk.StringVar(value="Société : —")
        self.dashboard_model_var = tk.StringVar(value="Modèle : —")
        self.dashboard_verapdf_var = tk.StringVar(value="veraPDF : —")
        self.dashboard_pdfa_var = tk.StringVar(value="PDF/A : —")
        values = [self.dashboard_company_var, self.dashboard_model_var, self.dashboard_verapdf_var, self.dashboard_pdfa_var]
        labels = ["🏢", "📄", "✔", "🧪"]
        for i, (ico, var) in enumerate(zip(labels, values)):
            lf = ttk.LabelFrame(cards, text=ico, style="Card.TLabelframe")
            lf.grid(row=0, column=i, sticky="nsew", padx=6, pady=4)
            ttk.Label(lf, textvariable=var, font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=14)
            cards.columnconfigure(i, weight=1)
        ttk.Button(self.tab_dashboard, text="CONVERTIR LES FACTURES", style="Primary.TButton", command=lambda: (self.show_page("convert"), self.convert_folder())).pack(fill="x", padx=16, pady=18)
        ttk.Label(self.tab_dashboard, text="Utilisez le menu à gauche pour modifier la société, valider un modèle ou consulter la validation.", foreground="#475569").pack(anchor="w", padx=16, pady=4)

    def _build_company_tab(self):
        self.form_vars: Dict[str, tk.StringVar] = {}
        for r, (key, label) in enumerate(FIELDS):
            ttk.Label(self.tab_company, text=label).grid(row=r, column=0, sticky="w", padx=8, pady=4)
            var = tk.StringVar()
            self.form_vars[key] = var
            ttk.Entry(self.tab_company, textvariable=var, width=78).grid(row=r, column=1, sticky="ew", padx=8, pady=4)
        self.tab_company.columnconfigure(1, weight=1)
        ttk.Button(self.tab_company, text="Enregistrer", command=self.save_company).grid(row=len(FIELDS), column=1, sticky="e", padx=8, pady=12)

    def _build_model_tab(self):
        self.model_pdf_var = tk.StringVar()
        row = 0
        ttk.Label(self.tab_model, text="Facture type PDF").grid(row=row, column=0, padx=8, pady=8, sticky="w")
        ttk.Entry(self.tab_model, textvariable=self.model_pdf_var).grid(row=row, column=1, padx=8, pady=8, sticky="ew")
        ttk.Button(self.tab_model, text="Choisir...", command=self.choose_model_pdf).grid(row=row, column=2, padx=8, pady=8)
        row += 1
        ttk.Button(self.tab_model, text="Analyser", command=self.analyze_model).grid(row=row, column=1, sticky="w", padx=8, pady=4)
        ttk.Button(self.tab_model, text="Valider le modèle", command=self.validate_model).grid(row=row, column=1, sticky="e", padx=8, pady=4)
        row += 1
        self.model_report = tk.Text(self.tab_model, wrap="word", font=("Consolas", 10))
        self.model_report.grid(row=row, column=0, columnspan=3, sticky="nsew", padx=8, pady=8)
        self.tab_model.columnconfigure(1, weight=1)
        self.tab_model.rowconfigure(row, weight=1)

    def _build_convert_tab(self):
        self.input_var = tk.StringVar(value=str(self.settings.get("default_input_dir", "factures")))
        self.output_var = tk.StringVar(value=str(self.settings.get("default_output_dir", "facturx")))
        self.todo_var = tk.StringVar(value=str(self.settings.get("default_todo_dir", "a_traiter")))
        intro = ttk.LabelFrame(self.tab_convert, text="Conversion", style="Card.TLabelframe")
        intro.grid(row=0, column=0, columnspan=3, sticky="ew", padx=8, pady=8)
        ttk.Label(intro, text="Déposez vos PDF dans le dossier Factures reçues, puis lancez la conversion.").grid(row=0, column=0, sticky="w", padx=8, pady=4)
        cards = [("📥 Factures reçues", self.input_var), ("📤 Factur-X générées", self.output_var), ("⚠ À vérifier", self.todo_var)]
        for r, (label, var) in enumerate(cards, start=1):
            frame = ttk.LabelFrame(self.tab_convert, text=label, style="Card.TLabelframe")
            frame.grid(row=r, column=0, columnspan=3, sticky="ew", padx=8, pady=6)
            entry = ttk.Entry(frame, textvariable=var)
            entry.grid(row=0, column=0, sticky="ew", padx=8, pady=8)
            entry.bind("<Enter>", lambda _e, v=var: self.set_status(v.get()))
            ttk.Button(frame, text="Parcourir", command=lambda v=var: self.choose_dir(v)).grid(row=0, column=1, padx=8, pady=8)
            frame.columnconfigure(0, weight=1)
        ttk.Button(self.tab_convert, text="CONVERTIR LES FACTURES", style="Primary.TButton", command=self.convert_folder).grid(row=4, column=0, columnspan=3, sticky="ew", padx=8, pady=16)
        self.convert_status_var = tk.StringVar(value="Prêt à convertir.")
        ttk.Label(self.tab_convert, textvariable=self.convert_status_var, style="Ok.TLabel").grid(row=5, column=0, columnspan=3, sticky="w", padx=8, pady=(4, 0))
        self.convert_progress = ttk.Progressbar(self.tab_convert, mode="indeterminate")
        self.convert_progress.grid(row=6, column=0, columnspan=3, sticky="ew", padx=8, pady=(2, 8))
        self.convert_report = tk.Text(self.tab_convert, height=12, wrap="word", font=("Consolas", 10), relief="flat", borderwidth=8)
        self.convert_report.grid(row=7, column=0, columnspan=3, sticky="nsew", padx=8, pady=8)
        self.tab_convert.columnconfigure(1, weight=1)
        self.tab_convert.rowconfigure(7, weight=1)

    def _build_validation_tab(self):
        ttk.Label(self.tab_validation, text="Validation locale : XML, XMP, PDF/A et veraPDF.").pack(anchor="w", padx=8, pady=8)
        self.validation_file = tk.StringVar()
        line = ttk.Frame(self.tab_validation)
        line.pack(fill="x", padx=8, pady=4)
        ttk.Entry(line, textvariable=self.validation_file).pack(side="left", fill="x", expand=True)
        ttk.Button(line, text="Choisir PDF...", command=self.choose_validation_pdf).pack(side="left", padx=4)
        ttk.Button(line, text="Valider", command=self.validate_pdf_minimal).pack(side="left", padx=4)
        ttk.Button(line, text="Ouvrir PDF", command=self.open_validation_pdf).pack(side="left", padx=4)

        board = ttk.LabelFrame(self.tab_validation, text="Résultat")
        board.pack(fill="x", padx=8, pady=8)
        self.validation_global = ttk.Label(board, text="Aucun fichier validé", font=("Segoe UI", 16, "bold"))
        self.validation_global.grid(row=0, column=0, columnspan=4, sticky="w", padx=10, pady=8)
        self.validation_labels: Dict[str, ttk.Label] = {}
        for i, key in enumerate(["xml", "xmp", "af", "pdfa", "profile", "verapdf"]):
            ttk.Label(board, text=key.upper()).grid(row=1 + i // 3, column=(i % 3) * 2, sticky="w", padx=10, pady=4)
            lab = ttk.Label(board, text="—")
            lab.grid(row=1 + i // 3, column=(i % 3) * 2 + 1, sticky="w", padx=4, pady=4)
            self.validation_labels[key] = lab

        self.validation_report = tk.Text(self.tab_validation, wrap="word", font=("Consolas", 10))
        self.validation_report.pack(fill="both", expand=True, padx=8, pady=8)
        self.validation_report.tag_configure("ok", foreground="green")
        self.validation_report.tag_configure("warn", foreground="#b36b00")
        self.validation_report.tag_configure("err", foreground="red")

    def _build_settings_tab(self):
        ttk.Label(self.tab_settings, text="Paramètres. Les composants techniques sont détectés automatiquement ; ils ne doivent être modifiés qu’en cas de dépannage.").grid(row=0, column=0, columnspan=4, sticky="w", padx=8, pady=8)
        self.default_input_setting = tk.StringVar(value=self.settings.get("default_input_dir", "factures"))
        self.default_output_setting = tk.StringVar(value=self.settings.get("default_output_dir", "facturx"))
        self.default_todo_setting = tk.StringVar(value=self.settings.get("default_todo_dir", "a_traiter"))
        self.verapdf_setting = tk.StringVar(value=self.settings.get("verapdf_path", ""))
        self.ghostscript_setting = tk.StringVar(value=self.settings.get("ghostscript_path", ""))
        self.icc_setting = tk.StringVar(value=self.settings.get("icc_profile_path", ""))

        rows = [
            ("📥 Factures reçues", self.default_input_setting, "dir"),
            ("📤 Factur-X générées", self.default_output_setting, "dir"),
            ("⚠ À vérifier", self.default_todo_setting, "dir"),
            ("veraPDF", self.verapdf_setting, "verapdf"),
            ("Ghostscript", self.ghostscript_setting, "ghostscript"),
            ("Profil ICC sRGB", self.icc_setting, "icc"),
        ]
        for r, (label, var, kind) in enumerate(rows, start=1):
            ttk.Label(self.tab_settings, text=label).grid(row=r, column=0, sticky="w", padx=8, pady=6)
            ttk.Entry(self.tab_settings, textvariable=var, width=78).grid(row=r, column=1, sticky="ew", padx=8, pady=6)
            if kind == "dir":
                ttk.Button(self.tab_settings, text="Choisir...", command=lambda v=var: self.choose_dir(v)).grid(row=r, column=2, padx=4, pady=6)
            elif kind == "verapdf":
                ttk.Button(self.tab_settings, text="Fichier...", command=self.choose_verapdf_file).grid(row=r, column=2, padx=4, pady=6)
                ttk.Button(self.tab_settings, text="Dossier...", command=self.choose_verapdf_dir).grid(row=r, column=3, padx=4, pady=6)
            elif kind == "ghostscript":
                ttk.Button(self.tab_settings, text="Fichier...", command=self.choose_ghostscript_file).grid(row=r, column=2, padx=4, pady=6)
                ttk.Button(self.tab_settings, text="Dossier...", command=self.choose_ghostscript_dir).grid(row=r, column=3, padx=4, pady=6)
            elif kind == "icc":
                ttk.Button(self.tab_settings, text="Fichier...", command=self.choose_icc_file).grid(row=r, column=2, padx=4, pady=6)
                ttk.Button(self.tab_settings, text="Dossier...", command=self.choose_icc_dir).grid(row=r, column=3, padx=4, pady=6)

        ttk.Button(self.tab_settings, text="Tester outils", command=self.test_runtime_tools).grid(row=8, column=1, sticky="w", padx=8, pady=12)
        ttk.Button(self.tab_settings, text="Enregistrer paramètres", command=self.save_settings).grid(row=8, column=1, sticky="e", padx=8, pady=12)
        self.settings_report = tk.Text(self.tab_settings, height=14, wrap="word", font=("Consolas", 10))
        self.settings_report.grid(row=9, column=0, columnspan=4, sticky="nsew", padx=8, pady=8)
        self.tab_settings.columnconfigure(1, weight=1)
        self.tab_settings.rowconfigure(9, weight=1)
        self.refresh_runtime_status()

    def _load_initial_company(self):
        self.refresh_company_list()
        last = self.settings.get("last_company", "")
        names = self.company_manager.list_names()
        if last in names:
            self.select_company(last)
        elif names:
            self.select_company(names[0])
        else:
            self.new_company(first_launch=True)

    def refresh_company_list(self):
        names = self.company_manager.list_names()
        self.company_combo["values"] = names

    def select_company(self, filename: str):
        if not filename:
            return
        self.current_company_file = filename
        self.current_company = self.company_manager.load(filename)
        self.company_combo.set(filename)
        self.settings.set("last_company", filename)
        self.populate_company_form()
        tmpl = self.current_company.get("invoice_template", {})
        self.model_pdf_var.set(tmpl.get("sample_file", ""))
        self.last_analysis = {}
        self.model_report.delete("1.0", "end")
        if tmpl.get("validated"):
            self.model_report.insert("1.0", f"Modèle déjà validé.\nScore : {tmpl.get('score', '')}\nDate : {tmpl.get('validated_at', '')}\nAucune nouvelle validation n'est demandée.")

    def populate_company_form(self):
        for key, _label in FIELDS:
            self.form_vars[key].set(str(self.current_company.get(key, "")))

    def form_to_company(self):
        for key, _label in FIELDS:
            self.current_company[key] = self.form_vars[key].get().strip()

    def save_company(self, show_message: bool = True):
        if not self.current_company_file:
            if show_message:
                messagebox.showwarning("Faxtur", "Aucune société sélectionnée.")
            return
        self.form_to_company()
        self.company_manager.save(self.current_company_file, self.current_company)
        self.set_status("Société enregistrée.")
        if show_message:
            messagebox.showinfo("Faxtur", "Société enregistrée.")

    def new_company(self, first_launch=False):
        win = tk.Toplevel(self)
        win.title("Nouvelle société")
        ttk.Label(win, text="Nom de la société").grid(row=0, column=0, padx=8, pady=8)
        name_var = tk.StringVar(value="Nouvelle société" if first_launch else "")
        ttk.Entry(win, textvariable=name_var, width=42).grid(row=0, column=1, padx=8, pady=8)
        def create():
            name = name_var.get().strip() or "Nouvelle société"
            filename = self.company_manager.new_company(name)
            self.refresh_company_list()
            self.select_company(filename)
            win.destroy()
        ttk.Button(win, text="Créer", command=create).grid(row=1, column=1, sticky="e", padx=8, pady=8)
        win.transient(self)
        win.grab_set()

    def duplicate_company(self):
        if not self.current_company_file:
            return
        new_file = self.company_manager.duplicate(self.current_company_file)
        self.refresh_company_list()
        self.select_company(new_file)

    def delete_company(self):
        if not self.current_company_file:
            return
        if messagebox.askyesno("Faxtur", f"Supprimer {self.current_company_file} ?"):
            self.company_manager.delete(self.current_company_file)
            self.current_company_file = ""
            self.current_company = {}
            self.refresh_company_list()
            names = self.company_manager.list_names()
            if names:
                self.select_company(names[0])

    def choose_model_pdf(self):
        p = filedialog.askopenfilename(title="Facture type", filetypes=[("PDF", "*.pdf")])
        if p:
            self.model_pdf_var.set(p)

    def analyze_model(self):
        if not self.current_company_file:
            messagebox.showwarning("Faxtur", "Choisir une société.")
            return
        self.form_to_company()
        pdf = Path(self.model_pdf_var.get())
        if not pdf.exists():
            messagebox.showerror("Faxtur", "PDF introuvable.")
            return
        try:
            self.last_analysis = engine.analyze_invoice_model(pdf, self.current_company)
            self.model_report.delete("1.0", "end")
            self.model_report.insert("1.0", engine.format_analysis_report(self.last_analysis))
        except Exception:
            self.model_report.delete("1.0", "end")
            self.model_report.insert("1.0", traceback.format_exc())

    def validate_model(self):
        if not self.last_analysis:
            messagebox.showwarning("Faxtur", "Analyser une facture type d'abord.")
            return
        self.form_to_company()
        self.current_company["invoice_template"] = {
            "validated": True,
            "validated_at": engine.datetime.now().isoformat(timespec="seconds"),
            "sample_file": self.model_pdf_var.get(),
            "score": self.last_analysis.get("score", 0),
            "fields": self.last_analysis.get("fields", {}),
            "notes": "Modèle validé manuellement par l'utilisateur.",
            "issuer_match": self.last_analysis.get("issuer_match", {}),
            "model_hints": {},
        }
        self.company_manager.save(self.current_company_file, self.current_company)
        self.set_status("Modèle enregistré.")
        self.refresh_header()

    def choose_dir(self, var: tk.StringVar):
        p = filedialog.askdirectory()
        if p:
            var.set(p)

    def convert_folder(self):
        if not self.current_company_file:
            messagebox.showwarning("Faxtur", "Choisir une société.")
            return
        # Sauvegarde silencieuse : ne pas interrompre la conversion avec une boîte
        # "Société enregistrée". Le message reste affiché uniquement depuis
        # l'onglet Société quand l'utilisateur clique sur Enregistrer.
        self.save_company(show_message=False)
        self.convert_status_var.set("Conversion en cours...")
        self.set_status("Conversion en cours...")
        self.convert_progress.start(10)
        self.update_idletasks()
        try:
            journal = engine.convert_folder(
                self.company_manager.path(self.current_company_file),
                self.input_var.get(),
                self.output_var.get(),
                self.todo_var.get(),
            )
            self.convert_report.delete("1.0", "end")
            self.convert_report.insert("1.0", f"Conversion terminée.\nJournal : {journal}\n")
            summary = self._show_journal_summary(Path(journal))
            self._validate_outputs_from_journal(Path(journal))
            self.convert_status_var.set("Conversion terminée.")
            self.set_status("Conversion terminée.")
            if summary:
                messagebox.showinfo("Faxtur", summary)
            else:
                messagebox.showinfo("Faxtur", f"Conversion terminée.\nJournal : {journal}")
        except Exception:
            self.convert_status_var.set("Erreur pendant la conversion.")
            self.set_status("Erreur pendant la conversion.")
            self.convert_report.delete("1.0", "end")
            self.convert_report.insert("1.0", traceback.format_exc())
        finally:
            self.convert_progress.stop()

    def _show_journal_summary(self, journal: Path):
        try:
            with journal.open("r", encoding="utf-8-sig", newline="") as f:
                rows = list(csv.DictReader(f, delimiter=";"))
            total = len(rows)
            ok = sum(1 for r in rows if r.get("statut") == "OK")
            todo = sum(1 for r in rows if r.get("statut") == "A_TRAITER")
            err = sum(1 for r in rows if r.get("statut") == "ERREUR")
            line = f"Résumé : {ok} OK / {todo} à traiter / {err} erreurs / {total} fichiers"
            self.convert_report.insert("end", f"\n{line}\n")
            return (
                "Conversion terminée.\n\n"
                f"Factures traitées : {total}\n"
                f"Conformes / générées : {ok}\n"
                f"À traiter : {todo}\n"
                f"Erreurs : {err}\n\n"
                f"Journal : {journal}"
            )
        except Exception:
            return ""

    def choose_validation_pdf(self):
        p = filedialog.askopenfilename(title="PDF Factur-X", filetypes=[("PDF", "*.pdf")])
        if p:
            self.validation_file.set(p)

    def validate_pdf_minimal(self):
        pdf = Path(self.validation_file.get())
        self.validation_report.delete("1.0", "end")
        if not pdf.exists():
            self.validation_report.insert("1.0", "PDF introuvable.", "err")
            return
        try:
            result = self.validate_facturx_pdf(pdf)
            self.display_validation_result(result)
            self.save_validation_report(result)
        except Exception:
            self.validation_report.insert("1.0", traceback.format_exc(), "err")

    def open_validation_pdf(self):
        pdf = Path(self.validation_file.get())
        if not pdf.exists():
            messagebox.showwarning("Faxtur", "Aucun PDF valide sélectionné.")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(pdf))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                run_hidden_subprocess(["open", str(pdf)], check=False)
            else:
                run_hidden_subprocess(["xdg-open", str(pdf)], check=False)
        except Exception as exc:
            messagebox.showerror("Faxtur", f"Impossible d'ouvrir le PDF : {exc}")

    def validate_facturx_pdf(self, pdf: Path) -> Dict[str, Any]:
        reader = engine.PdfReader(str(pdf))
        root = reader.trailer.get("/Root", {})
        names = str(root.get("/Names", ""))
        xmp_ok = "/Metadata" in root
        af_ok = "/AF" in root
        embedded_ok = "/EmbeddedFiles" in names or "factur-x.xml" in names

        vp = resolve_verapdf_path(self.settings.get("verapdf_path", ""), self.root_dir)
        verapdf: Dict[str, Any]
        if vp:
            try:
                verapdf = self.run_verapdf_structured(vp, pdf)
            except Exception as exc:
                verapdf = {"available": False, "path": str(vp), "compliant": None, "statement": f"veraPDF non exécuté : {exc}", "errors": []}
        else:
            verapdf = {"available": False, "path": "", "compliant": None, "statement": "veraPDF non configuré", "errors": []}

        pdfa_ok = verapdf.get("compliant") is True
        pdfa_unknown = verapdf.get("compliant") is None
        internal_ok = bool(xmp_ok and af_ok and embedded_ok)
        compliant = internal_ok and pdfa_ok
        if compliant:
            status = "CONFORME"
        elif internal_ok and pdfa_unknown:
            status = "VALIDATION_INCOMPLETE"
        else:
            status = "NON_CONFORME"

        return {
            "file": str(pdf),
            "date": datetime.now().isoformat(timespec="seconds"),
            "status": status,
            "compliant": compliant,
            "checks": {
                "xml": embedded_ok,
                "xmp": xmp_ok,
                "af": af_ok,
                "profile": embedded_ok and xmp_ok,
                "pdfa": pdfa_ok,
                "verapdf": verapdf.get("available") is True,
            },
            "embedded_names": names[:1000],
            "verapdf": verapdf,
        }

    def display_validation_result(self, result: Dict[str, Any]) -> None:
        checks = result.get("checks", {})
        status = result.get("status", "")
        if status == "CONFORME":
            title = "✔ FACTURE CONFORME"
            tag = "ok"
        elif status == "VALIDATION_INCOMPLETE":
            title = "⚠ VALIDATION INCOMPLÈTE"
            tag = "warn"
        else:
            title = "✗ FACTURE NON CONFORME"
            tag = "err"
        self.validation_global.configure(text=title)
        for key, lab in self.validation_labels.items():
            ok = checks.get(key)
            if ok is True:
                lab.configure(text="🟢 OK")
            elif ok is False:
                lab.configure(text="🔴 NON")
            else:
                lab.configure(text="⚪ N/T")

        self.validation_report.delete("1.0", "end")
        self.validation_report.insert("end", title + "\n", tag)
        self.validation_report.insert("end", "=" * 60 + "\n")
        self.validation_report.insert("end", f"Fichier : {result.get('file')}\n")
        self.validation_report.insert("end", f"Date    : {result.get('date')}\n\n")
        for label, key in [("XML embarqué", "xml"), ("XMP Factur-X", "xmp"), ("AFRelationship", "af"), ("Profil", "profile"), ("PDF/A", "pdfa"), ("veraPDF", "verapdf")]:
            ok = checks.get(key)
            mark = "OK" if ok else "NON TESTÉ" if key in {"pdfa", "verapdf"} and result.get("verapdf", {}).get("compliant") is None else "NON"
            self.validation_report.insert("end", f"{label:16s} : {mark}\n", "ok" if ok else "warn" if "TEST" in mark else "err")
        self.validation_report.insert("end", "\n")
        vp = result.get("verapdf", {})
        self.validation_report.insert("end", f"veraPDF : {vp.get('path', '') or 'non configuré'}\n")
        self.validation_report.insert("end", f"Message : {vp.get('statement', '')}\n")
        if vp.get("passed_rules") or vp.get("failed_rules"):
            self.validation_report.insert("end", f"Règles : {vp.get('passed_rules')} passées / {vp.get('failed_rules')} échouées\n")
            self.validation_report.insert("end", f"Contrôles : {vp.get('passed_checks')} réussis / {vp.get('failed_checks')} échoués\n")
        for err in vp.get("errors", [])[:10]:
            self.validation_report.insert("end", f"Erreur {err.get('clause')} : {err.get('description')}\n", "err")

    def save_validation_report(self, result: Dict[str, Any]) -> Path:
        log_dir = self.root_dir / "logs" / "validations"
        log_dir.mkdir(parents=True, exist_ok=True)
        stem = Path(result.get("file", "facture")).stem
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out = log_dir / f"{ts}_{stem}.json"
        with out.open("w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        return out

    def _validate_outputs_from_journal(self, journal: Path):
        try:
            with journal.open("r", encoding="utf-8-sig", newline="") as f:
                rows = list(csv.DictReader(f, delimiter=";"))
        except Exception:
            return
        ok_rows = [r for r in rows if r.get("statut") == "OK" and r.get("message")]
        if not ok_rows:
            return
        self.convert_report.insert("end", "\nValidation automatique des Factur-X générés :\n")
        summary: List[Dict[str, Any]] = []
        for r in ok_rows:
            pdf = Path(r.get("message", ""))
            if not pdf.exists():
                self.convert_report.insert("end", f"- {r.get('fichier')} : PDF généré introuvable\n")
                continue
            try:
                result = self.validate_facturx_pdf(pdf)
                self.save_validation_report(result)
                status = result.get("status")
                symbol = "🟢" if status == "CONFORME" else "🟠" if status == "VALIDATION_INCOMPLETE" else "🔴"
                self.convert_report.insert("end", f"- {pdf.name} : {symbol} {status}\n")
                summary.append({
                    "fichier": pdf.name,
                    "statut_validation": status,
                    "pdfa": result.get("checks", {}).get("pdfa"),
                    "xmp": result.get("checks", {}).get("xmp"),
                    "af": result.get("checks", {}).get("af"),
                    "verapdf": result.get("verapdf", {}).get("statement", ""),
                })
            except Exception as exc:
                self.convert_report.insert("end", f"- {pdf.name} : erreur validation {exc}\n")
        if summary:
            out_csv = Path(self.output_var.get()) / "journal_validation.csv"
            with out_csv.open("w", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(f, fieldnames=["fichier", "statut_validation", "pdfa", "xmp", "af", "verapdf"], delimiter=";")
                writer.writeheader()
                writer.writerows(summary)
            self.convert_report.insert("end", f"Journal validation : {out_csv}\n")

    def choose_verapdf_file(self):
        p = filedialog.askopenfilename(
            title="Choisir veraPDF",
            filetypes=[("veraPDF", "verapdf*.bat verapdf*.exe verapdf*.cmd"), ("Tous fichiers", "*.*")],
        )
        if p:
            self.verapdf_setting.set(p)
            self.refresh_runtime_status()

    def choose_verapdf_dir(self):
        p = filedialog.askdirectory(title="Choisir le dossier veraPDF")
        if p:
            self.verapdf_setting.set(p)
            self.refresh_runtime_status()

    def choose_ghostscript_file(self):
        p = filedialog.askopenfilename(
            title="Choisir Ghostscript",
            filetypes=[("Ghostscript", "gswin*.exe gs.exe gs"), ("Tous fichiers", "*.*")],
        )
        if p:
            self.ghostscript_setting.set(p)
            self.refresh_runtime_status()

    def choose_ghostscript_dir(self):
        p = filedialog.askdirectory(title="Choisir le dossier Ghostscript")
        if p:
            self.ghostscript_setting.set(p)
            self.refresh_runtime_status()

    def choose_icc_file(self):
        p = filedialog.askopenfilename(
            title="Choisir le profil ICC sRGB",
            filetypes=[("Profil ICC", "*.icc *.icm"), ("Tous fichiers", "*.*")],
        )
        if p:
            self.icc_setting.set(p)
            self.refresh_runtime_status()

    def choose_icc_dir(self):
        p = filedialog.askdirectory(title="Choisir le dossier contenant sRGB.icc")
        if p:
            self.icc_setting.set(p)
            self.refresh_runtime_status()

    def refresh_runtime_status(self):
        if not hasattr(self, "settings_report"):
            return

        # Auto-détection et auto-remplissage des champs v5.1.6.
        # L'utilisateur n'a plus besoin de recopier les chemins si les outils
        # sont dans runtime/ ou installés aux emplacements classiques.
        vp = resolve_verapdf_path(self.verapdf_setting.get() if hasattr(self, "verapdf_setting") else self.settings.get("verapdf_path", ""), self.root_dir)
        gs = resolve_ghostscript_path(self.ghostscript_setting.get() if hasattr(self, "ghostscript_setting") else self.settings.get("ghostscript_path", ""), self.root_dir)
        icc = resolve_icc_profile(self.icc_setting.get() if hasattr(self, "icc_setting") else self.settings.get("icc_profile_path", ""), self.root_dir)

        changed = False
        if vp and hasattr(self, "verapdf_setting") and self.verapdf_setting.get().strip() != str(vp):
            self.verapdf_setting.set(str(vp))
            self.settings.data["verapdf_path"] = str(vp)
            changed = True
        if gs and hasattr(self, "ghostscript_setting") and self.ghostscript_setting.get().strip() != str(gs):
            self.ghostscript_setting.set(str(gs))
            self.settings.data["ghostscript_path"] = str(gs)
            changed = True
        if icc and hasattr(self, "icc_setting") and self.icc_setting.get().strip() != str(icc):
            self.icc_setting.set(str(icc))
            self.settings.data["icc_profile_path"] = str(icc)
            changed = True
        if changed:
            self.settings.save()
            configure_runtime_environment(self.root_dir, self.settings.data)

        if hasattr(self, "dashboard_verapdf_var"):
            self.dashboard_verapdf_var.set("veraPDF : OK" if vp else "veraPDF : non trouvé")
        if hasattr(self, "dashboard_pdfa_var"):
            self.dashboard_pdfa_var.set("PDF/A : prêt" if (gs and icc) else "PDF/A : à vérifier")
        self.settings_report.delete("1.0", "end")
        self.settings_report.insert("end", "Diagnostic outils autonomes\n")
        self.settings_report.insert("end", "==========================\n")
        self.settings_report.insert("end", f"veraPDF     : {'OK ' + str(vp) if vp else 'NON TROUVÉ'}\n")
        self.settings_report.insert("end", f"Ghostscript : {'OK ' + str(gs) if gs else 'NON TROUVÉ'}\n")
        self.settings_report.insert("end", f"ICC sRGB    : {'OK ' + str(icc) if icc else 'NON TROUVÉ'}\n\n")
        self.settings_report.insert("end", "Les champs ci-dessus sont remplis automatiquement quand les outils sont détectés.\n")
        self.settings_report.insert("end", "Emplacements embarqués attendus :\n")
        self.settings_report.insert("end", "- runtime\\veraPDF\\verapdf.bat\n")
        self.settings_report.insert("end", "- runtime\\ghostscript\\bin\\gswin64c.exe\n")
        self.settings_report.insert("end", "- runtime\\icc\\sRGB.icc\n")

    def test_runtime_tools(self):
        self.save_settings(show_message=False)
        status = configure_runtime_environment(self.root_dir, self.settings.data)
        self.settings_report.delete("1.0", "end")
        self.settings_report.insert("end", "Test des outils\n")
        self.settings_report.insert("end", "===============\n")
        self.settings_report.insert("end", f"veraPDF     : {status.get('verapdf') or 'NON TROUVÉ'}\n")
        self.settings_report.insert("end", f"Ghostscript : {status.get('ghostscript') or 'NON TROUVÉ'}\n")
        self.settings_report.insert("end", f"ICC sRGB    : {status.get('icc') or 'NON TROUVÉ'}\n\n")
        vp = resolve_verapdf_path(self.settings.get("verapdf_path", ""), self.root_dir)
        if vp:
            try:
                proc = run_hidden_subprocess([str(vp), "--version"], capture_output=True, text=True, timeout=20, shell=vp.suffix.lower() in {".bat", ".cmd"})
                self.settings_report.insert("end", f"veraPDF --version : code {proc.returncode}\n{((proc.stdout or '') + (proc.stderr or ''))[:1200]}\n")
            except Exception:
                self.settings_report.insert("end", "Erreur test veraPDF :\n" + traceback.format_exc() + "\n")
        gs = resolve_ghostscript_path(self.settings.get("ghostscript_path", ""), self.root_dir)
        if gs:
            try:
                proc = run_hidden_subprocess([str(gs), "--version"], capture_output=True, text=True, timeout=20)
                self.settings_report.insert("end", f"Ghostscript --version : code {proc.returncode}\n{((proc.stdout or '') + (proc.stderr or ''))[:1200]}\n")
            except Exception:
                self.settings_report.insert("end", "Erreur test Ghostscript :\n" + traceback.format_exc() + "\n")

    def run_verapdf_structured(self, verapdf_path: Path, pdf: Path) -> Dict[str, Any]:
        """Exécute veraPDF et lit son rapport XML.

        Correction 5.1.5 : certaines installations veraPDF ne sortent pas du XML
        par défaut, ou ajoutent des lignes avant le XML. On force donc le format
        XML et on nettoie la sortie avant parsing.
        """
        use_shell = verapdf_path.suffix.lower() in {".bat", ".cmd"}
        attempts = [
            [str(verapdf_path), "--format", "xml", str(pdf)],
            [str(verapdf_path), "-f", "xml", str(pdf)],
            [str(verapdf_path), str(pdf)],
        ]
        last_out = ""
        last_code = None
        result: Dict[str, Any] = {
            "available": True,
            "path": str(verapdf_path),
            "returncode": None,
            "raw": "",
            "compliant": False,
            "statement": "",
            "passed_rules": "",
            "failed_rules": "",
            "passed_checks": "",
            "failed_checks": "",
            "errors": [],
        }

        env = os.environ.copy()
        # veraPDF installé en local peut dépendre d'un Java embarqué dans runtime/java.
        # On ne compte pas sur le PATH système Windows.
        java = resolve_java_path(self.settings.get("java_path", ""), self.root_dir)
        if java:
            env["JAVA_HOME"] = str(java.parent.parent)
            env["PATH"] = str(java.parent) + os.pathsep + env.get("PATH", "")

        for cmd in attempts:
            proc = run_hidden_subprocess(cmd, capture_output=True, text=True, timeout=120, shell=use_shell, env=env)
            out = (proc.stdout or "") + (proc.stderr or "")
            last_out, last_code = out, proc.returncode
            result["returncode"] = proc.returncode
            result["raw"] = out[:12000]
            if not out.strip():
                continue

            # Nettoyage : Java peut ajouter des warnings avant le XML.
            xml_start = out.find("<?xml")
            if xml_start < 0:
                xml_start = out.find("<report")
            xml_text = out[xml_start:] if xml_start >= 0 else out

            try:
                root = ET.fromstring(xml_text.encode("utf-8"))
                vr = root.find(".//validationReport")
                if vr is not None:
                    result["compliant"] = vr.attrib.get("isCompliant") == "true"
                    result["statement"] = vr.attrib.get("statement", "")
                    details = vr.find("details")
                    if details is not None:
                        result["passed_rules"] = details.attrib.get("passedRules", "")
                        result["failed_rules"] = details.attrib.get("failedRules", "")
                        result["passed_checks"] = details.attrib.get("passedChecks", "")
                        result["failed_checks"] = details.attrib.get("failedChecks", "")
                    for rule in root.findall(".//rule"):
                        if rule.attrib.get("status") == "failed":
                            result["errors"].append({
                                "clause": rule.attrib.get("clause", ""),
                                "failed_checks": rule.attrib.get("failedChecks", ""),
                                "description": rule.findtext("description", ""),
                            })
                    return result
            except Exception:
                # On essaie la variante de commande suivante.
                continue

        result["returncode"] = last_code
        result["raw"] = last_out[:12000]
        if last_out and ("java" in last_out.lower() and ("pas reconnu" in last_out.lower() or "not recognized" in last_out.lower())):
            result["statement"] = "veraPDF trouvé mais Java est absent du PATH. Corrigé en V5.1.7 : vérifiez que runtime\\java contient java.exe ou renseignez java_path dans config/settings.json. Début sortie : " + last_out[:300].replace("\n", " ")
        else:
            result["statement"] = "Sortie veraPDF non interprétable. Vérifiez que le chemin pointe vers verapdf.bat/verapdf.exe. Début sortie : " + (last_out[:300].replace("\n", " ") if last_out else "<vide>")
        return result

    def run_verapdf(self, verapdf_path: Path, pdf: Path) -> str:
        data = self.run_verapdf_structured(verapdf_path, pdf)
        msg = [
            "Validation veraPDF :",
            f"- Conforme PDF/A : {'OUI' if data.get('compliant') else 'NON'}",
            f"- Message : {data.get('statement', '')}",
            f"- Règles : {data.get('passed_rules', '')} passées / {data.get('failed_rules', '')} échouées",
            f"- Contrôles : {data.get('passed_checks', '')} réussis / {data.get('failed_checks', '')} échoués",
        ]
        for err in data.get("errors", [])[:5]:
            msg.append(f"- Erreur {err.get('clause')} ({err.get('failed_checks')} contrôles) : {err.get('description')}")
        return "\n".join(msg) + "\n"

    def save_settings(self, show_message: bool = True):
        self.settings.set("default_input_dir", self.default_input_setting.get().strip() or "factures")
        self.settings.set("default_output_dir", self.default_output_setting.get().strip() or "facturx")
        self.settings.set("default_todo_dir", self.default_todo_setting.get().strip() or "a_traiter")
        self.settings.set("verapdf_path", self.verapdf_setting.get().strip())
        self.settings.set("ghostscript_path", self.ghostscript_setting.get().strip())
        self.settings.set("icc_profile_path", self.icc_setting.get().strip())
        configure_runtime_environment(self.root_dir, self.settings.data)
        self.refresh_runtime_status()
        if show_message:
            messagebox.showinfo("Faxtur", "Paramètres enregistrés.")


def run_app(root_dir: Path) -> int:
    app = Faxtur(root_dir)
    app.mainloop()
    return 0
