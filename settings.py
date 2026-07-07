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

import json
import os
import shutil
from pathlib import Path
from typing import Any, Dict, Optional, Iterable
from engine.paths import VERAPDF, GHOSTSCRIPT, JAVA, ICC, SETTINGS_JSON

from engine.paths import APP_DIR, VERAPDF, GHOSTSCRIPT, ICC, JAVA, first_existing, icc_candidates

def get_user_desktop_path() -> Path:
    """Retourne le vrai Bureau de l'utilisateur courant.

    Sous Windows, cela respecte aussi les redirections OneDrive/Bureau
    lorsque la variable USERPROFILE/OneDrive est configurée.
    """
    if os.name == "nt":
        # 1) API Windows via ctypes : plus fiable que Path.home()/Desktop.
        try:
            import ctypes
            from ctypes import wintypes
            CSIDL_DESKTOPDIRECTORY = 0x10
            SHGFP_TYPE_CURRENT = 0
            buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
            ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_DESKTOPDIRECTORY, None, SHGFP_TYPE_CURRENT, buf)
            if buf.value:
                return Path(buf.value)
        except Exception:
            pass
        # 2) Fallbacks courants.
        for env_name in ("OneDriveCommercial", "OneDrive", "USERPROFILE"):
            base = os.environ.get(env_name)
            if base:
                for name in ("Bureau", "Desktop"):
                    candidate = Path(base) / name
                    if candidate.exists():
                        return candidate
    return Path.home() / "Desktop"


def default_work_dirs() -> Dict[str, str]:
    desktop = get_user_desktop_path()
    return {
        "default_input_dir": str(desktop / "Factures"),
        "default_output_dir": str(desktop / "Facture-X"),
        "default_todo_dir": str(desktop / "A traiter"),
    }


DEFAULT_SETTINGS = {
    "last_company": "",
    "default_input_dir": "",
    "default_output_dir": "",
    "default_todo_dir": "",
    "verapdf_path": "",
    "ghostscript_path": "",
    "icc_profile_path": "",
    "java_path": "",
    "prefer_embedded_tools": True,
}

class SettingsStore:
    def __init__(self, root: Path):
        self.root = Path(root)
        self.path = self.root / "config" / "settings.json"
        self.data: Dict[str, Any] = {}
        self.load()

    def load(self) -> Dict[str, Any]:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        if self.path.exists():
            try:
                self.data = {**DEFAULT_SETTINGS, **json.loads(self.path.read_text(encoding="utf-8"))}
            except Exception:
                self.data = dict(DEFAULT_SETTINGS)
        else:
            self.data = dict(DEFAULT_SETTINGS)
            self.save()
        self._migrate_work_dirs_to_user_desktop()
        return self.data

    def _migrate_work_dirs_to_user_desktop(self) -> None:
        """Évite de proposer AppData comme dossier métier.

        Les dossiers applicatifs restent dans LOCALAPPDATA, mais les factures
        doivent être sur le Bureau de l'utilisateur courant par défaut.
        On ne remplace que les anciennes valeurs techniques par défaut :
        - factures / facturx / a_traiter ;
        - chemins situés sous le dossier d'installation de l'application.
        Les chemins personnalisés par l'utilisateur sont conservés.
        """
        defaults = default_work_dirs()
        old_names = {
            "default_input_dir": {"", "factures"},
            "default_output_dir": {"", "facturx"},
            "default_todo_dir": {"", "a_traiter", "a traiter", "à traiter"},
        }
        changed = False
        root_resolved = self.root.resolve()
        for key, replacement in defaults.items():
            raw = str(self.data.get(key, "") or "").strip()
            replace = raw.lower() in old_names.get(key, set())
            if not replace and raw:
                try:
                    p = Path(raw)
                    # Ancien comportement : C:\...\Faxtur\factures etc.
                    if p.is_absolute() and root_resolved in p.resolve().parents:
                        replace = True
                except Exception:
                    pass
            if replace:
                self.data[key] = replacement
                try:
                    Path(replacement).mkdir(parents=True, exist_ok=True)
                except Exception:
                    pass
                changed = True
        if changed:
            self.save()

    def save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(json.dumps(self.data, ensure_ascii=False, indent=2), encoding="utf-8")

    def get(self, key: str, default: Any = None) -> Any:
        return self.data.get(key, default)

    def set(self, key: str, value: Any) -> None:
        self.data[key] = value
        self.save()


def _clean_path(raw_path: str | Path | None) -> Optional[Path]:
    if not raw_path:
        return None
    raw = str(raw_path).strip().strip('"')
    if not raw:
        return None
    return Path(raw)


def _first_existing(paths: Iterable[Path]) -> Optional[Path]:
    for p in paths:
        if p and p.exists() and p.is_file():
            return p
    return None


def resolve_verapdf_path(raw_path: str | Path | None, root: str | Path | None = None) -> Optional[Path]:
    """Retourne l'exécutable veraPDF si le chemin est un fichier ou un dossier.

    Priorité :
    1. chemin saisi par l'utilisateur ;
    2. veraPDF embarqué dans runtime/veraPDF ;
    3. emplacements Program Files ;
    4. PATH.
    """
    candidates = []
    p = _clean_path(raw_path)
    if p:
        if p.is_file():
            return p
        if p.is_dir():
            candidates.extend([
                p / "verapdf.bat",
                p / "verapdf.exe",
                p / "verapdf.cmd",
                p / "bin" / "verapdf.bat",
                p / "bin" / "verapdf.exe",
                p / "bin" / "verapdf.cmd",
            ])

    candidates.extend([VERAPDF])

    if root:
        r = Path(root)
        candidates.extend([
            r / "runtime" / "veraPDF" / "verapdf.bat",
            r / "runtime" / "veraPDF" / "verapdf.exe",
            r / "runtime" / "veraPDF" / "verapdf.cmd",
            r / "runtime" / "veraPDF" / "bin" / "verapdf.bat",
            r / "runtime" / "veraPDF" / "bin" / "verapdf.exe",
            r / "runtime" / "veraPDF" / "bin" / "verapdf.cmd",
        ])

    program_files = [os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")]
    for base in program_files:
        if base:
            b = Path(base)
            candidates.extend([
                b / "veraPDF" / "verapdf.bat",
                b / "veraPDF" / "verapdf.exe",
                b / "veraPDF" / "verapdf.cmd",
                b / "veraPDF" / "bin" / "verapdf.bat",
                b / "veraPDF" / "bin" / "verapdf.exe",
                b / "veraPDF" / "bin" / "verapdf.cmd",
                b / "vera" / "verapdf.bat",
                b / "vera" / "verapdf.exe",
            ])
    found = _first_existing(candidates)
    if found:
        return found
    for name in ("verapdf", "verapdf.bat", "verapdf.exe", "verapdf.cmd"):
        found_s = shutil.which(name)
        if found_s:
            return Path(found_s)
    return None


def resolve_ghostscript_path(raw_path: str | Path | None, root: str | Path | None = None) -> Optional[Path]:
    """Retourne gswin64c.exe / gswin32c.exe / gs si chemin fichier ou dossier."""
    candidates = []
    p = _clean_path(raw_path)
    if p:
        if p.is_file():
            return p
        if p.is_dir():
            candidates.extend([
                p / "gswin64c.exe",
                p / "gswin32c.exe",
                p / "gs.exe",
                p / "gs",
                p / "bin" / "gswin64c.exe",
                p / "bin" / "gswin32c.exe",
                p / "bin" / "gs.exe",
                p / "bin" / "gs",
            ])
    candidates.extend([VERAPDF])

    candidates.extend([GHOSTSCRIPT])

    if root:
        r = Path(root)
        candidates.extend([
            r / "runtime" / "ghostscript" / "bin" / "gswin64c.exe",
            r / "runtime" / "ghostscript" / "bin" / "gswin32c.exe",
            r / "runtime" / "ghostscript" / "bin" / "gs.exe",
            r / "runtime" / "ghostscript" / "bin" / "gs",
            r / "runtime" / "ghostscript" / "gswin64c.exe",
            r / "runtime" / "ghostscript" / "gswin32c.exe",
        ])
    for base in [os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")]:
        if base:
            gs_root = Path(base) / "gs"
            if gs_root.exists():
                for child in sorted(gs_root.glob("gs*"), reverse=True):
                    candidates.extend([
                        child / "bin" / "gswin64c.exe",
                        child / "bin" / "gswin32c.exe",
                    ])
    found = _first_existing(candidates)
    if found:
        return found
    for name in ("gswin64c", "gswin64c.exe", "gswin32c", "gswin32c.exe", "gs"):
        found_s = shutil.which(name)
        if found_s:
            return Path(found_s)
    return None


def resolve_icc_profile(raw_path: str | Path | None, root: str | Path | None = None) -> Optional[Path]:
    candidates = []
    p = _clean_path(raw_path)
    if p:
        if p.is_file():
            return p
        if p.is_dir():
            candidates.extend([p / "sRGB.icc", p / "srgb.icc", p / "sRGB_IEC61966-2-1_black_scaled.icc"])
    candidates.extend([VERAPDF])

    candidates.extend(icc_candidates())

    if root:
        r = Path(root)
        candidates.extend([
            r / "runtime" / "icc" / "sRGB.icc",
            r / "runtime" / "icc" / "srgb.icc",
            r / "engine" / "icc" / "sRGB.icc",
            r / "engine" / "icc" / "srgb.icc",
        ])
    candidates.extend([
        Path("/usr/share/color/icc/ghostscript/srgb.icc"),
        Path("/usr/share/color/icc/ghostscript/esrgb.icc"),
        Path("/usr/share/texlive/texmf-dist/tex/generic/colorprofiles/sRGB.icc"),
    ])
    return _first_existing(candidates)



def resolve_java_path(raw_path: str | Path | None, root: str | Path | None = None) -> Optional[Path]:
    """Retourne java.exe si disponible.

    Priorité : chemin saisi, runtime/java embarqué, PATH.
    Important pour veraPDF : verapdf.bat appelle "java". Sans Java dans le PATH,
    veraPDF échoue avec : '"java" n'est pas reconnu...'.
    """
    candidates = []
    p = _clean_path(raw_path)
    if p:
        if p.is_file():
            return p
        if p.is_dir():
            candidates.extend([
                p / "java.exe",
                p / "bin" / "java.exe",
            ])
            candidates.extend(sorted(p.glob("**/bin/java.exe"), reverse=True))
    candidates.extend([JAVA])

    if root:
        r = Path(root)
        java_root = r / "runtime" / "java"
        candidates.extend([
            java_root / "java.exe",
            java_root / "bin" / "java.exe",
        ])
        if java_root.exists():
            candidates.extend(sorted(java_root.glob("**/bin/java.exe"), reverse=True))
    found = _first_existing(candidates)
    if found:
        return found
    found_s = shutil.which("java") or shutil.which("java.exe")
    if found_s:
        return Path(found_s)
    return None

def configure_runtime_environment(root: str | Path, settings: Dict[str, Any]) -> Dict[str, str]:
    """Prépare PATH pour que le moteur gelé trouve les outils embarqués.

    Le moteur v4.7.3 appelle shutil.which("gs"/"gswin64c").
    On ne le modifie pas : on ajoute simplement les dossiers runtime/configurés au PATH.
    """
    root = Path(root) if root else APP_DIR
    gs = resolve_ghostscript_path(settings.get("ghostscript_path"), root)
    vp = resolve_verapdf_path(settings.get("verapdf_path"), root)
    icc = resolve_icc_profile(settings.get("icc_profile_path"), root)
    java = resolve_java_path(settings.get("java_path"), root)

    path_parts = []
    if gs:
        path_parts.append(str(gs.parent))
    if vp:
        path_parts.append(str(vp.parent))
    if java:
        path_parts.append(str(java.parent))
        os.environ["JAVA_HOME"] = str(java.parent.parent)
    # Dossiers runtime même vides : utile après installation portable.
    path_parts.extend([
        str(root / "runtime" / "ghostscript" / "bin"),
        str(root / "runtime" / "veraPDF"),
        str(root / "runtime" / "veraPDF" / "bin"),
        str(root / "runtime" / "java" / "bin"),
    ])
    current = os.environ.get("PATH", "")
    for part in reversed(path_parts):
        if part and part not in current:
            current = part + os.pathsep + current
    os.environ["PATH"] = current
    if icc:
        os.environ["FACTURX_ICC_PROFILE"] = str(icc)
    return {
        "ghostscript": str(gs or ""),
        "verapdf": str(vp or ""),
        "icc": str(icc or ""),
        "java": str(java or ""),
    }
