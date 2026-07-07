# ============================================================================
# Faxtur
# Copyright © 2026 Frédéric Brouard
#
# This Source Code Form is subject to the terms of the Mozilla Public License,
# v. 2.0. If a copy of the MPL was not distributed with this file, You can
# obtain one at https://mozilla.org/MPL/2.0/
# ============================================================================

from __future__ import annotations

from pathlib import Path
import sys
import shutil


def app_dir() -> Path:
    """
    Dossier racine de l'application.

    En mode source : .../source
    En mode exe PyInstaller : dossier contenant l'exe
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parents[1]


APP_DIR = app_dir()
BASE_DIR = APP_DIR
RUNTIME_DIR = APP_DIR / "runtime"


def first_existing(*paths: Path | str | None) -> Path | None:
    """
    Retourne le premier chemin existant, sinon None.
    Accepte aussi None pour simplifier les appels.
    """
    for p in paths:
        if not p:
            continue
        path = Path(p)
        if path.exists():
            return path
    return None


def from_path(executable_name: str) -> Path | None:
    """Cherche un exécutable dans le PATH système."""
    found = shutil.which(executable_name)
    return Path(found) if found else None


# ---------------------------------------------------------------------------
# Exécutables runtime
# ---------------------------------------------------------------------------

# IMPORTANT : veraPDF est souvent dans runtime/veraPDF/bin/verapdf.bat.
# Si on lance seulement un .bat isolé sans ses dossiers lib/, Java ne trouve pas
# org.verapdf.apps.GreenfieldCliWrapper.
VERAPDF = first_existing(
    RUNTIME_DIR / "veraPDF" / "bin" / "verapdf.bat",
    RUNTIME_DIR / "veraPDF" / "bin" / "verapdf.exe",
    RUNTIME_DIR / "veraPDF" / "verapdf.bat",
    RUNTIME_DIR / "veraPDF" / "verapdf.exe",
    from_path("verapdf"),
    from_path("verapdf.bat"),
    from_path("verapdf.exe"),
)

GHOSTSCRIPT = first_existing(
    RUNTIME_DIR / "ghostscript" / "bin" / "gswin64c.exe",
    RUNTIME_DIR / "ghostscript" / "bin" / "gswin32c.exe",
    RUNTIME_DIR / "ghostscript" / "bin" / "gs.exe",
    from_path("gswin64c"),
    from_path("gswin32c"),
    from_path("gs"),
)

JAVA = first_existing(
    RUNTIME_DIR / "java" / "bin" / "java.exe",
    RUNTIME_DIR / "jre" / "bin" / "java.exe",
    from_path("java"),
    from_path("java.exe"),
)

ICC = RUNTIME_DIR / "icc" / "sRGB.icc"

SETTINGS_JSON = APP_DIR / "settings.json"


def icc_candidates() -> list[Path]:
    """Profils ICC possibles pour la conversion PDF/A."""
    return [
        ICC,
        RUNTIME_DIR / "icc" / "sRGB.icc",
        RUNTIME_DIR / "icc" / "srgb.icc",
        RUNTIME_DIR / "icc" / "sRGB2014.icc",
        RUNTIME_DIR / "icc" / "sRGB_IEC61966-2-1_black_scaled.icc",
        APP_DIR / "sRGB.icc",
        APP_DIR / "srgb.icc",
        Path("C:/Windows/System32/spool/drivers/color/sRGB Color Space Profile.icm"),
        Path("C:/Windows/System32/spool/drivers/color/sRGB.icm"),
    ]
