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
import sys
from pathlib import Path

try:
    from engine.paths import APP_DIR
except Exception:
    APP_DIR = Path(__file__).resolve().parent

ROOT = APP_DIR
ENGINE_DIR = ROOT / "engine"
if str(ENGINE_DIR) not in sys.path:
    sys.path.insert(0, str(ENGINE_DIR))

# Prépare PATH avant tout appel au moteur gelé : Ghostscript/veraPDF embarqués ou configurés.
try:
    from settings import SettingsStore, configure_runtime_environment
    settings = SettingsStore(ROOT)
    runtime_status = configure_runtime_environment(ROOT, settings.data)
except Exception:
    runtime_status = {}

from gui import run_app

if __name__ == "__main__":
    raise SystemExit(run_app(ROOT))
