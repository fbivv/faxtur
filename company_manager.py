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
import re
import shutil
from pathlib import Path
from typing import Any, Dict, List, Optional

from engine import engine_v1 as engine


def slugify(name: str) -> str:
    s = (name or "societe").strip().lower()
    s = re.sub(r"[^a-z0-9]+", "_", s)
    s = re.sub(r"_+", "_", s).strip("_")
    return s or "societe"


class CompanyManager:
    def __init__(self, root: Path):
        self.root = Path(root)
        self.dir = self.root / "companies"
        self.dir.mkdir(parents=True, exist_ok=True)

    def list_files(self) -> List[Path]:
        return sorted(self.dir.glob("*.json"))

    def list_names(self) -> List[str]:
        return [p.name for p in self.list_files()]

    def path(self, filename: str) -> Path:
        p = Path(filename)
        if p.is_absolute():
            return p
        return self.dir / filename

    def load(self, filename: str) -> Dict[str, Any]:
        return engine.load_config(self.path(filename))

    def save(self, filename: str, data: Dict[str, Any]) -> Path:
        p = self.path(filename)
        engine.save_config(p, data)
        return p

    def create_filename(self, name: str) -> str:
        base = slugify(name)
        candidate = self.dir / f"{base}.json"
        i = 2
        while candidate.exists():
            candidate = self.dir / f"{base}_{i}.json"
            i += 1
        return candidate.name

    def new_company(self, name: str) -> str:
        cfg = json.loads(json.dumps(engine.DEFAULT_COMPANY))
        cfg["name"] = name
        filename = self.create_filename(name)
        self.save(filename, cfg)
        return filename

    def delete(self, filename: str) -> None:
        self.path(filename).unlink(missing_ok=True)

    def duplicate(self, filename: str) -> str:
        src = self.path(filename)
        new_name = src.stem + "_copie.json"
        dst = self.dir / new_name
        i = 2
        while dst.exists():
            new_name = f"{src.stem}_copie_{i}.json"
            dst = self.dir / new_name
            i += 1
        shutil.copy2(src, dst)
        return new_name
