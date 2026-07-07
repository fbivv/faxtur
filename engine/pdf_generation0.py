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
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional
from xml.sax.saxutils import escape

try:
    from pypdf import PdfReader, PdfWriter
    from pypdf.generic import (
        NameObject, TextStringObject, NumberObject, ArrayObject,
        ByteStringObject, DecodedStreamObject, DictionaryObject,
    )
except Exception as exc:  # pragma: no cover
    PdfReader = None
    PdfWriter = None
    NameObject = TextStringObject = NumberObject = ArrayObject = ByteStringObject = DecodedStreamObject = DictionaryObject = None
    PYPDF_IMPORT_ERROR = exc
else:
    PYPDF_IMPORT_ERROR = None

from engine.paths import GHOSTSCRIPT, first_existing, icc_candidates

APP_NAME = "Faxtur"
VERSION = "1.0.0"


def _run_subprocess_hidden(cmd, **kwargs):
    if sys.platform.startswith("win"):
        kwargs.setdefault("creationflags", getattr(subprocess, "CREATE_NO_WINDOW", 0))
        try:
            si = subprocess.STARTUPINFO()
            si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            kwargs.setdefault("startupinfo", si)
        except Exception:
            pass
    return subprocess.run(cmd, **kwargs)


def pdf_date_now() -> TextStringObject:
    return TextStringObject(datetime.utcnow().strftime("D:%Y%m%d%H%M%SZ"))

def xmp_packet(title: str) -> bytes:
    """XMP PDF/A-3 + Factur-X avec schema d'extension declare."""
    now = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    title_e = escape(title)
    producer_e = escape(APP_NAME + ' ' + VERSION)
    x = f"""<?xpacket begin="\ufeff" id="W5M0MpCehiHzreSzNTczkc9d"?>
<x:xmpmeta xmlns:x="adobe:ns:meta/">
  <rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">
    <rdf:Description rdf:about=""
      xmlns:pdfaid="http://www.aiim.org/pdfa/ns/id/"
      xmlns:pdfaExtension="http://www.aiim.org/pdfa/ns/extension/"
      xmlns:pdfaSchema="http://www.aiim.org/pdfa/ns/schema#"
      xmlns:pdfaProperty="http://www.aiim.org/pdfa/ns/property#">
      <pdfaid:part>3</pdfaid:part>
      <pdfaid:conformance>B</pdfaid:conformance>
      <pdfaExtension:schemas>
        <rdf:Bag>
          <rdf:li rdf:parseType="Resource">
            <pdfaSchema:schema>Factur-X PDFA Extension Schema</pdfaSchema:schema>
            <pdfaSchema:namespaceURI>urn:factur-x:pdfa:CrossIndustryDocument:invoice:1p0#</pdfaSchema:namespaceURI>
            <pdfaSchema:prefix>fx</pdfaSchema:prefix>
            <pdfaSchema:property>
              <rdf:Seq>
                <rdf:li rdf:parseType="Resource"><pdfaProperty:name>DocumentFileName</pdfaProperty:name><pdfaProperty:valueType>Text</pdfaProperty:valueType><pdfaProperty:category>external</pdfaProperty:category><pdfaProperty:description>name of the embedded XML invoice file</pdfaProperty:description></rdf:li>
                <rdf:li rdf:parseType="Resource"><pdfaProperty:name>DocumentType</pdfaProperty:name><pdfaProperty:valueType>Text</pdfaProperty:valueType><pdfaProperty:category>external</pdfaProperty:category><pdfaProperty:description>document type</pdfaProperty:description></rdf:li>
                <rdf:li rdf:parseType="Resource"><pdfaProperty:name>Version</pdfaProperty:name><pdfaProperty:valueType>Text</pdfaProperty:valueType><pdfaProperty:category>external</pdfaProperty:category><pdfaProperty:description>Factur-X version</pdfaProperty:description></rdf:li>
                <rdf:li rdf:parseType="Resource"><pdfaProperty:name>ConformanceLevel</pdfaProperty:name><pdfaProperty:valueType>Text</pdfaProperty:valueType><pdfaProperty:category>external</pdfaProperty:category><pdfaProperty:description>Factur-X conformance level</pdfaProperty:description></rdf:li>
              </rdf:Seq>
            </pdfaSchema:property>
          </rdf:li>
        </rdf:Bag>
      </pdfaExtension:schemas>
    </rdf:Description>
    <rdf:Description rdf:about="" xmlns:dc="http://purl.org/dc/elements/1.1/">
      <dc:title><rdf:Alt><rdf:li xml:lang="x-default">{title_e}</rdf:li></rdf:Alt></dc:title>
    </rdf:Description>
    <rdf:Description rdf:about="" xmlns:xmp="http://ns.adobe.com/xap/1.0/" xmp:CreateDate="{now}" xmp:ModifyDate="{now}" xmp:MetadataDate="{now}"/>
    <rdf:Description rdf:about="" xmlns:pdf="http://ns.adobe.com/pdf/1.3/" pdf:Producer="{producer_e}"/>
    <rdf:Description rdf:about="" xmlns:fx="urn:factur-x:pdfa:CrossIndustryDocument:invoice:1p0#">
      <fx:DocumentType>INVOICE</fx:DocumentType>
      <fx:DocumentFileName>factur-x.xml</fx:DocumentFileName>
      <fx:Version>1.0</fx:Version>
      <fx:ConformanceLevel>BASIC</fx:ConformanceLevel>
    </rdf:Description>
  </rdf:RDF>
</x:xmpmeta>
<?xpacket end="w"?>"""
    return x.encode("utf-8")

def _ghostscript_candidates() -> list[Path]:
    """Chemins possibles vers l'exécutable Ghostscript."""
    candidates: list[Path] = []

    if GHOSTSCRIPT:
        candidates.append(Path(GHOSTSCRIPT))

    for name in ("gswin64c", "gswin32c", "gs"):
        found = shutil.which(name)
        if found:
            candidates.append(Path(found))

    if os.name == "nt":
        for base in (Path("C:/Program Files/gs"), Path("C:/Program Files (x86)/gs")):
            if base.exists():
                candidates.extend(base.glob("gs*/bin/gswin64c.exe"))
                candidates.extend(base.glob("gs*/bin/gswin32c.exe"))

    # Déduplication en conservant l'ordre.
    result: list[Path] = []
    seen: set[str] = set()
    for c in candidates:
        try:
            key = str(c.resolve()).lower()
        except Exception:
            key = str(c).lower()
        if key not in seen:
            seen.add(key)
            result.append(c)
    return result


def _select_ghostscript() -> Path:
    """Sélectionne l'exécutable Ghostscript.

    Important : on ne bloque pas si gsdll64.dll n'est pas dans le même dossier
    que gswin64c.exe. Selon l'installation Windows, la DLL peut être trouvée
    par le PATH système ou être chargée autrement par Ghostscript.

    Si Ghostscript échoue réellement au lancement, l'erreur complète sera
    remontée par normalize_pdfa_with_ghostscript().
    """
    checked: list[str] = []

    for exe in _ghostscript_candidates():
        if exe.exists():
            return exe
        checked.append(f"{exe} : absent")

    raise FileNotFoundError(
        "Ghostscript introuvable.\n"
        "Chemins testés :\n- " + "\n- ".join(checked)
    )


def normalize_pdfa_with_ghostscript(input_pdf: str | Path) -> Path:
    """Convertit le PDF source en PDF/A-3B avec Ghostscript.

    Si Ghostscript échoue, on lève une erreur explicite. On ne revient jamais
    silencieusement au PDF original, sinon veraPDF signale ensuite polices non
    embarquées + DeviceRGB/DeviceGray.
    """
    gs_path = _select_ghostscript()

    icc_profile = first_existing(*icc_candidates())
    if not icc_profile:
        raise FileNotFoundError("Profil ICC introuvable : vérifiez runtime/icc/sRGB.icc")

    input_pdf = Path(input_pdf)
    if not input_pdf.exists():
        raise FileNotFoundError(f"PDF source introuvable : {input_pdf}")

    tmpdir = Path(tempfile.gettempdir())
    out_pdf = tmpdir / f"facturx_pdfa_{uuid.uuid4().hex}.pdf"
    ps_file = tmpdir / f"PDFA_def_{uuid.uuid4().hex}.ps"

    icc_ps = str(icc_profile).replace("\\", "/")
    ps_file.write_text(f"""%!
/ICCProfile ({icc_ps}) def
[ /_objdef {{icc_PDFA}} /type /stream /OBJ pdfmark
[ {{icc_PDFA}} << /N 3 >> /PUT pdfmark
[ {{icc_PDFA}} ICCProfile (r) file /PUT pdfmark
[ /_objdef {{OutputIntent_PDFA}} /type /dict /OBJ pdfmark
[ {{OutputIntent_PDFA}} <<
  /Type /OutputIntent
  /S /GTS_PDFA1
  /DestOutputProfile {{icc_PDFA}}
  /OutputConditionIdentifier (sRGB IEC61966-2.1)
  /Info (sRGB IEC61966-2.1)
>> /PUT pdfmark
[ {{Catalog}} << /OutputIntents [ {{OutputIntent_PDFA}} ] >> /PUT pdfmark
""", encoding="ascii")

    cmd = [
        str(gs_path),
        "-dBATCH",
        "-dNOPAUSE",
        "-dNOOUTERSAVE",
        "-dNOSAFER",
        "-sDEVICE=pdfwrite",
        "-dPDFA=3",
        "-dPDFACompatibilityPolicy=1",
        "-dEmbedAllFonts=true",
        "-dSubsetFonts=true",
        "-dCompressFonts=true",
        "-dUseCIEColor",
        "-sProcessColorModel=DeviceRGB",
        "-sColorConversionStrategy=RGB",
        "-sColorConversionStrategyForImages=RGB",
        "-sDefaultRGBProfile=" + str(icc_profile),
        "-sOutputICCProfile=" + str(icc_profile),
        f"-sOutputFile={out_pdf}",
        str(ps_file),
        str(input_pdf),
    ]

    env = os.environ.copy()
    gs_bin = str(gs_path.parent)
    env["PATH"] = gs_bin + os.pathsep + env.get("PATH", "")

    try:
        result = _run_subprocess_hidden(
            cmd,
            check=False,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=180,
            cwd=str(gs_path.parent),
            env=env,
        )
        if result.returncode != 0:
            err = (result.stderr or b"").decode("utf-8", errors="replace")
            out = (result.stdout or b"").decode("utf-8", errors="replace")
            raise RuntimeError(
                "Ghostscript a échoué pendant la conversion PDF/A.\n"
                f"Ghostscript utilisé : {gs_path}\n"
                + (err or out)[:4000]
            )

        if not out_pdf.exists() or out_pdf.stat().st_size == 0:
            raise RuntimeError("Ghostscript n'a pas produit de PDF/A de sortie.")

        return out_pdf
    finally:
        try:
            ps_file.unlink(missing_ok=True)
        except Exception:
            pass


def add_pdfa_output_intent(writer: PdfWriter) -> None:
    """Ajoute explicitement un OutputIntent sRGB au catalogue final pypdf.

    pypdf réécrit le fichier après Ghostscript pour ajouter factur-x.xml.
    Cette fonction remet un OutputIntent propre dans le PDF final, afin que
    veraPDF ne signale pas DeviceRGB/DeviceGray sans profil de sortie.
    """
    icc_profile = first_existing(*icc_candidates())
    if not icc_profile:
        raise FileNotFoundError("Profil ICC introuvable : vérifiez runtime/icc/sRGB.icc")

    icc_bytes = Path(icc_profile).read_bytes()
    icc_stream = DecodedStreamObject()
    icc_stream.set_data(icc_bytes)
    icc_stream.update({
        NameObject("/N"): NumberObject(3),
        NameObject("/Alternate"): NameObject("/DeviceRGB"),
    })
    icc_ref = writer._add_object(icc_stream)

    output_intent = DictionaryObject({
        NameObject("/Type"): NameObject("/OutputIntent"),
        NameObject("/S"): NameObject("/GTS_PDFA1"),
        NameObject("/OutputConditionIdentifier"): TextStringObject("sRGB IEC61966-2.1"),
        NameObject("/Info"): TextStringObject("sRGB IEC61966-2.1"),
        NameObject("/DestOutputProfile"): icc_ref,
    })
    writer._root_object[NameObject("/OutputIntents")] = ArrayObject([writer._add_object(output_intent)])

# ---------------------------------------------------------------------------
# XML Factur-X simple
# ---------------------------------------------------------------------------


def embed_xml_in_pdf(input_pdf: str | Path, output_pdf: str | Path, xml_bytes: bytes) -> None:
    """Ajoute factur-x.xml en pièce jointe PDF/A-3 de façon explicite.

    Correction v4.6 : on ne s'appuie plus sur writer.add_attachment(), car selon
    les versions de pypdf il crée un FileSpec direct et ne renseigne pas toujours
    tous les éléments vérifiés par veraPDF/FactPulse. On crée donc manuellement :
      - EmbeddedFile stream avec /Subtype /text#2Fxml et /Params ;
      - FileSpec indirect avec /F, /UF, /EF, /AFRelationship et /Subtype ;
      - Name tree /EmbeddedFiles ;
      - Catalog /AF pointant vers le même FileSpec indirect.
    """
    if PdfReader is None or PdfWriter is None:
        raise RuntimeError("pypdf absent. Installez : python -m pip install pypdf")

    normalized = normalize_pdfa_with_ghostscript(input_pdf)
    source_pdf = normalized

    reader = PdfReader(str(source_pdf))
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    # Remettre explicitement un OutputIntent dans le PDF final réécrit par pypdf.
    add_pdfa_output_intent(writer)

    title = Path(input_pdf).stem
    writer.add_metadata({
        "/Title": title,
        "/Subject": "Factur-X generated file",
        "/Producer": f"{APP_NAME} {VERSION}",
    })

    # XMP document + métadonnées Factur-X.
    meta = DecodedStreamObject()
    meta.set_data(xmp_packet(title))
    meta.update({NameObject("/Type"): NameObject("/Metadata"), NameObject("/Subtype"): NameObject("/XML")})
    writer._root_object[NameObject("/Metadata")] = writer._add_object(meta)

    filename = "factur-x.xml"
    mime_name = NameObject("/text/xml")

    # 1) Flux EmbeddedFile. Le Subtype doit être porté par ce flux.
    embedded_stream = DecodedStreamObject()
    embedded_stream.set_data(xml_bytes)
    embedded_stream.update({
        NameObject("/Type"): NameObject("/EmbeddedFile"),
        NameObject("/Subtype"): mime_name,
        NameObject("/Params"): DictionaryObject({
            NameObject("/Size"): NumberObject(len(xml_bytes)),
            NameObject("/ModDate"): pdf_date_now(),
        }),
    })
    embedded_ref = writer._add_object(embedded_stream)

    # 2) File specification indirect. On ajoute aussi /Subtype ici par prudence,
    # certains validateurs parlent de "file specification dictionary".
    filespec = DictionaryObject({
        NameObject("/Type"): NameObject("/Filespec"),
        NameObject("/F"): TextStringObject(filename),
        NameObject("/UF"): TextStringObject(filename),
        NameObject("/Desc"): TextStringObject("Factur-X Invoice Data"),
        NameObject("/AFRelationship"): NameObject("/Alternative"),
        NameObject("/EF"): DictionaryObject({
        NameObject("/F"): embedded_ref,
    }),
})
    filespec_ref = writer._add_object(filespec)

    # 3) Name tree /EmbeddedFiles.
    embedded_files_tree = DictionaryObject({
        NameObject("/Names"): ArrayObject([TextStringObject(filename), filespec_ref])
    })
    embedded_files_ref = writer._add_object(embedded_files_tree)
    names = DictionaryObject({NameObject("/EmbeddedFiles"): embedded_files_ref})
    writer._root_object[NameObject("/Names")] = writer._add_object(names)

    # 4) Associated Files au niveau Catalog : même FileSpec indirect.
    writer._root_object[NameObject("/AF")] = ArrayObject([filespec_ref])

    # Identifiant trailer requis PDF/A.
    writer._ID = ArrayObject([ByteStringObject(os.urandom(16)), ByteStringObject(os.urandom(16))])

    Path(output_pdf).parent.mkdir(parents=True, exist_ok=True)
    with open(output_pdf, "wb") as f:
        writer.write(f)

    try:
        Path(normalized).unlink(missing_ok=True)
    except Exception:
        pass

# ---------------------------------------------------------------------------
# Conversion dossier
# ---------------------------------------------------------------------------
