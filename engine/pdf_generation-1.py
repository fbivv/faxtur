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
      <fx:ConformanceLevel>BASIC WL</fx:ConformanceLevel>
    </rdf:Description>
  </rdf:RDF>
</x:xmpmeta>
<?xpacket end="w"?>"""
    return x.encode("utf-8")

def normalize_pdfa_with_ghostscript(input_pdf: str | Path) -> Optional[Path]:
    """Convertit le PDF source en PDF/A-3b avec OutputIntent sRGB.

    Correction v4.7.2 : recherche robuste d'un profil ICC, y compris :
    - profil embarqué dans ./icc/srgb.icc ;
    - profils Ghostscript Windows ;
    - profils Linux usuels.

    Sans ICC, Ghostscript peut produire un PDF visuellement correct mais veraPDF
    signale DeviceRGB/DeviceGray without OutputIntent. On refuse donc la
    normalisation si aucun ICC n'est trouvé.
    """
    gs = str(GHOSTSCRIPT) if GHOSTSCRIPT.exists() else (shutil.which("gs") or shutil.which("gswin64c") or shutil.which("gswin32c"))

    # Recherche Ghostscript Windows si non présent dans le PATH.
    if not gs and os.name == "nt":
        possible_gs = []
        for base in [Path("C:/Program Files/gs"), Path("C:/Program Files (x86)/gs")]:
            if base.exists():
                possible_gs += list(base.glob("gs*/bin/gswin64c.exe"))
                possible_gs += list(base.glob("gs*/bin/gswin32c.exe"))
        if possible_gs:
            gs = str(sorted(possible_gs)[-1])

    if not gs:
        return None

    icc_profile = first_existing(*icc_candidates())
    if not icc_profile:
        raise FileNotFoundError("Profil ICC introuvable")
    icc_profile = str(icc_profile)
    tmpdir = Path(tempfile.gettempdir())
    out_pdf = tmpdir / f"facturx_pdfa_{uuid.uuid4().hex}.pdf"
    ps_file = tmpdir / f"PDFA_def_{uuid.uuid4().hex}.ps"

    # Utiliser des slashs évite les problèmes d'échappement Windows dans PostScript.
    icc_ps = icc_profile.replace("\\", "/")
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
        gs,
        "-dBATCH", "-dNOPAUSE", "-dNOOUTERSAVE", "-dNOSAFER",
        "-sDEVICE=pdfwrite",
        "-dPDFA=3",
        "-dPDFACompatibilityPolicy=1",
        "-dEmbedAllFonts=true",
        "-dSubsetFonts=true",
        "-sProcessColorModel=DeviceRGB",
        "-sColorConversionStrategy=RGB",
        "-sColorConversionStrategyForImages=RGB",
        f"-sOutputICCProfile={icc_profile}",
        f"-sOutputFile={out_pdf}",
        str(ps_file), str(input_pdf),
    ]
    try:
        _run_subprocess_hidden(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=180)
        return out_pdf if out_pdf.exists() and out_pdf.stat().st_size > 0 else None
    except Exception:
        return None
    finally:
        try:
            ps_file.unlink(missing_ok=True)
        except Exception:
            pass

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
    source_pdf = normalized if normalized else Path(input_pdf)

    reader = PdfReader(str(source_pdf))
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    # Conserver les OutputIntents créés par Ghostscript pour PDF/A.
    try:
        src_root = reader.trailer["/Root"]
        if "/OutputIntents" in src_root:
            writer._root_object[NameObject("/OutputIntents")] = src_root["/OutputIntents"].clone(writer)
    except Exception:
        pass

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
        NameObject("/Desc"): TextStringObject("Factur-X XML invoice data"),
        NameObject("/AFRelationship"): NameObject("/Alternative"),
        NameObject("/Subtype"): mime_name,
        NameObject("/EF"): DictionaryObject({
            NameObject("/F"): embedded_ref,
            NameObject("/UF"): embedded_ref,
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

    if normalized:
        try:
            Path(normalized).unlink(missing_ok=True)
        except Exception:
            pass

# ---------------------------------------------------------------------------
# Conversion dossier
# ---------------------------------------------------------------------------
