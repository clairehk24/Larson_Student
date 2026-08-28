"""Apply repeatable accessibility remediations to Larson student assets.

The script intentionally changes only known course HTML and DOCX packages. It is
idempotent so it can be rerun after the downloadable activities are regenerated.
"""

from __future__ import annotations

import os
import tempfile
import xml.etree.ElementTree as ET
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


ROOT = Path(__file__).resolve().parents[1]
PAGES = ROOT / "pages"
DOWNLOADS = ROOT / "assets" / "downloads"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
DC_NS = "http://purl.org/dc/elements/1.1/"
CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DCTERMS_NS = "http://purl.org/dc/terms/"
DCTYPES_NS = "http://purl.org/dc/dcmitype/"
XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"

W = f"{{{W_NS}}}"
WP = f"{{{WP_NS}}}"
DC = f"{{{DC_NS}}}"

for prefix, uri in (
    ("w", W_NS),
    ("wp", WP_NS),
    ("cp", CP_NS),
    ("dc", DC_NS),
    ("dcterms", DCTERMS_NS),
    ("dcmitype", DCTYPES_NS),
    ("xsi", XSI_NS),
):
    ET.register_namespace(prefix, uri)


IMAGE_ALT_TEXT = {
    "simulation-1/presimulation-activity-2-anatomy-labeling.docx": [
        "Unlabeled lateral ankle diagram with leader lines for identifying the fibula, fifth metatarsal, anterior talofibular ligament, calcaneofibular ligament, and posterior talofibular ligament."
    ],
    "simulation-7/presimulation-activity-4-anatomy.docx": [
        "Posterior view of the elbow and forearm with unlabeled leader lines for identifying the radial, median, and ulnar nerves.",
        "Side-by-side medial and lateral views of the elbow with unlabeled leader lines for identifying bony landmarks and major ligaments.",
    ],
    "simulation-8/presimulation-activity-3-anatomy-labeling.docx": [
        "Unlabeled anterior cutaway diagram of the abdominal cavity with leader lines for identifying the pharynx, esophagus, liver, stomach, pancreas, small intestine, large intestine, and related organs."
    ],
    "simulation-15/postsimulation-activity-1-pulse-and-heart-rate.docx": [
        "Front and back views of a person showing pulse assessment locations: temporal, facial, carotid, brachial, radial, popliteal, posterior tibial, and dorsalis pedis arteries."
    ],
    "simulation-16/presimulation-activity-4-power-wheel.docx": [
        "Power wheel with Power at the center. Each wedge moves from greater privilege near the center to marginalization at the outer edge: college or university, high school, elementary education; able-bodied, some disability, significant disability; heterosexual, gay men, lesbian, bisexual, pansexual, or asexual; cisgender man, cisgender woman, transgender, intersex, or nonbinary; neurotypical, some neurodivergence, significant neurodivergence; slim, average, large body size; owns property, sheltered or renting, homeless; rich, middle class, poor; English, learned English, non-English monolingual; citizen, documented, undocumented; white, different shades closer to light skin or white, brown or dark skin."
    ],
}


def text_of_paragraph(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(f".//{W}t")).strip()


def parse_xml(data: bytes) -> ET.Element:
    # Preserve Word's original prefixes because mc:Ignorable names them by prefix.
    for _, (prefix, uri) in ET.iterparse(BytesIO(data), events=("start-ns",)):
        try:
            ET.register_namespace(prefix or "", uri)
        except ValueError:
            pass
    return ET.fromstring(data)


def document_title(document: ET.Element, fallback: str) -> str:
    body = document.find(f"{W}body")
    if body is not None:
        for paragraph in body.findall(f"{W}p"):
            value = text_of_paragraph(paragraph)
            if value:
                return value
    return fallback


def ensure_heading_one(document: ET.Element) -> bool:
    body = document.find(f"{W}body")
    if body is None:
        return False
    for paragraph in body.findall(f"{W}p"):
        if not text_of_paragraph(paragraph):
            continue
        properties = paragraph.find(f"{W}pPr")
        if properties is None:
            properties = ET.Element(f"{W}pPr")
            paragraph.insert(0, properties)
        style = properties.find(f"{W}pStyle")
        if style is None:
            style = ET.Element(f"{W}pStyle")
            properties.insert(0, style)
        changed = style.get(f"{W}val") != "Heading1"
        style.set(f"{W}val", "Heading1")
        return changed
    return False


def ensure_table_accessibility(document: ET.Element, title: str) -> int:
    changed = 0
    for number, table in enumerate(document.findall(f".//{W}tbl"), start=1):
        properties = table.find(f"{W}tblPr")
        if properties is None:
            properties = ET.Element(f"{W}tblPr")
            table.insert(0, properties)

        caption = properties.find(f"{W}tblCaption")
        if caption is None:
            caption = ET.SubElement(properties, f"{W}tblCaption")
        caption_value = f"{title} — table {number}"
        if caption.get(f"{W}val") != caption_value:
            caption.set(f"{W}val", caption_value)
            changed += 1

        description = properties.find(f"{W}tblDescription")
        if description is None:
            description = ET.SubElement(properties, f"{W}tblDescription")
        description_value = "Worksheet table. Read the cells from left to right across each row."
        if description.get(f"{W}val") != description_value:
            description.set(f"{W}val", description_value)
            changed += 1

        rows = table.findall(f"{W}tr")
        if rows:
            row_properties = rows[0].find(f"{W}trPr")
            if row_properties is None:
                row_properties = ET.Element(f"{W}trPr")
                rows[0].insert(0, row_properties)
            if row_properties.find(f"{W}tblHeader") is None:
                ET.SubElement(row_properties, f"{W}tblHeader")
                changed += 1
    return changed


def ensure_image_alternatives(document: ET.Element, relative_path: str) -> int:
    drawings = document.findall(f".//{WP}docPr")
    alternatives = IMAGE_ALT_TEXT.get(relative_path, [])
    if drawings and len(drawings) != len(alternatives):
        raise ValueError(
            f"Expected {len(alternatives)} image descriptions for {relative_path}, "
            f"found {len(drawings)} drawings"
        )
    changed = 0
    for drawing, alternative in zip(drawings, alternatives):
        if drawing.get("descr") != alternative:
            drawing.set("descr", alternative)
            changed += 1
    return changed


def set_core_title(core: ET.Element, title: str) -> bool:
    node = core.find(f"{DC}title")
    if node is None:
        node = ET.SubElement(core, f"{DC}title")
    changed = (node.text or "").strip() != title
    node.text = title
    return changed


def rewrite_docx(path: Path) -> tuple[int, str]:
    relative_path = path.relative_to(DOWNLOADS).as_posix()
    with ZipFile(path, "r") as source:
        document = parse_xml(source.read("word/document.xml"))
        core = parse_xml(source.read("docProps/core.xml"))
        title = document_title(document, path.stem.replace("-", " ").title())

        changes = int(set_core_title(core, title))
        changes += int(ensure_heading_one(document))
        changes += ensure_table_accessibility(document, title)
        changes += ensure_image_alternatives(document, relative_path)

        document_bytes = ET.tostring(document, encoding="utf-8", xml_declaration=True)
        core_bytes = ET.tostring(core, encoding="utf-8", xml_declaration=True)

        entries = []
        for item in source.infolist():
            if item.filename == "word/document.xml":
                entries.append((item, document_bytes))
            elif item.filename == "docProps/core.xml":
                entries.append((item, core_bytes))
            else:
                entries.append((item, source.read(item.filename)))

    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{path.stem}-", suffix=".docx", dir=path.parent
    )
    os.close(descriptor)
    temporary_path = Path(temporary_name)
    try:
        with ZipFile(temporary_path, "w", compression=ZIP_DEFLATED) as target:
            for item, content in entries:
                target.writestr(item, content)
        os.replace(temporary_path, path)
    finally:
        temporary_path.unlink(missing_ok=True)
    return changes, title


def remediate_html() -> int:
    changed_files = 0
    replacements = {
        '<span class="file-icon">': '<span class="file-icon" aria-hidden="true">',
        '<span class="download-arrow">': '<span class="download-arrow" aria-hidden="true">',
    }
    for path in sorted(PAGES.glob("*.html")):
        original = path.read_text(encoding="utf-8")
        updated = original
        for old, new in replacements.items():
            updated = updated.replace(old, new)
        if updated != original:
            path.write_text(updated, encoding="utf-8", newline="")
            changed_files += 1
    return changed_files


def main() -> None:
    html_count = remediate_html()
    docx_count = 0
    change_count = 0
    for path in sorted(DOWNLOADS.rglob("*.docx")):
        changes, title = rewrite_docx(path)
        docx_count += 1
        change_count += changes
        print(f"DOCX: {path.relative_to(ROOT)} — {title}")
    print(
        f"Remediated {html_count} HTML files and {docx_count} DOCX files "
        f"({change_count} DOCX accessibility updates)."
    )


if __name__ == "__main__":
    main()
