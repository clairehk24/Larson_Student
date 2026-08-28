"""Apply repeatable accessibility remediations to Larson student assets.

DOCX parts are edited as raw XML bytes. This deliberately preserves every
namespace declaration and compatibility attribute produced by Microsoft Word.
Use ``--base-ref`` only to rebuild packages from a known Git revision.
"""

from __future__ import annotations

import argparse
import html
import os
import re
import subprocess
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
W = f"{{{W_NS}}}"
WP = f"{{{WP_NS}}}"
DC = f"{{{DC_NS}}}"


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


def xml_attribute(value: str) -> bytes:
    return html.escape(value, quote=True).encode("utf-8")


def text_of_paragraph(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(f".//{W}t")).strip()


def document_title(document_xml: bytes, fallback: str) -> str:
    document = ET.fromstring(document_xml)
    body = document.find(f"{W}body")
    if body is not None:
        for paragraph in body.findall(f"{W}p"):
            value = text_of_paragraph(paragraph)
            if value:
                return value
    return fallback


def set_core_title(core_xml: bytes, title: str) -> tuple[bytes, int]:
    replacement = b"<dc:title>" + html.escape(title).encode("utf-8") + b"</dc:title>"
    pattern = rb"<dc:title(?:\s[^>]*)?>.*?</dc:title>|<dc:title\s*/>"
    updated, count = re.subn(pattern, replacement, core_xml, count=1, flags=re.DOTALL)
    if count == 0:
        closing = b"</cp:coreProperties>"
        if closing not in core_xml:
            raise ValueError("DOCX core properties have no insertion point for dc:title")
        updated = core_xml.replace(closing, replacement + closing, 1)
    elif count != 1:
        raise ValueError("DOCX core properties contain multiple dc:title elements")
    return updated, int(updated != core_xml)


def ensure_heading_one(document_xml: bytes) -> tuple[bytes, int]:
    paragraph = re.search(rb"<w:p(?:\s[^>]*)?>.*?</w:p>", document_xml, re.DOTALL)
    if paragraph is None:
        return document_xml, 0
    first = paragraph.group(0)
    updated_first, count = re.subn(
        rb'(<w:pStyle\b[^>]*\bw:val=")Heading2(")',
        rb"\1Heading1\2",
        first,
        count=1,
    )
    if not count:
        return document_xml, 0
    return document_xml[: paragraph.start()] + updated_first + document_xml[paragraph.end() :], 1


def ensure_image_alternatives(
    document_xml: bytes, relative_path: str
) -> tuple[bytes, int]:
    alternatives = IMAGE_ALT_TEXT.get(relative_path, [])
    matches = list(re.finditer(rb"<wp:docPr\b[^>]*/>", document_xml))
    if matches and len(matches) != len(alternatives):
        raise ValueError(
            f"Expected {len(alternatives)} image descriptions for {relative_path}, "
            f"found {len(matches)} drawings"
        )
    if not matches:
        return document_xml, 0
    pieces: list[bytes] = []
    position = 0
    changes = 0
    for match, alternative in zip(matches, alternatives):
        tag = match.group(0)
        description = b'descr="' + xml_attribute(alternative) + b'"'
        if re.search(rb"\bdescr=", tag):
            updated = re.sub(rb'descr="[^"]*"', description, tag, count=1)
        else:
            updated = tag[:-2].rstrip() + b" " + description + b"/>"
        pieces.extend((document_xml[position : match.start()], updated))
        position = match.end()
        changes += int(updated != tag)
    pieces.append(document_xml[position:])
    return b"".join(pieces), changes


def remediate_table(table: bytes, title: str, number: int) -> tuple[bytes, int]:
    changes = 0
    caption = (
        b'<w:tblCaption w:val="'
        + xml_attribute(f"{title} — table {number}")
        + b'"/><w:tblDescription w:val="Worksheet table. Read the cells from left to right across each row."/>'
    )
    properties = re.search(rb"<w:tblPr(?:\s[^>]*)?>", table)
    if properties is None:
        opening = re.search(rb"<w:tbl(?:\s[^>]*)?>", table)
        if opening is None:
            raise ValueError("Malformed Word table")
        insertion = b"<w:tblPr>" + caption + b"</w:tblPr>"
        table = table[: opening.end()] + insertion + table[opening.end() :]
        changes += 1
    elif b"<w:tblCaption" not in table:
        properties_end = table.find(b"</w:tblPr>", properties.end())
        if properties_end < 0:
            raise ValueError("Malformed Word table properties")
        change_marker = table.find(b"<w:tblPrChange", properties.end(), properties_end)
        insertion_at = change_marker if change_marker >= 0 else properties_end
        table = table[:insertion_at] + caption + table[insertion_at:]
        changes += 1

    first_row = re.search(rb"<w:tr(?:\s[^>]*)?>", table)
    if first_row is not None:
        row_end = table.find(b"</w:tr>", first_row.end())
        row = table[first_row.start() : row_end if row_end >= 0 else len(table)]
        if b"<w:tblHeader" not in row:
            row_properties = re.search(rb"<w:trPr(?:\s[^>]*)?>", row)
            if row_properties is None:
                insertion = b"<w:trPr><w:tblHeader/></w:trPr>"
                table = table[: first_row.end()] + insertion + table[first_row.end() :]
            else:
                if row_properties.group(0).rstrip().endswith(b"/>"):
                    replacement = row_properties.group(0).rstrip()[:-2] + b"><w:tblHeader/></w:trPr>"
                    insertion_at = first_row.start() + row_properties.start()
                    absolute_end = first_row.start() + row_properties.end()
                    table = table[:insertion_at] + replacement + table[absolute_end:]
                    return table, changes + 1
                properties_end = row.find(b"</w:trPr>", row_properties.end())
                if properties_end < 0:
                    raise ValueError("Malformed Word table-row properties")
                later_property = re.search(
                    rb"<w:(?:tblCellSpacing|jc|hidden|ins|del|trPrChange)\b",
                    row[row_properties.end() : properties_end],
                )
                relative_at = (
                    row_properties.end() + later_property.start()
                    if later_property is not None
                    else properties_end
                )
                insertion_at = first_row.start() + relative_at
                table = table[:insertion_at] + b"<w:tblHeader/>" + table[insertion_at:]
            changes += 1
    return table, changes


def ensure_table_accessibility(document_xml: bytes, title: str) -> tuple[bytes, int]:
    tables = list(re.finditer(rb"<w:tbl(?:\s[^>]*)?>.*?</w:tbl>", document_xml, re.DOTALL))
    if not tables:
        return document_xml, 0
    pieces: list[bytes] = []
    position = 0
    changes = 0
    for number, match in enumerate(tables, start=1):
        updated, table_changes = remediate_table(match.group(0), title, number)
        pieces.extend((document_xml[position : match.start()], updated))
        position = match.end()
        changes += table_changes
    pieces.append(document_xml[position:])
    return b"".join(pieces), changes


def validate_compatibility_namespaces(document_xml: bytes) -> None:
    root_start = document_xml.find(b"<w:document")
    root_end = document_xml.find(b">", root_start)
    if root_start < 0 or root_end < 0:
        raise ValueError("Word document root element is missing")
    root = document_xml[root_start : root_end + 1]
    match = re.search(rb'mc:Ignorable="([^"]+)"', root)
    if match:
        for prefix in match.group(1).split():
            if b"xmlns:" + prefix + b"=" not in root:
                raise ValueError(f"mc:Ignorable refers to undeclared namespace {prefix!r}")
    ET.fromstring(document_xml)


def git_file(revision: str, path: Path) -> bytes:
    relative = path.relative_to(ROOT).as_posix()
    return subprocess.check_output(["git", "show", f"{revision}:{relative}"], cwd=ROOT)


def rewrite_docx(path: Path, base_ref: str | None = None) -> tuple[int, str]:
    relative_path = path.relative_to(DOWNLOADS).as_posix()
    package = git_file(base_ref, path) if base_ref else path.read_bytes()
    with ZipFile(BytesIO(package), "r") as source:
        document_xml = source.read("word/document.xml")
        core_xml = source.read("docProps/core.xml")
        title = document_title(document_xml, path.stem.replace("-", " ").title())

        core_xml, changes = set_core_title(core_xml, title)
        document_xml, count = ensure_heading_one(document_xml)
        changes += count
        document_xml, count = ensure_table_accessibility(document_xml, title)
        changes += count
        document_xml, count = ensure_image_alternatives(document_xml, relative_path)
        changes += count
        validate_compatibility_namespaces(document_xml)

        entries = []
        for item in source.infolist():
            if item.filename == "word/document.xml":
                entries.append((item, document_xml))
            elif item.filename == "docProps/core.xml":
                entries.append((item, core_xml))
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
        with ZipFile(temporary_path) as check:
            if check.testzip() is not None:
                raise ValueError(f"CRC failure while rebuilding {path}")
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
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--base-ref",
        help="Rebuild DOCX packages from this Git revision before applying fixes",
    )
    args = parser.parse_args()
    html_count = remediate_html()
    docx_count = 0
    change_count = 0
    for path in sorted(DOWNLOADS.rglob("*.docx")):
        changes, title = rewrite_docx(path, args.base_ref)
        docx_count += 1
        change_count += changes
        print(f"DOCX: {path.relative_to(ROOT)} — {title}")
    print(
        f"Remediated {html_count} HTML files and {docx_count} DOCX files "
        f"({change_count} DOCX accessibility updates)."
    )


if __name__ == "__main__":
    main()
