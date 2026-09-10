"""Report inline formatting used in reader-facing manuscript content.

This complements validate_manuscript_text.py, whose text-only comparison cannot
detect emphasis lost while the Word manuscripts are converted to HTML pages.
"""

from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET
from html.parser import HTMLParser

from generate_simulation_downloads import SIMULATIONS, is_begin_marker
from validate_manuscript_text import (
    END_DOWNLOAD,
    INTRO_SOURCE,
    PRODUCTION_TAG,
    docx_blocks,
    strip_production_tag,
)


ROOT = Path(__file__).resolve().parents[1]
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def property_enabled(properties, name):
    element = properties.find(W + name) if properties is not None else None
    if element is None:
        return False
    value = element.get(W + "val")
    return value not in {"0", "false", "off", "none"}


def paragraph_runs(path):
    with ZipFile(path) as archive:
        document = ET.fromstring(archive.read("word/document.xml"))
    paragraphs = []
    for paragraph in document.iter(W + "p"):
        runs = []
        for run in paragraph.iter(W + "r"):
            text = "".join(node.text or "" for node in run.iter(W + "t"))
            if not text:
                continue
            properties = run.find(W + "rPr")
            styles = set()
            if property_enabled(properties, "i") or property_enabled(properties, "iCs"):
                styles.add("italic")
            if property_enabled(properties, "b") or property_enabled(properties, "bCs"):
                styles.add("bold")
            underline = properties.find(W + "u") if properties is not None else None
            if underline is not None and underline.get(W + "val", "single") != "none":
                styles.add("underline")
            vertical = properties.find(W + "vertAlign") if properties is not None else None
            if vertical is not None and vertical.get(W + "val") in {"superscript", "subscript"}:
                styles.add(vertical.get(W + "val"))
            runs.append((text, frozenset(styles)))
        text = "".join(text for text, _styles in runs)
        if text:
            paragraphs.append((text, runs))
    return paragraphs


def trim_runs(runs, start, end):
    result = []
    position = 0
    for text, styles in runs:
        run_end = position + len(text)
        left = max(start, position)
        right = min(end, run_end)
        if left < right:
            result.append((text[left - position:right - position], styles))
        position = run_end
    return result


def visible_runs(text, runs):
    visible = strip_production_tag(text)
    if not visible:
        return []
    if text.startswith("PHOTO HERE") and "<a>" in text:
        start = text.index("<a>") + len("<a>")
    else:
        match = PRODUCTION_TAG.match(text)
        start = match.end() if match else 0
    return trim_runs(runs, start, start + len(visible))


def page_paragraphs(path, activities=(), resources=()):
    paragraphs = paragraph_runs(path)
    texts = [text for text, _runs in paragraphs]
    output = []
    in_download = False
    for index, (text, runs) in enumerate(paragraphs):
        if text.startswith("Navigation menu/button"):
            continue
        if text.startswith("\\qqBEGIN downloadable content"):
            entries = [(button, title) for button, title, _filename in activities]
            entries += [(resource, resource) for resource in resources]
            button_name = next(
                button
                for button, title in entries
                if is_begin_marker(
                    text,
                    button,
                    title,
                    texts[index + 1] if index + 1 < len(texts) else None,
                )
            )
            output.append((button_name, [(button_name, frozenset({"bold"}))]))
            in_download = True
            continue
        if text == END_DOWNLOAD:
            in_download = False
            continue
        if in_download and text.endswith(END_DOWNLOAD):
            in_download = False
            continue
        if in_download:
            continue
        visible = strip_production_tag(text)
        if visible.strip():
            visible_formatting = visible_runs(text, runs)
            production_tag = PRODUCTION_TAG.match(text)
            if (production_tag and production_tag.group() in {
                "<cn>",
                "<ct>",
                "<a>",
                "<title>",
                "<lh>",
            }) or visible.strip() == "Student Simulation Summary":
                visible_formatting = [
                    (run_text, styles | {"bold"})
                    for run_text, styles in visible_formatting
                ]
            output.append((visible, visible_formatting))
    return output


def styled_spans(runs):
    spans = []
    for text, styles in runs:
        if not styles:
            continue
        if spans and spans[-1][1] == styles:
            spans[-1] = (spans[-1][0] + text, styles)
        else:
            spans.append((text, styles))
    return spans


class FormattedHTMLParser(HTMLParser):
    block_tags = {"h1", "h2", "h3", "p", "li"}
    italic_tags = {"em", "i"}
    bold_tags = {"strong", "b"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.in_manuscript = False
        self.current_tag = None
        self.current_runs = []
        self.style_stack = []
        self.blocks = []

    def handle_starttag(self, tag, attrs):
        attributes = dict(attrs)
        if "data-manuscript" in attributes:
            self.in_manuscript = True
        if self.in_manuscript and self.current_tag is None and (
            tag in self.block_tags or "data-manuscript-block" in attributes
        ):
            self.current_tag = tag
            self.current_runs = []
            implicit = set()
            if tag in {"h1", "h2", "h3", "strong", "b"}:
                implicit.add("bold")
            if "kicker" in attributes.get("class", "").split():
                implicit.add("bold")
            self.style_stack.append((tag, implicit))
            return
        if self.current_tag is not None:
            styles = set()
            if tag in self.italic_tags:
                styles.add("italic")
            if tag in self.bold_tags:
                styles.add("bold")
            if tag == "u":
                styles.add("underline")
            if tag == "sup":
                styles.add("superscript")
            if tag == "sub":
                styles.add("subscript")
            self.style_stack.append((tag, styles))

    def handle_data(self, data):
        if self.current_tag is None:
            return
        styles = frozenset(
            style
            for _tag, active_styles in self.style_stack
            for style in active_styles
        )
        self.current_runs.append((data, styles))

    def handle_endtag(self, tag):
        if self.current_tag is None:
            return
        for index in range(len(self.style_stack) - 1, -1, -1):
            if self.style_stack[index][0] == tag:
                del self.style_stack[index]
                break
        if tag == self.current_tag:
            text = "".join(run_text for run_text, _styles in self.current_runs)
            self.blocks.append((text, self.current_runs))
            self.current_tag = None
            self.current_runs = []
            self.style_stack = []


def html_paragraphs(path):
    parser = FormattedHTMLParser()
    parser.feed(path.read_text(encoding="utf-8"))
    return parser.blocks


def formatting_map(runs):
    return tuple(
        style
        for text, styles in runs
        for style in [styles] * len(text)
    )


def describe_mismatch(text, expected, actual):
    descriptions = []
    start = None
    previous = None
    for index, (wanted, found) in enumerate(zip(expected, actual)):
        difference = (wanted, found)
        if difference != previous:
            if start is not None and previous[0] != previous[1]:
                descriptions.append((start, index, previous))
            start = index
            previous = difference
    if start is not None and previous[0] != previous[1]:
        descriptions.append((start, len(text), previous))
    return descriptions


def compare_page(label, source, page):
    html = html_paragraphs(page)
    if [text for text, _runs in source] != [text for text, _runs in html]:
        print(f"FAIL {label}: text blocks differ; run validate_manuscript_text.py")
        return False
    valid = True
    for block_index, ((text, source_runs), (_html_text, html_runs)) in enumerate(
        zip(source, html), start=1
    ):
        expected = formatting_map(source_runs)
        actual = formatting_map(html_runs)
        if expected != actual:
            valid = False
            print(f"FAIL {label} block {block_index}: {text!r}")
            for start, end, (wanted, found) in describe_mismatch(text, expected, actual):
                print(
                    f"  {text[start:end]!r}: expected {sorted(wanted)}, "
                    f"found {sorted(found)}"
                )
    if valid:
        print(f"PASS {label}")
    return valid


def report(label, paragraphs):
    found = False
    for block_index, (text, runs) in enumerate(paragraphs, start=1):
        spans = styled_spans(runs)
        if spans:
            if not found:
                print(label)
                found = True
            print(f"  block {block_index}: {text!r}")
            for span, styles in spans:
                print(f"    {','.join(sorted(styles))}: {span!r}")
    return found


def main():
    results = []
    introduction = page_paragraphs(INTRO_SOURCE)
    report("introduction source formatting", introduction)
    results.append(
        compare_page("pages/introduction.html formatting", introduction, ROOT / "pages" / "introduction.html")
    )
    for simulation in SIMULATIONS:
        source = page_paragraphs(
            simulation["source"],
            simulation["activities"],
            simulation.get("resources", ()),
        )
        report(f"{simulation['id']} source formatting", source)
        results.append(
            compare_page(
                f"pages/{simulation['id']}.html formatting",
                source,
                simulation["page"],
            )
        )
    return 0 if all(results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
