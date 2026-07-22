"""Turn markdown into OneNote page XML.

The conversion is deliberately narrow: headings, paragraphs, bullet and
numbered lists, and inline bold/italic/code/links. Anything else is carried
through as plain text. It is the inverse of the subset that xml_parser
produces, and it is lossy in both directions — see apply_markdown for how that
is contained.
"""

import html
import re
import xml.etree.ElementTree as ET

NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
NSMAP = {"one": NS}

ET.register_namespace("one", NS)

_HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")
_BULLET_RE = re.compile(r"^(\s*)[-*+]\s+(.*)$")
_NUMBER_RE = re.compile(r"^(\s*)\d+[.)]\s+(.*)$")

_LINK_RE = re.compile(r"\[([^\]]*)\]\(([^)]+)\)")
_BOLD_RE = re.compile(r"\*\*(.+?)\*\*", re.S)
_ITALIC_RE = re.compile(r"(?<!\*)\*([^*]+?)\*(?!\*)")
_CODE_RE = re.compile(r"`([^`]+)`")


def _q(tag: str) -> str:
    return f"{{{NS}}}{tag}"


def inline_to_onenote(text: str) -> str:
    """Convert inline markdown to the HTML subset OneNote stores inside one:T."""
    out = html.escape(text, quote=False)
    # Links first: their label may itself contain emphasis.
    out = _LINK_RE.sub(lambda m: f'<a href="{html.escape(m.group(2), quote=True)}">{m.group(1)}</a>', out)
    out = _BOLD_RE.sub(r'<span style="font-weight:bold">\1</span>', out)
    out = _ITALIC_RE.sub(r'<span style="font-style:italic">\1</span>', out)
    out = _CODE_RE.sub(r'<span style="font-family:Consolas">\1</span>', out)
    return out


# OneNote rejects a QuickStyleDef that is missing any of these, with
# "Required attribute 'font' is missing." and friends.
_QUICK_STYLE_DEFAULTS = {
    "fontColor": "automatic",
    "highlightColor": "automatic",
    "font": "Calibri",
    "fontSize": "11.0",
    "spaceBefore": "0.0",
    "spaceAfter": "0.0",
}

# Heading sizes roughly matching OneNote's built-in h1..h6.
_HEADING_SIZES = {"h1": "16.0", "h2": "14.0", "h3": "12.0", "h4": "11.0", "h5": "11.0", "h6": "11.0"}


class _QuickStyles:
    """Look up (or create) the QuickStyleDef index for a style name."""

    def __init__(self, page: ET.Element):
        self.page = page
        self.defs = page.findall("one:QuickStyleDef", NSMAP)
        self.by_name = {qs.get("name", ""): qs.get("index", "") for qs in self.defs}
        indices = [int(i) for i in self.by_name.values() if i.isdigit()]
        self._next = max(indices, default=-1) + 1

    def _base_attrs(self) -> dict[str, str]:
        """Start from the page's body style so new styles look native."""
        attrs = dict(_QUICK_STYLE_DEFAULTS)
        for qs in self.defs:
            if qs.get("name") == "p":
                for key in _QUICK_STYLE_DEFAULTS:
                    if qs.get(key):
                        attrs[key] = qs.get(key, "")
                break
        return attrs

    def index_for(self, name: str) -> str:
        if name in self.by_name:
            return self.by_name[name]

        index = str(self._next)
        self._next += 1

        attrs = self._base_attrs()
        attrs["index"] = index
        attrs["name"] = name
        if name in _HEADING_SIZES:
            attrs["fontSize"] = _HEADING_SIZES[name]
            attrs["bold"] = "true"

        qs = ET.Element(_q("QuickStyleDef"), attrs)
        # QuickStyleDef elements must come before the page body.
        self.page.insert(len(self.by_name), qs)
        self.defs.append(qs)
        self.by_name[name] = index
        return index


def _make_oe(text: str, styles: _QuickStyles, style: str | None, list_kind: str | None) -> ET.Element:
    oe = ET.Element(_q("OE"))
    if style:
        oe.set("quickStyleIndex", styles.index_for(style))
    if list_kind == "bullet":
        lst = ET.SubElement(oe, _q("List"))
        ET.SubElement(lst, _q("Bullet"), {"bullet": "2"})
    elif list_kind == "number":
        lst = ET.SubElement(oe, _q("List"))
        ET.SubElement(lst, _q("Number"), {"numberSequence": "0", "numberFormat": "##"})
    t = ET.SubElement(oe, _q("T"))
    t.text = inline_to_onenote(text)
    return oe


def markdown_to_oes(markdown: str, page: ET.Element) -> list[ET.Element]:
    """Build the one:OE elements for a markdown body.

    Indented list items become nested one:OEChildren, matching how xml_parser
    renders nesting back out as indentation.
    """
    styles = _QuickStyles(page)
    roots: list[ET.Element] = []
    # (indent_width, element_to_append_children_to)
    stack: list[tuple[int, ET.Element]] = []

    for raw in markdown.splitlines():
        if not raw.strip():
            continue

        style: str | None = None
        list_kind: str | None = None
        indent = 0
        text = raw.strip()

        heading = _HEADING_RE.match(raw.strip())
        bullet = _BULLET_RE.match(raw)
        number = _NUMBER_RE.match(raw)

        if heading:
            style = f"h{len(heading.group(1))}"
            text = heading.group(2)
        elif bullet:
            indent = len(bullet.group(1))
            list_kind = "bullet"
            text = bullet.group(2)
        elif number:
            indent = len(number.group(1))
            list_kind = "number"
            text = number.group(2)
        else:
            indent = len(raw) - len(raw.lstrip())

        oe = _make_oe(text, styles, style, list_kind)

        while stack and stack[-1][0] >= indent:
            stack.pop()
        if stack:
            parent_oe = stack[-1][1]
            children = parent_oe.find("one:OEChildren", NSMAP)
            if children is None:
                children = ET.SubElement(parent_oe, _q("OEChildren"))
            children.append(oe)
        else:
            roots.append(oe)
        stack.append((indent, oe))

    return roots


def apply_markdown(page_xml: str, markdown: str, mode: str = "append") -> str:
    """Return updated page XML with the markdown applied.

    The whole page is round-tripped rather than sent as a fragment: every
    existing element keeps its objectID, so OneNote matches and leaves it
    alone, and only the elements added here are new. That is what makes
    "append" genuinely non-destructive — a fragment containing just the new
    content would replace everything under the outline.

    Args:
        mode: "append" adds to the end of the first outline; "replace" clears
            that outline's content first, discarding anything it held —
            including images and ink, which this converter cannot rebuild.
    """
    if mode not in ("append", "replace"):
        raise ValueError(f"mode must be 'append' or 'replace', got {mode!r}")

    root = ET.fromstring(page_xml)
    outline = root.find("one:Outline", NSMAP)
    if outline is None:
        outline = ET.SubElement(root, _q("Outline"))

    children = outline.find("one:OEChildren", NSMAP)
    if children is None:
        children = ET.SubElement(outline, _q("OEChildren"))

    if mode == "replace":
        for child in list(children):
            children.remove(child)

    for oe in markdown_to_oes(markdown, root):
        children.append(oe)

    return ET.tostring(root, encoding="unicode")


def set_page_title(page_xml: str, title: str) -> str:
    """Return page XML with its title replaced."""
    root = ET.fromstring(page_xml)
    t = root.find("one:Title//one:T", NSMAP)
    if t is None:
        title_el = ET.SubElement(root, _q("Title"))
        oe = ET.SubElement(title_el, _q("OE"))
        t = ET.SubElement(oe, _q("T"))
    t.text = inline_to_onenote(title)
    return ET.tostring(root, encoding="unicode")
