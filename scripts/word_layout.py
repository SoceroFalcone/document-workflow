"""Apply or check the document-workflow preset on a DOCX copy."""

import argparse
from collections import Counter
from io import BytesIO
import json
from pathlib import Path
import re
import sys
from zipfile import BadZipFile

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.style import WD_STYLE_TYPE
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.opc.exceptions import PackageNotFoundError
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Cm, Twips
from lxml import etree

from word_crossrefs import ParagraphText, field_parts, managed_number_field


DEFAULTS = Path(__file__).resolve().parents[1] / "references" / "default-word.json"
PROFILES = {
    "academic-paper": DEFAULTS.with_name("academic-paper-word.json"),
    "thesis": DEFAULTS.with_name("thesis-word.json"),
}
RULE_PRIORITY = ("skill default", "academic-paper profile", "thesis profile", "project", "template", "task")
PARAGRAPH_ORDER = (
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
    "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
    "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
    "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing",
    "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
)
FONT_ORDER = (
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
    "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid",
    "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs",
    "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs",
    "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange",
)
PARAGRAPH_SWITCHES = {
    "widowControl": "widow_control",
    "keepNext": "keep_with_next",
    "keepLines": "keep_lines_together",
    "pageBreakBefore": "page_break_before",
    "suppressLineNumbers": "suppress_line_numbers",
    "suppressAutoHyphens": "suppress_hyphenation",
}
ALIGNMENTS = {"left": "left", "center": "center", "right": "right", "justify": "both"}
TIGHT_WRAPS = {"none", "allLines", "firstAndLastLine", "firstLineOnly", "lastLineOnly"}
COMPLEX_CONTENT = {
    qn(tag) for tag in (
        "w:drawing", "w:pict", "w:object", "w:sdt", "w:ins", "w:del", "w:moveFrom",
        "w:moveTo", "w:pPrChange", "w:rPrChange", "w:fldChar", "w:fldSimple",
        "w:instrText", "m:oMath", "m:oMathPara",
    )
}


def read_config(path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def overlay(target, values, source, origins, prefix=""):
    if not isinstance(values, dict):
        raise ValueError(f"{prefix or source} must be a JSON object")
    for name, value in values.items():
        field = f"{prefix}.{name}" if prefix else name
        if name not in target:
            raise ValueError(f"Unknown option: {field}")
        if isinstance(target[name], dict):
            overlay(target[name], value, source, origins, field)
        else:
            target[name] = value
            origins[field] = source


def number(value, field, minimum, maximum, units):
    if type(value) not in (int, float) or not minimum <= value <= maximum:
        raise ValueError(f"Invalid number for {field}: {value!r}")
    return round(value * units)


def style_roles(config):
    names = {}
    for role, rules in config["roles"].items():
        styles = rules["styles"]
        if not isinstance(styles, list) or any(
            not isinstance(name, str) or not name.strip() for name in styles
        ):
            raise ValueError(f"roles.{role}.styles must be an array of style names")
        for name in styles:
            normalized = name.strip().casefold()
            if normalized in names and names[normalized] != role:
                raise ValueError(f"Style name is assigned to two roles: {name}")
            names[normalized] = role
    return names


def resolve_role_fonts(config, sources):
    for name in ("east_asia", "latin"):
        global_source = sources[f"font.{name}"]
        option = f"{name}_font"
        for role, rules in config["roles"].items():
            field = f"roles.{role}.{option}"
            if rules[option] is None or RULE_PRIORITY.index(sources[field]) < RULE_PRIORITY.index(global_source):
                rules[option] = config["font"][name]
                sources[field] = global_source


def expected_properties(config):
    font = config["font"]
    for name in ("east_asia", "latin"):
        if not isinstance(font[name], str) or not font[name].strip():
            raise ValueError(f"font.{name} must be a font name")
    color = font["color"]
    if not isinstance(color, str) or not re.fullmatch(r"[a-fA-F0-9]{6}", color):
        raise ValueError("font.color must be a six-digit RGB hex value")

    settings = config["paragraph"]
    spacing = {
        "line": number(settings["line_spacing_multiple"], "line_spacing_multiple", 0.1, 10, 240),
        "lineRule": "auto",
    }
    for key in ("before", "after"):
        amount = number(settings[f"space_{key}_pt"], f"space_{key}_pt", 0, 1000, 20)
        spacing[key] = amount
        spacing[f"{key}Autospacing"] = 0
        if amount == 0:
            spacing[f"{key}Lines"] = 0
    common = {"spacing": spacing, "snapToGrid": {"val": 0}, "contextualSpacing": {"val": 0}}
    for tag, key in PARAGRAPH_SWITCHES.items():
        if type(settings[key]) is not bool:
            raise ValueError(f"paragraph.{key} must be true or false")
        common[tag] = {"val": int(settings[key])}
    wrap = settings["text_box_tight_wrap"]
    if not isinstance(wrap, str) or wrap not in TIGHT_WRAPS:
        raise ValueError("Invalid paragraph.text_box_tight_wrap")
    common["textboxTightWrap"] = {"val": wrap}

    by_role = {}
    for role, rules in config["roles"].items():
        alignment = rules["alignment"]
        if not isinstance(alignment, str) or alignment not in ALIGNMENTS:
            raise ValueError(f"Invalid roles.{role}.alignment")
        indent = number(rules["first_line_chars"], f"roles.{role}.first_line_chars", 0, 20, 100)
        ind = {"firstLineChars": indent}
        if indent == 0:
            ind["firstLine"] = 0
        size = number(rules["font_size_pt"], f"roles.{role}.font_size_pt", 1, 1638, 2)
        bold = rules["bold"]
        if bold is not None and type(bold) is not bool:
            raise ValueError(f"roles.{role}.bold must be true, false or null")
        remove_bottom_border = rules.get("remove_bottom_border", False)
        if type(remove_bottom_border) is not bool:
            raise ValueError(f"roles.{role}.remove_bottom_border must be true or false")
        fonts = {}
        for name in ("east_asia", "latin"):
            option = f"{name}_font"
            value = rules.get(option)
            if value is None:
                value = font[name]
            if not isinstance(value, str) or not value.strip():
                raise ValueError(f"roles.{role}.{option} must be a font name or null")
            fonts[name] = value
        paragraph = {**common, "jc": {"val": ALIGNMENTS[alignment]}, "ind": ind}
        run = {
            "rFonts": {"eastAsia": fonts["east_asia"], "ascii": fonts["latin"], "hAnsi": fonts["latin"]},
            "sz": {"val": size},
            "szCs": {"val": size},
            "color": {"val": color.upper()},
        }
        if bold is not None:
            run.update({"b": {"val": int(bold)}, "bCs": {"val": int(bold)}})
        by_role[role] = (paragraph, run, remove_bottom_border)
    return by_role


def word_tag(name):
    return qn(f"w:{name}")


def set_properties(element, order, properties, clear):
    changed = set()
    for tag, attributes in properties.items():
        child = element.find(word_tag(tag))
        if child is None:
            child = OxmlElement(f"w:{tag}")
            successors = [word_tag(name) for name in order[order.index(tag) + 1:]]
            following = next((node for node in element if node.tag in successors), None)
            if following is None:
                element.append(child)
            else:
                following.addprevious(child)
            changed.add(tag)
        for name in clear.get(tag, ()):
            if name not in attributes and child.attrib.pop(word_tag(name), None) is not None:
                changed.add(tag)
        for name, value in attributes.items():
            key = word_tag(name)
            if child.get(key) != str(value):
                child.set(key, str(value))
                changed.add(tag)
    return changed


def keyword_label_runs(paragraph):
    view = ParagraphText(paragraph)
    label = re.match(r"^[ \t\u3000]*关键词[ \t\u3000]*[：:]", view.text)
    if label is None or not view.replaceable(*label.span()):
        return set(), 0
    view, splits = view.split_text_boundaries(label.span())
    return {part.nodes[0] for part in view.cover(*label.span())}, splits


def apply_paragraph(paragraph, expected, keyword_label=None):
    paragraph_properties, run_properties, remove_bottom_border = expected
    ppr = paragraph.get_or_add_pPr()
    changes = set_properties(
        ppr, PARAGRAPH_ORDER, paragraph_properties,
        {"spacing": ("beforeLines", "afterLines"), "ind": ("firstLine", "hanging", "hangingChars")},
    )
    if remove_bottom_border:
        changes.update(set_properties(ppr, PARAGRAPH_ORDER, {"pBdr": {}}, {}))
        borders = ppr.find(word_tag("pBdr"))
        if set_properties(
            borders, ("top", "left", "bottom", "right", "between", "bar"),
            {"bottom": {"val": "nil"}}, {},
        ):
            changes.add("pBdr")
    marker = ppr.find(word_tag("rPr"))
    if marker is None:
        marker = OxmlElement("w:rPr")
        ppr.insert_element_before(marker, "w:sectPr", "w:pPrChange")
    label_runs, splits = keyword_label_runs(paragraph) if keyword_label is not None else (set(), 0)
    if splits:
        changes.add("keyword label run boundary")
    # Match the paragraph mark's font to the font used for the character-unit indent.
    clear = {"rFonts": ("asciiTheme", "hAnsiTheme", "eastAsiaTheme"),
             "color": ("themeColor", "themeTint", "themeShade")}
    changes.update(set_properties(marker, FONT_ORDER, run_properties, clear))
    for run in paragraph.iter(word_tag("r")):
        fmt = keyword_label if run in label_runs else run_properties
        changes.update(set_properties(run.get_or_add_rPr(), FONT_ORDER, fmt, clear))
    return sorted(changes)


def paragraph_style_chain(paragraph, styles_by_id, normal):
    ppr = paragraph.pPr
    style = styles_by_id.get(ppr.style) if ppr is not None and ppr.style else normal
    visited = set()
    while style is not None and style.style_id not in visited:
        visited.add(style.style_id)
        yield style
        style = style.base_style


def review_reason(paragraph, styles, spacing):
    special = {node for node in paragraph.iter() if node.tag in COMPLEX_CONTENT}
    if special:
        managed = {
            child
            for part in field_parts(list(paragraph)) if managed_number_field(part)
            for node in part.nodes
            for child in node.iter()
            if child.tag in {qn("w:fldChar"), qn("w:fldSimple"), qn("w:instrText")}
        }
        if special - managed:
            return "complex paragraph (field, object, equation, control, or revision)"
    style_properties = [style.element.pPr for style in styles]
    for ppr in [paragraph.pPr, *style_properties]:
        if ppr is None:
            continue
        if ppr.find(word_tag("framePr")) is not None:
            return "framed paragraph"
        if ppr.numPr is not None and ppr.numPr.numId is not None and ppr.numPr.numId.val:
            return "numbered or bulleted paragraph"
    for ppr in style_properties:
        if ppr is None or ppr.spacing is None:
            continue
        if any(spacing[key] and ppr.spacing.get(word_tag(f"{key}Lines"), "0") != "0"
               for key in ("before", "after")):
            return "inherited line-unit paragraph spacing conflicts with point override"
    return None


def page_properties(page):
    for key, value in page.items():
        minimum = 1 if key in {"width_cm", "height_cm"} else 0
        if type(value) not in (int, float) or not minimum <= value <= 55.88:
            raise ValueError(f"Invalid page.{key}: {value!r}")
    if (page["left_margin_cm"] + page["right_margin_cm"] >= page["width_cm"]
            or page["top_margin_cm"] + page["bottom_margin_cm"] >= page["height_cm"]):
        raise ValueError("Margins must leave room for text")
    return {
        "page_width": Cm(page["width_cm"]),
        "page_height": Cm(page["height_cm"]),
        **{f"{side}_margin": Cm(page[f"{side}_margin_cm"])
           for side in ("top", "bottom", "left", "right")},
    }


def review_thesis_page(page, location, review):
    # Compare in Word's saved units so an exact 25 mm setting survives rounding.
    dimensions = (page["page_width"], page["page_height"])
    a4 = [Cm(21).twips, Cm(29.7).twips]
    if None in dimensions or sorted(length.twips for length in dimensions) != a4:
        review.append({"location": location, "reason": "A4 is recommended for printed theses"})

    for side, minimum_mm in (("top", 25), ("left", 25), ("bottom", 20), ("right", 20)):
        margin = page[f"{side}_margin"]
        if margin is not None and margin.twips >= Cm(minimum_mm / 10).twips:
            continue
        actual = "not set" if margin is None else f"{Twips(margin.twips).mm:.2f} mm"
        review.append({
            "location": location,
            "reason": f"{side.capitalize()} margin is {actual}. Printed theses require at least {minimum_mm} mm.",
        })


def inspect_document(document, config, profile=None, check_source=False):
    expected = expected_properties(config)
    aliases = style_roles(config)
    page = page_properties(config["page"])
    changes, review = [], []
    if profile == "thesis" and not check_source:
        review_thesis_page(page, "selected page settings", review)
    counts = Counter()
    styles_by_id = {style.style_id: style for style in document.styles}
    normal = document.styles.default(WD_STYLE_TYPE.PARAGRAPH)

    def process(paragraph, location, role=None):
        styles = list(paragraph_style_chain(paragraph, styles_by_id, normal))
        reason = review_reason(paragraph, styles, expected["body"][0]["spacing"])
        if reason:
            review.append({"location": location, "reason": reason})
            return
        style_name = (styles[0].name or styles[0].style_id) if styles else "(unknown)"
        role = role or aliases.get(style_name.strip().casefold())
        if role is None:
            review.append({"location": location, "reason": f"unmapped style: {style_name}"})
            return
        outline = paragraph.pPr.find(word_tag("outlineLvl")) if paragraph.pPr is not None else None
        if role == "body" and outline is not None:
            level = int(outline.get(word_tag("val"), "9"))
            if level in range(3):
                role = f"heading_{level + 1}"
            elif level in range(3, 9):
                review.append({"location": location, "reason": f"unmapped outline level: {level + 1}"})
                return
        counts[role] += 1
        keyword_label = expected["keywords_cn_label"][1] if role == "keywords_cn" else None
        properties = apply_paragraph(paragraph, expected[role], keyword_label)
        if properties:
            changes.append({"location": location, "role": role, "properties": properties})

    body = document.element.body
    for index, paragraph in enumerate(body.findall(word_tag("p")), 1):
        process(paragraph, f"paragraph {index}")
    for tag in ("tbl", "sdt", "customXml", "altChunk", "ins", "del"):
        count = len(body.findall(word_tag(tag)))
        if count:
            review.append({"location": "main document", "reason": f"{count} {tag} block(s) skipped"})

    for rel in document.part.rels.values():
        if rel.reltype != RELATIONSHIP_TYPE.FOOTNOTES or rel.is_external:
            continue
        part = rel.target_part
        notes = parse_xml(part.blob)
        before = len(changes)
        for note in notes.findall(word_tag("footnote")):
            if note.get(word_tag("type"), "normal") != "normal":
                continue
            location = f"footnote {note.get(word_tag('id'))}"
            for index, paragraph in enumerate(note.findall(word_tag("p")), 1):
                process(paragraph, f"{location}, paragraph {index}", "footnote")
            if any(child.tag != word_tag("p") for child in note):
                review.append({"location": location, "reason": "non-paragraph block skipped"})
        if len(changes) > before:
            part._blob = etree.tostring(notes, encoding="UTF-8", xml_declaration=True, standalone=True)

    sections = document.sections
    dimensions = {(section.page_width, section.page_height) for section in sections}
    keep_page_settings = (
        len(dimensions) > 1
        or any(width and height and width > height for width, height in dimensions)
        or any(section.orientation == WD_ORIENT.LANDSCAPE for section in sections)
    )
    if keep_page_settings:
        review.append({"location": "page setup", "reason": "landscape or mixed page sizes kept"})
    for index, section in enumerate(sections, 1):
        location = f"section {index}"
        if profile == "thesis" and (check_source or keep_page_settings):
            actual_page = {name: getattr(section, name) for name in page}
            review_thesis_page(actual_page, location, review)
        if keep_page_settings:
            continue
        updated = []
        for name, expected_value in page.items():
            actual = getattr(section, name)
            if actual is None or actual.twips != expected_value.twips:
                setattr(section, name, expected_value)
                updated.append(name)
        if updated:
            changes.append({"location": location, "properties": updated})
    return {"checked_paragraphs": dict(counts), "differences": changes, "manual_review": review}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("operation", choices=("check", "apply"))
    parser.add_argument("document", type=Path)
    parser.add_argument("--out", type=Path, help="new DOCX path, required for apply")
    parser.add_argument("--profile", choices=tuple(PROFILES), help="paper layout profile")
    parser.add_argument("--project", type=Path, help="project JSON overrides")
    parser.add_argument("--template", type=Path, help="JSON rules extracted from the selected template")
    parser.add_argument("--task", type=Path, help="overrides for this task")
    args = parser.parse_args()

    try:
        if args.document.suffix.lower() != ".docx":
            raise ValueError("Input must be a DOCX")
        if args.operation == "apply":
            if args.out is None or args.out.suffix.lower() != ".docx":
                raise ValueError("apply requires a new .docx path via --out")
            if args.out.exists() or args.out.resolve() == args.document.resolve():
                raise ValueError("Choose an output path that does not already exist")
        elif args.out is not None:
            raise ValueError("--out is for apply")

        config = read_config(DEFAULTS)
        profile = read_config(PROFILES[args.profile]) if args.profile else {}
        for role, rules in profile.get("roles", {}).items():
            config["roles"].setdefault(role, rules.copy())
        for rules in config["roles"].values():
            rules.setdefault("latin_font", None)
            rules.setdefault("east_asia_font", None)
        sources = {}
        overlay(config, config, "skill default", sources)
        if profile:
            overlay(config, profile, f"{args.profile} profile", sources)
        for source, path in (("project", args.project), ("template", args.template), ("task", args.task)):
            if path is not None:
                overlay(config, read_config(path), source, sources)
        resolve_role_fonts(config, sources)

        document = Document(args.document)
        report = inspect_document(document, config, args.profile, check_source=args.operation == "check")
        report["input"] = str(args.document)
        report["rule_sources"] = sources
        if args.profile is not None:
            report["profile"] = args.profile
        if args.operation == "check":
            print(json.dumps(report, ensure_ascii=False, indent=2))
            return int(bool(report["differences"]))

        output = BytesIO()
        document.save(output)
        output.seek(0)
        verified = inspect_document(Document(output), config, args.profile)
        if verified["differences"]:
            raise ValueError("The generated DOCX still has differences in checked properties")
        with args.out.open("xb") as destination:
            try:
                destination.write(output.getbuffer())
                destination.flush()
            except OSError:
                destination.close()
                args.out.unlink()
                raise
        report["changes"] = report.pop("differences")
        report["output"] = str(args.out)
        report["verified"] = True
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 0
    except (OSError, ValueError, BadZipFile, PackageNotFoundError, etree.XMLSyntaxError) as error:
        print(f"word_layout: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
