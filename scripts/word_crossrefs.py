"""Link figure, table, and bibliography numbers within a DOCX."""

import argparse
from collections import Counter, defaultdict
from copy import deepcopy
from dataclasses import dataclass
from io import BytesIO
import json
from pathlib import Path
import re
import sys
from uuid import uuid4
from zipfile import BadZipFile

from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml.etree import XMLSyntaxError


SEQUENCES = {"图": "DWFigure", "表": "DWTable", "文献": "DWBibliography"}
BY_SEQUENCE = {name.casefold(): kind for kind, name in SEQUENCES.items()}
BOOKMARK = re.compile(r"DWRef_[0-9a-f]{32}", re.I)
NUMBER = r"[1-9][0-9]{0,8}"
GAP = r"[ \t\u00a0\u3000]"
CAPTION = re.compile(rf"^\s*(?P<label>[图表]{GAP}*)(?P<number>{NUMBER})(?![0-9]|[.．－–—-]{GAP}*[0-9]){GAP}*\S")
ENTRY = re.compile(rf"^{GAP}*\[(?P<number>{NUMBER})\](?={GAP}*\S)", re.M)
FIGURE_MENTION = re.compile(
    rf"(?<![附续量视地插列挂])(?P<label>[图表]{GAP}*)(?P<number>{NUMBER})(?![0-9]|[.\uFF0E－–—~～-]{GAP}*[0-9])"
)
SEPARATORS = r"[,，、;；\-\u2013\u2014～~]"
CITATION = re.compile(rf"\[{GAP}*(?P<items>{NUMBER}(?:{GAP}*{SEPARATORS}{GAP}*{NUMBER})*){GAP}*\]")
RANGE_MARK = re.compile(r"[-–—～~]")
REFERENCE_TITLE = re.compile(r"(?:(?:[一二三四五六七八九十]+、|\d+[、.])\s*)?(?:参考文献|references|bibliography)[:：]?", re.I)

T = qn("w:t")
RUN = qn("w:r")
RUN_PROPS = qn("w:rPr")
START = qn("w:bookmarkStart")
END = qn("w:bookmarkEnd")
ID = qn("w:id")
NAME = qn("w:name")
FIELD_CHAR = qn("w:fldChar")
FIELD_TYPE = qn("w:fldCharType")
FIELD_INSTRUCTION = qn("w:instrText")
FIELD_LOCK = qn("w:fldLock")
XML_SPACE = qn("xml:space")
INLINE_TEXT = {qn("w:br"): "\n", qn("w:cr"): "\n", qn("w:tab"): "\t"}
ZERO_WIDTH = {START, END, qn("w:pPr"), qn("w:proofErr"), qn("w:lastRenderedPageBreak")}
STORY_ANCESTORS = {qn(f"w:{tag}") for tag in ("body", "tbl", "tr", "tc", "document")}
BREAK = "\uFFFC"


def node_text(node):
    if node.tag == T:
        return node.text or ""
    return INLINE_TEXT.get(node.tag, "" if node.tag in ZERO_WIDTH else BREAK)


def set_text(node, value):
    node.text = value
    if value[:1].isspace() or value[-1:].isspace():
        node.set(XML_SPACE, "preserve")


@dataclass
class Part:
    """One direct run, inline object, or complete Word field."""

    nodes: tuple
    text: str
    plain: bool = False
    code: str = ""
    results: tuple = ()
    locked: bool = False
    start: int = 0

    @property
    def stop(self):
        return self.start + len(self.text)

    @property
    def words(self):
        return self.code.split()


def field_parts(nodes):
    index = 0
    while index < len(nodes):
        node = nodes[index]
        index += 1
        if node.tag == qn("w:fldSimple"):
            code = node.get(qn("w:instr"), "")
            results = tuple(node.iter(T))
            yield Part((node,), "".join(t.text or "" for t in results), code=code,
                       results=results, locked=node.get(FIELD_LOCK) in {"1", "true", "on"})
            continue

        opening = node.find(FIELD_CHAR) if node.tag == RUN else None
        if opening is not None and opening.get(FIELD_TYPE) == "begin":
            # Word saves native fields as a sequence of runs. Keep nested/foreign fields opaque.
            start = index - 1
            depth = 1
            separated = False
            instruction, results = [], []
            nested = False
            while index < len(nodes) and depth:
                current = nodes[index]
                index += 1
                if current.tag != RUN:
                    continue
                for item in current:
                    if item.tag == FIELD_CHAR:
                        action = item.get(FIELD_TYPE)
                        if action == "begin":
                            depth += 1
                            nested = True
                        elif action == "end":
                            depth -= 1
                        elif action == "separate" and depth == 1:
                            separated = True
                    elif depth == 1 and not separated and item.tag == FIELD_INSTRUCTION:
                        instruction.append(item.text or "")
                    elif depth == 1 and separated and item.tag == T:
                        results.append(item)
            code = "".join(instruction)
            safe = not nested and depth == 0 and separated
            display = "".join(t.text or "" for t in results) if safe else BREAK
            yield Part(tuple(nodes[start:index]), display, code=code,
                       results=tuple(results) if safe else (),
                       locked=opening.get(FIELD_LOCK) in {"1", "true", "on"})
        elif node.tag == RUN:
            content = "".join(node_text(item) for item in node if item.tag != RUN_PROPS)
            yield Part((node,), content, plain=BREAK not in content)
        else:
            yield Part((node,), node_text(node))


def managed_number_field(part):
    if not part.results or part.locked or len(part.words) < 2:
        return False
    operation, argument = part.words[:2]
    return (operation.casefold() == "seq" and argument.casefold() in BY_SEQUENCE
            or operation.casefold() == "ref" and BOOKMARK.fullmatch(argument)
            and any(word.casefold() == r"\h" for word in part.words[2:]))


def clone_fragment(run, start, stop):
    """Copy the selected run text, preserving the run's formatting and line breaks."""
    fragment = deepcopy(run)
    offset = 0
    for child in list(fragment):
        if child.tag == RUN_PROPS:
            continue
        value = node_text(child)
        left, right = max(start - offset, 0), min(stop - offset, len(value))
        offset += len(value)
        if left >= right:
            fragment.remove(child)
        elif child.tag == T:
            set_text(child, value[left:right])
    return fragment


class ParagraphText:
    def __init__(self, element):
        self.element = element
        self.parts = list(field_parts(list(element)))
        offset = 0
        for part in self.parts:
            part.start = offset
            offset = part.stop
        self.text = "".join(part.text for part in self.parts)

    def cover(self, start, stop):
        return [p for p in self.parts if p.start < stop and p.stop > start and p.text]

    def replaceable(self, start, stop):
        parts = self.cover(start, stop)
        if not parts or not all(part.plain for part in parts):
            return False
        first, last = (self.element.index(part.nodes[0]) for part in (parts[0], parts[-1]))
        return all(child.tag == RUN for child in self.element[first:last + 1])

    def properties(self, position):
        return self.cover(position, position + 1)[0].nodes[0].find(RUN_PROPS)

    def split_text_boundaries(self, positions):
        """Expose character boundaries as run boundaries while keeping Word fields intact."""
        view = self
        splits = 0
        for position in sorted(set(positions), reverse=True):
            part = next((p for p in view.cover(position, position + 1)
                         if p.start < position < p.stop), None)
            if part is None:
                continue
            if not part.plain:
                raise ValueError("A formatting boundary crosses protected Word content")
            run = part.nodes[0]
            offset = position - part.start
            run.addprevious(clone_fragment(run, 0, offset))
            run.addprevious(clone_fragment(run, offset, len(part.text)))
            self.element.remove(run)
            view = ParagraphText(self.element)
            splits += 1
        return view, splits


    def replace(self, start, stop, replacement):
        if not self.replaceable(start, stop):
            raise ValueError("The number crosses a Word field, bookmark, or other protected content")
        parts = self.cover(start, stop)
        first, last = parts[0], parts[-1]
        index = self.element.index(first.nodes[0])
        if start > first.start:
            replacement.insert(0, clone_fragment(first.nodes[0], 0, start - first.start))
        if stop < last.stop:
            replacement.append(clone_fragment(last.nodes[0], stop - last.start, len(last.text)))
        for part in parts:
            self.element.remove(part.nodes[0])
        for offset, node in enumerate(replacement):
            self.element.insert(index + offset, node)


def new_run(properties, tag, value=None):
    run = OxmlElement("w:r")
    if properties is not None:
        run.append(deepcopy(properties))
    child = OxmlElement(f"w:{tag}")
    run.append(child)
    if value is not None:
        set_text(child, value)
    return run, child


def native_field(code, value, properties):
    output = []
    for token in ("begin", "code", "separate", "result", "end"):
        if token in {"code", "result"}:
            run, _ = new_run(properties, "instrText" if token == "code" else "t",
                             f" {code} " if token == "code" else value)
        else:
            run, marker = new_run(properties, "fldChar")
            marker.set(FIELD_TYPE, token)
        output.append(run)
    return output


@dataclass
class Target:
    kind: str
    label: str
    old_number: int
    paragraph: int
    start: int
    stop: int
    number_at: int
    bookmark: str
    field: Part | None = None
    number: int = 0

    @property
    def value(self):
        return f"{self.label}{self.number}"


@dataclass
class Reference:
    paragraph: int
    start: int
    stop: int
    targets: tuple


def style_id(element):
    style = element.find(f"{qn('w:pPr')}/{qn('w:pStyle')}")
    return "" if style is None else style.get(qn("w:val"), "").casefold()


def caption_location(element, kind, caption_styles):
    if style_id(element) in caption_styles:
        return True
    adjacent = element.getprevious() if kind == "图" else element.getnext()
    if kind == "图":
        return adjacent is not None and adjacent.tag == qn("w:p") and any(
            child.tag in {qn("w:drawing"), qn("w:pict")} for child in adjacent.iter()
        )
    return adjacent is not None and adjacent.tag == qn("w:tbl")


def read_target(view, part, kind, paragraph):
    nodes = list(view.element)
    at = nodes.index(part.nodes[0])
    before = nodes[:at]
    after = nodes[nodes.index(part.nodes[-1]) + 1:]
    ending = next((node for node in after if node.tag not in
                   {qn("w:proofErr"), qn("w:lastRenderedPageBreak")}), None)
    if ending is None or ending.tag != END:
        return None
    opening = next((node for node in reversed(before)
                    if node.tag == START and node.get(ID) == ending.get(ID)), None)
    if opening is None:
        return None
    name = opening.get(NAME, "")
    if not BOOKMARK.fullmatch(name) or not re.fullmatch(NUMBER, part.text):
        return None
    prefix = "".join(item.text for item in view.parts
                     if nodes.index(opening) < nodes.index(item.nodes[0]) < at)
    if (kind == "文献" and prefix) or (kind != "文献" and not re.fullmatch(rf"{kind}{GAP}*", prefix)):
        return None
    return Target(kind, prefix, int(part.text), paragraph, part.start, part.stop,
                  part.start, name, part)


def citation_numbers(match, available):
    """Expand ranges into individually linked entries so each identity survives renumbering."""
    value = match.group("items")
    chain = re.split(r"[,，、;；]\s*", value)
    numbers = []
    for item in chain:
        ends = RANGE_MARK.split(item)
        if len(ends) == 1:
            numbers.append(int(ends[0]))
        elif len(ends) == 2:
            first, last = map(int, ends)
            if last < first or last - first + 1 > len(available):
                return None
            numbers.extend(range(first, last + 1))
        else:
            return None
    return numbers


def change_result(field, value):
    set_text(field.results[0], value)
    for node in field.results[1:]:
        node.text = ""


def inventory(document):
    body = document.element.body
    elements = [p for p in body.iter(qn("w:p"))
                if p.getparent().tag in {qn("w:body"), qn("w:tc")}
                and all(parent.tag in STORY_ANCESTORS for parent in p.iterancestors())]
    views = [ParagraphText(element) for element in elements]
    # Captions and bibliography entries belong to the main story's top-level paragraphs.
    captions = {"caption", "题注"}
    captions.update(style.style_id.casefold() for style in document.styles
                    if style.name and style.name.casefold() in captions)
    heading_styles = section_heading_styles(document)
    targets, source_views, refs, issues = [], [], [], []
    blocked = set()
    in_bibliography = False
    has_bibliography = False

    def problem(paragraph, kind, message, block=False):
        issues.append({"paragraph": paragraph + 1, "kind": kind, "reason": message})
        if block:
            blocked.add(kind)

    for index, view in enumerate(views):
        text = view.text
        top_level = view.element.getparent() is body
        if top_level and bibliography_heading(text):
            in_bibliography = has_bibliography = True
            continue
        if top_level and in_bibliography and bibliography_ends(view, heading_styles):
            in_bibliography = False

        for part in view.parts:
            type_and_name = part.words[:2]
            if len(type_and_name) != 2:
                continue
            operation, name = (word.casefold() for word in type_and_name)
            if operation == "seq" and name in BY_SEQUENCE:
                kind = BY_SEQUENCE[name]
                target = read_target(view, part, kind, index)
                if target is None:
                    problem(index, kind, "编号域的书签缺失、结构发生变化或结果无法识别。", block=True)
                else:
                    targets.append(target)
            elif operation == "ref" and BOOKMARK.fullmatch(type_and_name[1]):
                refs.append((index, part, type_and_name[1]))

        caption = CAPTION.match(text) if top_level else None
        if caption and not caption_location(view.element, caption["label"][0], captions):
            caption = None
        if caption:
            candidates = [(caption.group("label")[0], caption, "label")]
        elif top_level and in_bibliography:
            candidates = [("文献", match, "number") for match in ENTRY.finditer(text)]
        else:
            candidates = []
            if not style_id(view.element).startswith(("toc", "目录")):
                source_views.append((index, view))

        for kind, match, start_name in candidates:
            start, stop = match.start(start_name), match.end("number")
            parts = view.cover(start, stop)
            if any(p.words and p.words[0].casefold() == "seq" and len(p.words) > 1
                   and BY_SEQUENCE.get(p.words[1].casefold()) == kind for p in parts):
                continue
            if not view.replaceable(start, stop):
                problem(index, kind, "现有编号含 Word 域或特殊格式，请沿用原有编号体系。", block=True)
                continue
            targets.append(Target(kind, match.groupdict().get("label", ""), int(match["number"]),
                                  index, start, stop, match.start("number"), f"DWRef_{uuid4().hex}"))

    targets.sort(key=lambda target: (target.paragraph, target.start))
    counters = Counter()
    by_number, by_bookmark = {}, {}
    duplicate_numbers = set()
    for target in targets:
        counters[target.kind] += 1
        target.number = counters[target.kind]
        key = target.kind, target.old_number
        if key in by_number:
            duplicate_numbers.add(key)
            if target.field is None or by_number[key].field is None:
                problem(target.paragraph, target.kind, f"编号 {target.old_number} 重复，引用目标不唯一。",
                        block=True)
        by_number[key] = target
        key = target.bookmark.casefold()
        if key in by_bookmark:
            problem(target.paragraph, target.kind, "目标书签重复，请移动原条目或为副本重建书签。",
                    block=True)
        by_bookmark[key] = target

    pending = []
    bibliography_numbers = {number for kind, number in by_number if kind == "文献"}
    for index, view in source_views:
        for match in FIGURE_MENTION.finditer(view.text):
            start, stop = match.start(), match.end()
            if not view.replaceable(start, stop):
                continue
            kind, number = match["label"][0], int(match["number"])
            target = by_number.get((kind, number))
            if target is None or (kind, number) in duplicate_numbers:
                problem(index, kind, f"「{match[0]}」缺少唯一可识别的题注目标。", block=True)
            else:
                pending.append(Reference(index, start, stop, (target,)))
        if not has_bibliography:
            continue
        for match in CITATION.finditer(view.text):
            start, stop = match.span("items")
            if RANGE_MARK.search(match["items"]):
                if not view.replaceable(start, stop):
                    continue
                numbers = citation_numbers(match, bibliography_numbers)
                if numbers is None or any(number not in bibliography_numbers
                                          or ("文献", number) in duplicate_numbers for number in numbers):
                    problem(index, "文献", f"「{match[0]}」的文献目标缺失、重复或区间无法对应。", block=True)
                else:
                    pending.append(Reference(index, start, stop,
                                             tuple(by_number["文献", number] for number in numbers)))
                continue
            for item in re.finditer(NUMBER, match["items"]):
                begin, end = start + item.start(), start + item.end()
                if not view.replaceable(begin, end):
                    continue
                number = int(item[0])
                key = "文献", number
                if key not in by_number or key in duplicate_numbers:
                    problem(index, "文献", f"「{match[0]}」中的 {number} 缺少唯一可识别的文献目标。",
                            block=True)
                else:
                    pending.append(Reference(index, begin, end, (by_number[key],)))

    stale = []
    for target in targets:
        if target.field is None:
            continue
        if target.field.locked or not target.field.results:
            problem(target.paragraph, target.kind, "编号域已锁定或缺少结果。", block=True)
        elif target.field.text != str(target.number):
            stale.append((target.kind, target.field, str(target.number)))
    for index, part, bookmark in refs:
        target = by_bookmark.get(bookmark.casefold())
        if target is None or target.field is None:
            problem(index, "引用", f"书签 {bookmark} 的目标已丢失，请核对所引项目。")
        elif part.locked or not part.results or r"\h" not in [word.casefold() for word in part.words]:
            problem(index, target.kind, "引用域缺少可更新的超链接或已被锁定。", block=True)
        elif part.text != target.value:
            stale.append((target.kind, part, target.value))

    return views, targets, pending, stale, blocked, issues, len(refs)


def apply_fields(views, targets, pending, stale, blocked, document):
    existing_ids = (node.get(ID, "") for node in document.element.body.iter(START))
    next_id = max((int(value) for value in existing_ids if value.isdecimal()), default=-1) + 1
    edits = defaultdict(list)
    for target in targets:
        if target.field is None and target.kind not in blocked:
            edits[target.paragraph].append((target.start, target.stop, target))
    for reference in pending:
        if reference.targets[0].kind not in blocked:
            edits[reference.paragraph].append((reference.start, reference.stop, reference))

    for index, changes in edits.items():
        for start, stop, action in sorted(changes, key=lambda edit: edit[0], reverse=True):
            view = ParagraphText(views[index].element)
            if isinstance(action, Target):
                opening, ending = OxmlElement("w:bookmarkStart"), OxmlElement("w:bookmarkEnd")
                opening.set(ID, str(next_id))
                ending.set(ID, str(next_id))
                next_id += 1
                opening.set(NAME, action.bookmark)
                pieces = [opening]
                if action.label:
                    pieces.append(new_run(view.properties(start), "t", action.label)[0])
                pieces.extend(native_field(rf"SEQ {SEQUENCES[action.kind]} \* CHARFORMAT",
                                           str(action.number), view.properties(action.number_at)))
                pieces.append(ending)
            else:
                properties = view.properties(start)
                pieces = []
                for target in action.targets:
                    if pieces:
                        pieces.append(new_run(properties, "t", ",")[0])
                    pieces.extend(native_field(rf"REF {target.bookmark} \h \* CHARFORMAT",
                                               target.value, properties))
            view.replace(start, stop, pieces)

    # Updating an earlier field can change the text offsets of still-pending edits.
    for kind, field, value in stale:
        if kind not in blocked:
            change_result(field, value)


def bibliography_heading(text):
    return bool(REFERENCE_TITLE.fullmatch(re.sub(r"\s+", "", text)))


def bibliography_ends(view, heading_styles):
    return (style_id(view.element) in heading_styles
            or bool(re.match(r"^\s*(?:附录|Appendix\b)", view.text, re.I)))


def section_heading_styles(document):
    return {style.style_id.casefold() for style in document.styles
            if style.name and re.fullmatch(r"(?:heading|标题)\s*\d+", style.name, re.I)}


def split_entry_paragraph(view, starts):
    """Split separate numbered entries at their original line breaks."""
    if len(starts) == 1:
        return [view.element]

    line_break_tags = {qn("w:br"), qn("w:cr")}
    for begin in starts[1:]:
        offset = begin - 1
        parts = view.cover(offset, begin)
        if len(parts) != 1 or not parts[0].plain:
            raise ValueError("无法在文献编号前辨识可拆分的换行。")
        part = parts[0]
        position = part.start
        for child in part.nodes[0]:
            if child.tag == RUN_PROPS:
                continue
            if position == offset and child.tag in line_break_tags:
                if child.get(qn("w:type"), "textWrapping") != "textWrapping":
                    raise ValueError("文献条目跨越手动分页或分栏，请先确认条目边界。")
                break
            position += len(node_text(child))
        else:
            raise ValueError("未找到可安全拆分的文献条目换行。")

    # End positions leave out the manual line break that becomes a paragraph boundary.
    limits = [0, *starts[1:]]
    stops = [*(offset - 1 for offset in starts[1:]), len(view.text)]
    paragraphs = []
    for ordinal, (begin, end) in enumerate(zip(limits, stops)):
        paragraph = deepcopy(view.element)
        for child in list(paragraph):
            if child.tag != qn("w:pPr"):
                paragraph.remove(child)
        if ordinal:
            for name in ("paraId", "textId"):
                paragraph.attrib.pop(qn(f"w14:{name}"), None)
        for part in view.parts:
            if part.nodes[0].tag == qn("w:pPr"):
                continue
            if begin <= part.start and part.stop <= end:
                paragraph.extend(deepcopy(node) for node in part.nodes)
            elif part.start < end and part.stop > begin:
                if not part.plain:
                    raise ValueError("文献条目换行穿过 Word 域或特殊对象。")
                paragraph.append(clone_fragment(part.nodes[0], begin - part.start, end - part.start))
        paragraphs.append(paragraph)
    return paragraphs


def bibliography_entry_blocks(body, views, entries, heading_styles):
    """Collect paragraphs that move together with each bookmarked entry."""
    by_paragraph = defaultdict(list)
    by_element = {view.element: view for view in views}
    for target in entries:
        by_paragraph[views[target.paragraph].element].append(target)
    elements = list(body)
    first = elements.index(views[entries[0].paragraph].element)
    last = elements.index(views[entries[-1].paragraph].element)
    for following in elements[last + 1:]:
        view = by_element.get(following)
        if view is None or not view.text.strip() or bibliography_ends(view, heading_styles):
            break
        last += 1
    region = elements[first:last + 1]
    blocks = {}
    current = None
    hard_break_types = {"page", "column"}
    for element in region:
        view = by_element.get(element)
        if view is None or bibliography_ends(view, heading_styles):
            raise ValueError("参考文献不是可安全重排的连续段落。")
        if (element.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None
                or any(node.get(qn("w:type")) in hard_break_types for node in element.iter(qn("w:br")))):
            raise ValueError("参考文献之间有分页、分栏或分节边界，请先确认其归属。")

        targets = by_paragraph.get(element)
        if not targets:
            if current is None:
                raise ValueError("未能确定参考文献续段的所属条目。")
            current.append(element)
            continue

        start_by_number = {match.start("number"): match.start() for match in ENTRY.finditer(view.text)}
        starts = [start_by_number.get(target.start) for target in targets]
        if None in starts or view.text[:starts[0]].strip():
            raise ValueError("参考文献条目应以独立括号编号起行。")
        paragraphs = split_entry_paragraph(view, starts)
        for target, paragraph in zip(targets, paragraphs):
            current = [paragraph]
            blocks[target.bookmark.casefold()] = current

    # Check that moving an entry cannot separate the two ends of a Word range.
    ranges = (("bookmarkStart", "bookmarkEnd"), ("commentRangeStart", "commentRangeEnd"),
              ("moveFromRangeStart", "moveFromRangeEnd"), ("moveToRangeStart", "moveToRangeEnd"),
              ("permStart", "permEnd"))
    for block in blocks.values():
        for start, stop in ranges:
            openings = Counter(node.get(ID) for p in block for node in p.iter(qn(f"w:{start}")))
            closings = Counter(node.get(ID) for p in block for node in p.iter(qn(f"w:{stop}")))
            if openings != closings:
                raise ValueError("有书签、批注或修订范围跨越参考文献条目，请先确认。")
    return first, region, blocks


def sort_bibliography(document, views, targets, *, update):
    """Order entries by the first main-text REF to their stable bookmarks."""
    body = document.element.body
    entries = [target for target in targets if target.kind == "文献"]
    titles = [index for index, view in enumerate(views)
              if view.element.getparent() is body and bibliography_heading(view.text)]
    if not entries:
        if titles:
            raise ValueError("参考文献区没有可用于首引排序的数字条目。")
        return {"entries": 0, "cited": 0, "uncited": 0, "moved": 0, "input_positions_in_new_order": []}

    if len(titles) != 1 or any(target.paragraph <= titles[0] for target in entries):
        raise ValueError("首引排序需要一个能与正文明确对应的参考文献区。")
    if any(views[target.paragraph].element.getparent() is not body for target in entries):
        raise ValueError("参考文献条目必须位于主文档的参考文献段落中。")

    heading_styles = section_heading_styles(document)
    for view in views[titles[0] + 1:entries[-1].paragraph + 1]:
        if view.element.getparent() is body and bibliography_ends(view, heading_styles):
            raise ValueError("有参考文献编号位于参考文献区之外。")
    by_bookmark = {target.bookmark.casefold(): target for target in entries}
    first_cited = {}
    for view in views[:titles[0]]:
        if style_id(view.element).startswith(("toc", "目录")):
            continue
        for part in view.parts:
            words = part.words
            if len(words) >= 2 and words[0].casefold() == "ref" and words[1].casefold() in by_bookmark:
                first_cited.setdefault(words[1].casefold(), None)

    order = [*first_cited, *(name for name in by_bookmark if name not in first_cited)]
    moved = sum(name != old for name, old in zip(order, by_bookmark))
    report = {
        "entries": len(entries),
        "cited": len(first_cited),
        "uncited": len(entries) - len(first_cited),
        "moved": moved,
        "input_positions_in_new_order": [by_bookmark[name].number for name in order],
    }
    if not moved:
        return report

    first, region, blocks = bibliography_entry_blocks(body, views, entries, heading_styles)
    if update:
        for element in region:
            body.remove(element)
        for offset, element in enumerate(p for name in order for p in blocks[name]):
            body.insert(first + offset, element)
    return report


def superscript_citations(views, targets, *, update):
    """Set the complete [n] or [n,m] marker, including REF code formatting."""
    names = {target.bookmark.casefold() for target in targets if target.kind == "文献"}
    changed = 0
    for view in views:
        if style_id(view.element).startswith(("toc", "目录")):
            continue
        if view.element.getparent().tag == qn("w:body") and bibliography_heading(view.text):
            break
        spans = []
        for match in CITATION.finditer(view.text):
            parts = view.cover(*match.span())
            if not any(len(part.words) > 1 and part.words[0].casefold() == "ref"
                       and part.words[1].casefold() in names for part in parts):
                continue
            if any(run.rPr is None or run.rPr.superscript is not True
                   for part in parts for node in part.nodes for run in node.iter(RUN)):
                spans.append(match.span())
        changed += len(spans)
        if update and spans:
            view, _ = view.split_text_boundaries(bound for span in spans for bound in span)
            for start, stop in spans:
                for part in view.cover(start, stop):
                    for node in part.nodes:
                        for run in node.iter(RUN):
                            run.get_or_add_rPr().superscript = True
    return changed


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("operation", choices=("apply", "check"))
    parser.add_argument("input", type=Path)
    parser.add_argument("--out", type=Path)
    parser.add_argument("--sort-bibliography", action="store_true",
                        help="order bibliography entries by first citation in the main text")
    parser.add_argument("--superscript-citations", action="store_true",
                        help="superscript linked numeric bibliography citations in the main text")
    args = parser.parse_args()

    try:
        if args.input.suffix.casefold() != ".docx":
            raise ValueError("Input must be a DOCX")
        if args.operation == "apply":
            if args.out is None or args.out.suffix.casefold() != ".docx":
                raise ValueError("apply requires a new DOCX via --out")
            if args.out.exists() or args.input.resolve() == args.out.resolve():
                raise ValueError("Choose a new output path")
        elif args.out is not None:
            raise ValueError("--out is for apply")

        document = Document(args.input)
        views, targets, pending, stale, blocked, issues, ref_count = inventory(document)
        report = {
            "targets": dict(Counter(target.kind for target in targets)),
            "active_references": ref_count,
            "pending_targets": sum(target.field is None for target in targets),
            "pending_references": sum(len(reference.targets) for reference in pending),
            "outdated_fields": len(stale),
            "review": issues,
        }
        needs_changes = bool(report["pending_targets"] or pending or stale or issues)
        if args.operation == "check" and args.superscript_citations:
            markers = superscript_citations(views, targets, update=False)
            report["citation_markers_needing_superscript"] = markers
            needs_changes = needs_changes or bool(markers)
        sorted_in_memory = False
        if args.sort_bibliography and not (blocked or issues):
            # Link plain citations first so ordering never depends on a number that is about to change.
            apply_fields(views, targets, pending, stale, blocked, document)
            sorted_in_memory = True
            linked_views, linked_targets, unlinked, _, errors, linked_issues, _ = inventory(document)
            if unlinked or errors or linked_issues:
                raise ValueError("首引排序前未能核对全部引用目标。")
            order = sort_bibliography(document, linked_views, linked_targets, update=args.operation == "apply")
            report["bibliography_order"] = order
            needs_changes = needs_changes or bool(order["moved"])
            if args.operation == "apply" and order["moved"]:
                _, _, detached, new_numbers, errors, linked_issues, _ = inventory(document)
                if detached or errors or linked_issues:
                    raise ValueError("重排后无法核对文献书签。")
                for kind, field, value in new_numbers:
                    change_result(field, value)
                report["outdated_fields"] += len(new_numbers)
        if args.operation == "check":
            print(json.dumps(report, ensure_ascii=False, indent=2))
            return int(needs_changes)
        if blocked or (args.sort_bibliography and issues):
            report["blocked"] = sorted(blocked)
            print(json.dumps(report, ensure_ascii=False, indent=2))
            return 1

        if not sorted_in_memory:
            apply_fields(views, targets, pending, stale, blocked, document)
        if args.superscript_citations:
            linked_views, linked_targets, *_ = inventory(document)
            report["citation_markers_superscripted"] = superscript_citations(
                linked_views, linked_targets, update=True,
            )
        output = BytesIO()
        document.save(output)
        output.seek(0)
        reopened = Document(output)
        final_views, verified, remaining, outdated, failures, review, count = inventory(reopened)
        if (remaining or outdated or failures or any(target.field is None for target in verified)
                or len(verified) != len(targets) or count != ref_count + report["pending_references"]):
            raise ValueError("The resulting cross-references could not be verified")
        if args.sort_bibliography and sort_bibliography(reopened, final_views, verified, update=False)["moved"]:
            raise ValueError("保存后的参考文献与正文首引顺序不一致。")
        if args.superscript_citations and superscript_citations(final_views, verified, update=False):
            raise ValueError("保存后的正文文献标记仍有上标差异。")
        with args.out.open("xb") as destination:
            try:
                destination.write(output.getbuffer())
                destination.flush()
            except OSError:
                destination.close()
                args.out.unlink()
                raise
        report.update({
            "output": str(args.out),
            "active_references": count,
            "created_targets": report.pop("pending_targets"),
            "created_references": report.pop("pending_references"),
            "refreshed_fields": report.pop("outdated_fields"),
            "review": review,
        })
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return int(bool(review))
    except (OSError, ValueError, BadZipFile, PackageNotFoundError, XMLSyntaxError) as error:
        print(f"word_crossrefs: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
