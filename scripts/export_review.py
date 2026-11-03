"""Export an applicant's annotated draft and its linked simulated peer review."""

import argparse
import io
import json
import re
import shutil
import sys
import tempfile
from copy import deepcopy
from pathlib import Path, PurePosixPath
from zipfile import ZipFile, ZIP_DEFLATED

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from lxml import etree


REVIEWERS = ("专家一", "专家二", "专家三", "专家四", "专家五")
REPORT_FONT = "宋体"
LEGACY_SPECIALTIES = {
    "公共卫生管理": "公共卫生专家（卫生服务与政策）",
    "护理管理": "护理学专家（护理管理）",
    "药事管理": "药学管理专家（药事管理）",
}
LEGACY_PANELS = ("学科与科学价值", "研究方法", "实施可行性")
COUNTS = {"A": 8, "B": 8, "C": 8, "D": 8, "E": 6, "F": 10, "G": 8, "H": 10, "X": 4}
GRADES = ("优", "良", "中", "差")
FUNDING = {"A": "优先资助", "B": "可资助", "C": "不予资助"}
QUESTIONS = (
    "一、请评述该申请项目是否面向经济社会发展需要或国家需求背后的基础科学问题，请详细阐述判断理由。",
    "二、请评述申请所提出的科学问题的创新性与预期成果的科学价值。",
    "三、请评述申请人的创新潜力与研究方案的可行性。如有可能，请对进一步完善研究方案提出意见。",
    "四、其他建议。",
)
GROUPS = {"A": "标题选择", "B": "背景和意义", "C": "文献回顾", "D": "研究内容",
          "E": "关键科学问题", "F": "研究方法", "G": "研究基础与工作条件",
          "H": "其他部分与提交核对", "X": "跨章节核对", "M": "适用的研究方法专项"}
CORE_RULES = tuple(f"{group}{n:02d}" for group, count in COUNTS.items() for n in range(1, count + 1))
METHOD_RULES = tuple(f"M{n:02d}" for n in range(1, 9))


def required(obj, *keys):
    for key in keys:
        if not isinstance(obj.get(key), str) or not obj[key].strip():
            raise ValueError(f"缺少文字或为空：{key}")


def rule_id(value):
    match = re.fullmatch(r"([A-HX])(\d{2})-(\d{2,})", value)
    if not match or not (1 <= int(match[2]) <= COUNTS[match[1]]) or int(match[3]) == 0:
        raise ValueError(f"无效的问题编号：{value}")
    return value


def uses_checklist(data):
    return (any(key in data for key in ("overall_grade", "overall_funding", "checklist"))
            or any("opinions" in p or "grade" in p for p in data["panels"]))


def check_refs(refs, ids, location):
    if not isinstance(refs, list) or len(refs) != len(set(refs)) or not set(refs) <= ids:
        raise ValueError(f"{location} 引用了无效或重复的问题编号")


def check_ratings(grade, funding, location):
    if grade not in GRADES:
        raise ValueError(f"{location} 综合评价必须为优、良、中或差")
    if funding not in FUNDING:
        raise ValueError(f"{location} 资助意见必须为 A、B 或 C")


def validate_form(data, ids):
    check_ratings(data.get("overall_grade"), data.get("overall_funding"), "总体")
    for i, p in enumerate(data["panels"], 1):
        check_ratings(p.get("grade"), p.get("funding"), f"第{i}份评议")
        opinions = p.get("opinions")
        if (not isinstance(opinions, list) or len(opinions) != len(QUESTIONS)
                or not all(isinstance(t, str) and t.strip() for t in opinions)):
            raise ValueError(f"第{i}份评议需按顺序回答四项具体评价意见")
    rows = data.get("checklist")
    if not isinstance(rows, list):
        raise ValueError("整稿评价需提供逐项 checklist")
    seen, linked = set(), set()
    for row in rows:
        required(row, "rule", "status", "location", "evidence")
        key = row["rule"]
        if key not in CORE_RULES + METHOD_RULES or key in seen:
            raise ValueError(f"清单编号无效或重复：{key}")
        seen.add(key)
        if row["status"] not in ("通过", "需修改", "待核实", "不适用"):
            raise ValueError(f"{key} 状态无效")
        if row["status"] in ("需修改", "待核实"):
            required(row, "action")
        elif not isinstance(row.get("action", ""), str):
            raise ValueError(f"{key} 的 action 应为文本")
        check_refs(row.get("issue_ids"), ids, key)
        linked.update(row["issue_ids"])
    missing = [key for key in CORE_RULES if key not in seen]
    if missing:
        raise ValueError(f"缺少清单行：{'、'.join(missing)}")
    if ids - linked:
        raise ValueError(f"活动批注未在清单引用：{'、'.join(sorted(ids - linked))}")


def validate_legacy(data, ids):
    required(data, "grading_basis", "differences")
    for strength in data["strengths"]:
        required(strength, "location", "text")
    priorities = data["priority_ids"]
    if len(priorities) > 3 or len(set(priorities)) != len(priorities) or not set(priorities) <= ids:
        raise ValueError("首页优先项应为不重复的有效编号，最多三项")
    if ids and not priorities:
        raise ValueError("存在活动问题时请指定首页处理顺序")
    for panel in data["panels"]:
        required(panel, "understanding", "strongest", "judgment", "funding", "reason")
        fields = panel.get("form_fields", [])
        if not isinstance(fields, list):
            raise ValueError("form_fields 应为按原表顺序排列的字段列表")
        for field in fields:
            required(field, "label", "answer")


def validate(data, *, legacy=False):
    required(data, "title", "context", "summary")
    for key in ("issues", "panels"):
        if not isinstance(data.get(key), list):
            raise ValueError(f"{key} 应为列表")
    ids = set()
    for issue in data["issues"]:
        required(issue, "id", "location", "problem", "impact", "action", "done")
        issue_id = rule_id(issue["id"])
        if issue_id in ids:
            raise ValueError(f"重复编号：{issue_id}")
        ids.add(issue_id)
        if issue.get("severity") not in ("关键问题", "重要问题", "局部问题"):
            raise ValueError(f"{issue_id} 缺有效影响程度")
        if issue.get("status") not in ("需修改", "待核实"):
            raise ValueError(f"{issue_id} 活动问题状态须为需修改或待核实")
        if issue.get("basis") not in ("年度官方要求", "研究方法判断", "团队写作经验"):
            raise ValueError(f"{issue_id} 缺有效依据类别")
        required(issue["anchor"], "quote")
        if issue["basis"] == "年度官方要求":
            required(issue, "source")
        if "comment" in issue:
            required(issue, "comment")
    form = uses_checklist(data)
    if legacy:
        perspectives = {p.get("perspective") for p in data["panels"]}
        historical = perspectives == set(LEGACY_PANELS) or (form and perspectives == set(LEGACY_SPECIALTIES))
        if len(data["panels"]) != 3 or not historical:
            raise ValueError("--legacy 仅用于复现已有的整组三人评议，请核对原有身份")
    elif not form or len(data["panels"]) != len(REVIEWERS):
        raise ValueError("新评须提供五份四问评议。仅复现已有三人旧稿时使用 --legacy")
    covered = set()
    for i, panel in enumerate(data["panels"], 1):
        check_refs(panel["issue_ids"], ids, f"第{i}份评议")
        covered.update(panel["issue_ids"])
    if form:
        validate_form(data, ids)
    else:
        validate_legacy(data, ids)
    if any(i["id"] not in covered for i in data["issues"] if i["severity"] != "局部问题"):
        raise ValueError("关键和重要问题须在至少一份模拟意见中引用")
    prior_ids = set()
    for prior in data.get("rechecks", []):
        required(prior, "id", "location", "note")
        key = rule_id(prior["id"])
        if key in prior_ids or prior.get("status") not in ("已解决", "部分解决", "仍待处理", "待核实"):
            raise ValueError(f"复查记录无效：{key}")
        prior_ids.add(key)
        if prior["status"] == "已解决" and key in ids:
            raise ValueError(f"已解决的问题应退出活动清单：{key}")
        if prior["status"] in ("部分解决", "仍待处理") and key not in ids:
            raise ValueError(f"仍需修改的复查项应以原编号在新稿定位：{key}")


def paragraphs(doc):
    # Document order, including table cells. Nested drawing/textbox content
    # and revision wrappers need a different Word story / revision view.
    return [Paragraph(p, doc._body) for p in doc.element.body.xpath(
        ".//w:p[not(ancestor::w:p or ancestor::w:del or ancestor::w:ins)]"
    )]


def plain_run(run):
    return all(child.tag in (qn("w:rPr"), qn("w:t")) for child in run._r)


def split_run(run, offset):
    """Split textual runs only, keeping both halves' character properties."""
    if not plain_run(run):
        raise ValueError("锚点切入域、图片、换行或特殊对象，请选附近独立文本并注明原位置")
    text = run.text
    second = Run(deepcopy(run._r), run._parent)
    run._r.addnext(second._r)
    run.text = text[:offset]
    second.text = text[offset:]
    return second


def anchor_runs(paragraph, quote):
    # python-docx includes hyperlink text in Paragraph.text but omits it
    # from Paragraph.runs. Map offsets using paragraph content in order.
    start = paragraph.text.index(quote)
    end = start + len(quote)
    cursor = 0
    selected = []
    for item in paragraph.iter_inner_content():
        for run in ([item] if isinstance(item, Run) else item.runs):
            stop = cursor + len(run.text)
            if stop > start and cursor < end and stop > cursor:
                if run._r.getparent() is not paragraph._p or not plain_run(run):
                    raise ValueError("选区包含链接、域或复合 run，请选正文纯文本作为邻近锚点")
                if end < stop:
                    split_run(run, end - cursor)
                if start > cursor:
                    run = split_run(run, start - cursor)
                selected.append(run)
            cursor = stop
    if not selected or "".join(r.text for r in selected) != quote:
        raise ValueError("原文范围无法完整匹配，请重新定位")
    # Do not accidentally span invisible drawings, fields or revisions.
    first, last = selected[0]._r, selected[-1]._r
    siblings = list(paragraph._p)
    for e in siblings[siblings.index(first): siblings.index(last) + 1]:
        if e.tag == qn("w:r"):
            r = Run(e, paragraph)
            if not plain_run(r) and not e.xpath("./w:commentReference"):
                raise ValueError("原句跨图形、域或其他对象，请缩短引文")
        elif e.tag not in {qn("w:commentRangeStart"), qn("w:commentRangeEnd"),
                           qn("w:bookmarkStart"), qn("w:bookmarkEnd")}:
            raise ValueError("原句跨链接或修订等对象，请改用单一普通文本锚点")
    return selected


def resolve(doc_paragraphs, anchor):
    quote = anchor["quote"]
    number = anchor.get("paragraph")
    if number is None:
        matches = [p for p in doc_paragraphs if quote in p.text]
    elif type(number) is int and 1 <= number <= len(doc_paragraphs):
        matches = [doc_paragraphs[number - 1]]
    else:
        raise ValueError(f"无效的正文段号 {number}")
    if len(matches) != 1 or matches[0].text.find(quote) < 0 or (
        matches[0].text.find(quote, matches[0].text.find(quote) + 1) >= 0
    ):
        raise ValueError("原句缺失或重复，请延长原文短引或用 --list-paragraphs 指定段号")
    return anchor_runs(matches[0], quote)


def annotate(doc, runs, issue):
    note = issue.get("comment") or f"{issue['problem']} {issue['action']}"
    lines = [f"{issue['id']}　{note}"]
    if issue.get("source"):
        lines.append(issue["source"])
    comment = doc.add_comment(runs, text="\n".join(lines), author="申报前模拟评阅", initials="评")
    for p in comment.paragraphs:
        for run in p.runs:
            set_song_font(run._r.get_or_add_rPr().get_or_add_rFonts())


def save_with_comments(doc, source, target):
    """Patch only the two Word parts and, for a first comment, their OPC links."""
    body = doc.part
    rel = next(r for r in body.rels.values() if r.reltype == RT.COMMENTS)
    body_path = PurePosixPath(str(body.partname).lstrip("/"))
    comments_path = str(rel.target_part.partname)
    rels_path = str(body_path.parent / "_rels" / (body_path.name + ".rels"))
    patches = {str(body_path): body.blob, comments_path.lstrip("/"): rel.target_part.blob}
    with ZipFile(source) as original, ZipFile(target, "w", compression=ZIP_DEFLATED) as output:
        if any(name.startswith("_xmlsignatures/") for name in original.namelist()):
            raise ValueError("签名稿的签名无法随修改保留，请提供可编辑的未签名副本")
        if comments_path.lstrip("/") not in original.namelist():
            for path, tag, attributes in (
                (rels_path, "Relationship", {"Id": rel.rId, "Type": RT.COMMENTS, "Target": rel.target_ref}),
                ("[Content_Types].xml", "Override",
                 {"PartName": comments_path, "ContentType": rel.target_part.content_type}),
            ):
                if path in original.namelist():
                    root = etree.fromstring(original.read(path))
                else:
                    ns = "http://schemas.openxmlformats.org/package/2006/relationships"
                    root = etree.Element(f"{{{ns}}}Relationships", nsmap={None: ns})
                etree.SubElement(root, etree.QName(etree.QName(root).namespace, tag), attributes)
                patches[path] = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
        for item in original.infolist():
            data = patches.pop(item.filename, None)
            output.writestr(item, original.read(item) if data is None else data)
        for name, data in patches.items():
            output.writestr(name, data)


def set_song_font(fonts):
    for attribute in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme", "csTheme"):
        fonts.attrib.pop(qn(f"w:{attribute}"), None)
    for attribute in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn(f"w:{attribute}"), REPORT_FONT)


def finish_report_fonts(doc):
    """Unify a generated document, including inherited table styles and theme fallbacks."""
    for fonts in doc.styles.element.xpath(".//w:rFonts"):
        set_song_font(fonts)
    for run in doc.element.xpath(".//w:r[w:t]"):
        set_song_font(run.get_or_add_rPr().get_or_add_rFonts())
    theme = doc.part.part_related_by(RT.THEME)
    root = etree.fromstring(theme.blob)
    drawing_ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    for face in root.xpath("//a:fontScheme//*[@typeface]", namespaces=drawing_ns):
        face.set("typeface", REPORT_FONT)
    theme._blob = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
    return doc


def style_doc():
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Mm(210), Mm(297)
    section.left_margin = section.right_margin = Mm(24)
    for name, size in (("Normal", 11), ("Title", 21), ("Heading 1", 15), ("Heading 2", 12)):
        style = doc.styles[name]
        style.font.name = REPORT_FONT
        style.font.size = Pt(size)
        style.font.color.rgb = RGBColor.from_string("202020")
        if name == "Title":
            for border in style._element.xpath("./w:pPr/w:pBdr"):
                border.getparent().remove(border)
    doc.styles["Normal"].paragraph_format.space_after = Pt(5)
    doc.styles["Normal"].paragraph_format.line_spacing = 1.25
    return doc


def annotate_pdf(source, data):
    import pypdfium2 as pdfium
    doc = style_doc()
    doc.add_paragraph("申请书修改批注摘录", style="Title")
    doc.add_paragraph(data["title"])
    doc.add_paragraph(f"来源　{source.name}　　{data['context']}")
    doc.add_paragraph("页码按原 PDF 含封面的物理页顺序计。批注锚定在对应页的引文。")
    with pdfium.PdfDocument(source) as pdf:
        for issue in data["issues"]:
            anchor = issue["anchor"]
            page_no, quote = anchor.get("page"), anchor["quote"]
            if type(page_no) is not int or not 1 <= page_no <= len(pdf):
                raise ValueError(f"{issue['id']} 的 PDF 物理页码无效")
            page = pdf[page_no - 1]
            page_text = "".join(page.get_textpage().get_text_bounded().split())
            if "".join(quote.split()) not in page_text:
                if anchor.get("visual_verified") is not True:
                    raise ValueError(f"{issue['id']} 引文未在物理页 {page_no} 找到，需核对页图")
                img = io.BytesIO()
                bitmap = page.render(scale=1.5)
                bitmap.to_pil().save(img, format="PNG")
                doc.add_picture(img, width=Mm(138))
            doc.add_heading(f"原 PDF 第 {page_no} 页　{issue['location']}", 2)
            annotate(doc, doc.add_paragraph().add_run(quote), issue)
            page.close()
    if not data["issues"]:
        doc.add_paragraph("本轮无待定位的修改事项。")
    return finish_report_fonts(doc)


def legacy_report(data):
    doc = style_doc()
    doc.add_paragraph("国家自然科学基金申请前模拟函评结果", style="Title")
    doc.add_paragraph(data["title"])
    doc.add_paragraph(f"{data['context']}　旧版复现")
    doc.add_paragraph(data["summary"])
    doc.add_heading("值得保留的论证", 1)
    for s in data["strengths"][:2]:
        doc.add_paragraph(f"{s['text']}（{s['location']}）")
    if not data["strengths"]:
        doc.add_paragraph("现有材料尚不足以定位可确认的优势。")
    issues = {i["id"]: i for i in data["issues"]}
    doc.add_heading("优先修改顺序", 1)
    for key in data["priority_ids"]:
        i = issues[key]
        doc.add_paragraph(f"{key}　{i['location']}　{i['action']}")
    if not data["priority_ids"]:
        doc.add_paragraph("本轮无优先修改事项。")
    doc.add_heading("评阅视角的分歧", 1)
    doc.add_paragraph(data["differences"])
    doc.add_paragraph(f"档位依据　{data['grading_basis']}")
    for index, panel in enumerate(data["panels"], 1):
        heading = doc.add_heading(f"模拟函评 {index}　{panel['perspective']}", 1)
        if index == 1:
            heading.paragraph_format.page_break_before = True
        doc.add_paragraph(panel["understanding"])
        doc.add_paragraph(f"最强依据　{panel['strongest']}")
        for key in panel["issue_ids"]:
            i = issues[key]
            doc.add_paragraph(f"{key}　{i['problem']}　{i['impact']}")
        if not panel["issue_ids"]:
            doc.add_paragraph("该视角下未发现能由所给材料定位的主要疑虑。")
        doc.add_paragraph(f"模拟综合判断　{panel['judgment']}　　模拟资助建议　{panel['funding']}")
        doc.add_paragraph(panel["reason"])
        for field in panel.get("form_fields", []):
            doc.add_heading(field["label"], 2)
            doc.add_paragraph(field["answer"])
    if issues:
        doc.add_heading("原稿批注索引", 1)
        for key, i in issues.items():
            doc.add_paragraph(f"{key}　{i['severity']}　{i['location']}　{i['problem']}")
    return doc


def rule_titles():
    refs = Path(__file__).resolve().parents[1] / "references"
    checklist = (refs / "review-checklist.md").read_text(encoding="utf-8")
    methods = (refs / "method-checks.md").read_text(encoding="utf-8")
    names = dict(re.findall(r"^\| ([A-HX]\d{2})\s+([^|\n]+?)\s*\|", checklist, re.M))
    names.update(re.findall(r"^## (M\d{2})\s+([^\n]+)", methods, re.M))
    if not set(CORE_RULES + METHOD_RULES) <= names.keys():
        raise ValueError("技能参考文件缺少检查项名称，请核对规则编号及标题")
    return names


def review_table(doc, widths, headings, rows):
    table = doc.add_table(rows=0, cols=len(widths))
    table.autofit = False
    table.style = "Light Shading Accent 1"
    for c, width in zip(table.columns, widths):
        c.width = Mm(width)
    for i, values in enumerate([headings, *rows]):
        row = table.add_row()
        properties = row._tr.get_or_add_trPr()
        properties.append(OxmlElement("w:cantSplit"))
        if i == 0:
            properties.append(OxmlElement("w:tblHeader"))
        for cell, value, width in zip(row.cells, values, widths):
            cell.width = Mm(width)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.1
            if i == 0:
                p.paragraph_format.keep_with_next = True
            run = p.add_run(value)
            run.bold = i == 0
            run.font.size = Pt(10)


def funding_label(value):
    return f"{FUNDING[value]}（{value}）"


def form_report(data, *, legacy=False):
    doc = style_doc()
    doc.sections[0].top_margin = doc.sections[0].bottom_margin = Mm(20)
    doc.add_paragraph("国家自然科学基金申请前模拟函评结果", style="Title")
    doc.add_paragraph(data["title"])
    doc.add_paragraph(f"{data['context']}　旧版复现" if legacy else data["context"])
    doc.add_heading("总体意见", 1)
    doc.add_paragraph(f"综合评价　{data['overall_grade']}　　　资助意见　{funding_label(data['overall_funding'])}")
    doc.add_paragraph(data["summary"])
    doc.add_heading("各评阅意见", 2)
    panels = data["panels"]
    panel_names = REVIEWERS
    if legacy:
        panel_names = [LEGACY_SPECIALTIES.get(p["perspective"], p["perspective"]) for p in panels]
    review_table(doc, (80, 34, 48), ("模拟评阅", "综合评价", "资助意见"),
                 [(name, p["grade"], funding_label(p["funding"])) for name, p in zip(panel_names, panels)])
    for i, (name, panel) in enumerate(zip(panel_names, panels), 1):
        heading = doc.add_heading(f"模拟函评 {i}　{name}" if legacy else name, 1)
        heading.paragraph_format.page_break_before = i == 1
        ratings = doc.add_paragraph(f"综合评价　{panel['grade']}　　　资助意见　{funding_label(panel['funding'])}")
        ratings.paragraph_format.keep_with_next = True
        for question, answer in zip(QUESTIONS, panel["opinions"]):
            doc.add_heading(question, 2)
            for part in answer.splitlines():
                if part.strip():
                    para = doc.add_paragraph(part.strip())
                    para.paragraph_format.widow_control = True
        if panel["issue_ids"]:
            doc.paragraphs[-1].paragraph_format.keep_with_next = True
            doc.add_paragraph("相关批注　" + "、".join(panel["issue_ids"]))
    doc.add_heading("逐项评阅清单", 1)
    names = rule_titles()
    entries = {r["rule"]: r for r in data["checklist"]}
    for group, label in GROUPS.items():
        rows = []
        for key in sorted(k for k in entries if k.startswith(group)):
            r = entries[key]
            refs = "、".join(r["issue_ids"])
            action = r.get("action") or "—"
            if refs:
                action = f"{action}\n{refs}"
            rows.append((f"{key}　{names[key]}", r["status"], r["location"],
                         r["evidence"], action))
        if rows:
            doc.add_heading(f"{group}　{label}", 2)
            review_table(doc, (32, 18, 26, 47, 39),
                         ("检查项", "结果", "原稿位置", "原文依据与判断", "修改建议及批注"), rows)
    return doc


def report(data, *, legacy=False):
    validate(data, legacy=legacy)
    doc = form_report(data, legacy=legacy) if uses_checklist(data) else legacy_report(data)
    if data.get("rechecks"):
        doc.add_heading("二轮复查", 1)
        review_table(doc, (22, 22, 35, 83), ("原编号", "状态", "新稿位置", "复查意见"),
                     [(r["id"], r["status"], r["location"], r["note"]) for r in data["rechecks"]])
    return finish_report_fonts(doc)


def export(source, review, destination, *, legacy=False):
    data = json.loads(review.read_text(encoding="utf-8-sig"))
    if destination.exists():
        raise ValueError("目标目录已存在，请使用新的输出目录")
    review_document = report(data, legacy=legacy)
    with tempfile.TemporaryDirectory(dir=destination.parent) as tmp:
        bundle = Path(tmp) / "bundle"
        bundle.mkdir()
        suffix = source.suffix.lower()
        if suffix == ".pdf":
            annotated = annotate_pdf(source, data)
            name = "申请书修改批注摘录.docx"
            annotated.save(bundle / name)
        elif suffix == ".docx":
            annotated = Document(source)
            before = paragraphs(annotated)
            for issue in data["issues"]:
                try:
                    annotate(annotated, resolve(before, issue["anchor"]), issue)
                except ValueError as exc:
                    raise ValueError(f"{issue['id']}：{exc}") from exc
            name = "申请书修改批注版.docx"
            if data["issues"]:
                save_with_comments(annotated, source, bundle / name)
            else:
                shutil.copyfile(source, bundle / name)
        else:
            raise ValueError("源文件须为 DOCX 或 PDF")
        review_document.save(bundle / "国家自然科学基金申请前模拟函评结果.docx")
        for output in bundle.iterdir():
            with ZipFile(output) as archive:
                if archive.testzip():
                    raise ValueError(f"文件完整性检查失败：{output.name}")
            Document(output)
        bundle.rename(destination)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--list-paragraphs", type=Path, metavar="DOCX")
    parser.add_argument("--legacy", action="store_true", help="按已有三人评议记录复现旧版报告")
    parser.add_argument("paths", nargs="*", type=Path, metavar="PATH")
    args = parser.parse_args()
    try:
        if args.list_paragraphs and not args.paths and not args.legacy:
            for n, p in enumerate(paragraphs(Document(args.list_paragraphs)), 1):
                if p.text:
                    print(f"{n}\t{p.text}")
        elif len(args.paths) == 3 and not args.list_paragraphs:
            export(*args.paths, legacy=args.legacy)
        else:
            parser.error("请提供 源文件 审阅JSON 新输出目录，或 --list-paragraphs 源文件.docx")
    except (OSError, ValueError, KeyError, TypeError, AttributeError) as exc:
        sys.exit(f"导出失败：{exc}")


if __name__ == "__main__":
    main()
