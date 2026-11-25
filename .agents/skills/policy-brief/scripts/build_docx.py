"""Build paired editable briefs from reviewed prose and a compatible DOCX template.

Python 3.11+ standard library only. This is a layout helper, not a fact checker.
"""
from __future__ import annotations

import argparse
import io
import json
import re
import sys
from hashlib import sha256
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, BadZipFile, ZipFile

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
ET.register_namespace("w", W)
STYLE_IDS = {"Title", "Normal", "Heading1", "Heading2", "Heading3", "Heading4"}
DEFAULT_TEMPLATE = Path(__file__).resolve().parent.parent / "assets" / "report-template.docx"
MARKER = re.compile(r"\[E\d{3}(?:[、,，;；\s]+E\d{3})*\]|〔E\d{3}(?:[、,，;；\s]+E\d{3})*〕")
BODY = re.compile(rb"(<w:body(?:\s[^>]*)?>)(.*?)(</w:body>)", re.S)
SECTION = re.compile(rb"(<w:sectPr(?:\s[^>]*)?(?:/>|>.*?</w:sectPr>))\s*$", re.S)


def digest(data: bytes) -> str:
    return sha256(data).hexdigest()


def check_display_fields(root: ET.Element) -> None:
    """Only page-count fields are allowed before any Word Fields.Update call."""
    allowed = {"PAGE", "NUMPAGES", "SECTIONPAGES"}
    stack: list[list[str]] = []

    def verify(instruction: str) -> None:
        command = instruction.strip().split(maxsplit=1)
        if not command or command[0].upper() not in allowed:
            raise ValueError("模板含非页码类字段；请在新副本中核实并专门处理，不能自动更新")

    for node in root.iter():
        if node.tag == f"{{{W}}}fldSimple":
            verify(node.get(f"{{{W}}}instr", ""))
        elif node.tag == f"{{{W}}}fldChar":
            kind = node.get(f"{{{W}}}fldCharType")
            if kind == "begin":
                stack.append([])
            elif kind == "end":
                if not stack:
                    raise ValueError("模板字段边界不完整")
                verify("".join(stack.pop()))
        elif node.tag == f"{{{W}}}instrText":
            if not stack:
                raise ValueError("模板字段指令缺少边界")
            stack[-1].append(node.text or "")
    if stack:
        raise ValueError("模板字段边界不完整")


def inspect_template(template: Path) -> tuple[dict[str, bytes], bytes]:
    raw = template.read_bytes()
    with ZipFile(io.BytesIO(raw)) as archive:
        names = archive.namelist()
        if len(names) != len(set(names)):
            raise ValueError("模板 ZIP 存在重名成员")
        if len(names) > 1000 or sum(info.file_size for info in archive.infolist()) > 100_000_000:
            raise ValueError("模板包过大；请使用经过整理的文字模板")
        for info in archive.infolist():
            path = PurePosixPath(info.filename)
            if path.is_absolute() or ".." in path.parts or "\\" in info.filename:
                raise ValueError("模板 ZIP 路径不安全")
            if info.file_size > 20_000_000:
                raise ValueError("模板成员过大；请使用经过整理的文字模板")
        parts = {name: archive.read(name) for name in names}
    for required in ("[Content_Types].xml", "word/document.xml", "word/styles.xml"):
        if required not in parts:
            raise ValueError("模板缺少 " + required)
    forbidden = ("vbaproject", "comments", "embeddings/", "activex/")
    for name, data in parts.items():
        if any(token in name.lower() for token in forbidden):
            raise ValueError("模板含宏、批注或嵌入对象；请先在新副本中处理")
        if name.endswith((".xml", ".rels")):
            if b"<!DOCTYPE" in data or b"<!ENTITY" in data:
                raise ValueError("模板含不支持的 XML 声明")
            parsed = ET.fromstring(data)
            if name.startswith("word/"):
                check_display_fields(parsed)
            if name.endswith(".rels") and any(x.get("TargetMode") == "External" for x in parsed):
                raise ValueError("模板含外部关系；请先核实并在新副本中处理")
            if name.startswith("word/") and any(
                node.tag in {f"{{{W}}}{x}" for x in ("ins", "del", "vanish", "commentReference")}
                for node in parsed.iter()
            ):
                raise ValueError("模板含修订、隐藏文字或批注；不能静默选择内容状态")
    document = ET.fromstring(parts["word/document.xml"])
    body = document.find("w:body", NS)
    if body is None or len(document.findall(".//w:sectPr", NS)) != 1:
        raise ValueError("本工具仅支持单节文字模板；复杂模板需专门处理")
    if not list(body) or list(body)[-1].tag != f"{{{W}}}sectPr":
        raise ValueError("模板缺少正文末尾的节属性")
    styles = ET.fromstring(parts["word/styles.xml"])
    found = {node.get(f"{{{W}}}styleId") for node in styles.findall("w:style", NS)}
    if not STYLE_IDS <= found:
        raise ValueError("模板须包含 Title、Normal、Heading1 至 Heading4 样式")
    matches = list(BODY.finditer(parts["word/document.xml"]))
    if len(matches) != 1 or SECTION.search(matches[0].group(2)) is None:
        raise ValueError("模板正文 XML 不属于支持的 w:body 格式；请专门适配")
    return parts, raw


def parse_source(raw: bytes) -> list[tuple[str, str]]:
    text = raw.decode("utf-8-sig")
    rows: list[tuple[str, str]] = []
    for line_number, original in enumerate(text.splitlines(), 1):
        if any(ord(char) < 32 and char != "\t" for char in original) or "\t" in original:
            raise ValueError(f"第 {line_number} 行含制表或控制字符")
        line = original.strip()
        if not line:
            continue
        if re.match(r"^(?:---+|\*\*\*+|___+)\s*$", line):
            raise ValueError("输入须仅含已审核的标题与正文；分隔线后的工作记录请另存，不会自动截断")
        heading = re.match(r"^(#{1,5})\s+(.+)$", line)
        if heading:
            level = len(heading.group(1))
            style = "Title" if level == 1 else "Heading" + str(level - 1)
            content = heading.group(2)
        else:
            style, content = "Normal", line
        if re.search(r"(^#{1,}|[`|<>]|!\[|\]\(|^[-+*]\s|^>\s)", content) or (style == "Normal" and re.match(r"^\d+[.)]\s", content)):
            raise ValueError(f"第 {line_number} 行含不支持的 Markdown；请转为经核对的普通段落")
        # Paired ** markers are editorial emphasis; the paragraph style controls output.
        content = re.sub(r"\*\*([^*]+)\*\*", r"\1", content)
        if "*" in content or re.search(r"https?://|_{1,2}[^_]+_{1,2}|~~", content):
            raise ValueError(f"第 {line_number} 行含未配对强调或网址；来源链接保留在证据台账")
        without_markers = MARKER.sub("", content)
        if re.search(r"E\d+", without_markers) or "E000" in content:
            raise ValueError(f"第 {line_number} 行证据编号格式不支持；使用 [E001] 或 〔E001、E002〕")
        rows.append((style, content))
    if not rows or rows[0][0] != "Title" or sum(style == "Title" for style, _ in rows) != 1:
        raise ValueError("首个非空行须为唯一的 # 主标题")
    if len(rows) < 2 or not any(style == "Normal" for style, _ in rows):
        raise ValueError("没有可排版正文")
    if not any(MARKER.search(text) for _, text in rows):
        raise ValueError("正式报送稿输入须保留证据编号；本工具不会替代证据审核")
    return rows


def replace_body(original: bytes, rows: list[tuple[str, str]]) -> bytes:
    match = BODY.search(original)
    assert match is not None
    section_match = SECTION.search(match.group(2))
    assert section_match is not None
    paragraphs = []
    for style, text in rows:
        paragraph = ET.Element(f"{{{W}}}p")
        properties = ET.SubElement(paragraph, f"{{{W}}}pPr")
        ET.SubElement(properties, f"{{{W}}}pStyle", {f"{{{W}}}val": style})
        run = ET.SubElement(paragraph, f"{{{W}}}r")
        element = ET.SubElement(run, f"{{{W}}}t", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"})
        element.text = text
        serialized = ET.tostring(paragraph, encoding="utf-8")
        # Root namespaces are retained byte-for-byte, including mc:Ignorable prefixes.
        serialized = serialized.replace((' xmlns:w="' + W + '"').encode(), b"", 1)
        paragraphs.append(serialized)
    result = original[:match.start(2)] + b"".join(paragraphs) + section_match.group(1) + original[match.end(2):]
    parsed = ET.fromstring(result)
    actual = ["".join(node.itertext()) for node in parsed.findall("w:body/w:p", NS)]
    if actual != [text for _, text in rows]:
        raise ValueError("生成正文与输入不一致")
    return result


def make_package(parts: dict[str, bytes], rows: list[tuple[str, str]]) -> bytes:
    modified = dict(parts)
    modified["word/document.xml"] = replace_body(parts["word/document.xml"], rows)
    buffer = io.BytesIO()
    with ZipFile(buffer, "w", ZIP_DEFLATED) as archive:
        for name, data in modified.items():
            archive.writestr(name, data)
    data = buffer.getvalue()
    with ZipFile(io.BytesIO(data)) as archive:
        if any(archive.read(name) != value for name, value in parts.items() if name != "word/document.xml"):
            raise ValueError("正文之外的模板部件被改变")
    return data


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", required=True, type=Path, help="UTF-8 已审核正文 Markdown")
    parser.add_argument("--out-dir", required=True, type=Path, help="尚不存在的新输出目录")
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE, help="兼容的单节 DOCX 模板")
    args = parser.parse_args()
    try:
        source, template, target = args.source.resolve(), args.template.resolve(), args.out_dir.resolve()
        if target.exists():
            raise ValueError("输出目录已存在；请改用新的任务版本目录，不能覆盖")
        source_raw = source.read_bytes()
        rows = parse_source(source_raw)
        parts, template_raw = inspect_template(template)
        clean_rows = [(style, MARKER.sub("", text)) for style, text in rows]
        rendered = {
            "04-报送清稿.docx": make_package(parts, clean_rows),
            "04-报送核对稿.docx": make_package(parts, rows),
        }
        if source.read_bytes() != source_raw or template.read_bytes() != template_raw:
            raise ValueError("输入在处理期间变化，停止交付")
        report = {
            "source_name": source.name,
            "source_sha256": digest(source_raw),
            "template_name": template.name,
            "template_sha256": digest(template_raw),
            "paragraphs": len(rows),
            "evidence_ids": sorted(set(re.findall(r"E\d{3}", "\n".join(text for _, text in rows)))),
            "clean_text_difference": "仅移除约定格式的 E 证据标记；保留原数字、措辞及顺序",
            "content_review": "脚本不审核事实、权限或证据充分性；须已有独立审核记录",
            "visual_review": "未完成；须更新字段、渲染并逐页检查",
            "files": {name: {"sha256": digest(data), "other_template_parts_unchanged": True} for name, data in rendered.items()},
        }
        # All parsing and in-memory validation finish before creating any outputs.
        target.mkdir(parents=True, exist_ok=False)
        for name, data in rendered.items():
            with (target / name).open("xb") as handle:
                handle.write(data)
        with (target / "build-check.json").open("x", encoding="utf-8") as handle:
            json.dump(report, handle, ensure_ascii=False, indent=2)
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 0
    except (OSError, ValueError, BadZipFile, ET.ParseError) as error:
        print("未生成可交付稿：" + str(error), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
