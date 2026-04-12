#!/usr/bin/env python3
"""
Extract Format Schema from DOCX - Extract styles and page setup to YAML

用法:
    python scripts/extract_schema.py <模板文档> [--output <输出文件>]

示例:
    python scripts/extract_schema.py templates/government/gov-001-template.docx
    python scripts/extract_schema.py templates/government/gov-001-template.docx -o format.yaml
"""

import argparse
import sys
import zipfile
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import List, Optional

import yaml
from lxml import etree


# Word ML namespace
NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NSMAP = {'w': NS}


def qn(tag):
    """Qualified name with namespace."""
    return f'{{{NS}}}{tag}'


@dataclass
class FontDef:
    """Font definition for a style."""
    ascii: Optional[str] = None
    east_asia: Optional[str] = None
    size_pt: Optional[float] = None
    bold: bool = False
    italic: bool = False
    color: Optional[str] = None

    def to_dict(self):
        result = {}
        if self.ascii:
            result['ascii'] = self.ascii
        if self.east_asia:
            result['east_asia'] = self.east_asia
        if self.size_pt:
            result['size_pt'] = self.size_pt
        if self.bold:
            result['bold'] = True
        if self.italic:
            result['italic'] = True
        if self.color:
            result['color'] = self.color
        return result


@dataclass
class SpacingDef:
    """Line spacing definition."""
    before_pt: Optional[float] = None  # before paragraph spacing
    after_pt: Optional[float] = None    # after paragraph spacing
    line_pt: Optional[float] = None     # line spacing
    line_rule: Optional[str] = None     # auto, exact, or atLeast (None if not set)

    def to_dict(self):
        result = {}
        if self.before_pt is not None:
            result['before_pt'] = self.before_pt
        if self.after_pt is not None:
            result['after_pt'] = self.after_pt
        if self.line_pt is not None:
            result['line_pt'] = self.line_pt
        return result


@dataclass
class IndentDef:
    """Indentation definition."""
    left_pt: Optional[float] = None
    right_pt: Optional[float] = None
    first_line_pt: Optional[float] = None

    def to_dict(self):
        result = {}
        if self.left_pt is not None:
            result['left_pt'] = self.left_pt
        if self.right_pt is not None:
            result['right_pt'] = self.right_pt
        if self.first_line_pt is not None:
            result['first_line_pt'] = self.first_line_pt
        return result


@dataclass
class StyleDef:
    """Style definition extracted from styles.xml."""
    style_id: str
    name: str
    font: Optional[FontDef] = None
    alignment: Optional[str] = None  # left, center, right, both (justified)
    spacing: Optional[SpacingDef] = None
    indent: Optional[IndentDef] = None
    based_on: Optional[str] = None

    def to_dict(self):
        result = {
            'name': self.name,
        }
        if self.based_on:
            result['based_on'] = self.based_on
        if self.font:
            font_dict = self.font.to_dict()
            if font_dict:
                result['font'] = font_dict
        if self.alignment:
            result['alignment'] = self.alignment
        if self.spacing:
            result['spacing'] = self.spacing.to_dict()
        if self.indent:
            indent_dict = self.indent.to_dict()
            if indent_dict:
                result['indent'] = indent_dict
        return result


@dataclass
class PageSetup:
    """Page setup extracted from sectPr."""
    page_width_pt: float
    page_height_pt: float
    top_margin_pt: float
    bottom_margin_pt: float
    left_margin_pt: float
    right_margin_pt: float
    header_pt: float = 9.0
    footer_pt: float = 9.0

    def to_dict(self):
        return {
            'page_width_pt': self.page_width_pt,
            'page_height_pt': self.page_height_pt,
            'top_margin_pt': self.top_margin_pt,
            'bottom_margin_pt': self.bottom_margin_pt,
            'left_margin_pt': self.left_margin_pt,
            'right_margin_pt': self.right_margin_pt,
            'header_pt': self.header_pt,
            'footer_pt': self.footer_pt,
        }


@dataclass
class FormatSchema:
    """Complete format schema for a document."""
    source: str
    styles: List[StyleDef]
    page_setup: PageSetup

    def to_dict(self):
        return {
            'source': self.source,
            'styles': [s.to_dict() for s in self.styles],
            'page_setup': self.page_setup.to_dict(),
        }


def dxa_to_pt(dxa: int) -> float:
    """Convert twentieths of a point (dxa) to points (pt)."""
    return dxa / 20.0


def half_pts_to_pt(hps: int) -> float:
    """Convert half-points (hps) to points. Font size uses half-points."""
    return hps / 2.0


def twips_to_pt(twips: int) -> float:
    """Convert twips to points. 1 twip = 1/20 point."""
    return twips / 20.0


def extract_font(rPr: etree._Element) -> Optional[FontDef]:
    """Extract font information from run properties."""
    if rPr is None:
        return None

    font = FontDef()

    # Extract rFonts
    rFonts = rPr.find(qn('rFonts'))
    if rFonts is not None:
        font.ascii = rFonts.get(qn('ascii')) or rFonts.get(qn('asciiTheme'))
        font.east_asia = rFonts.get(qn('eastAsia')) or rFonts.get(qn('eastAsiaTheme'))

    # Extract size (sz is in half-points, not dxa)
    sz = rPr.find(qn('sz'))
    if sz is not None:
        val = sz.get(qn('val'))
        if val:
            font.size_pt = half_pts_to_pt(int(val))

    # Extract bold
    if rPr.find(qn('b')) is not None:
        font.bold = True

    # Extract italic
    if rPr.find(qn('i')) is not None:
        font.italic = True

    # Extract color
    color = rPr.find(qn('color'))
    if color is not None:
        font.color = color.get(qn('val'))

    # Check if any font properties exist
    if font.ascii or font.east_asia or font.size_pt or font.bold or font.italic or font.color:
        return font
    return None


def extract_spacing(pPr: etree._Element) -> Optional[SpacingDef]:
    """Extract spacing information from paragraph properties."""
    if pPr is None:
        return None

    spacing = SpacingDef()

    spacing_elem = pPr.find(qn('spacing'))
    if spacing_elem is not None:
        before = spacing_elem.get(qn('before'))
        after = spacing_elem.get(qn('after'))
        line = spacing_elem.get(qn('line'))
        line_rule_val = spacing_elem.get(qn('lineRule'))

        if before:
            spacing.before_pt = dxa_to_pt(int(before))
        if after:
            spacing.after_pt = dxa_to_pt(int(after))
        if line:
            spacing.line_pt = dxa_to_pt(int(line))
        if line_rule_val:
            spacing.line_rule = line_rule_val

    # Check if any spacing properties exist
    if spacing.before_pt is not None or spacing.after_pt is not None or spacing.line_pt is not None:
        return spacing
    return None


def extract_indent(pPr: etree._Element) -> Optional[IndentDef]:
    """Extract indentation information from paragraph properties."""
    if pPr is None:
        return None

    indent = IndentDef()

    ind = pPr.find(qn('ind'))
    if ind is not None:
        left = ind.get(qn('left'))
        right = ind.get(qn('right'))
        firstLine = ind.get(qn('firstLine'))

        if left:
            indent.left_pt = dxa_to_pt(int(left))
        if right:
            indent.right_pt = dxa_to_pt(int(right))
        if firstLine:
            indent.first_line_pt = dxa_to_pt(int(firstLine))

    # Check if any indent properties exist
    if indent.left_pt is not None or indent.right_pt is not None or indent.first_line_pt is not None:
        return indent
    return None


def extract_alignment(pPr: etree._Element) -> Optional[str]:
    """Extract alignment from paragraph properties."""
    if pPr is None:
        return None

    jc = pPr.find(qn('jc'))
    if jc is not None:
        return jc.get(qn('val'))
    return None


def extract_styles(styles_xml: bytes) -> List[StyleDef]:
    """Extract paragraph styles from styles.xml content."""
    root = etree.fromstring(styles_xml)
    styles = []

    # Find all paragraph styles
    for style_elem in root.findall(f'.//{qn("style")}'):
        style_type = style_elem.get(qn('type'))
        if style_type != 'paragraph':
            continue

        style_id = style_elem.get(qn('styleId'))
        if not style_id:
            continue

        # Get style name
        name_elem = style_elem.find(qn('name'))
        name = name_elem.get(qn('val')) if name_elem is not None else style_id

        # Get basedOn
        based_on_elem = style_elem.find(qn('basedOn'))
        based_on = based_on_elem.get(qn('val')) if based_on_elem is not None else None

        # Get paragraph properties
        pPr = style_elem.find(qn('pPr'))

        # Get run properties (for font info)
        rPr = style_elem.find(qn('rPr'))

        style_def = StyleDef(
            style_id=style_id,
            name=name,
            font=extract_font(rPr),
            alignment=extract_alignment(pPr),
            spacing=extract_spacing(pPr),
            indent=extract_indent(pPr),
            based_on=based_on,
        )

        styles.append(style_def)

    return styles


def extract_page_setup(document_xml: bytes) -> PageSetup:
    """Extract page setup from document.xml sectPr."""
    root = etree.fromstring(document_xml)

    # Find sectPr
    sectPr = root.find(f'.//{qn("sectPr")}')
    if sectPr is None:
        # Default page setup if no sectPr found
        return PageSetup(
            page_width_pt=612.0,  # Letter width in points
            page_height_pt=792.0,  # Letter height in points
            top_margin_pt=72.0,
            bottom_margin_pt=72.0,
            left_margin_pt=72.0,
            right_margin_pt=72.0,
        )

    # Get page size
    pgSz = sectPr.find(qn('pgSz'))
    if pgSz is not None:
        page_width = int(pgSz.get(qn('w'), 12240))  # Default to letter width in dxa
        page_height = int(pgSz.get(qn('h'), 15840))  # Default to letter height in dxa
    else:
        page_width = 12240
        page_height = 15840

    # Get page margins
    pgMar = sectPr.find(qn('pgMar'))
    if pgMar is not None:
        top_margin = int(pgMar.get(qn('top'), 1440))
        bottom_margin = int(pgMar.get(qn('bottom'), 1440))
        left_margin = int(pgMar.get(qn('left'), 1440))
        right_margin = int(pgMar.get(qn('right'), 1440))
        header = int(pgMar.get(qn('header'), 720))
        footer = int(pgMar.get(qn('footer'), 720))
    else:
        top_margin = bottom_margin = left_margin = right_margin = 1440
        header = footer = 720

    return PageSetup(
        page_width_pt=dxa_to_pt(page_width),
        page_height_pt=dxa_to_pt(page_height),
        top_margin_pt=dxa_to_pt(top_margin),
        bottom_margin_pt=dxa_to_pt(bottom_margin),
        left_margin_pt=dxa_to_pt(left_margin),
        right_margin_pt=dxa_to_pt(right_margin),
        header_pt=twips_to_pt(header),
        footer_pt=twips_to_pt(footer),
    )


def extract_schema(doc_path: str) -> FormatSchema:
    """Extract format schema from a DOCX file."""
    source = doc_path

    with zipfile.ZipFile(doc_path, 'r') as zf:
        # Read styles.xml
        styles_xml = zf.read('word/styles.xml')

        # Read document.xml for sectPr
        document_xml = zf.read('word/document.xml')

    # Extract styles and page setup
    styles = extract_styles(styles_xml)
    page_setup = extract_page_setup(document_xml)

    return FormatSchema(
        source=source,
        styles=styles,
        page_setup=page_setup,
    )


def format_yaml(schema: FormatSchema) -> str:
    """Format schema as human-readable YAML."""
    data = schema.to_dict()

    # Custom representer for better formatting
    def str_representer(dumper, data):
        if '\n' in data:
            return dumper.represent_scalar('tag:yaml.org,2002:str', data, style='|')
        return dumper.represent_scalar('tag:yaml.org,2002:str', data)

    yaml.add_representer(str, str_representer)

    return yaml.dump(
        data,
        default_flow_style=False,
        allow_unicode=True,
        sort_keys=False,
        width=120,
        indent=2,
    )


def main():
    parser = argparse.ArgumentParser(
        description="从DOCX模板提取格式Schema (样式和页面设置)"
    )
    parser.add_argument(
        'input',
        help='输入DOCX模板文档路径'
    )
    parser.add_argument(
        '-o', '--output',
        help='输出YAML文件路径 (默认输出到stdout)'
    )

    args = parser.parse_args()

    try:
        # Extract schema
        schema = extract_schema(args.input)

        # Format as YAML
        yaml_output = format_yaml(schema)

        if args.output:
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write(yaml_output)
            print(f"✅ Schema已提取到: {args.output}")
        else:
            print(yaml_output)

    except Exception as e:
        print(f"❌ 错误: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
