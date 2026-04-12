#!/usr/bin/env python3
"""
DFGP Apply - 应用格式到目标文档

用法:
    python scripts/apply.py <目标文档> <模板> <输出文件>

示例:
    python scripts/apply.py input.docx templates/government/gov-001-template.docx output.docx
"""

import argparse
import sys
from pathlib import Path

# 添加父目录到路径以便导入dfgp_tool
sys.path.insert(0, str(Path(__file__).parent.parent))


def apply_format(input_path: str, template_path: str, output_path: str):
    """应用模板格式到目标文档"""
    # 导入灌注工具
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    
    doc = Document(input_path)
    template = Document(template_path)
    
    # 获取模板的页面设置
    tpl_section = template.sections[0]
    
    # 应用到目标文档
    section = doc.sections[0]
    section.top_margin = tpl_section.top_margin
    section.bottom_margin = tpl_section.bottom_margin
    section.left_margin = tpl_section.left_margin
    section.right_margin = tpl_section.right_margin
    
    # TODO: 实现完整的格式应用逻辑
    # 目前仅复制页面设置
    
    doc.save(output_path)
    print(f"✅ 格式已应用: {output_path}")


def main():
    parser = argparse.ArgumentParser(description="应用DFGP格式到目标文档")
    parser.add_argument("input", help="输入目标文档路径")
    parser.add_argument("template", help="模板文档路径")
    parser.add_argument("output", help="输出文档路径")
    
    args = parser.parse_args()
    
    try:
        apply_format(args.input, args.template, args.output)
    except Exception as e:
        print(f"❌ 错误: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
