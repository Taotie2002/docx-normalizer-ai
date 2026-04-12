#!/usr/bin/env python3
"""
DFGP Extract - 从模板文档提取格式规范

用法:
    python scripts/extract.py <模板文档> [--output <输出文件>]

示例:
    python scripts/extract.py templates/government/gov-001-template.docx
    python scripts/extract.py templates/government/gov-001-template.docx -o format.json
"""

import argparse
import json
import sys
from pathlib import Path

# 添加父目录到路径以便导入dfgp_tool
sys.path.insert(0, str(Path(__file__).parent.parent))


def extract_format(doc_path: str) -> dict:
    """从文档提取格式信息"""
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    
    doc = Document(doc_path)
    
    result = {
        "source": doc_path,
        "sections": [],
        "paragraphs": []
    }
    
    # 提取页面设置
    section = doc.sections[0]
    result["page_setup"] = {
        "top_margin_cm": section.top_margin.cm,
        "bottom_margin_cm": section.bottom_margin.cm,
        "left_margin_cm": section.left_margin.cm,
        "right_margin_cm": section.right_margin.cm,
    }
    
    # 提取段落格式
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        
        p_info = {
            "index": i,
            "text": text[:50] + "..." if len(text) > 50 else text,
            "alignment": str(para.alignment),
        }
        
        # 提取字体信息
        if para.runs:
            run = para.runs[0]
            p_info["font"] = {
                "name": run.font.name,
                "size_pt": run.font.size.pt if run.font.size else None,
            }
        
        # 提取缩进
        fi = para.paragraph_format.first_line_indent
        if fi:
            p_info["first_line_indent_pt"] = fi.pt
        
        result["paragraphs"].append(p_info)
    
    return result


def main():
    parser = argparse.ArgumentParser(description="从模板文档提取DFGP格式")
    parser.add_argument("input", help="输入模板文档路径")
    parser.add_argument("-o", "--output", help="输出JSON文件路径")
    
    args = parser.parse_args()
    
    try:
        result = extract_format(args.input)
        
        if args.output:
            with open(args.output, 'w', encoding='utf-8') as f:
                json.dump(result, f, ensure_ascii=False, indent=2)
            print(f"✅ 格式已提取到: {args.output}")
        else:
            print(json.dumps(result, ensure_ascii=False, indent=2))
            
    except Exception as e:
        print(f"❌ 错误: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
