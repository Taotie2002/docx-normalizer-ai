#!/usr/bin/env python3
"""
DFGP Apply - 应用格式到目标文档

用法:
    python scripts/apply.py <目标文档> <模板文档> <输出文件>

示例:
    python scripts/apply.py input.docx templates/government/gov-001-template.docx output.docx
"""

import argparse
import sys
import os
from pathlib import Path

# 添加父目录到路径以便导入dfgp_tool
sys.path.insert(0, str(Path(__file__).parent.parent))

# 导入dfgp_tool的核心函数
from scripts.dfgp_tool import 灌注公文


def main():
    parser = argparse.ArgumentParser(description="应用DFGP格式到目标文档")
    parser.add_argument("input", help="输入目标文档路径")
    parser.add_argument("template", help="模板文档路径（用于页码）")
    parser.add_argument("output", help="输出文档路径")
    
    args = parser.parse_args()
    
    if not os.path.exists(args.input):
        print(f"❌ 错误: 输入文件不存在: {args.input}")
        sys.exit(1)
    
    if not os.path.exists(args.template):
        print(f"❌ 错误: 模板文件不存在: {args.template}")
        sys.exit(1)
    
    try:
        # 调用dfgp_tool的灌注函数
        灌注公文(args.input, args.output, args.template)
        print(f"✅ 格式已应用: {args.output}")
    except Exception as e:
        print(f"❌ 错误: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
