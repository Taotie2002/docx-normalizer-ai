#!/usr/bin/env python3
"""
Backup - 文档备份与完整性校验

F-12: SHA-256计算原文件hash
F-13: 创建backup/目录存储备份，生成batch_manifest.json记录备份信息

Usage:
    python scripts/backup.py <input_docx> [backup_dir]
    python scripts/backup.py input.docx                    # 默认 backup/
    python scripts/backup.py input.docx my_backups/       # 自定义目录
"""

import argparse
import hashlib
import json
import os
import shutil
import sys
import uuid
from datetime import datetime
from pathlib import Path

# 添加父目录到路径
sys.path.insert(0, str(Path(__file__).parent.parent))


def sha256_file(path: str) -> str:
    """计算文件的SHA-256哈希值"""
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(8192), b''):
            h.update(chunk)
    return h.hexdigest()


def backup_document(input_path: str, backup_dir: str = "backup") -> dict:
    """
    备份单个文档
    
    Returns:
        manifest entry dict
    """
    input_path = os.path.abspath(input_path)
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"文件不存在: {input_path}")

    backup_path = Path(backup_dir)
    backup_path.mkdir(parents=True, exist_ok=True)

    # 生成备份文件名：原名 + 时间戳 + uuid前8位
    basename = Path(input_path).stem
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    short_uuid = uuid.uuid4().hex[:8]
    backup_filename = f"{basename}_{timestamp}_{short_uuid}.docx"
    backup_file_path = backup_path / backup_filename

    # 复制文件
    shutil.copy2(input_path, backup_file_path)

    # 计算SHA-256
    file_hash = sha256_file(backup_file_path)
    file_size = os.path.getsize(backup_file_path)

    entry = {
        "id": short_uuid,
        "original_path": input_path,
        "backup_path": str(backup_file_path),
        "backup_filename": backup_filename,
        "sha256": file_hash,
        "size_bytes": file_size,
        "backup_time": datetime.now().isoformat(timespec='seconds'),
        "original_sha256": sha256_file(input_path),
    }

    return entry


def batch_backup(input_paths: list, backup_dir: str = "backup") -> dict:
    """
    批量备份多个文档，维护统一的manifest
    
    Returns:
        full manifest dict
    """
    manifest_path = Path(backup_dir) / "batch_manifest.json"
    
    # 加载已有manifest或创建新的
    if manifest_path.exists():
        with open(manifest_path, 'r', encoding='utf-8') as f:
            manifest = json.load(f)
    else:
        manifest = {
            "version": "1.0",
            "created": datetime.now().isoformat(timespec='seconds'),
            "backups": [],
        }

    batch_id = uuid.uuid4().hex[:8]
    batch_time = datetime.now().isoformat(timespec='seconds')
    entries = []

    for path in input_paths:
        try:
            entry = backup_document(path, backup_dir)
            entry["batch_id"] = batch_id
            entries.append(entry)
            manifest["backups"].append(entry)
        except FileNotFoundError as e:
            print(f"⚠️  跳过: {e}")

    manifest["last_batch_id"] = batch_id
    manifest["last_batch_time"] = batch_time
    manifest["total_backups"] = len(manifest["backups"])

    # 保存manifest
    with open(manifest_path, 'w', encoding='utf-8') as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    print(f"📦 批次 {batch_id} 备份完成: {len(entries)} 个文件")
    print(f"   Manifest: {manifest_path}")
    for e in entries:
        print(f"   ✅ {e['original_path']} → {e['backup_filename']} [{e['sha256'][:12]}...]")

    return manifest


def main():
    parser = argparse.ArgumentParser(description="文档备份工具 - SHA-256校验 + manifest记录")
    parser.add_argument("inputs", nargs='+', help="要备份的文档路径")
    parser.add_argument("--backup-dir", "-d", default="backup", help="备份目录 (默认: backup/)")
    parser.add_argument("--manifest-only", action='store_true', help="仅显示manifest，不执行备份")
    args = parser.parse_args()

    if args.manifest_only:
        manifest_path = Path(args.backup_dir) / "batch_manifest.json"
        if manifest_path.exists():
            with open(manifest_path, 'r', encoding='utf-8') as f:
                print(f.read())
        else:
            print("No manifest found.")
        return

    manifest = batch_backup(args.inputs, args.backup_dir)
    print(f"\n✅ 备份完成，共 {manifest['total_backups']} 个备份")


if __name__ == "__main__":
    main()
