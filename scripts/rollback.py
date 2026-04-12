#!/usr/bin/env python3
"""
Rollback - 根据manifest恢复文档

F-13: 回滚功能：根据manifest恢复指定版本

Usage:
    python scripts/rollback.py --list                        # 列出所有备份
    python scripts/rollback.py --latest                      # 恢复最新批次
    python scripts/rollback.py --batch <batch_id>            # 恢复指定批次
    python scripts/rollback.py --id <uuid>                   # 恢复指定备份
    python scripts/rollback.py --restore <backup_path> <target_path>  # 直接指定路径恢复
"""

import argparse
import hashlib
import json
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))


def sha256_file(path: str) -> str:
    """计算文件的SHA-256哈希值"""
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(8192), b''):
            h.update(chunk)
    return h.hexdigest()


def load_manifest(backup_dir: str = "backup") -> dict:
    """加载manifest文件"""
    manifest_path = Path(backup_dir) / "batch_manifest.json"
    if not manifest_path.exists():
        raise FileNotFoundError(f"Manifest不存在: {manifest_path}")
    with open(manifest_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def list_backups(manifest: dict):
    """列出所有备份"""
    backups = manifest.get("backups", [])
    if not backups:
        print("暂无备份记录")
        return

    print(f"\n{'ID':<10} {'批次':<10} {'时间':<22} {'原文件':<40} {'SHA256':<14}")
    print("-" * 100)
    for b in backups:
        orig = Path(b['original_path']).name
        print(f"{b['id']:<10} {b.get('batch_id', 'N/A'):<10} {b['backup_time']:<22} {orig:<40} {b['sha256'][:12]:<14}")


def verify_backup(backup_entry: dict, backup_dir: str = "backup") -> bool:
    """验证备份文件完整性"""
    backup_path = Path(backup_dir) / backup_entry['backup_filename']
    if not backup_path.exists():
        print(f"   ❌ 备份文件不存在: {backup_path}")
        return False
    
    current_hash = sha256_file(str(backup_path))
    expected_hash = backup_entry['sha256']
    
    if current_hash != expected_hash:
        print(f"   ❌ SHA-256校验失败!")
        print(f"   期望: {expected_hash}")
        print(f"   实际: {current_hash}")
        return False
    
    print(f"   ✅ SHA-256校验通过")
    return True


def rollback_entry(backup_entry: dict, backup_dir: str = "backup") -> bool:
    """恢复单个备份"""
    backup_filename = backup_entry['backup_filename']
    original_path = backup_entry['original_path']
    
    backup_path = Path(backup_dir) / backup_filename
    
    print(f"\n🔄 恢复: {original_path}")
    print(f"   来自: {backup_path}")
    
    # 完整性校验
    if not verify_backup(backup_entry, backup_dir):
        print("   ❌ 完整性校验失败，拒绝恢复")
        return False
    
    # 确保目标目录存在
    target_dir = Path(original_path).parent
    target_dir.mkdir(parents=True, exist_ok=True)
    
    # 恢复前先备份当前文件（如果存在且与目标不同）
    if os.path.exists(original_path):
        current_hash = sha256_file(original_path)
        if current_hash != backup_entry.get('original_sha256'):
            # 当前文件已修改，先保存当前状态到backup目录
            corrupt_name = f".corrupt_{Path(original_path).name}_{backup_entry['id']}"
            corrupt_backup = Path(backup_dir) / corrupt_name
            shutil.copy2(original_path, str(corrupt_backup))
            print(f"   ⚠️  当前文件已被修改，已备份到: {corrupt_backup}")
    
    # 执行恢复
    shutil.copy2(str(backup_path), original_path)
    
    # 再次校验
    restored_hash = sha256_file(original_path)
    if restored_hash != backup_entry['sha256']:
        print("   ❌ 恢复后校验失败!")
        return False
    
    print(f"   ✅ 恢复成功: {original_path}")
    return True


def rollback_batch(batch_id: str, backup_dir: str = "backup") -> dict:
    """恢复指定批次的所有备份"""
    manifest = load_manifest(backup_dir)
    
    entries = [b for b in manifest.get("backups", []) if b.get("batch_id") == batch_id]
    if not entries:
        raise ValueError(f"未找到批次: {batch_id}")
    
    print(f"\n📦 开始恢复批次 {batch_id}，共 {len(entries)} 个文件")
    
    results = {}
    for entry in entries:
        try:
            results[entry['id']] = rollback_entry(entry, backup_dir)
        except Exception as e:
            print(f"   ❌ 恢复失败: {e}")
            results[entry['id']] = False
    
    success = sum(1 for v in results.values() if v)
    print(f"\n✅ 批次 {batch_id} 恢复完成: {success}/{len(entries)} 成功")
    return results


def rollback_latest(backup_dir: str = "backup") -> dict:
    """恢复最新批次"""
    manifest = load_manifest(backup_dir)
    batch_id = manifest.get("last_batch_id")
    if not batch_id:
        raise ValueError("没有可用的备份批次")
    return rollback_batch(batch_id, backup_dir)


def rollback_by_id(backup_id: str, backup_dir: str = "backup") -> bool:
    """恢复指定ID的单个备份"""
    manifest = load_manifest(backup_dir)
    entry = next((b for b in manifest.get("backups", []) if b['id'] == backup_id), None)
    if not entry:
        raise ValueError(f"未找到备份: {backup_id}")
    return rollback_entry(entry, backup_dir)


def rollback_direct(backup_path: str, target_path: str) -> bool:
    """直接从备份路径恢复到目标路径"""
    if not os.path.exists(backup_path):
        raise FileNotFoundError(f"备份文件不存在: {backup_path}")
    
    print(f"\n🔄 直接恢复")
    print(f"   备份: {backup_path}")
    print(f"   目标: {target_path}")
    
    # SHA-256校验
    h = hashlib.sha256()
    with open(backup_path, 'rb') as f:
        for chunk in iter(lambda: f.read(8192), b''):
            h.update(chunk)
    print(f"   SHA-256: {h.hexdigest()}")
    
    target_dir = Path(target_path).parent
    target_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(backup_path, target_path)
    
    print(f"   ✅ 恢复成功: {target_path}")
    return True


def main():
    parser = argparse.ArgumentParser(description="文档回滚工具 - 根据manifest恢复文档")
    parser.add_argument("--list", "-l", action='store_true', help="列出所有备份")
    parser.add_argument("--latest", action='store_true', help="恢复最新批次")
    parser.add_argument("--batch", "-b", type=str, help="恢复指定批次")
    parser.add_argument("--id", "-i", type=str, help="恢复指定ID的单个备份")
    parser.add_argument("--restore", nargs=2, metavar=("backup_path", "target_path"), help="直接指定路径恢复")
    parser.add_argument("--backup-dir", "-d", default="backup", help="备份目录 (默认: backup/)")
    args = parser.parse_args()

    if args.list:
        manifest = load_manifest(args.backup_dir)
        list_backups(manifest)
        return

    if args.latest:
        rollback_latest(args.backup_dir)
        return

    if args.batch:
        rollback_batch(args.batch, args.backup_dir)
        return

    if args.id:
        rollback_by_id(args.id, args.backup_dir)
        return

    if args.restore:
        backup_path, target_path = args.restore
        rollback_direct(backup_path, target_path)
        return

    parser.print_help()


if __name__ == "__main__":
    main()
