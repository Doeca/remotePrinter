#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
清理数据库中对应文件不存在的记录
检查records表中的记录，如果对应的文件不在files_record.txt中，则删除该记录
"""

import sqlite3
import sys
import os
from pathlib import Path

# 指定数据库路径
DB_PATH = "/Users/doeca/Documents/remotePrinter/logs/remotePrinter.db"
FILES_RECORD_PATH = "/Users/doeca/Documents/remotePrinter/debugFiles/files_record.txt"

# 设置要检查的最新记录数量
RECORDS_LIMIT = 1000


def load_existing_files(files_record_path):
    """从files_record.txt加载所有文件内容（作为一个大字符串）"""
    if not os.path.exists(files_record_path):
        print(f"错误: 文件 {files_record_path} 不存在")
        return None
    
    with open(files_record_path, 'r', encoding='utf-8') as f:
        # 读取整个文件内容
        content = f.read()
    
    print(f"✓ 从 files_record.txt 加载了文件内容")
    return content


def extract_instance_id_from_record_key(record_key):
    """
    从record_key提取实例ID
    
    record_key格式可能有多种:
    1. {instanceID}_COMPLETED
    2. 直接是instanceID
    3. {instanceID}_{status}
    
    提取规则：如果包含下划线，取下划线之前的部分；否则直接返回
    """
    # 如果包含_COMPLETED或其他状态后缀，去除它
    if '_' in record_key:
        # 提取实例ID部分（下划线之前的部分）
        instance_id = record_key.rsplit('_', 1)[0]
    else:
        instance_id = record_key
    
    return instance_id


def clean_missing_records(db_path, files_content, limit=2000, dry_run=True):
    """
    清理不存在的文件对应的记录
    
    Args:
        db_path: 数据库路径
        files_content: files_record.txt的文件内容（字符串）
        limit: 检查最新的多少条记录
        dry_run: 是否仅模拟运行（不实际删除）
    """
    if not os.path.exists(db_path):
        print(f"错误: 数据库文件 {db_path} 不存在")
        return
    
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # 获取最新的N条记录
        cursor.execute('''
            SELECT id, record_key, created_at 
            FROM records 
            ORDER BY created_at DESC 
            LIMIT ?
        ''', (limit,))
        
        records = cursor.fetchall()
        print(f"\n检查最新的 {len(records)} 条记录...")
        print("=" * 80)
        
        to_delete = []
        existing_count = 0
        
        for record_id, record_key, created_at in records:
            # 提取实例ID
            instance_id = extract_instance_id_from_record_key(record_key)
            
            # 检查实例ID是否出现在files_record.txt中
            if instance_id not in files_content:
                to_delete.append((record_id, record_key, instance_id, created_at))
            else:
                existing_count += 1
        
        print(f"\n统计结果:")
        print(f"  文件存在的记录: {existing_count}")
        print(f"  文件缺失的记录: {len(to_delete)}")
        
        if to_delete:
            print(f"\n{'模拟' if dry_run else '准备'}删除以下记录:")
            print("-" * 80)
            for idx, (rec_id, rec_key, instance_id, created_at) in enumerate(to_delete[:10], 1):
                print(f"  {idx}. ID={rec_id}, Key={rec_key}")
                print(f"     实例ID={instance_id}, 创建时间={created_at}")
            
            if len(to_delete) > 10:
                print(f"  ... 还有 {len(to_delete) - 10} 条记录")
            
            if not dry_run:
                # 执行删除
                ids_to_delete = [rec[0] for rec in to_delete]
                placeholders = ','.join('?' * len(ids_to_delete))
                cursor.execute(f'DELETE FROM records WHERE id IN ({placeholders})', ids_to_delete)
                conn.commit()
                print(f"\n✓ 已删除 {cursor.rowcount} 条记录")
            else:
                print(f"\n⚠ 模拟运行模式 - 未实际删除记录")
                print(f"  如需实际删除，请设置 dry_run=False")
        else:
            print("\n✓ 所有记录对应的文件都存在，无需清理")
        
        conn.close()
        
        return {
            'total_checked': len(records),
            'existing': existing_count,
            'to_delete': len(to_delete),
            'deleted': 0 if dry_run else len(to_delete)
        }
        
    except Exception as e:
        print(f"错误: {e}")
        import traceback
        traceback.print_exc()
        return None


def main():
    """主函数"""
    print("=" * 80)
    print("数据库记录清理工具")
    print("=" * 80)
    print(f"数据库路径: {DB_PATH}")
    print(f"文件列表: {FILES_RECORD_PATH}")
    print(f"检查数量: 最新 {RECORDS_LIMIT} 条记录")
    print("=" * 80)
    
    # 加载存在的文件列表
    existing_files = load_existing_files(FILES_RECORD_PATH)
    if existing_files is None:
        return 1
    
    # 先进行模拟运行
    print("\n【模拟运行】")
    result = clean_missing_records(DB_PATH, existing_files, RECORDS_LIMIT, dry_run=True)
    
    if result and result['to_delete'] > 0:
        print("\n" + "=" * 80)
        choice = input("\n是否确认删除这些记录？(yes/no): ").strip().lower()
        
        if choice in ['yes', 'y']:
            print("\n【正式执行】")
            result = clean_missing_records(DB_PATH, existing_files, RECORDS_LIMIT, dry_run=False)
            print("\n✓ 清理完成")
        else:
            print("\n✗ 取消操作")
    
    return 0


if __name__ == '__main__':
    sys.exit(main())
