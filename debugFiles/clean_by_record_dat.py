#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从record.dat读取记录，检查最新200个数据的实例ID是否在files_record.txt中
如果不存在，则从数据库中删除包含该实例ID的记录
"""

import sqlite3
import sys
import os
import json
from pathlib import Path

# 指定路径
RECORD_DAT_PATH = "/Users/doeca/Documents/remotePrinter/debugFiles/record.dat"
FILES_RECORD_PATH = "/Users/doeca/Documents/remotePrinter/debugFiles/files_record.txt"
DB_PATH = "/Users/doeca/Documents/remotePrinter/logs/remotePrinter.db"

# 检查最新的记录数量
RECORDS_LIMIT = 2000


def load_files_record(files_record_path):
    """从files_record.txt加载所有文件内容（作为一个大字符串）"""
    if not os.path.exists(files_record_path):
        print(f"错误: 文件 {files_record_path} 不存在")
        return None
    
    with open(files_record_path, 'r', encoding='utf-8') as f:
        # 读取整个文件内容
        content = f.read()
    
    print(f"✓ 从 files_record.txt 加载了文件内容")
    return content


def load_record_dat(record_dat_path, limit=200):
    """从record.dat加载最新的N条记录（JSON数组格式）"""
    if not os.path.exists(record_dat_path):
        print(f"错误: 文件 {record_dat_path} 不存在")
        return None
    
    try:
        with open(record_dat_path, 'r', encoding='utf-8') as f:
            # 读取JSON数组
            content = f.read().strip()
            
            # 如果内容为空
            if not content:
                print("警告: record.dat 文件为空")
                return []
            
            # 尝试解析JSON数组
            records = json.loads(content)
            
            if not isinstance(records, list):
                print(f"错误: record.dat 不是一个数组格式，而是 {type(records)}")
                return None
    except json.JSONDecodeError as e:
        print(f"错误: 解析JSON失败: {e}")
        return None
    except Exception as e:
        print(f"错误: 读取文件失败: {e}")
        return None
    
    # 取最新的N条（数组最后的是最新的）
    recent_records = records[-limit:] if len(records) > limit else records
    
    print(f"✓ 从 record.dat 加载了 {len(recent_records)} 条记录（总共{len(records)}条，取最新{limit}条）")
    return recent_records


def extract_instance_id(record):
    """
    从记录中提取实例ID
    
    记录格式可能有：
    1. instance_id_STATUS (如: bsXCPX7FRe2xgBM3zuX2pg03401729225473_COMPLETED)
    2. 直接就是instance_id
    
    提取规则：如果包含下划线，取下划线之前的部分（排除纯数字的状态后缀）
    """
    if '_' in record:
        # 分割并获取实例ID部分（下划线之前）
        parts = record.rsplit('_', 1)
        # parts[0]是instance_id, parts[1]是状态(COMPLETED/RUNNING等)
        return parts[0]
    else:
        return record


def clean_missing_records_from_db(db_path, instance_ids_to_check, files_content, dry_run=True):
    """
    清理数据库中对应文件不存在的记录
    
    Args:
        db_path: 数据库路径
        instance_ids_to_check: 要检查的实例ID列表
        files_content: files_record.txt的文件内容（字符串）
        dry_run: 是否仅模拟运行（不实际删除）
    """
    if not os.path.exists(db_path):
        print(f"错误: 数据库文件 {db_path} 不存在")
        return None
    
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        print(f"\n检查 {len(instance_ids_to_check)} 个实例ID...")
        print("=" * 80)
        
        to_delete = []
        existing_count = 0
        
        for instance_id in instance_ids_to_check:
            # 检查实例ID是否出现在files_record.txt中
            if instance_id not in files_content:
                to_delete.append(instance_id)
            else:
                existing_count += 1
        
        print(f"\n统计结果:")
        print(f"  文件存在的实例: {existing_count}")
        print(f"  文件缺失的实例: {len(to_delete)}")
        
        if to_delete:
            print(f"\n{'模拟' if dry_run else '准备'}删除以下实例ID相关的记录:")
            print("-" * 80)
            
            # 显示前10个要删除的实例ID
            for idx, instance_id in enumerate(to_delete[:10], 1):
                print(f"  {idx}. {instance_id}")
            
            if len(to_delete) > 10:
                print(f"  ... 还有 {len(to_delete) - 10} 个实例ID")
            
            # 先统计会删除多少条记录
            total_to_delete = 0
            for instance_id in to_delete:
                try:
                    cursor.execute(
                        "SELECT COUNT(*) FROM records WHERE record_key = ? OR record_key LIKE ?",
                        (instance_id, f"{instance_id}_%")
                    )
                    result = cursor.fetchone()
                    if result:
                        total_to_delete += result[0]
                except Exception as e:
                    print(f"警告: 统计实例 {instance_id} 时出错: {e}")
                    continue
            
            if dry_run:
                print(f"\n⚠ 模拟运行模式 - 将删除 {total_to_delete} 条数据库记录")
                print(f"  如需实际删除，请设置 dry_run=False")
            else:
                # 执行删除 - 删除records表中包含这些instance_id的记录
                deleted_count = 0
                for instance_id in to_delete:
                    try:
                        # record_key可能是 instance_id_STATUS 格式
                        # 使用LIKE查询删除所有包含该instance_id的记录
                        cursor.execute(
                            "DELETE FROM records WHERE record_key = ? OR record_key LIKE ?",
                            (instance_id, f"{instance_id}_%")
                        )
                        deleted_count += cursor.rowcount
                    except Exception as e:
                        print(f"警告: 删除实例 {instance_id} 时出错: {e}")
                        continue
                
                conn.commit()
                print(f"\n✓ 已删除 {deleted_count} 条记录（预计{total_to_delete}条）")
        else:
            print("\n✓ 所有实例对应的文件都存在，无需清理")
        
        conn.close()
        
        return {
            'total_checked': len(instance_ids_to_check),
            'existing': existing_count,
            'to_delete': len(to_delete)
        }
        
    except Exception as e:
        print(f"错误: {e}")
        import traceback
        traceback.print_exc()
        return None


def main():
    """主函数"""
    print("=" * 80)
    print("数据库记录清理工具 (基于record.dat)")
    print("=" * 80)
    print(f"record.dat路径: {RECORD_DAT_PATH}")
    print(f"files_record.txt路径: {FILES_RECORD_PATH}")
    print(f"数据库路径: {DB_PATH}")
    print(f"检查数量: 最新 {RECORDS_LIMIT} 条记录")
    print("=" * 80)
    
    # 加载files_record.txt内容
    files_content = load_files_record(FILES_RECORD_PATH)
    if files_content is None:
        return 1
    
    # 加载record.dat中的最新记录
    records = load_record_dat(RECORD_DAT_PATH, RECORDS_LIMIT)
    if records is None:
        return 1
    
    # 提取实例ID（去重）
    instance_ids = set()
    for record in records:
        instance_id = extract_instance_id(record)
        instance_ids.add(instance_id)
    
    instance_ids_list = list(instance_ids)
    print(f"\n提取到 {len(instance_ids_list)} 个唯一实例ID")
    
    # 先进行模拟运行
    print("\n【模拟运行】")
    result = clean_missing_records_from_db(DB_PATH, instance_ids_list, files_content, dry_run=False)
    
    if result and result['to_delete'] > 0:
        print("\n" + "=" * 80)
        choice = input("\n是否确认删除这些记录？(yes/no): ").strip().lower()
        
        if choice in ['yes', 'y']:
            print("\n【正式执行】")
            result = clean_missing_records_from_db(DB_PATH, instance_ids_list, files_content, dry_run=False)
            print("\n✓ 清理完成")
        else:
            print("\n✗ 取消操作")
    
    return 0


if __name__ == '__main__':
    sys.exit(main())
