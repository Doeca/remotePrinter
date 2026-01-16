#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从 record.dat 文件重置生产模式数据库的 records 表

使用方法：
    python reset_records_from_dat.py

注意：
    - 此脚本仅操作生产模式数据库（logs/remotePrinter.db）
    - 会完全清空 records 表后重新导入数据
    - 执行前会要求确认
"""

import os
import json
import sqlite3
from logsys import logger


def reset_records_from_dat():
    """从 record.dat 重置生产数据库的 records 表"""
    
    # 强制使用生产数据库路径
    PROD_DB_PATH = os.path.join("logs", "remotePrinter.db")
    RECORD_FILE = "./record.dat"
    
    # 检查 record.dat 文件是否存在
    if not os.path.exists(RECORD_FILE):
        logger.error(f"{RECORD_FILE} 文件不存在")
        print(f"错误：{RECORD_FILE} 文件不存在")
        return False
    
    # 检查生产数据库是否存在
    if not os.path.exists(PROD_DB_PATH):
        logger.error(f"生产数据库 {PROD_DB_PATH} 不存在")
        print(f"错误：生产数据库 {PROD_DB_PATH} 不存在")
        return False
    
    try:
        # 读取 record.dat 文件
        with open(RECORD_FILE, "r", encoding="utf-8") as f:
            content = f.read().strip()
            if not content:
                logger.error(f"{RECORD_FILE} 文件为空")
                print(f"错误：{RECORD_FILE} 文件为空")
                return False
            records = json.loads(content)
        
        if not isinstance(records, list):
            logger.error(f"{RECORD_FILE} 格式错误，应该是一个列表")
            print(f"错误：{RECORD_FILE} 格式错误，应该是一个列表")
            return False
        
        record_count = len(records)
        print(f"\n找到 {record_count} 条记录")
        print(f"目标数据库：{PROD_DB_PATH}")
        
        # 确认操作
        confirm = input("\n警告：此操作将清空生产数据库的 records 表并重新导入数据。\n确认继续？(yes/no): ")
        if confirm.lower() != 'yes':
            print("操作已取消")
            return False
        
        # 连接生产数据库
        conn = sqlite3.connect(PROD_DB_PATH)
        cursor = conn.cursor()
        
        # 开始事务
        cursor.execute('BEGIN TRANSACTION')
        
        try:
            # 清空 records 表
            cursor.execute('DELETE FROM records')
            deleted_count = cursor.rowcount
            logger.info(f"已清空 records 表，删除了 {deleted_count} 条记录")
            print(f"已清空 records 表，删除了 {deleted_count} 条记录")
            
            # 重置自增ID
            cursor.execute('DELETE FROM sqlite_sequence WHERE name="records"')
            
            # 导入新数据
            success_count = 0
            duplicate_count = 0
            
            for record in records:
                try:
                    cursor.execute('INSERT INTO records (record_key) VALUES (?)', (record,))
                    success_count += 1
                except sqlite3.IntegrityError:
                    # 重复的记录
                    duplicate_count += 1
            
            # 提交事务
            conn.commit()
            
            print(f"\n导入完成：")
            print(f"  - 成功导入：{success_count} 条")
            if duplicate_count > 0:
                print(f"  - 跳过重复：{duplicate_count} 条")
            
            logger.info(f"成功从 {RECORD_FILE} 重置 records 表，导入 {success_count} 条记录")
            
            # 验证数据
            cursor.execute('SELECT COUNT(*) FROM records')
            final_count = cursor.fetchone()[0]
            print(f"  - 最终记录数：{final_count} 条")
            
            return True
            
        except Exception as e:
            # 回滚事务
            conn.rollback()
            logger.error(f"导入失败，已回滚: {e}")
            print(f"错误：导入失败，已回滚: {e}")
            return False
        
        finally:
            conn.close()
    
    except json.JSONDecodeError as e:
        logger.error(f"{RECORD_FILE} JSON 解析失败: {e}")
        print(f"错误：{RECORD_FILE} JSON 解析失败: {e}")
        return False
    
    except Exception as e:
        logger.error(f"重置失败: {e}")
        print(f"错误：重置失败: {e}")
        return False


if __name__ == "__main__":
    print("=" * 60)
    print("从 record.dat 重置生产数据库 records 表")
    print("=" * 60)
    
    success = reset_records_from_dat()
    
    if success:
        print("\n✓ 重置成功")
    else:
        print("\n✗ 重置失败")
