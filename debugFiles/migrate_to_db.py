#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
数据迁移脚本
将 record.dat、ids_cache.json 和 task 文件夹的数据迁移到 SQLite 数据库
"""

import os
import json
import db
from logsys import logger


def migrate_records():
    """迁移 record.dat 到数据库"""
    record_file = "./record.dat"
    if not os.path.exists(record_file):
        logger.info("record.dat 文件不存在，跳过迁移")
        return
    
    try:
        with open(record_file, "r") as f:
            records = json.loads(f.read())
        
        count = 0
        for record in records:
            if db.add_record(record):
                count += 1
        
        logger.info(f"成功迁移 {count} 条记录从 record.dat")
        
        # 备份原文件
        backup_file = record_file + ".bak"
        os.rename(record_file, backup_file)
        logger.info(f"已将原文件备份为 {backup_file}")
        
    except Exception as e:
        logger.error(f"迁移 record.dat 失败: {e}")


def migrate_ids_cache():
    """迁移 ids_cache.json 到数据库"""
    cache_file = "./ids_cache.json"
    if not os.path.exists(cache_file):
        logger.info("ids_cache.json 文件不存在，跳过迁移")
        return
    
    try:
        with open(cache_file, "r") as f:
            cache_data = json.loads(f.read())
        
        count = 0
        for process_code, instances in cache_data.items():
            for instance_id, timestamp in instances.items():
                if db.add_ids_cache(process_code, instance_id, timestamp):
                    count += 1
        
        logger.info(f"成功迁移 {count} 条缓存记录从 ids_cache.json")
        
        # 备份原文件
        backup_file = cache_file + ".bak"
        os.rename(cache_file, backup_file)
        logger.info(f"已将原文件备份为 {backup_file}")
        
    except Exception as e:
        logger.error(f"迁移 ids_cache.json 失败: {e}")


def migrate_tasks():
    """迁移 task 文件夹到数据库"""
    task_dir = "./task"
    if not os.path.exists(task_dir):
        logger.info("task 文件夹不存在，跳过迁移")
        return
    
    try:
        count = 0
        for filename in os.listdir(task_dir):
            if not filename.endswith(".task"):
                continue
            
            filepath = os.path.join(task_dir, filename)
            try:
                with open(filepath, "r", encoding="gbk") as f:
                    task_data = json.loads(f.read())
                
                # 提取任务信息
                instance_id = task_data.get('pid')
                status = task_data.get('status')
                task_type = task_data.get('task')
                title = task_data.get('title', '')
                
                if instance_id and status and task_type:
                    if db.add_task(instance_id, status, task_type, title):
                        count += 1
                        logger.debug(f"迁移任务: {filename}")
                else:
                    logger.warning(f"任务文件格式不完整: {filename}")
                    
            except Exception as e:
                logger.error(f"迁移任务文件 {filename} 失败: {e}")
                continue
        
        logger.info(f"成功迁移 {count} 个任务从 task 文件夹")
        
        # 备份原文件夹
        backup_dir = task_dir + "_bak"
        if os.path.exists(backup_dir):
            import shutil
            shutil.rmtree(backup_dir)
        os.rename(task_dir, backup_dir)
        logger.info(f"已将原文件夹备份为 {backup_dir}")
        
        # 重新创建空的 task 文件夹（为了兼容性）
        os.makedirs(task_dir)
        
    except Exception as e:
        logger.error(f"迁移 task 文件夹失败: {e}")


def main():
    """执行所有迁移"""
    logger.info("=" * 50)
    logger.info("开始数据迁移到 SQLite 数据库")
    logger.info("=" * 50)
    
    # 初始化数据库
    db.init_db()
    
    # 执行迁移
    migrate_records()
    migrate_ids_cache()
    migrate_tasks()
    
    logger.info("=" * 50)
    logger.info("数据迁移完成！")
    logger.info("=" * 50)
    logger.info("原始文件已备份，如确认数据正确可手动删除备份文件")


if __name__ == "__main__":
    main()
