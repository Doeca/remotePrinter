#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
将 files 目录下的所有文件名写入到 files_record.txt 中
"""

import os
from pathlib import Path


def list_files_to_record(files_dir='files', output_file='files_record.txt'):
    """
    遍历指定目录下的所有文件，将文件名写入输出文件
    
    Args:
        files_dir: 要遍历的目录，默认为 'files'
        output_file: 输出文件名，默认为 'files_record.txt'
    """
    # 获取脚本所在目录
    script_dir = Path(__file__).parent
    files_path = script_dir / files_dir
    output_path = script_dir / output_file
    
    # 检查目录是否存在
    if not files_path.exists():
        print(f"错误: 目录 {files_path} 不存在")
        return
    
    if not files_path.is_dir():
        print(f"错误: {files_path} 不是一个目录")
        return
    
    # 获取所有文件
    all_files = []
    for item in files_path.iterdir():
        if item.is_file():
            all_files.append(item.name)
    
    # 排序文件名
    all_files.sort()
    
    # 写入文件
    with open(output_path, 'w', encoding='utf-8') as f:
        for filename in all_files:
            f.write(filename + '\n')
    
    print(f"✓ 成功将 {len(all_files)} 个文件名写入到 {output_path}")
    print(f"\n前10个文件:")
    for i, filename in enumerate(all_files[:10], 1):
        print(f"  {i}. {filename}")
    
    if len(all_files) > 10:
        print(f"  ... 还有 {len(all_files) - 10} 个文件")


if __name__ == '__main__':
    list_files_to_record()
