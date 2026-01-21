#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
根据流程code和审批编号(business_id)查找实例id
"""

import sys
import os

# 添加父目录到系统路径，以便导入dingLib
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import dingLib
from logsys import logger


def find_instance_by_business_id(process_code, business_id):
    """
    根据流程code和审批编号查找实例id
    
    Args:
        process_code: 流程code，例如 'PROC-203976C0-5A6E-4943-B716-5043B7F4262C'
        business_id: 审批编号，例如 '202411211215000279665'
    
    Returns:
        str: 找到的实例id，如果未找到返回None
    """
    logger.info(f"开始查找实例: process_code={process_code}, business_id={business_id}")
    
    try:
        # 获取该流程下的所有实例id列表
        # 默认获取COMPLETED状态，如果需要其他状态可以修改p_statuses参数
        instances_data = dingLib.getInstances(process_code, p_statuses=["COMPLETED", "RUNNING"])
        instance_list = instances_data.get('list', [])
        
        logger.info(f"共获取到 {len(instance_list)} 个实例")
        
        # 遍历所有实例，查找匹配的business_id
        for idx, instance_id in enumerate(instance_list, 1):
            logger.debug(f"正在检查第 {idx}/{len(instance_list)} 个实例: {instance_id}")
            
            # 获取实例详情
            detail = dingLib.getDetail(instance_id)
            
            if detail:
                # 转换为字典格式
                detail_dict = dingLib.class_to_dict(detail)
                current_business_id = detail_dict.get('business_id', '')
                
                # 比对business_id
                if current_business_id == business_id:
                    logger.info(f"✓ 找到匹配的实例: {instance_id}")
                    logger.info(f"  标题: {detail_dict.get('title', 'N/A')}")
                    logger.info(f"  状态: {detail_dict.get('status', 'N/A')}")
                    logger.info(f"  创建时间: {detail_dict.get('create_time', 'N/A')}")
                    return instance_id
                else:
                    logger.debug(f"  business_id不匹配: {current_business_id}")
            else:
                logger.warning(f"无法获取实例 {instance_id} 的详情")
        
        logger.warning(f"未找到匹配的实例: business_id={business_id}")
        return None
        
    except Exception as e:
        logger.error(f"查找实例时出错: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return None


def main():
    """
    主函数 - 支持命令行调用
    """
    if len(sys.argv) < 3:
        print("使用方法:")
        print(f"  python {sys.argv[0]} <process_code> <business_id>")
        print("\n示例:")
        print(f"  python {sys.argv[0]} PROC-203976C0-5A6E-4943-B716-5043B7F4262C 202411211215000279665")
        sys.exit(1)
    
    process_code = sys.argv[1]
    business_id = sys.argv[2]
    
    print(f"\n{'='*60}")
    print(f"流程Code: {process_code}")
    print(f"审批编号: {business_id}")
    print(f"{'='*60}\n")
    
    instance_id = find_instance_by_business_id(process_code, business_id)
    
    if instance_id:
        print(f"\n✓ 成功找到实例ID: {instance_id}\n")
        return 0
    else:
        print(f"\n✗ 未找到匹配的实例\n")
        return 1


if __name__ == '__main__':
    # 示例用法（如果不提供命令行参数，可以直接修改这里进行测试）
    if len(sys.argv) == 1:
        # 测试数据
        test_process_code = "PROC-203976C0-5A6E-4943-B716-5043B7F4262C"
        test_business_id = "202411211215000279665"
        
        print("使用测试数据运行...")
        instance_id = find_instance_by_business_id(test_process_code, test_business_id)
        
        if instance_id:
            print(f"\n找到的实例ID: {instance_id}")
        else:
            print("\n未找到匹配的实例")
    else:
        sys.exit(main())


# python find_instance_by_businessid.py PROC-203976C0-5A6E-4943-B716-5043B7F4262C 202601141732000379642