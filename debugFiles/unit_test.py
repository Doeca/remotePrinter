import sys
import os

# 親ディレクトリをパスに追加
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import dingLib
import json


def class_to_dict(obj):
    if isinstance(obj, dict):
        return {k: class_to_dict(v) for k, v in obj.items()}
    elif hasattr(obj, "_ast"):
        return class_to_dict(obj._ast())
    elif hasattr(obj, "__iter__") and not isinstance(obj, str):
        return [class_to_dict(v) for v in obj]
    elif hasattr(obj, "__dict__"):
        return {k: class_to_dict(v) for k, v in obj.__dict__.items() if not k.startswith('_')}
    else:
        return obj


# PROC-203976C0-5A6E-4943-B716-5043B7F4262C 合作档口营业款付款申请(新版)
# res = dingLib.getInstances(
#     "PROC-4AD43F2F-8B07-4D46-9D1B-5DCD4342627E", ['COMPLETED'])

for i in  ['QgrXPaClQcKZBAsvUCOiZw03401768378823']:
    res = json.dumps(class_to_dict(dingLib.getDetail(i)), ensure_ascii=False).encode("utf-8")
    open(f"./test{i}.json", "wb").write(res)


# 2026-01-23 17:17:55,412 - my_logger - INFO - 开始查找实例: process_code=PROC-203976C0-5A6E-4943-B716-5043B7F4262C, business_id=202601141620000236974
# 2026-01-23 17:18:13,178 - my_logger - INFO - 共获取到 75 个实例
# 2026-01-23 17:18:22,011 - my_logger - INFO - ✓ 找到匹配的实例: QgrXPaClQcKZBAsvUCOiZw03401768378823
# 2026-01-23 17:18:22,012 - my_logger - INFO -   标题: 何小清提交的合作档口营业款付款申请
# 2026-01-23 17:18:22,012 - my_logger - INFO -   状态: COMPLETED
# 2026-01-23 17:18:22,012 - my_logger - INFO -   创建时间: 2026-01-14T16:20Z

# ✓ 成功找到实例ID: 


# SELECT * FROM records WHERE record_key LIKE '%QgrXPaClQcKZBAsvUCOiZw03401768378823%';