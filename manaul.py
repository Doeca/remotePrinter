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


res = dingLib.getInstances(
    "PROC-203976C0-5A6E-4943-B716-5043B7F4262C", ['RUNNING'])

for i in res.list:
    data = dingLib.getDetail(i)
    continue_handle = False
    for vd in data.operation_records:
        # print(vd.show_name,vd.result)
        if vd.show_name == '店经理' and vd.result == 'AGREE':
            continue_handle = True
            break
    if not continue_handle:
        # print("[Task3] 店经理暂未同意，不继续处理")
        continue
    fs = open(f"./task/{i}_RUNNING.task", "w")
    fs.write(json.dumps(
        {'pid': i, 'status': "RUNNING", 'task': 3, 'title': data.title}, ensure_ascii=False))
    fs.close()
    print(f"新增任务:{i}")
