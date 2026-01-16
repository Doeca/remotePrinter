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
    "PROC-C7373528-790F-4A9A-B5FE-FD9564658B4E", ['COMPLETED'])


records = json.loads(s=open("./record.dat", "r").read())
for instanceID in res['list']:
    if f'{instanceID}_COMPLETED' in records:
        continue

    records.append(f'{instanceID}_COMPLETED')

json.dump(records, open("./record.dat", "w"), ensure_ascii=False)