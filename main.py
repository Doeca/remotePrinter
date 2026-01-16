import threading
import time
import os
import shutil
import task3
import task5
import task6
import task7
import dingLib
from logsys import logger
import db

# 常量定义
CACHE_CLEAN_INTERVAL = 1000  # 每1000次循环清理一次缓存
CHECK_INTERVAL_MINUTES = 15  # 检查间隔（分钟）
WORK_HOUR_START = 8  # 工作时间开始
WORK_HOUR_END = 20   # 工作时间结束

ALL_TASKS = [
    {"code": "PROC-203976C0-5A6E-4943-B716-5043B7F4262C",
     "task": 3, "statuses": ["COMPLETED", "RUNNING"], "results": ["agree", ""], "name": "合作档口营业款付款申请(新版)"},
    {"code": "PROC-4AD43F2F-8B07-4D46-9D1B-5DCD4342627E",
     "task": 5, "statuses": ["COMPLETED"], "results": ["agree"], "name": "付款信息变更"},
    {"code": "PROC-33B576A9-B68C-4338-A135-CE48B7C67E2E",
     "task": 6, "statuses": ["COMPLETED"], "results": ["agree"], "name": "开票申请"},
    {"code": "PROC-C7373528-790F-4A9A-B5FE-FD9564658B4E",
     "task": 7, "statuses": ["COMPLETED"], "results": ["agree"], "name": "开票红冲申请"}
]


def class_to_dict(obj):
    """将类对象递归转换为字典"""
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


def fetch_and_create_tasks():
    """获取钉钉任务实例并创建待处理任务"""
    for task in ALL_TASKS:
        logger.info(f"任务{task.get('name', 'Unknown')}开始获取列表")
        instances_result = dingLib.getInstances(task.get('code'), task.get('statuses', []))
        instance_list = instances_result.get('list', []) if instances_result else []
        
        for instanceID in instance_list:
            if db.check_record_exists(f'{instanceID}_COMPLETED'):
                # 如果只有一个状态，或者完成状态都已经处理过了，就不再处理
                # 但是有可能会有一个情况就是，RUNNING我还没有处理，但是COMPLETED已经处理了
                # 但是这个情况我也无法得知，所以只能这样了
                continue
            
            data = class_to_dict(dingLib.getDetail(instanceID))
            if not data:
                logger.warning(f"无法获取实例详情: {instanceID}")
                continue
            
            status = data.get('status')
            if not status:
                logger.warning(f"实例 {instanceID} 缺少 status 字段")
                continue

            # 如果可以有多个状态，判断当前状态是否已经处理过
            if len(task.get('statuses', [])) > 1:
                if db.check_record_exists(f'{instanceID}_{status}'):
                    logger.debug(
                        f"【Stage1】任务：{data.get('title', 'Unknown')}的{status}任务已经处理过，不处理")
                    continue

            # 检查状态是否在目标状态列表中
            if status not in task.get('statuses', []):
                # TERMINATED状态需要标记为已处理
                if status == "TERMINATED":
                    db.add_record(f"{instanceID}_{status}")
                    logger.debug(
                        f"【Stage2】任务: {data.get('title', 'Unknown')} 状态: {status}，已终止，标记为已处理")
                else:
                    logger.debug(
                        f"【Stage2】任务: {data.get('title', 'Unknown')} 状态: {status}，不在目标状态列表中，跳过")
                continue

            result = data.get('result')
            if result not in task.get('results', []):
                db.add_record(f"{instanceID}_{status}")
                logger.debug(
                    f"【Stage3】任务: {data.get('title', 'Unknown')} 结果: {result}，不在目标结果中，标记为已处理")
                continue

            # 调试模式下直接标记为已处理
            if os.path.isfile(".debug"):
                db.add_record(f"{instanceID}_{status}")
                logger.debug(
                    f"[调试模式] 标记为已处理 - 实例ID: {instanceID}_{status}, 任务类型: {task.get('task')}, 标题: {data.get('title', 'Unknown')}")
                continue
            
            # 创建新任务
            title = data.get('title', 'Unknown')
            task_type = task.get('task')
            if not task_type:
                logger.warning(f"任务 {task.get('name')} 缺少 task 字段")
                continue
            
            if not db.check_task_exists(instanceID, status):
                db.add_task(instanceID, status, task_type, title)
                db.add_record(f"{instanceID}_{status}")
                logger.info(
                    f"新任务 - PID: {instanceID}, 类型: {task_type}, 标题: {title}")
            else:
                logger.debug(f"重复任务，已跳过 - PID: {instanceID}")


def process_pending_tasks():
    """处理所有待处理的任务"""
    tasks = db.get_all_tasks()
    for task_data in tasks:
        # 使用 try-except 确保单个任务失败不会阻塞其他任务
        try:
            pid = task_data.get('pid', 'Unknown')
            task_type = task_data.get('task')
            title = task_data.get('title', 'Unknown')
            status = task_data.get('status', '')
            
            logger.info(f"处理任务 - PID: {pid}, 类型: {task_type}, 标题: {title}")
            # 任务处理函数映射
            task_handlers = {
                3: lambda: task3.handle_all(pid, status, title),
                5: lambda: task5.handle_all(pid, title),
                6: lambda: task6.handle_all(pid, title),
                7: lambda: task7.handle_all(pid, title)
            }
            
            handler = task_handlers.get(task_type)
            if handler:
                res = handler()
                if res:
                    task_id = task_data.get('id')
                    if task_id:
                        db.delete_task(task_id)
                        logger.info(f"任务完成并删除 - ID: {task_id}")
            else:
                logger.warning(f"未知的任务类型: {task_type}")
        except Exception as ex:
            logger.error(f"处理任务时出错 - PID: {task_data.get('pid', 'Unknown')}\n错误: {ex}")

class MainTask(threading.Thread):
    """主任务线程，负责定期检查和处理钉钉任务"""
    
    def __init__(self, thread_id, name):
        super().__init__()
        self.thread_id = thread_id
        self.name = name
        self.cycle_count = 0

    def clean_cache(self):
        """清理缓存目录"""
        try:
            logger.info("开始缓存清理任务")
            if os.path.exists("./cache"):
                shutil.rmtree("./cache")
            os.makedirs("./cache")
            logger.info("缓存清理任务完成")
        except Exception as ex:
            logger.error(f"缓存清理失败: {ex}")

    def is_working_hours(self):
        """检查是否在工作时间内"""
        current_hour = time.localtime().tm_hour
        return WORK_HOUR_START <= current_hour <= WORK_HOUR_END

    def process_tasks_safely(self):
        """安全地处理任务（带异常捕获）"""
        try:
            fetch_and_create_tasks()
            process_pending_tasks()
        except Exception as ex:
            logger.error(f"处理任务时出错: {ex}")

    def run(self):
        """主循环"""
        logger.info(f"启动线程: {self.name}")
        
        while True:
            self.cycle_count += 1
            
            # 定期清理缓存
            if self.cycle_count >= CACHE_CLEAN_INTERVAL:
                self.clean_cache()
                self.cycle_count = 0
            
            logger.debug("开始检查任务")
            
            # 仅在工作时间处理任务
            if self.is_working_hours():
                if os.path.isfile(".debug"):
                    # 调试模式：不捕获异常
                    fetch_and_create_tasks()
                    process_pending_tasks()
                else:
                    # 生产模式：捕获异常
                    self.process_tasks_safely()
            else:
                logger.debug("非工作时间，跳过任务检查")
            
            logger.debug("任务检查完成")
            time.sleep(CHECK_INTERVAL_MINUTES * 60)


def init_directories():
    """初始化必要的目录"""
    directories = ['cache', 'files', 'logs']
    for directory in directories:
        if not os.path.exists(directory):
            os.makedirs(directory)
            logger.info(f"创建目录: {directory}")
        else:
            logger.debug(f"目录已存在: {directory}")


if __name__ == "__main__":
    # 初始化必要的目录
    init_directories()
    
    main_thread = MainTask(1, "TaskProcessor")
    main_thread.run()
