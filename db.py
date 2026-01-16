import sqlite3
import os
import json
from logsys import logger

# 根据运行模式选择不同的数据库文件
def get_db_path():
    """根据是否存在.test文件返回不同的数据库路径"""
    if os.path.isfile(".debug"):
        # 调试/测试模式使用单独的数据库
        return os.path.join("logs", "remotePrinter_test.db")
    else:
        # 生产模式使用正式数据库
        return os.path.join("logs", "remotePrinter.db")

DB_PATH = get_db_path()


def init_db():
    """初始化数据库，创建所有需要的表"""
    # 确保logs目录存在
    log_dir = os.path.dirname(DB_PATH)
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # 创建 records 表 - 替代 record.dat
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            record_key TEXT UNIQUE NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # 创建索引以提高查询性能
    cursor.execute('''
        CREATE INDEX IF NOT EXISTS idx_record_key ON records(record_key)
    ''')
    
    # 创建 ids_cache 表 - 替代 ids_cache.json
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS ids_cache (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            process_code TEXT NOT NULL,
            instance_id TEXT NOT NULL,
            timestamp INTEGER NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(process_code, instance_id)
        )
    ''')
    
    # 创建索引以提高查询性能
    cursor.execute('''
        CREATE INDEX IF NOT EXISTS idx_process_code ON ids_cache(process_code)
    ''')
    cursor.execute('''
        CREATE INDEX IF NOT EXISTS idx_timestamp ON ids_cache(timestamp)
    ''')
    
    # 创建 tasks 表 - 替代 task 文件夹
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            instance_id TEXT NOT NULL,
            status TEXT NOT NULL,
            task_type INTEGER NOT NULL,
            title TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(instance_id, status)
        )
    ''')
    
    # 创建索引以提高查询性能
    cursor.execute('''
        CREATE INDEX IF NOT EXISTS idx_instance_status ON tasks(instance_id, status)
    ''')
    
    conn.commit()
    conn.close()
    logger.info("数据库初始化完成")


# ========== Records 操作 ==========

def add_record(record_key: str) -> bool:
    """添加一条记录"""
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('INSERT OR IGNORE INTO records (record_key) VALUES (?)', (record_key,))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        logger.error(f"添加记录失败: {e}")
        return False


def check_record_exists(record_key: str) -> bool:
    """检查记录是否存在"""
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('SELECT 1 FROM records WHERE record_key = ? LIMIT 1', (record_key,))
        result = cursor.fetchone()
        conn.close()
        return result is not None
    except Exception as e:
        logger.error(f"检查记录失败: {e}")
        return False


def get_all_records() -> list:
    """获取所有记录"""
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('SELECT record_key FROM records')
        results = [row[0] for row in cursor.fetchall()]
        conn.close()
        return results
    except Exception as e:
        logger.error(f"获取所有记录失败: {e}")
        return []


# ========== IDs Cache 操作 ==========

def add_ids_cache(process_code: str, instance_id: str, timestamp: int) -> bool:
    """添加或更新ID缓存"""
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT OR REPLACE INTO ids_cache (process_code, instance_id, timestamp) 
            VALUES (?, ?, ?)
        ''', (process_code, instance_id, timestamp))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        logger.error(f"添加ID缓存失败: {e}")
        return False


def get_ids_cache_by_process(process_code: str) -> dict:
    """获取指定process_code的所有缓存"""
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT instance_id, timestamp FROM ids_cache 
            WHERE process_code = ?
        ''', (process_code,))
        results = {row[0]: row[1] for row in cursor.fetchall()}
        conn.close()
        return results
    except Exception as e:
        logger.error(f"获取ID缓存失败: {e}")
        return {}


def check_ids_cache_exists(process_code: str, instance_id: str) -> bool:
    """检查ID缓存是否存在"""
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT 1 FROM ids_cache 
            WHERE process_code = ? AND instance_id = ? 
            LIMIT 1
        ''', (process_code, instance_id))
        result = cursor.fetchone()
        conn.close()
        return result is not None
    except Exception as e:
        logger.error(f"检查ID缓存失败: {e}")
        return False


def delete_old_ids_cache(process_code: str, timestamp_threshold: int) -> int:
    """删除指定process_code下时间戳小于threshold的缓存"""
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('''
            DELETE FROM ids_cache 
            WHERE process_code = ? AND timestamp < ?
        ''', (process_code, timestamp_threshold))
        deleted_count = cursor.rowcount
        conn.commit()
        conn.close()
        logger.debug(f"清除过期缓存 {deleted_count} 条")
        return deleted_count
    except Exception as e:
        logger.error(f"删除过期缓存失败: {e}")
        return 0


def get_all_ids_cache() -> dict:
    """获取所有ID缓存，返回嵌套字典格式"""
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('SELECT process_code, instance_id, timestamp FROM ids_cache')
        results = {}
        for row in cursor.fetchall():
            process_code, instance_id, timestamp = row
            if process_code not in results:
                results[process_code] = {}
            results[process_code][instance_id] = timestamp
        conn.close()
        return results
    except Exception as e:
        logger.error(f"获取所有ID缓存失败: {e}")
        return {}


# ========== Tasks 操作 ==========

def add_task(instance_id: str, status: str, task_type: int, title: str) -> bool:
    """添加一个任务"""
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT OR IGNORE INTO tasks (instance_id, status, task_type, title) 
            VALUES (?, ?, ?, ?)
        ''', (instance_id, status, task_type, title))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        logger.error(f"添加任务失败: {e}")
        return False


def get_all_tasks() -> list:
    """获取所有任务"""
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT id, instance_id, status, task_type, title FROM tasks 
            ORDER BY created_at ASC
        ''')
        results = []
        for row in cursor.fetchall():
            results.append({
                'id': row[0],
                'pid': row[1],
                'status': row[2],
                'task': row[3],
                'title': row[4]
            })
        conn.close()
        return results
    except Exception as e:
        logger.error(f"获取所有任务失败: {e}")
        return []


def delete_task(task_id: int) -> bool:
    """删除任务"""
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('DELETE FROM tasks WHERE id = ?', (task_id,))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        logger.error(f"删除任务失败: {e}")
        return False


def check_task_exists(instance_id: str, status: str) -> bool:
    """检查任务是否存在"""
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT 1 FROM tasks 
            WHERE instance_id = ? AND status = ? 
            LIMIT 1
        ''', (instance_id, status))
        result = cursor.fetchone()
        conn.close()
        return result is not None
    except Exception as e:
        logger.error(f"检查任务失败: {e}")
        return False


def get_latest_record_by_task_type(task_type: int) -> dict:
    """获取指定任务类型的最新一条记录（基于COMPLETED状态）"""
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        # 从records表中查找该任务类型对应的最新COMPLETED记录
        # record_key格式为: {instanceID}_COMPLETED
        cursor.execute('''
            SELECT record_key FROM records 
            WHERE record_key LIKE '%_COMPLETED'
            ORDER BY created_at DESC
        ''')
        
        # 需要通过instanceID关联到tasks表来确认任务类型
        for row in cursor.fetchall():
            record_key = row[0]
            instance_id = record_key.replace('_COMPLETED', '')
            
            # 查询这个instance_id的任务类型
            cursor2 = conn.cursor()
            cursor2.execute('''
                SELECT instance_id, status, task_type, title 
                FROM tasks 
                WHERE instance_id = ? AND task_type = ?
                ORDER BY created_at DESC
                LIMIT 1
            ''', (instance_id, task_type))
            
            task_row = cursor2.fetchone()
            if task_row:
                conn.close()
                return {
                    'pid': task_row[0],
                    'status': task_row[1],
                    'task': task_row[2],
                    'title': task_row[3]
                }
        
        # 如果没有找到COMPLETED的，尝试找任何状态的
        cursor.execute('''
            SELECT instance_id, status, task_type, title 
            FROM tasks 
            WHERE task_type = ?
            ORDER BY created_at DESC
            LIMIT 1
        ''', (task_type,))
        
        task_row = cursor.fetchone()
        conn.close()
        
        if task_row:
            return {
                'pid': task_row[0],
                'status': task_row[1],
                'task': task_row[2],
                'title': task_row[3]
            }
        return None
    except Exception as e:
        logger.error(f"获取最新任务记录失败 (task_type={task_type}): {e}")
        return None


def get_latest_records_for_all_task_types() -> list:
    """获取所有任务类型的最新一条记录"""
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        # 获取所有不同的任务类型
        cursor.execute('SELECT DISTINCT task_type FROM tasks ORDER BY task_type')
        task_types = [row[0] for row in cursor.fetchall()]
        conn.close()
        
        results = []
        for task_type in task_types:
            record = get_latest_record_by_task_type(task_type)
            if record:
                results.append(record)
        
        return results
    except Exception as e:
        logger.error(f"获取所有任务类型的最新记录失败: {e}")
        return []


# ========== 数据库回滚操作 ==========

def get_records_count_by_date(date_string: str) -> dict:
    """获取指定日期之后各表的记录数量"""
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        counts = {}
        
        # 统计records表
        cursor.execute('''
            SELECT COUNT(*) FROM records 
            WHERE created_at >= datetime(?)
        ''', (date_string,))
        counts['records'] = cursor.fetchone()[0]
        
        # 统计ids_cache表
        cursor.execute('''
            SELECT COUNT(*) FROM ids_cache 
            WHERE created_at >= datetime(?)
        ''', (date_string,))
        counts['ids_cache'] = cursor.fetchone()[0]
        
        # 统计tasks表
        cursor.execute('''
            SELECT COUNT(*) FROM tasks 
            WHERE created_at >= datetime(?)
        ''', (date_string,))
        counts['tasks'] = cursor.fetchone()[0]
        
        conn.close()
        return counts
    except Exception as e:
        logger.error(f"统计记录数量失败: {e}")
        return {}


def rollback_records_by_date(date_string: str, dry_run: bool = True) -> dict:
    """回滚指定日期之后的所有记录
    
    Args:
        date_string: 日期字符串，格式如 '2026-01-16' 或 '2026-01-16 10:30:00'
        dry_run: 是否仅模拟运行（不实际删除），默认True
    
    Returns:
        包含删除统计信息的字典
    """
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        result = {
            'date': date_string,
            'dry_run': dry_run,
            'deleted': {}
        }
        
        # 删除records表记录
        if dry_run:
            cursor.execute('''
                SELECT COUNT(*) FROM records 
                WHERE created_at >= datetime(?)
            ''', (date_string,))
            result['deleted']['records'] = cursor.fetchone()[0]
        else:
            cursor.execute('''
                DELETE FROM records 
                WHERE created_at >= datetime(?)
            ''', (date_string,))
            result['deleted']['records'] = cursor.rowcount
        
        # 删除ids_cache表记录
        if dry_run:
            cursor.execute('''
                SELECT COUNT(*) FROM ids_cache 
                WHERE created_at >= datetime(?)
            ''', (date_string,))
            result['deleted']['ids_cache'] = cursor.fetchone()[0]
        else:
            cursor.execute('''
                DELETE FROM ids_cache 
                WHERE created_at >= datetime(?)
            ''', (date_string,))
            result['deleted']['ids_cache'] = cursor.rowcount
        
        # 删除tasks表记录
        if dry_run:
            cursor.execute('''
                SELECT COUNT(*) FROM tasks 
                WHERE created_at >= datetime(?)
            ''', (date_string,))
            result['deleted']['tasks'] = cursor.fetchone()[0]
        else:
            cursor.execute('''
                DELETE FROM tasks 
                WHERE created_at >= datetime(?)
            ''', (date_string,))
            result['deleted']['tasks'] = cursor.rowcount
        
        if not dry_run:
            conn.commit()
            logger.info(f"回滚完成 - 删除 records: {result['deleted']['records']}, "
                       f"ids_cache: {result['deleted']['ids_cache']}, "
                       f"tasks: {result['deleted']['tasks']}")
        else:
            logger.info(f"模拟回滚 - 将删除 records: {result['deleted']['records']}, "
                       f"ids_cache: {result['deleted']['ids_cache']}, "
                       f"tasks: {result['deleted']['tasks']}")
        
        conn.close()
        return result
    except Exception as e:
        logger.error(f"回滚操作失败: {e}")
        return {'error': str(e)}


def get_date_range_info() -> dict:
    """获取数据库中记录的日期范围信息"""
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        info = {}
        
        # records表日期范围
        cursor.execute('''
            SELECT MIN(created_at), MAX(created_at), COUNT(*) 
            FROM records
        ''')
        row = cursor.fetchone()
        info['records'] = {
            'min_date': row[0],
            'max_date': row[1],
            'count': row[2]
        }
        
        # ids_cache表日期范围
        cursor.execute('''
            SELECT MIN(created_at), MAX(created_at), COUNT(*) 
            FROM ids_cache
        ''')
        row = cursor.fetchone()
        info['ids_cache'] = {
            'min_date': row[0],
            'max_date': row[1],
            'count': row[2]
        }
        
        # tasks表日期范围
        cursor.execute('''
            SELECT MIN(created_at), MAX(created_at), COUNT(*) 
            FROM tasks
        ''')
        row = cursor.fetchone()
        info['tasks'] = {
            'min_date': row[0],
            'max_date': row[1],
            'count': row[2]
        }
        
        conn.close()
        return info
    except Exception as e:
        logger.error(f"获取日期范围信息失败: {e}")
        return {}


# 初始化数据库
init_db()
