import logging
import os
import sys
from logging.handlers import TimedRotatingFileHandler

# 创建logs目录（如果不存在）
log_dir = "logs"
if not os.path.exists(log_dir):
    os.makedirs(log_dir)

# 创建logger对象
logger = logging.getLogger("my_logger")
logger.setLevel(logging.DEBUG)  # 设置日志的最低级别


# 创建一个StreamHandler，用于输出到控制台（设置UTF-8编码以避免Unicode错误）
# 重新配置stdout使用UTF-8编码，避免cp932编码错误
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

console_handler = logging.StreamHandler(sys.stdout)
if os.path.exists(".debug"):
    console_handler.setLevel(logging.DEBUG)  # 设置控制台输出日志的最低级别为DEBUG
else:
    console_handler.setLevel(logging.INFO)  # 设置控制台输出日志的最低级别为INFO

# 创建一个TimedRotatingFileHandler，设置日志文件每3天分割一次
handler = TimedRotatingFileHandler(os.path.join(log_dir, "debug_log.log"), when="D", interval=3, backupCount=5, encoding='utf-8')
handler.setLevel(logging.DEBUG)  # 设置Handler的日志级别

# 创建一个TimedRotatingFileHandler，设置日志文件每7天分割一次
suc_rec = TimedRotatingFileHandler(os.path.join(log_dir, "record_log.log"), when="D", interval=7, backupCount=5, encoding='utf-8')
suc_rec.setLevel(logging.CRITICAL)  # 设置Handler的日志级别

# 创建日志格式器
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
console_handler.setFormatter(formatter)
# 将Handler添加到logger中
logger.addHandler(handler)
logger.addHandler(suc_rec)
logger.addHandler(console_handler)

# # 记录不同级别的日志
# logger.debug("This is a debug message.")
# logger.info("This is an info message.")
# logger.warning("This is a warning message.")
# logger.error("This is an error message.")
# logger.critical("This is a critical message.")