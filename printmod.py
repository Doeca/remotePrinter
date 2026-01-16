import os
import platform
from logsys import logger

# 条件导入 Windows 特定的库
IS_WINDOWS = platform.system() == 'Windows'

if IS_WINDOWS:
    try:
        import win32api
        import win32con
        import win32print
    except ImportError:
        print("Warning: win32 libraries not available, printing functions will be disabled")
        IS_WINDOWS = False

GHOSTSCRIPT_PATH = "C:\\drivers\\GHOSTSCRIPT\\bin\\gswin32.exe"
GSPRINT_PATH = "C:\\drivers\\GSPRINT\\gsprint.exe"

# YOU CAN PUT HERE THE NAME OF YOUR SPECIFIC PRINTER INSTEAD OF DEFAULT


def file_print(filename, device_name='佳能-发票打印机'):
    ff = os.path.abspath(filename)
    logger.info(f"[通知] 执行发票打印，文件路径：{ff}")
    if os.path.isfile(".noprint"):
        logger.debug(f"jump since .noprint")
        return
    
    if not IS_WINDOWS:
        logger.info(f"[调试模式] 非Windows环境，跳过实际打印操作 - 设备: {device_name}")
        return
    
    win32api.ShellExecute(0, 'open', GSPRINT_PATH, '-ghostscript "'+GHOSTSCRIPT_PATH +
                          '" -dPDFFitPage -printer "'+device_name+f'" "{ff}"', '.', 0)


def a4_print(filename, device_name='柯尼卡-A4打印机'):
    ff = os.path.abspath(filename)
    logger.info(f"[通知] 执行打印，文件路径：{ff}")
    if os.path.isfile(".noprint"):
        logger.debug(f"jump since .noprint")
        return
    
    if not IS_WINDOWS:
        logger.info(f"[调试模式] 非Windows环境，跳过实际打印操作 - 设备: {device_name}")
        return
    
    win32api.ShellExecute(0, 'open', GSPRINT_PATH, '-ghostscript "'+GHOSTSCRIPT_PATH +
                          '" -dDuplex=true -dPDFFitPage -sPAPERSIZE=a4 -printer "'+device_name+f'" "{ff}"', '.', 0)
    return


def a4_print_singleside(filename, device_name='柯尼卡-A4打印机'):
    ff = os.path.abspath(filename)
    logger.info(f"[通知] 执行单面打印，文件路径：{ff}")
    if os.path.isfile(".noprint"):
        logger.debug(f"jump since .noprint")
        return
    
    if not IS_WINDOWS:
        logger.info(f"[调试模式] 非Windows环境，跳过实际打印操作 - 设备: {device_name}")
        return
    
    win32api.ShellExecute(0, 'open', GSPRINT_PATH, '-ghostscript "'+GHOSTSCRIPT_PATH +
                          '" -dDuplex=false -dPDFFitPage -sPAPERSIZE=a4 -printer "'+device_name+f'" "{ff}"', '.', 0)
    return
