import os
import base64
import magic
import requests
import json
import subprocess
import platform
from PIL import Image
from PyPDF2 import PdfMerger
from reportlab.lib.pagesizes import mm
from reportlab.pdfgen import canvas
from logsys import logger

def mergeImages(pdfp, files, wid=240, hei=140):
    valid_files = []
    for fi in files:
        try:
            with open(fi, 'rb') as f:
                contents = f.read()
            img_format = magic.from_buffer(contents, mime=True).split('/')[-1]
            if img_format == 'heic':
                ctr = base64.b64encode(contents).decode("utf-8")
                open("./cache/text.txt", "w").write(ctr)
                resp = requests.post("http://tun4.cquluna.top:20062/convert", headers={"Content-Type": "application/json"},
                                     data=json.dumps({"data": ctr}))
                open(fi, 'wb').write(base64.b64decode(resp.text))
            tmp = Image.open(fi)
            if (tmp.height > tmp.width) != (hei > wid):
                tmp = tmp.transpose(Image.ROTATE_90)
            tmp.save(fi)
            valid_files.append(fi)
        except Exception as ex:
            logger.warning(f"画像処理失敗、スキップ: {fi} - エラー: {ex}")
            continue

    # reportlab で画像を PDF に変換
    try:
        c = canvas.Canvas(pdfp, pagesize=(wid*mm, hei*mm))
        
        for fi in valid_files:
            try:
                img = Image.open(fi)
                img_width, img_height = img.size
                
                # ページサイズに合わせて画像を配置
                c.drawImage(fi, 0, 0, width=wid*mm, height=hei*mm, preserveAspectRatio=True)
                c.showPage()
                
            except Exception as ex:
                logger.warning(f"PDF変換失敗、スキップ: {fi} - エラー: {ex}")
                continue
        
        c.save()
        logger.info(f"PDF生成完了: {pdfp}")
        
    except Exception as e:
        logger.error(f"PDF生成失敗: {e}")
        raise


def mergePdfs(pdfp, files):
    file_merger = PdfMerger()
    for fi in files:
        file_merger.append(fi)
    file_merger.write(os.path.abspath(pdfp))
    logger.debug(os.path.abspath(pdfp))


def pdfexport(pdfp, prefix):
    """
    PDFをPNG画像に変換（Ghostscript 使用）
    
    Args:
        pdfp: PDFファイルパス
        prefix: 出力画像のプレフィックス
    
    Returns:
        生成された画像ファイルパスのリスト
    """
    if not os.path.exists(pdfp):
        logger.error(f"PDF file not found: {pdfp}")
        return []
    
    # Ghostscript パスを設定
    if platform.system() == 'Windows':
        gs_path = os.path.abspath("./drivers/GHOSTSCRIPT/bin/gswin64c.exe")
        if not os.path.exists(gs_path):
            gs_path = os.path.abspath("./drivers/GHOSTSCRIPT/bin/gswin32c.exe")
    else:
        gs_path = "gs"  # macOS/Linux
    
    if not os.path.exists(gs_path) and platform.system() == 'Windows':
        logger.error(f"Ghostscript not found: {gs_path}")
        return []
    
    try:
        # ページ数を取得（Ghostscript で取得 - PyMuPDF を完全回避）
        cmd_count = [
            gs_path,
            "-q",
            "-dNODISPLAY",
            "-dNOSAFER",
            "-c",
            f"({pdfp}) (r) file runpdfbegin pdfpagecount = quit"
        ]
        
        result = subprocess.run(cmd_count, capture_output=True, text=True, timeout=10)
        if result.returncode == 0:
            page_count = int(result.stdout.strip())
        else:
            # フォールバック: PyPDF2 で取得
            from PyPDF2 import PdfReader
            reader = PdfReader(pdfp)
            page_count = len(reader.pages)
        
        logger.info(f"PDF読み込み成功: {pdfp}, ページ数: {page_count}")
        
        res = []
        
        for pg in range(page_count):
            output_path = f"{prefix}_{pg}.png"
            
            # Ghostscript コマンド
            cmd = [
                gs_path,
                "-dNOPAUSE",
                "-dBATCH",
                "-dSAFER",
                "-sDEVICE=png16m",
                "-r150",  # DPI 150
                f"-dFirstPage={pg+1}",
                f"-dLastPage={pg+1}",
                f"-sOutputFile={output_path}",
                pdfp
            ]
            
            logger.info(f"开始处理第 {pg+1}/{page_count} 页")
            logger.debug(f"Ghostscript 命令: {' '.join(cmd)}")
            
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=30
            )
            
            if result.returncode != 0:
                logger.error(f"✗ 页面 {pg+1} 转换失败: {result.stderr}")
                continue
            
            if os.path.exists(output_path):
                res.append(output_path)
                logger.info(f"✓ 页面 {pg+1}/{page_count} 转换成功: {output_path}")
            else:
                logger.error(f"✗ 页面 {pg+1} 输出文件未生成: {output_path}")
        
        logger.info(f"PDF変換完了: {len(res)} 枚の画像を生成")
        return res
        
    except Exception as e:
        logger.error(f"PDF変換失敗: {type(e).__name__}: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return []
