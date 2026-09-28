# -*- coding: utf-8 -*-
"""
word-deliverable / render_pdf_uno.py
用 LibreOffice UNO API 打开 .docx，更新目录(TOC)与所有字段(PAGE/NUMPAGES)，
然后导出为 PDF。

必须用 LibreOffice 自带 Python 运行：
    "C:\\Program Files\\LibreOffice\\program\\python.exe" render_pdf_uno.py <input.docx> <output.pdf>

原理：
  1. 启动 soffice 监听 socket
  2. UNO 连接 → 打开文档(隐藏)
  3. doc.refresh() 更新所有字段
  4. 遍历 DocumentIndexes 更新所有目录
  5. storeToURL 导出 PDF
  6. 关闭文档、终止 soffice
"""
import os
import sys
import time
import subprocess
import socket

SOFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"


def _free_port():
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port


def main():
    if len(sys.argv) < 3:
        print("用法: python render_pdf_uno.py <input.docx> <output.pdf>")
        sys.exit(1)

    docx_path = os.path.abspath(sys.argv[1])
    pdf_path = os.path.abspath(sys.argv[2])
    if not os.path.exists(docx_path):
        print(f"文件不存在: {docx_path}")
        sys.exit(1)

    import uno
    from com.sun.star.beans import PropertyValue

    port = _free_port()
    accept_str = f"socket,host=127.0.0.1,port={port};urp;"

    # 启动 soffice 监听
    proc = subprocess.Popen(
        [SOFFICE, "--headless", "--invisible", "--nocrashreport",
         "--nodefault", "--nologo", "--nofirststartwizard",
         "--norestore", f"--accept={accept_str}"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )

    try:
        # 等待连接就绪
        localContext = uno.getComponentContext()
        resolver = localContext.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", localContext)
        ctx = None
        for i in range(30):
            try:
                ctx = resolver.resolve(
                    f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(0.5)
        if ctx is None:
            print("错误: 无法连接到 LibreOffice UNO")
            sys.exit(2)

        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

        # 打开文档
        url = uno.systemPathToFileUrl(docx_path)
        hidden = PropertyValue()
        hidden.Name = "Hidden"
        hidden.Value = True
        doc = desktop.loadComponentFromURL(url, "_blank", 0, (hidden,))

        if doc is None:
            print("错误: 无法打开文档")
            sys.exit(3)

        # 更新所有字段
        try:
            doc.refresh()
        except Exception as e:
            print(f"  refresh 警告: {e}")

        # 更新所有目录(TOC)
        try:
            indexes = doc.getDocumentIndexes()
            for i in range(indexes.getCount()):
                indexes.getByIndex(i).update()
        except Exception as e:
            print(f"  indexes 警告: {e}")

        # 再次刷新字段（TOC 更新后页码可能变化）
        try:
            doc.refresh()
        except Exception:
            pass

        # 导出 PDF
        pdf_url = uno.systemPathToFileUrl(pdf_path)
        filter_name = PropertyValue()
        filter_name.Name = "FilterName"
        filter_name.Value = "writer_pdf_Export"
        doc.storeToURL(pdf_url, (filter_name,))

        doc.close(True)
        print(f"OK PDF (UNO): {pdf_path}")
    except Exception as e:
        print(f"UNO 失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(4)
    finally:
        try:
            desktop.terminate()
        except Exception:
            pass
        try:
            proc.wait(timeout=10)
        except Exception:
            proc.kill()


if __name__ == "__main__":
    main()
