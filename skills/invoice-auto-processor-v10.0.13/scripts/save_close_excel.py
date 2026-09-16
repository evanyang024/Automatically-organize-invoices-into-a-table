"""
Excel 锁住时的处理脚本 — 用 pywin32 COM 保存并关闭
===== 前提: pywin32 已安装 =====
"""
import win32com.client
import os
import time

def save_and_close_excel(filepath, timeout=10):
    """找到打开指定文件的 Excel 实例，保存并关闭"""
    filename = os.path.basename(filepath)
    
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except:
        print("未找到运行中的 Excel 实例")
        return False

    found = False
    for wb in excel.Workbooks:
        if os.path.basename(wb.FullName) == filename:
            print(f"找到: {wb.FullName}")
            wb.Save()
            found = True

    if found:
        time.sleep(1)  # 等保存完成
        excel.Quit()
        print("Excel 已保存并关闭")
    else:
        print(f"未找到打开的文件: {filename}")

    return found

# 用法:
# save_and_close_excel("E:/桌面E/待整理发票/output/发票台账.xlsx")
