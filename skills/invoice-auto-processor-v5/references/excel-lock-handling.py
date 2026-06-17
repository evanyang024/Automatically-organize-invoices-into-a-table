# Excel 锁处理 — pywin32 COM 自动化

当程序因 Excel 打开而报 Permission denied 时，执行以下步骤：

```python
import win32com.client

def save_and_close_excel(filename="发票台账.xlsx"):
    """连接运行中的 Excel，保存指定工作簿，关闭 Excel"""
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except:
        print("Excel 未运行")
        return False

    for wb in excel.Workbooks:
        if filename in wb.Name:
            wb.Save()
            print(f"已保存: {wb.Name}")
            break

    excel.Quit()
    print("Excel 已关闭")
    return True
```

**使用时机**：`append_invoice_rows` 或 `process_existing_files` 抛出 `PermissionError` 时自动调用。

**注意**：不要用 `taskkill /F /IM EXCEL.EXE`，会丢失用户未保存的修改。
