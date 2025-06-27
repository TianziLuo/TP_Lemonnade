import os
import win32com.client as win32

def open_excel_file(filepath):

    if not os.path.exists(filepath):
        print(f"❌ File does not exist: {filepath}")
        return None

    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = True
        wb = excel.Workbooks.Open(filepath)
        print(f"✅ Successfully opened: {filepath}")
        return wb
    except Exception as e:
        print(f"❌ Failed to open: {filepath}\nError: {e}")
        return None

'''
if __name__ == "__main__":
    open_excel_file(r"C:\Frank\2.2_店小秘.xlsx")
'''
