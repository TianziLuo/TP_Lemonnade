
from pathlib import Path
from utils.open_excel_utils import open_excel_file
from utils.SKU_mapping import SKU_out
from utils.copy_file_utils import copy_file_to_dirs
from utils.create_file_utils import generate_tp_csv, update_inventory_excel
from utils.TP_upload_utils import teapplix_upload
from utils.download_rename import rename_DXM

downloads = Path.home() / "Downloads"

# === Step 1 ===
def step_1_1(): open_excel_file(r"C:\Frank\1.1_核心.xlsx")
def step_1_2(): SKU_out()
def step_1_3():
    copy_file_to_dirs(
        r"C:\Users\monica\Downloads\上传易仓SKU映射关系.xls",
        [r"C:\Frank\易仓-TP"]
    )

# === Step 2 ===
def step_2_1(): open_excel_file(r"C:\Frank\2.1_易仓管理.xlsx")
def step_2_2():
    generate_tp_csv(
        source_path=r"C:\Frank\2.1_易仓管理.xlsx",
        sheet_name="易仓进TP",
        template_path=r"C:\Template\TP-Upload.csv",
        output_path=downloads / "TP-Upload.csv"
    )

def step_2_3a():
    teapplix_upload(
        username="wayfaircolourtree",
        email="wayfair.colourtree@gmail.com",
        password="Colourtree168!!",
        csv_path=str(downloads / "TP-Upload.csv")
    )

def step_2_3b():
    teapplix_upload(
        username="colourtree",
        email="colourtreeusa@gmail.com",
        password="Colourtree168!",
        csv_path=str(downloads / "TP-Upload.csv")
    )

def step_2_4():
    copy_file_to_dirs(
        str(downloads / "TP-Upload.csv"),
        [
            r"C:\Frank\易仓-TP",
            r"C:\ACT\公用核心\upload_1.1 TP 库存更新"
        ]
    )

def step_2_5():
    open_excel_file(r"C:\Frank\易仓-TP\SKUINV.xlsx")

def step_2_6():
    copy_file_to_dirs(
        r"C:\Frank\易仓-TP\SKUINV.xlsx",
        [r"\\MICHAEL\ctshippingapp\SHIPDOC\INVUPLOAD"]
    )

# === Step 3 ===
def step_3_1(): rename_DXM()
def step_3_2(): open_excel_file(r"C:\Frank\2.2_店小秘.xlsx")

def step_3_3():
    output = downloads / "店小秘 更新库存.xlsx"
    update_inventory_excel(
        source_path=r"C:\Frank\2.2_店小秘.xlsx",
        sheet_name="盘点",
        template_path=r"C:\Template\店小秘 更新库存.xlsx",
        output_path=output
    )
def step_3_4():
    copy_file_to_dirs(
        str(downloads / "店小秘 更新库存.xlsx"),
        [
            r"C:\Frank\原始数据\店小秘+TP+订单+盘点",
            r"C:\ACT\公用核心\Upload_2.1 店小秘  库存更新"
        ]
    )
