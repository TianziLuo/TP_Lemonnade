import shutil
from pathlib import Path
from typing import Union, List

def copy_file_to_dirs(source: Union[str, Path], target_dirs: List[Union[str, Path]]) -> List[Path]:
    """

    Parameters:
        source (str | Path): Path to the source file.
        target_dirs (list of str | Path): List of target directory paths.

    Returns:
        List[Path]: List of successfully copied file paths.
    """
    copied_paths = []
    source_path = Path(source)

    if not source_path.is_file():
        print(f"❌ Source file does not exist: {source_path}")
        return copied_paths

    for dir_path in target_dirs:
        target_dir = Path(dir_path)
        try:
            target_dir.mkdir(parents=True, exist_ok=True)
            dest_file = target_dir / source_path.name
            shutil.copy2(source_path, dest_file)
            copied_paths.append(dest_file)
        except Exception as e:
            print(f"❌ Failed to copy to {target_dir}: {e}")

    if copied_paths:
        print("✅ Copied:")
        for path in copied_paths:
            print(f"→ {path}")

    return copied_paths

r'''
if __name__ == "__main__":
    copy_file_to_dirs(
    r"C:\Users\monica\Downloads\店小秘 更新库存.xlsx",
    [
        r"C:\Frank\原始数据\店小秘+TP+订单+盘点",
        r"C:\ACT\公用核心\Upload_2.1 店小秘  库存更新"
    ]
)
'''