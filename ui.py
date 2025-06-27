import tkinter as tk
from tkinter import messagebox
import tp_tasks

# 统一配色
BG_COLOR      = "#FFFAEE"  # 米白背景
BTN_COLOR     = "#FFF4B2"  # 淡黄按钮
BTN_HOVER     = "#B2DFF7"  # 按钮悬停色
LABEL_BG      = "#CBE8F3"  # 浅蓝标题背景
TEXT_COLOR    = "#020A1B"  # 主文字颜色
TITLE_COLOR   = "#055E61"  # Header 标题颜色

def run_safe(func, name):
    try:
        func()
        messagebox.showinfo("完成", f"{name} 成功 ✅")
    except Exception as e:
        messagebox.showerror("错误", f"{name} 失败 ❌\n{e}")

def add_section(frame, title, steps):
    section = tk.LabelFrame(frame, text=title, padx=10, pady=8,
                            bg=LABEL_BG, fg=TEXT_COLOR, font=("Segoe UI", 11, "bold"))
    section.pack(pady=10, fill="x", padx=10)
    for label, func in steps:
        btn = tk.Button(section, text=label, width=35, bg=BTN_COLOR, fg=TEXT_COLOR, relief="groove",
                        activebackground=BTN_HOVER, font=("Segoe UI", 10),
                        command=lambda f=func, l=label: run_safe(f, l))
        btn.pack(pady=2)

def create_ui():
    root = tk.Tk()
    root.title("Inventory Update_lemonnade☀️")
    root.configure(bg=BG_COLOR)
    root.geometry("450x725")
    root.resizable(False, False)

    # ===== Header 标题 =====
    header = tk.Label(root, text="Inventory update _ lemonnade🍹", font=("Segoe UI", 16, "bold"),
                      bg=BG_COLOR, fg=TITLE_COLOR)
    header.pack(pady=5)

    # ===== 主框架 =====
    frame = tk.Frame(root, bg=BG_COLOR)
    frame.pack(pady=10, fill="both", expand=True)

    # SKU Section
    add_section(frame, "🍋 SKU Mapping 🍋", [
        ("Open 1.1", tp_tasks.step_1_1),
        ("Generate SKU Mapping", tp_tasks.step_1_2),
        ("Copy SKU File", tp_tasks.step_1_3),
    ])

    # TP Upload Section
    add_section(frame, "🍋 TP Upload 🍋", [
        ("Open 2.1", tp_tasks.step_2_1),
        ("Generate TP.csv", tp_tasks.step_2_2),
        ("Upload New TP", tp_tasks.step_2_3a),
        ("Upload Old TP", tp_tasks.step_2_3b),
        ("Copy TP.csv", tp_tasks.step_2_4),
        ("Open SKUINV", tp_tasks.step_2_5),
        ("Copy SKUINV File", tp_tasks.step_2_6),
    ])

    # DXM Section
    add_section(frame, "🍋 DXM 🍋", [
        ("Rename DXM", tp_tasks.step_3_1),
        ("Open 2.2", tp_tasks.step_3_2),
        ("Generate Inventory Update", tp_tasks.step_3_3),
        ("Copy Inventory Update", tp_tasks.step_3_4),
    ])

    root.mainloop()

if __name__ == "__main__":
    create_ui()
