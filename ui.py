import tkinter as tk
from tkinter import messagebox
import tp_tasks


BG_COLOR      = "#FFFAEE"  
BTN_COLOR     = "#FFF4B2"  
BTN_HOVER     = "#B2DFF7"  
LABEL_BG      = "#CBE8F3"  
TEXT_COLOR    = "#020A1B"  
TITLE_COLOR   = "#055E61"  

def run_safe(func, name):
    try:
        func()
        messagebox.showinfo("Success", f"{name} completed ✅")
    except Exception as e:
        messagebox.showerror("Error", f"{name} failed ❌\n{e}")

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
    root.geometry("450x525")
    root.resizable(False, False)

    # ===== Header  =====
    header = tk.Label(root, text="Inventory update _ lemonnade🍹", font=("Segoe UI", 16, "bold"),
                      bg=BG_COLOR, fg=TITLE_COLOR)
    header.pack(pady=5)

    # ===== Main Frame =====
    frame = tk.Frame(root, bg=BG_COLOR)
    frame.pack(pady=10, fill="both", expand=True)

    # SKU Section
    add_section(frame, "🍋 SKU Mapping 🍋", [
        ("Open 1.1", tp_tasks.step_1_1),
        ("Generate & Copy SKU Mapping", tp_tasks.step_1_2),
    ])

    # TP Upload Section
    add_section(frame, "🍋 TP Upload 🍋", [
        ("Open 2.1", tp_tasks.step_2_1),
        ("Generate & Upload & Copy TP.csv", tp_tasks.step_2_2),
        ("Open SKUINV", tp_tasks.step_2_4),
        ("Copy SKUINV", tp_tasks.step_2_5),
    ])

    # DXM Section
    add_section(frame, "🍋 DXM 🍋", [
        ("Rename & Open", tp_tasks.step_3_1),
        ("Generate & Copy Inventory Update", tp_tasks.step_3_3),
    ])

    root.mainloop()

if __name__ == "__main__":
    create_ui()
