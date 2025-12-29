# 測試腳本 - 捕獲錯誤
import sys
import os
import traceback

APP_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(APP_DIR)
sys.path.insert(0, APP_DIR)

log_file = os.path.join(APP_DIR, "startup_error.log")

try:
    from main_modern import main
    main()
except Exception as e:
    with open(log_file, "w", encoding="utf-8") as f:
        f.write(f"Error: {e}\n\n")
        f.write(traceback.format_exc())
    
    import tkinter as tk
    from tkinter import messagebox
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("啟動錯誤", f"錯誤已記錄到:\n{log_file}\n\n{e}")
    root.destroy()
