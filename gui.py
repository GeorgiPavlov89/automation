# gui.py
import subprocess, sys, os, signal, tkinter as tk
from tkinter import messagebox

worker_proc = None

def resolve_worker_path():
    base = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(__file__)
    exe_path = os.path.join(base, "Worker.exe")
    if os.path.exists(exe_path):
        return [exe_path]
    # dev режим – стартираме Python скрипта
    return [sys.executable, os.path.join(base, "main.py")]

def start_worker():
    global worker_proc
    if worker_proc and worker_proc.poll() is None:
        messagebox.showinfo("Info", "Автоматизацията вече работи.")
        return
    try:
        worker_proc = subprocess.Popen(resolve_worker_path(), cwd=os.path.dirname(os.path.abspath(__file__)))
        status.set(f"Статус: RUNNING (PID {worker_proc.pid})")
    except Exception as e:
        messagebox.showerror("Грешка", f"Не мога да стартирам: {e}")

def stop_worker():
    global worker_proc
    if not worker_proc or worker_proc.poll() is not None:
        messagebox.showinfo("Info", "Няма стартирана автоматизация.")
        return
    try:
        worker_proc.terminate()
        try:
            worker_proc.wait(timeout=10)
        except Exception:
            if os.name == "nt":
                subprocess.call(["taskkill", "/F", "/T", "/PID", str(worker_proc.pid)])
            else:
                os.kill(worker_proc.pid, signal.SIGKILL)
        status.set("Статус: STOPPED")
        worker_proc = None
    except Exception as e:
        messagebox.showerror("Грешка", f"Не мога да спра: {e}")

def on_close():
    try: 
        if worker_proc and worker_proc.poll() is None:
            stop_worker()
    finally:
        root.destroy()

root = tk.Tk()
root.title("Automation Control")
root.geometry("320x160")
status = tk.StringVar(value="Статус: STOPPED")

tk.Label(root, textvariable=status, font=("Segoe UI", 11)).pack(pady=10)
tk.Button(root, text="Start", width=14, command=start_worker).pack(pady=6)
tk.Button(root, text="Stop", width=14, command=stop_worker).pack(pady=2)
root.protocol("WM_DELETE_WINDOW", on_close)
root.mainloop()
