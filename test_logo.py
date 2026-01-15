*** DELETE FILE
import tkinter as tk
from PIL import Image, ImageTk
import os, sys

root = tk.Tk()
root.title("Test Logo")
root.geometry("400x400")

# Szukaj logo.png w katalogu EXE, potem w _MEIPASS
if getattr(sys, 'frozen', False):
    exe_dir = os.path.dirname(sys.executable)
    meipass = getattr(sys, '_MEIPASS', None)
    base_dirs = [exe_dir]
    if meipass and meipass != exe_dir:
        base_dirs.append(meipass)
else:
    base_dirs = [os.path.dirname(os.path.abspath(__file__))]

logo_path = None
for d in base_dirs:
    if d:
        p = os.path.join(d, "logo.png")
        if os.path.exists(p):
            logo_path = p
            break

if logo_path:
    img = Image.open(logo_path)
    photo = ImageTk.PhotoImage(img)
    label = tk.Label(root, image=photo)
    label.pack(expand=True)
else:
    label = tk.Label(root, text="Nie znaleziono logo.png", fg="red")
    label.pack(expand=True)

root.mainloop()
