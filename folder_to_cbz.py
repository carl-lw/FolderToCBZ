import os
import zipfile
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD

selected_folders = []

# -------------------- Utility Functions --------------------

def normalize_folder_name(folder_path):    
    
    """Returns the clean basename of a folder"""

    folder_path = folder_path.strip('"')
    folder_path = os.path.normpath(folder_path)
    folder_name = os.path.basename(folder_path)
    return folder_name.strip()

def rename_folder(folder_name, regex_find=None, regex_replace="", prefix="", suffix=""):
    
    """Implement renaming and use of regex"""

    folder_name = normalize_folder_name(folder_name)
    
    if regex_find:
        try:
            folder_name = re.sub(regex_find, regex_replace, folder_name)
        except re.error as e:
            print(f"Invalid regex: {e}")
    
    folder_name = f"{prefix}{folder_name}{suffix}"
    return folder_name

# -------------------- GUI Functionality --------------------

def add_folders():
    folder = filedialog.askdirectory()
    if folder and folder not in selected_folders:
        selected_folders.append(folder)
        folder_list.insert(tk.END, folder)

def drop_event(event):
    paths = root.splitlist(event.data)
    for path in paths:
        if os.path.isdir(path) and path not in selected_folders:
            selected_folders.append(path)
            folder_list.insert(tk.END, path)

def clear_folders():
    selected_folders.clear()
    folder_list.delete(0, tk.END)

def build_renamed_map():

    """Return list of tuples: (folder_path, original_name, new_name)"""

    regex_find = find_entry.get()
    regex_replace = replace_entry.get()
    prefix = prefix_entry.get()
    suffix = suffix_entry.get()

    renamed = []
    for folder in selected_folders:
        original_name = normalize_folder_name(folder)
        new_name = rename_folder(
            original_name,
            regex_find=regex_find if regex_find else None,
            regex_replace=regex_replace,
            prefix=prefix,
            suffix=suffix
        )
        renamed.append((folder, original_name, new_name))
    return renamed

def test_rename():
    renamed = build_renamed_map()
    if not renamed:
        return
    preview = "\n".join([f"{orig}  →  {new}" for _, orig, new in renamed])
    print("=== Preview ===")
    print(preview)
    messagebox.showinfo("Preview renaming", preview)

def compress_worker():
    renamed = build_renamed_map()
    if not renamed:
        finish_processing()
        return

    progress["maximum"] = len(renamed)
    progress["value"] = 0

    for folder_path, orig, new_name in renamed:
        parent = os.path.dirname(os.path.normpath(folder_path))
        zip_path = os.path.join(parent, f"{new_name}.zip")
        cbz_path = os.path.join(parent, f"{new_name}.cbz")

        print(f"Compressing: {orig}  →  {new_name}")

        file_count = sum(len(files) for _, _, files in os.walk(folder_path))
        processed = 0

        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
            for root_dir, _, files in os.walk(folder_path):
                for file in files:
                    filepath = os.path.join(root_dir, file)
                    arcname = os.path.relpath(filepath, folder_path)
                    zipf.write(filepath, arcname)
                    processed += 1
                    status_label.config(text=f"{new_name}: {processed}/{file_count} files")
                    root.update_idletasks()

        if os.path.exists(cbz_path):
            os.remove(cbz_path)
        os.rename(zip_path, cbz_path)
        progress.step()
        root.update_idletasks()

    finish_processing()

def finish_processing():
    convert_btn.config(state="normal")
    add_btn.config(state="normal")
    clear_btn.config(state="normal")
    test_btn.config(state="normal")
    progress["value"] = 0
    status_label.config(text="✅ Done!")
    messagebox.showinfo("Complete", "All folders processed.")

def compress_to_cbz():
    if not selected_folders:
        messagebox.showwarning("Warning", "No folders selected.")
        return

    convert_btn.config(state="disabled")
    add_btn.config(state="disabled")
    clear_btn.config(state="disabled")
    test_btn.config(state="disabled")

    status_label.config(text="Starting...")
    threading.Thread(target=compress_worker, daemon=True).start()

# -------------------- GUI --------------------

root = TkinterDnD.Tk()
root.title("Folder to CBZ")
root.geometry("750x600")

tk.Label(root, text="Drag folders here or click Add Folder", font=("Arial", 12)).pack(pady=5)

folder_list = tk.Listbox(root, width=100, height=12)
folder_list.pack(pady=5)
folder_list.drop_target_register(DND_FILES)
folder_list.dnd_bind("<<Drop>>", drop_event)

btn_frame = tk.Frame(root)
btn_frame.pack(pady=8)

add_btn = tk.Button(btn_frame, text="Add Folder", command=add_folders)
add_btn.grid(row=0, column=0, padx=6)

clear_btn = tk.Button(btn_frame, text="Clear List", command=clear_folders)
clear_btn.grid(row=0, column=1, padx=6)

regex_frame = tk.Frame(root)
regex_frame.pack(pady=10)

tk.Label(regex_frame, text="Regex Find:").grid(row=0, column=0, padx=6, sticky="e")
find_entry = tk.Entry(regex_frame, width=50)
find_entry.grid(row=0, column=1, padx=6)

tk.Label(regex_frame, text="Replace With:").grid(row=1, column=0, padx=6, sticky="e")
replace_entry = tk.Entry(regex_frame, width=50)
replace_entry.grid(row=1, column=1, padx=6)

tk.Label(regex_frame, text="Prefix:").grid(row=2, column=0, padx=6, sticky="e")
prefix_entry = tk.Entry(regex_frame, width=50)
prefix_entry.grid(row=2, column=1, padx=6)

tk.Label(regex_frame, text="Suffix:").grid(row=3, column=0, padx=6, sticky="e")
suffix_entry = tk.Entry(regex_frame, width=50)
suffix_entry.grid(row=3, column=1, padx=6)

test_btn = tk.Button(root, text="Preview renaming", command=test_rename)
test_btn.pack(pady=6)

convert_btn = tk.Button(root, text="Convert to .cbz", command=compress_to_cbz, font=("Arial", 12))
convert_btn.pack(pady=12)

progress = ttk.Progressbar(root, orient="horizontal", length=700, mode="determinate")
progress.pack(pady=6)

status_label = tk.Label(root, text="", font=("Arial", 10))
status_label.pack(pady=6)

root.mainloop()
