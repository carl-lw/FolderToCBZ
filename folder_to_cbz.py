#!/usr/bin/env python3

""" Folder to CBZ/CBR/PDF converter with GUI. """

import os
import re
import sys
import shutil
import zipfile
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Optional drag & drop
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except Exception:
    DND_AVAILABLE = False
    from tkinter import Tk as TkinterDnD

# img2pdf for PDF
try:
    import img2pdf
    IMG2PDF_AVAILABLE = True
except Exception:
    IMG2PDF_AVAILABLE = False

selected_folders = []

# ---------- Utilities ----------

def normalize_folder_name(folder_path):
    if not folder_path:
        return ""
    folder_path = folder_path.strip().strip('"')
    folder_path = os.path.normpath(folder_path)
    folder_name = os.path.basename(folder_path)
    return folder_name.strip()

def drop_event(event):
    paths = root.splitlist(event.data)
    for path in paths:
        if os.path.isdir(path) and path not in selected_folders:
            selected_folders.append(path)
            folder_list.insert(tk.END, path)

digit_re = re.compile(r'(\d+)')

def center_window(parent, window):
    window.update_idletasks()  # Ensure geometry is ready

    # Parent geometry
    px = parent.winfo_x()
    py = parent.winfo_y()
    pw = parent.winfo_width()
    ph = parent.winfo_height()

    # Window geometry
    ww = window.winfo_width()
    wh = window.winfo_height()

    # Compute centered position
    x = px + (pw // 2) - (ww // 2)
    y = py + (ph // 2) - (wh // 2)

    window.geometry(f"+{x}+{y}")

def natural_sort_key(s):
    parts = digit_re.split(s)
    key = []
    for p in parts:
        if p.isdigit():
            key.append(int(p))
        else:
            key.append(p.lower())
    return key

def find_rar_executable():
    candidates = ['rar', 'rar.exe', 'rar5', 'unrar']
    for name in candidates:
        path = shutil.which(name)
        if path:
            return path

    common_dirs = [
        r"C:\Program Files\WinRAR",
        r"C:\Program Files (x86)\WinRAR",
    ]

    for folder in common_dirs:
        for name in candidates:
            full = os.path.join(folder, name)
            if os.path.isfile(full):
                return full
    
    return None

RAR_PATH = find_rar_executable()

# ---------- Renaming / mapping ----------

def rename_folder(folder_name, regex_find=None, regex_replace="", prefix="", suffix="", literal=False):
    orig = normalize_folder_name(folder_name)
    new_name = orig
    if regex_find:
        try:
            if literal:
                regex_find = re.escape(regex_find)
            new_name = re.sub(regex_find, regex_replace, new_name, flags=re.UNICODE)
        except re.error as e:
            print(f"[rename] Invalid regex '{regex_find}': {e}", file=sys.stderr)
    return f"{prefix}{new_name}{suffix}"

def build_renamed_map(selected, regex_find, regex_replace, prefix, suffix, literal=False):
    out = []
    for folder in selected:
        orig = normalize_folder_name(folder)
        new = rename_folder(orig, regex_find=regex_find if regex_find else None,
                            regex_replace=regex_replace, prefix=prefix, suffix=suffix,
                            literal=literal)
        out.append((folder, orig, new))
    return out

# ---------- Compression / export helpers ----------

def write_cbz_fast(folder_path, output_cbz_path, compress=True):
    try:
        folder_path = os.path.normpath(folder_path)
        if os.path.exists(output_cbz_path):
            os.remove(output_cbz_path)
        compression = zipfile.ZIP_DEFLATED if compress else zipfile.ZIP_STORED
        all_files = []
        for root, _, files in os.walk(folder_path):
            for f in files:
                all_files.append(os.path.join(root, f))
        with zipfile.ZipFile(output_cbz_path, "w", compression=compression) as zf:
            for fpath in all_files:
                arcname = os.path.relpath(fpath, folder_path)
                zf.write(fpath, arcname)
        return True
    except Exception as e:
        print(f"[write_cbz_fast] Exception: {e}", file=sys.stderr)
        return False

def create_cbr_with_rar(rar_path, folder_path, output_cbr_path):
    try:
        folder_path_norm = os.path.normpath(folder_path)
        parent = os.path.dirname(folder_path_norm) or "."
        output_name = os.path.basename(output_cbr_path)
        folder_basename = os.path.basename(folder_path_norm)
        cmd = [rar_path, 'a', '-r', output_name, folder_basename]
        proc = subprocess.run(cmd, cwd=parent, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if proc.returncode == 0:
            return True
        else:
            print(f"[create_cbr] rar failed: {proc.returncode}\nSTDOUT:{proc.stdout}\nSTDERR:{proc.stderr}", file=sys.stderr)
            return False
    except Exception as e:
        print(f"[create_cbr] Exception: {e}", file=sys.stderr)
        return False

def make_pdf_from_images(folder_path, output_pdf_path):
    if not IMG2PDF_AVAILABLE:
        return False, "img2pdf not installed"
    exts = ('.jpg', '.jpeg', '.png', '.tif', '.tiff', '.webp', '.bmp')
    images = []
    for root_dir, _, files in os.walk(folder_path):
        for fname in files:
            if fname.lower().endswith(exts):
                images.append(os.path.join(root_dir, fname))
    if not images:
        return False, "No image files found"
    images.sort(key=lambda p: natural_sort_key(os.path.relpath(p, folder_path)))
    try:
        with open(output_pdf_path, "wb") as f:
            f.write(img2pdf.convert(images))
        return True, None
    except Exception as e:
        print(f"[make_pdf] Exception: {e}", file=sys.stderr)
        return False, str(e)

# ---------- GUI & threading worker ----------

def remove_folder_from_list(path):
    norm = os.path.normpath(path)
    for i in range(folder_list.size()):
        if os.path.normpath(folder_list.get(i)) == norm:
            folder_list.delete(i)
            break
    for j, f in enumerate(list(selected_folders)):
        if os.path.normpath(f) == norm:
            selected_folders.pop(j)
            break

def open_multi_folder_dialog():
    dlg = tk.Toplevel(root)
    dlg.title("Select Folders")
    dlg.geometry("600x400")
    dlg.transient(root)
    dlg.grab_set()
    
    tree = ttk.Treeview(dlg)
    tree.pack(fill="both", expand=True)

    center_window(root, dlg)

    # Populate tree with directories
    def insert_node(parent, path):
        try:
            for entry in sorted(os.listdir(path)):
                full = os.path.join(path, entry)
                if os.path.isdir(full):
                    node = tree.insert(parent, "end", text=entry, values=[full])
                    tree.insert(node, "end", text="", values=["dummy"])  # dummy child
        except PermissionError:
            pass

    # Fill root dirs
    drives = [os.path.abspath(os.sep)] if os.name != 'nt' else [f"{d}:/" for d in "ABCDEFGHIJKLMNOPQRSTUVWXYZ" if os.path.exists(f"{d}:/")]
    for d in drives:
        node = tree.insert("", "end", text=d, values=[d])
        tree.insert(node, "end", text="", values=["dummy"])

    # Expand dynamically
    def on_open(event):
        node = tree.focus()
        children = tree.get_children(node)
        if children and tree.item(children[0], "values")[0] == "dummy":
            tree.delete(children[0])
            path = tree.item(node, "values")[0]
            insert_node(node, path)

    tree.bind("<<TreeviewOpen>>", on_open)
    tree.config(selectmode="extended")  # allow multi-select

    selected = []
    def on_ok():
        for node in tree.selection():
            path = tree.item(node, "values")[0]
            if os.path.isdir(path):
                selected.append(path)
        dlg.destroy()
        for folder in selected:
            norm = os.path.normpath(folder)
            if norm not in selected_folders:
                selected_folders.append(norm)
                folder_list.insert(tk.END, norm)

    btn_frame = tk.Frame(dlg)
    btn_frame.pack(fill="x")
    tk.Button(btn_frame, text="OK", command=on_ok).pack(side="right", padx=4, pady=4)
    tk.Button(btn_frame, text="Cancel", command=dlg.destroy).pack(side="right", padx=4, pady=4)

def on_add_folder():
    open_multi_folder_dialog()

def on_remove_selected():
    sel = list(folder_list.curselection())
    if not sel:
        return
    for idx in reversed(sel):
        try:
            val = folder_list.get(idx)
            norm = os.path.normpath(val)
            for i, f in enumerate(selected_folders):
                if os.path.normpath(f) == norm:
                    selected_folders.pop(i)
                    break
            folder_list.delete(idx)
        except Exception:
            pass

def on_clear_list():
    selected_folders.clear()
    folder_list.delete(0, tk.END)

def open_options():
    dlg = tk.Toplevel(root)
    dlg.title("Options")
    dlg.geometry("360x250")
    dlg.resizable(False, False)
    dlg.transient(root)
    dlg.grab_set()

    center_window(root, dlg)

    tk.Label(dlg, text="Output format:", font=("Arial", 11)).pack(anchor="w", padx=12, pady=(10,2))
    formats_frame = tk.Frame(dlg)
    formats_frame.pack(anchor="w", padx=12)

    def make_radio(text, value):
        rb = tk.Radiobutton(formats_frame, text=text, variable=output_format_var, value=value)
        rb.pack(anchor="w", pady=2)
        return rb

    rb_cbz = make_radio("CBZ (zip, .cbz)", "cbz")
    rb_cbr = make_radio("CBR (RAR, .cbr)", "cbr")
    rb_pdf = make_radio("PDF (images → PDF)", "pdf")

    if not RAR_PATH:
        rb_cbr.config(state="disabled")
        if output_format_var.get() == "cbr":
            output_format_var.set("cbz")
    if not IMG2PDF_AVAILABLE:
        rb_pdf.config(state="disabled")
        if output_format_var.get() == "pdf":
            output_format_var.set("cbz")

    tk.Label(dlg, text=" ", height=1).pack()
    chk_delete = tk.Checkbutton(dlg, text="Delete source folder after successful conversion", variable=delete_after_var)
    chk_delete.pack(anchor="w", padx=12, pady=4)

    chk_literal = tk.Checkbutton(dlg, text="Treat Regex Find as literal string", variable=literal_regex_var)
    chk_literal.pack(anchor="w", padx=12, pady=4)

    btn_frame = tk.Frame(dlg)
    btn_frame.pack(fill="x", pady=12)
    tk.Button(btn_frame, text="OK", command=dlg.destroy, width=10).pack(side="right", padx=8)
    tk.Button(btn_frame, text="Cancel", command=dlg.destroy, width=10).pack(side="right")

def test_preview():
    mapping = build_renamed_map(selected_folders, find_entry.get(), replace_entry.get(),
                                prefix_entry.get(), suffix_entry.get(), literal=literal_regex_var.get())
    if not mapping:
        messagebox.showinfo("Preview", "No folders selected or nothing to preview.")
        return
    lines = [f"{orig}  →  {new}" for _, orig, new in mapping]
    preview = "\n".join(lines)
    print("=== Preview ===")
    print(preview)
    messagebox.showinfo("Preview renaming", preview)

def compress_worker():
    mapping = build_renamed_map(selected_folders, find_entry.get(), replace_entry.get(),
                                prefix_entry.get(), suffix_entry.get(), literal=literal_regex_var.get())
    if not mapping:
        root.after(0, lambda: messagebox.showwarning("Warning", "No folders to process."))
        root.after(0, finish_processing)
        return

    total = len(mapping)
    root.after(0, lambda: progress.config(maximum=total, value=0))
    fmt = output_format_var.get()
    delete_after = delete_after_var.get()

    for idx, (folder_path, orig, new_name) in enumerate(mapping, start=1):
        parent = os.path.dirname(os.path.normpath(folder_path)) or os.getcwd()
        cbz_path = os.path.join(parent, f"{new_name}.cbz")
        cbr_path = os.path.join(parent, f"{new_name}.cbr")
        pdf_path = os.path.join(parent, f"{new_name}.pdf")

        root.after(0, lambda n=new_name: status_label.config(text=f"Processing: {n}"))

        success = False
        err_msg = None

        if fmt == "cbz":
            success = write_cbz_fast(folder_path, cbz_path)
            if not success:
                err_msg = "Failed to create CBZ"
        elif fmt == "cbr":
            if RAR_PATH:
                success = create_cbr_with_rar(RAR_PATH, folder_path, cbr_path)
                if not success:
                    root.after(0, lambda: messagebox.showwarning(
                        "RAR failed",
                        f"Failed to create RAR using {RAR_PATH} for '{orig}'. Falling back to CBZ."
                    ))
                    success = write_cbz_fast(folder_path, cbz_path)
                    if not success:
                        err_msg = "Failed to create CBR (and fallback CBZ failed)"
            else:
                root.after(0, lambda: messagebox.showwarning(
                    "RAR not available",
                    f"RAR binary not found. Creating CBZ instead for '{orig}'."
                ))
                success = write_cbz_fast(folder_path, cbz_path)
                if not success:
                    err_msg = "Failed to create CBZ"
        elif fmt == "pdf":
            ok, reason = make_pdf_from_images(folder_path, pdf_path)
            success = ok
            if not ok:
                err_msg = reason or "PDF creation failed"

        if success and delete_after:
            try:
                shutil.rmtree(folder_path)
                root.after(0, lambda path=folder_path: remove_folder_from_list(path))
            except Exception as e:
                print(f"[delete] Failed to delete {folder_path}: {e}", file=sys.stderr)
                root.after(0, lambda: messagebox.showwarning("Delete failed", f"Could not delete {folder_path}: {e}"))

        if not success:
            msg = f"Failed: {orig}\nReason: {err_msg or 'unknown'}"
            root.after(0, lambda m=msg: messagebox.showwarning("Conversion error", m))

        root.after(0, lambda val=idx: progress.config(value=val))

    root.after(0, finish_processing)

def finish_processing():
    convert_btn.config(state="normal")
    add_plus_btn.config(state="normal")
    remove_btn.config(state="normal")
    clear_btn.config(state="normal")
    options_btn.config(state="normal")
    test_btn.config(state="normal")
    progress["value"] = 0
    status_label.config(text="✅ Finished!")
    messagebox.showinfo("Completed", "All folders processed successfully!")

def start_conversion():
    if not selected_folders:
        messagebox.showwarning("Warning", "No folders selected.")
        return
    convert_btn.config(state="disabled")
    add_plus_btn.config(state="disabled")
    remove_btn.config(state="disabled")
    clear_btn.config(state="disabled")
    options_btn.config(state="disabled")
    test_btn.config(state="disabled")
    status_label.config(text="Starting...")
    threading.Thread(target=compress_worker, daemon=True).start()

# ---------- GUI-builder ----------

root = TkinterDnD.Tk() if DND_AVAILABLE else TkinterDnD()
root.title("FolderToCBZ(CBR/PDF)")
root.geometry("780x680")

tk.Label(root, text="Drag folders here or use the + button to add", font=("Arial", 12)).pack(pady=6)

folder_list = tk.Listbox(root, width=110, height=14, selectmode=tk.EXTENDED)
folder_list.pack(pady=6)

if DND_AVAILABLE:
    folder_list.drop_target_register(DND_FILES)
    folder_list.dnd_bind("<<Drop>>", lambda e: (drop_event(e) if DND_AVAILABLE else None))

btn_frame = tk.Frame(root)
btn_frame.pack(pady=6)

add_plus_btn = tk.Button(btn_frame, text="+", width=3, font=("Arial", 12, "bold"), command=on_add_folder)
add_plus_btn.grid(row=0, column=0, padx=6)

remove_btn = tk.Button(btn_frame, text="−", width=3, font=("Arial", 12, "bold"), command=on_remove_selected)
remove_btn.grid(row=0, column=1, padx=6)

clear_btn = tk.Button(btn_frame, text="Clear List", command=on_clear_list)
clear_btn.grid(row=0, column=2, padx=10)

options_btn = tk.Button(btn_frame, text="Options", command=open_options)
options_btn.grid(row=0, column=3, padx=10)

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

output_format_var = tk.StringVar(value="cbz")
delete_after_var = tk.BooleanVar(value=False)
literal_regex_var = tk.BooleanVar(value=False)

test_btn = tk.Button(root, text="Preview renaming", command=test_preview)
test_btn.pack(pady=6)

convert_btn = tk.Button(root, text="Convert", command=start_conversion, font=("Arial", 12))
convert_btn.pack(pady=12)

progress = ttk.Progressbar(root, orient="horizontal", length=720, mode="determinate")
progress.pack(pady=6)

status_label = tk.Label(root, text="", font=("Arial", 10))
status_label.pack(pady=6)

if not DND_AVAILABLE:
    print("[startup] tkinterdnd2 not installed; drag & drop disabled.")
if not IMG2PDF_AVAILABLE:
    print("[startup] img2pdf not installed; PDF export disabled.")
if not RAR_PATH:
    print("[startup] 'rar' not found on PATH; CBR export disabled.")

root.mainloop()
