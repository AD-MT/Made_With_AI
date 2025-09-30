import os
import shutil
import re
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext

def parse_pasted_list(pasted_text):
    if not pasted_text.strip():
        return []
    text = pasted_text.replace('\n', ',').replace('\r', ',')
    if '"' in text or "'" in text:
        pattern = r'(?:"([^"]*)")|(?:\'([^\']*)\')|([^,\s][^,]*[^,\s])'
        items = []
        for match in re.finditer(pattern, text):
            item = next((g for g in match.groups() if g is not None), "").strip()
            if item:
                items.append(item)
    else:
        items = [item.strip() for item in text.split(',') if item.strip()]
    return items

def bulk_copy_files(source_dir, dest_dir, phrases, logger, mode="folders"):
    dumped_files = set()
    for phrase in phrases:
        file_found = False
        for root, dirs, files in os.walk(source_dir):
            for filename in files:
                if phrase in filename:
                    file_found = True
                    source_file_path = os.path.join(root, filename)
                    if mode == "folders":
                        # Copy to all matching destination folders
                        for dest_root, dest_dirs, dest_files in os.walk(dest_dir):
                            for dest_folder in dest_dirs:
                                if phrase in dest_folder:
                                    dest_folder_path = os.path.join(dest_root, dest_folder)
                                    dest_file_path = os.path.join(dest_folder_path, filename)
                                    try:
                                        shutil.copy(source_file_path, dest_file_path)
                                        logger(f"Copied '{filename}' to '{dest_folder_path}'")
                                    except Exception as e:
                                        logger(f"Failed to copy '{filename}': {e}")
                                    break
                    elif mode == "dump":
                        dest_file_path = os.path.join(dest_dir, filename)
                        if dest_file_path in dumped_files:
                            continue
                        try:
                            shutil.copy(source_file_path, dest_file_path)
                            logger(f"Copied '{filename}' to '{dest_dir}'")
                            dumped_files.add(dest_file_path)
                        except Exception as e:
                            logger(f"Failed to copy '{filename}': {e}")
        if not file_found:
            logger(f"No file found containing the phrase '{phrase}'")

def rename_files_in_folder(folder_path, logger):
    for root, dirs, files in os.walk(folder_path):
        files.sort()
        files = [file for file in files if not file.endswith('.ini')]
        current_folder_name = os.path.basename(root)
        for index, file_name in enumerate(files, start=1):
            old_file_path = os.path.join(root, file_name)
            new_file_name = f"{current_folder_name}_{index}{os.path.splitext(file_name)[1]}"
            new_file_path = os.path.join(root, new_file_name)
            if os.path.exists(new_file_path):
                logger(f"File {new_file_name} already exists. Skipping...")
                continue
            os.rename(old_file_path, new_file_path)
            logger(f"Renamed '{old_file_path}' to '{new_file_path}'")

def create_folders(path, folders, logger):
    os.makedirs(path, exist_ok=True)
    for folder in folders:
        try:
            os.makedirs(os.path.join(path, folder), exist_ok=True)
            logger(f"Created folder: {folder}")
        except Exception as e:
            logger(f"Failed to create {folder}: {e}")
    logger(f'All {len(folders)} folders processed.')

def tag_folders_doc_check(base_path, logger):
    logger("Document check script started")
    if not os.path.exists(base_path): logger(f"Path does not exist: {base_path}"); return
    for commodity_category in os.listdir(base_path):
        commodity_category_path = os.path.join(base_path, commodity_category)
        if not os.path.isdir(commodity_category_path): continue
        for part_number in os.listdir(commodity_category_path):
            part_number_path = os.path.join(commodity_category_path, part_number)
            if not os.path.isdir(part_number_path): continue
            base_part_number = part_number.split('_')[0]
            contains_docs = any(
                fn.lower().endswith(('.pdf', '.doc', '.docx'))
                for fn in os.listdir(part_number_path)
            )
            suffix = '_r' if contains_docs else '_nfy'
            new_folder_name = base_part_number + suffix
            if new_folder_name == part_number: logger(f"No renaming needed for {part_number_path}"); continue
            new_folder_path = os.path.join(commodity_category_path, new_folder_name)
            if os.path.exists(new_folder_path): logger(f"Target exists, skipping: {new_folder_path}"); continue
            try:
                os.rename(part_number_path, new_folder_path)
                logger(f"Renamed {part_number_path} -> {new_folder_path}")
            except Exception as e:
                logger(f"Failed to rename {part_number_path}: {e}")

def tag_folders_image_check(base_path, logger):
    logger("Image check script started")
    if not os.path.exists(base_path): logger(f"Path does not exist: {base_path}"); return
    try:
        for commodity_category in os.listdir(base_path):
            commodity_category_path = os.path.join(base_path, commodity_category)
            if os.path.isdir(commodity_category_path):
                for part_number in os.listdir(commodity_category_path):
                    part_number_path = os.path.join(commodity_category_path, part_number)
                    if os.path.isdir(part_number_path):
                        contains_images = any(
                            file_name.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff')) 
                            for file_name in os.listdir(part_number_path))
                        if contains_images:
                            new_folder_name = part_number + '_r'
                        else:
                            new_folder_name = part_number + '_nfy'
                        new_folder_path = os.path.join(commodity_category_path, new_folder_name)
                        try:
                            os.rename(part_number_path, new_folder_path)
                            logger(f"Renamed {part_number_path} to {new_folder_path}")
                        except Exception as e:
                            logger(f"Failed to rename {part_number_path}: {e}")
    except Exception as e:
        logger(f"An error occurred: {e}")

class App:
    def __init__(self, root):
        self.root = root
        root.title("File and Folder Management Tool")
        root.geometry('700x600')

        self.option_var = tk.StringVar()
        self.options = [
            "File Hunter - Copy files to matching folders",
            "File Renamer - Rename files based on folder name",
            "Folder Generator - Create multiple folders",
            "Folder Marker (Document Check) - Tag folders based on document presence",
            "Folder Marker (Image Check) - Tag folders based on image presence",
        ]

        # --- STATUS LABEL ---
        self.status_var = tk.StringVar()
        self.status_var.set("Ready.")
        self.status_label = tk.Label(
            root,
            textvariable=self.status_var,
            fg="#ba9800",
            font=("Segoe UI", 11, "italic")
        )
        self.status_label.pack(anchor=tk.W, padx=10, pady=(5, 0))

        instruction_label = tk.Label(
            root,
            text="Click here to select tool:",
            fg='#115099',
            font=("Segoe UI", 11, 'bold')
        )
        instruction_label.pack(anchor=tk.W, padx=10, pady=(8, 0))

        self.option_var.set("-- Click here to select tool --")
        option_menu = tk.OptionMenu(
            root,
            self.option_var,
            *self.options,
            command=self.show_frame
        )
        option_menu.config(
            width=58,
            font=("Segoe UI", 10, "bold"),
            bg="#e0f0ff",
            activebackground="#cde6fa",
            highlightthickness=2,
            relief=tk.RAISED,
            borderwidth=2
        )
        option_menu["menu"].config(
            font=("Segoe UI", 10),
            bg="#f8fcfe", activebackground="#e0f0ff"
        )
        option_menu.pack(fill=tk.X, padx=8, pady=(0, 8))

        self.frames = [
            self._build_filehunter(),
            self._build_renamer(),
            self._build_foldergen(),
            self._build_doctagger(),
            self._build_imagetagger()
        ]
        for frame in self.frames:
            frame.pack_forget()

        self.logbox = scrolledtext.ScrolledText(
            root, height=14, font=("Consolas", 10), background="#f1f1f1", relief=tk.SUNKEN
        )
        self.logbox.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        self.logbox.config(state='disabled')

        tk.Label(root, text="--- anegrete @ OutlierAI ---", fg="#666").pack(side=tk.BOTTOM, pady=2)

    def run_in_thread(self, btn, func):
        self.status_var.set("Working... Please be patient.")
        btn.config(state=tk.DISABLED)
        def thread_target():
            try:
                func()
            finally:
                self.root.after(0, lambda: [
                    self.status_var.set("Done!"),
                    btn.config(state=tk.NORMAL)
                ])
        threading.Thread(target=thread_target, daemon=True).start()

    def log(self, txt):
        self.logbox.config(state='normal')
        self.logbox.insert(tk.END, txt+'\n')
        self.logbox.see(tk.END)
        self.logbox.config(state='disabled')

    def clearlog(self):
        self.logbox.config(state='normal')
        self.logbox.delete(1.0, tk.END)
        self.logbox.config(state='disabled')

    def hide_all(self):
        for frame in self.frames: frame.pack_forget()

    def show_frame(self, option):
        self.clearlog()
        self.status_var.set("Ready.")
        self.hide_all()
        try:
            idx = self.options.index(option)
            self.frames[idx].pack(fill=tk.X, pady=6, padx=10)
        except Exception:
            pass

    def style_button(self, b):
        b.config(
            bg="#0066cc", fg="white", activebackground="#004080", activeforeground="white",
            font=("Segoe UI", 10, "bold"), relief=tk.RAISED, borderwidth=3, highlightthickness=0, padx=7, pady=1, cursor="hand2"
        )
        def on_enter(e): b.config(bg="#004080")
        def on_leave(e): b.config(bg="#0066cc")
        b.bind("<Enter>", on_enter)
        b.bind("<Leave>", on_leave)

    def _build_filehunter(self):
        f = tk.Frame(self.root, bd=2, relief=tk.RIDGE)
        src, dst = tk.StringVar(), tk.StringVar()
        copy_mode = tk.StringVar(value="folders")

        tk.Label(f, text="Source Directory:", font=("Segoe UI", 10, "bold")).grid(row=0,column=0,sticky=tk.W, padx=2, pady=2)
        tk.Entry(f, textvariable=src, width=50).grid(row=0,column=1, padx=2)
        b_src = tk.Button(f, text="...", command=lambda: src.set(filedialog.askdirectory()))
        b_src.grid(row=0,column=2)
        self.style_button(b_src)

        tk.Label(f, text="Destination Directory:", font=("Segoe UI", 10, "bold")).grid(row=1,column=0,sticky=tk.W, padx=2, pady=2)
        tk.Entry(f, textvariable=dst, width=50).grid(row=1,column=1, padx=2)
        b_dst = tk.Button(f, text="...", command=lambda: dst.set(filedialog.askdirectory()))
        b_dst.grid(row=1,column=2)
        self.style_button(b_dst)

        mode_frame = tk.Frame(f)
        tk.Label(mode_frame, text="Mode:", font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT)
        rb1 = tk.Radiobutton(mode_frame, text="Copy to matching subfolders (default)", variable=copy_mode, value="folders", font=("Segoe UI", 10))
        rb2 = tk.Radiobutton(mode_frame, text="Dump all files directly into destination folder", variable=copy_mode, value="dump", font=("Segoe UI", 10))
        rb1.pack(side=tk.LEFT, padx=4)
        rb2.pack(side=tk.LEFT, padx=4)
        mode_frame.grid(row=2,column=0,columnspan=3,sticky=tk.W,pady=5)

        tk.Label(f, text="Phrases (paste from Excel, comma or line-separated):", font=("Segoe UI", 10, "bold")).grid(row=3,column=0,sticky=tk.W, pady=(10,0))
        tbox = tk.Text(f, width=60, height=3)
        tbox.grid(row=4,column=0,columnspan=3)

        br = tk.Button(f, text="Run")
        def run():
            self.clearlog()
            phrases_list = parse_pasted_list(tbox.get(1.0,tk.END))
            mode_val = copy_mode.get()
            def task():
                self.log(f"Processing {len(phrases_list)} phrases in mode: {mode_val}")
                bulk_copy_files(src.get(), dst.get(), phrases_list, self.log, mode_val)
                self.log("Done.")
            self.run_in_thread(br, task)
        br.config(command=run)
        br.grid(row=5, column=2, sticky=tk.E, pady=10)
        self.style_button(br)
        return f

    def _build_renamer(self):
        f = tk.Frame(self.root, bd=2, relief=tk.RIDGE)
        folder = tk.StringVar()
        tk.Label(f, text="Folder Path:", font=("Segoe UI", 10, "bold")).grid(row=0,column=0,sticky=tk.W, padx=2, pady=2)
        tk.Entry(f, textvariable=folder, width=52).grid(row=0,column=1)
        b_folder = tk.Button(f, text="...", command=lambda: folder.set(filedialog.askdirectory()))
        b_folder.grid(row=0,column=2)
        self.style_button(b_folder)

        br = tk.Button(f, text="Run")
        def run():
            self.clearlog()
            def task():
                rename_files_in_folder(folder.get(), self.log)
                self.log("Done.")
            self.run_in_thread(br, task)
        br.config(command=run)
        br.grid(row=1,column=2,pady=7,sticky=tk.E)
        self.style_button(br)
        return f

    def _build_foldergen(self):
        f = tk.Frame(self.root, bd=2, relief=tk.RIDGE)
        path = tk.StringVar()
        tk.Label(f, text="Parent Directory:", font=("Segoe UI", 10, "bold")).grid(row=0,column=0,sticky=tk.W, padx=2, pady=2)
        tk.Entry(f, textvariable=path, width=52).grid(row=0,column=1)
        b_path = tk.Button(f, text="...", command=lambda: path.set(filedialog.askdirectory()))
        b_path.grid(row=0,column=2)
        self.style_button(b_path)

        tk.Label(f, text="Folder names (comma or line separated):", font=("Segoe UI", 10, "bold")).grid(row=1, column=0, sticky=tk.W, pady=(10,0))
        tbox = tk.Text(f, width=56, height=3)
        tbox.grid(row=2,column=0,columnspan=3)

        br = tk.Button(f, text="Run")
        def run():
            self.clearlog()
            names = parse_pasted_list(tbox.get(1.0,tk.END))
            def task():
                create_folders(path.get(), names, self.log)
                self.log("Done.")
            self.run_in_thread(br, task)
        br.config(command=run)
        br.grid(row=3,column=2,pady=7,sticky=tk.E)
        self.style_button(br)
        return f

    def _build_doctagger(self):
        f = tk.Frame(self.root, bd=2, relief=tk.RIDGE)
        folder = tk.StringVar()
        tk.Label(f, text="Commodity Library Base Path:", font=("Segoe UI", 10, "bold")).grid(row=0,column=0,sticky=tk.W, padx=2, pady=2)
        tk.Entry(f, textvariable=folder, width=52).grid(row=0,column=1)
        b_folder = tk.Button(f, text="...", command=lambda: folder.set(filedialog.askdirectory()))
        b_folder.grid(row=0,column=2)
        self.style_button(b_folder)

        br = tk.Button(f, text="Run")
        def run():
            self.clearlog()
            def task():
                tag_folders_doc_check(folder.get(), self.log)
                self.log("Done.")
            self.run_in_thread(br, task)
        br.config(command=run)
        br.grid(row=1,column=2,pady=7,sticky=tk.E)
        self.style_button(br)
        return f

    def _build_imagetagger(self):
        f = tk.Frame(self.root, bd=2, relief=tk.RIDGE)
        folder = tk.StringVar()
        tk.Label(f, text="Commodity Library Base Path:", font=("Segoe UI", 10, "bold")).grid(row=0,column=0,sticky=tk.W, padx=2, pady=2)
        tk.Entry(f, textvariable=folder, width=52).grid(row=0,column=1)
        b_folder = tk.Button(f, text="...", command=lambda: folder.set(filedialog.askdirectory()))
        b_folder.grid(row=0,column=2)
        self.style_button(b_folder)

        br = tk.Button(f, text="Run")
        def run():
            self.clearlog()
            def task():
                tag_folders_image_check(folder.get(), self.log)
                self.log("Done.")
            self.run_in_thread(br, task)
        br.config(command=run)
        br.grid(row=1,column=2,pady=7,sticky=tk.E)
        self.style_button(br)
        return f

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()