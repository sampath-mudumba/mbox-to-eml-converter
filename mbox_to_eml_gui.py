#!/usr/bin/env python3
"""
Generic MBOX→EML Batch Import GUI
=================================

– Generic conversion: any .mbox file → .eml files  
– Batch import preparation: small batches avoid client limits  
– Handles Windows & cross-platform path issues  
– Detailed instructions & verification scripts  
– Simple GUI with Tkinter

Usage:
    python mbox_to_eml_gui_modified.py

Author: AI Assistant
Date: August 2025


Requirements for mbox_to_eml_gui_modified.py
============================================

Python 3.6 or newer

Standard library modules (no external dependencies):
– tkinter (GUI framework)
– mailbox (to read MBOX files)
– email and email.header (to parse and decode messages)
– email.generator.Generator (to write EML files)
– shutil (to copy files)
– pathlib (for cross-platform path handling)
– re (for filename sanitization)
– datetime (to timestamp instructions)

A graphical desktop environment (Windows, macOS, or Linux) with Tkinter installed and enabled

Sufficient file system permissions to read the source .mbox file and write to the target output directory

At least 100 MB of free disk space for intermediate and batch folders (depending on the total size of your emails)
"""

import os
import mailbox
import shutil
from email.generator import Generator
from email.header import decode_header
import email
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from datetime import datetime

class MboxToEmlBatchGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MBOX→EML Batch Import Tool")
        self.geometry("500x380")
        self.mbox_path = tk.StringVar()
        self.output_dir = tk.StringVar(value="eml_batches")
        self.batch_size = tk.IntVar(value=50)
        self.batch_mb = tk.IntVar(value=100)

        tk.Label(self, text="Select MBOX File:").pack(anchor="w", padx=10, pady=5)
        tk.Entry(self, textvariable=self.mbox_path, width=50).pack(padx=10)
        tk.Button(self, text="Browse…", command=self.browse_mbox).pack(padx=10, pady=5)

        tk.Label(self, text="Output Directory:").pack(anchor="w", padx=10, pady=5)
        tk.Entry(self, textvariable=self.output_dir, width=50).pack(padx=10)
        tk.Button(self, text="Browse…", command=self.browse_output).pack(padx=10, pady=5)

        tk.Label(self, text="Batch Size (msgs):").pack(anchor="w", padx=10, pady=5)
        tk.Entry(self, textvariable=self.batch_size, width=10).pack(padx=10)

        tk.Label(self, text="Max Batch Size (MB):").pack(anchor="w", padx=10, pady=5)
        tk.Entry(self, textvariable=self.batch_mb, width=10).pack(padx=10)

        tk.Button(self, text="Run Conversion", command=self.run).pack(pady=15)

        self.log = tk.Text(self, height=6)
        self.log.pack(fill="both", padx=10, pady=5)

    def browse_mbox(self):
        path = filedialog.askopenfilename(filetypes=[("MBOX files","*.mbox"),("All","*.*")])
        if path:
            self.mbox_path.set(path)

    def browse_output(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir.set(path)

    def log_msg(self, msg):
        self.log.insert("end", msg+"\n")
        self.log.see("end")
        self.update()

    def sanitize(self, text, max_len=50):
        text = re.sub(r'[<>:"/\\|?*]', '_', text)
        text = re.sub(r'\s+', ' ', text).strip()
        return text[:max_len] or "No_Subject"

    def decode_subject(self, raw_subj):
        try:
            parts = decode_header(raw_subj)
            decoded = ''.join(
                part.decode(enc or 'utf-8', errors='ignore') if isinstance(part, bytes) else part
                for part, enc in parts
            )
            return decoded
        except Exception:
            return "No_Subject"

    def run(self):
        mbox_file = Path(self.mbox_path.get())
        out_dir = Path(self.output_dir.get())
        bs = self.batch_size.get()
        mb = self.batch_mb.get()

        if not mbox_file.exists():
            messagebox.showerror("Error","MBOX file not found")
            return
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
        except:
            messagebox.showerror("Error","Cannot create output directory")
            return

        # Step 1: Convert .mbox → .eml
        mbox = mailbox.mbox(str(mbox_file))
        eml_dir = out_dir / "all_eml"
        eml_dir.mkdir(exist_ok=True)
        self.log_msg("Converting .mbox to .eml…")
        for i, msg in enumerate(mbox, 1):
            raw_subj = msg.get('Subject','No_Subject')
            subj = self.decode_subject(raw_subj)
            safe_subj = self.sanitize(subj)
            fn = f"{i:05d}_{safe_subj}.eml"
            path = eml_dir / fn
            with open(path, "w", encoding="utf-8", errors="replace") as f:
                Generator(f).flatten(msg)
            if i % 50 == 0:
                self.log_msg(f"  Converted {i} messages")
        total = i
        self.log_msg(f"✓ Converted {total} messages")

        # Step 2: Batch splitting
        self.log_msg("Creating import batches…")
        eml_files = sorted(eml_dir.glob("*.eml"), key=lambda p: p.stat().st_size)
        batches=[]
        cur,cur_size=[],0
        maxb=mb*1024*1024
        for f in eml_files:
            sz=f.stat().st_size
            if len(cur)>=bs or cur_size+sz>maxb:
                batches.append(cur); cur,cur_size=[],0
            cur.append(f); cur_size+=sz
        if cur: batches.append(cur)
        self.log_msg(f"→ {len(batches)} batches created")

        # Step 3: Copy to batch folders
        self.log_msg("Copying to batch folders…")
        batch_dirs=[]
        for idx,b in enumerate(batches,1):
            bd=out_dir/f"batch_{idx:03d}_{len(b)}msg"
            bd.mkdir(exist_ok=True)
            for e in b:
                shutil.copy2(str(e), str(bd/e.name))
            batch_dirs.append(str(bd))
            self.log_msg(f"  Batch {idx:03d}: {len(b)} messages → {bd.name}")

        # Step 4: Write instructions
        inst=out_dir/"IMPORT_INSTRUCTIONS.txt"
        with open(inst, "w", encoding="utf-8") as f:
            f.write(f"MBOX→EML Batch Import Instructions\n")
            f.write(f"Generated: {datetime.now()}\n\n")
            f.write(f"Total messages: {total}\n")
            f.write(f"Batches: {len(batch_dirs)}\n\n")
            f.write("Import one batch at a time:\n")
            for i,bd in enumerate(batch_dirs,1):
                f.write(f"Batch {i}: {Path(bd).name}\n")
        self.log_msg(f"✓ Instructions: {inst.name}")

        # Step 5: Verification script
        ver=out_dir/"verify_import_success.py"
        with open(ver, "w", encoding="utf-8") as f:
            f.write(f"""#!/usr/bin/env python3
\"\"\"Import Verification Script
Expected messages: {total}
\"\"\"
print("Expected messages: {total}")
print("-- Check your email client folder has this count --")
""")
        self.log_msg(f"✓ Verification: {ver.name}")

        messagebox.showinfo("Done","Conversion & batching complete!\nSee log for details.")

if __name__=="__main__":
    app = MboxToEmlBatchGUI()
    app.mainloop()