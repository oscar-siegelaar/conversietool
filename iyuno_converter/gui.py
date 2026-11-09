"""Small desktop window: choose (or drop) a SyncScript export, check the cover data and
characters, then save the Iyuno Word document."""
from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from . import __version__
from .convert import gender_code
from .pipeline import convert, convert_folder, default_cover, default_output_name
from .reader import Episode, read_export

COVER_FIELDS = [
    ("original_title", "Original title"),
    ("translated_title", "Translated title"),
    ("series_title", "Series title"),
    ("translated_series_title", "Translated series title"),
    ("season", "Season #"),
    ("episode", "Episode #"),
    ("language", "Language"),
    ("date", "Date (empty = today)"),
    ("translated_by", "Translated by"),
    ("auth_code", "SDI auth code (from Iyuno)"),
]
GENDERS = ("M", "V", "MV", "")


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.src: Path | None = None
        self.ep: Episode | None = None
        root.title(f"SyncScript → Iyuno Word  v{__version__}")
        root.geometry("900x680")

        top = ttk.Frame(root, padding=10)
        top.pack(fill="x")
        self.drop = ttk.Label(
            top, text="Drop a SyncScript export (.xlsx) here, or click to choose one",
            anchor="center", relief="groove", padding=18,
        )
        self.drop.pack(fill="x")
        self.drop.bind("<Button-1>", lambda e: self.choose())
        self._enable_drop()

        body = ttk.Frame(root, padding=(10, 0))
        body.pack(fill="both", expand=True)

        form = ttk.LabelFrame(body, text="Cover page", padding=8)
        form.pack(fill="x", pady=(8, 4))
        self.vars: dict[str, tk.StringVar] = {}
        for i, (key, label) in enumerate(COVER_FIELDS):
            ttk.Label(form, text=label).grid(row=i // 2, column=(i % 2) * 2, sticky="w", padx=4, pady=2)
            v = tk.StringVar()
            ttk.Entry(form, textvariable=v, width=38).grid(row=i // 2, column=(i % 2) * 2 + 1, sticky="we", padx=4, pady=2)
            self.vars[key] = v
        form.columnconfigure(1, weight=1)
        form.columnconfigure(3, weight=1)

        chars = ttk.LabelFrame(body, text="Characters – double-click to edit original name or gender", padding=8)
        chars.pack(fill="both", expand=True, pady=4)
        self.tree = ttk.Treeview(chars, columns=("name", "original", "gender"), show="headings", height=12)
        for col, title, w in (("name", "Script name", 260), ("original", "Original name", 360), ("gender", "Gender", 80)):
            self.tree.heading(col, text=title)
            self.tree.column(col, width=w, anchor="w")
        sb = ttk.Scrollbar(chars, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self.edit_cell)

        bottom = ttk.Frame(root, padding=10)
        bottom.pack(fill="x")
        self.status = ttk.Label(bottom, text="")
        self.status.pack(side="left")
        self.go = ttk.Button(bottom, text="Make Word document…", command=self.save, state="disabled")
        self.go.pack(side="right")
        ttk.Button(bottom, text="Convert whole folder…", command=self.batch).pack(side="right", padx=8)
        self.clear_btn = ttk.Button(bottom, text="Cancel / other file", command=self.reset, state="disabled")
        self.clear_btn.pack(side="right")

    # ------------------------------------------------------------ loading
    DROP_TEXT = "Drop a SyncScript export (.xlsx) here, or click to choose one"

    def reset(self):
        """Forget the current export so another one can be loaded."""
        self.src = self.ep = None
        self.drop.configure(text=self.DROP_TEXT)
        for v in self.vars.values():
            v.set("")
        self.tree.delete(*self.tree.get_children())
        self.go.configure(state="disabled")
        self.clear_btn.configure(state="disabled")
        self.status.configure(text="")

    def _enable_drop(self):
        try:
            from tkinterdnd2 import DND_FILES  # optional
            self.drop.drop_target_register(DND_FILES)
            self.drop.dnd_bind("<<Drop>>", lambda e: self.load(self.root.tk.splitlist(e.data)[0]))
        except Exception:
            pass

    def choose(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if path:
            self.load(path)

    def load(self, path: str):
        try:
            ep = read_export(path)
        except Exception as exc:  # show any read problem to the user
            messagebox.showerror("Cannot read file", str(exc))
            return
        self.src, self.ep = Path(path), ep
        self.drop.configure(text=f"{self.src.name}  –  {len(ep.rows)} rows   (click to choose another file)")
        cover = default_cover(ep)
        for key, _ in COVER_FIELDS:
            self.vars[key].set(getattr(cover, key))
        self.tree.delete(*self.tree.get_children())
        for name, info in sorted(ep.characters.items(), key=lambda kv: kv[0].lower()):
            if name.lower().startswith("insert"):
                continue
            self.tree.insert("", "end", iid=name, values=(name, info.original_name or name, gender_code(info.gender)))
        self.go.configure(state="normal")
        self.clear_btn.configure(state="normal")
        self.status.configure(text="Check the cover data, the SDI auth code and the genders.")

    # ------------------------------------------------------------ editing
    def edit_cell(self, event):
        item, col = self.tree.identify_row(event.y), self.tree.identify_column(event.x)
        if not item or col == "#1":
            return
        x, y, w, h = self.tree.bbox(item, col)
        idx = int(col[1:]) - 1
        value = self.tree.item(item, "values")[idx]
        if col == "#3":
            widget = ttk.Combobox(self.tree, values=GENDERS, state="readonly")
        else:
            widget = ttk.Entry(self.tree)
        widget.place(x=x, y=y, width=w, height=h)
        widget.insert(0, value) if col != "#3" else widget.set(value)
        widget.focus()

        def done(_=None):
            vals = list(self.tree.item(item, "values"))
            vals[idx] = widget.get()
            self.tree.item(item, values=vals)
            widget.destroy()

        widget.bind("<Return>", done)
        widget.bind("<FocusOut>", done)
        if col == "#3":
            widget.bind("<<ComboboxSelected>>", done)

    # ------------------------------------------------------------ saving
    def save(self):
        assert self.ep and self.src
        cover = default_cover(self.ep)
        for key, _ in COVER_FIELDS:
            setattr(cover, key, self.vars[key].get().strip())
        if not cover.auth_code and not messagebox.askyesno(
            "No SDI auth code", "The SDI auth code is empty. Iyuno's macro and system need it. Continue anyway?"
        ):
            return
        for iid in self.tree.get_children():
            name, original, gender = self.tree.item(iid, "values")
            info = self.ep.characters[name]
            info.original_name = original
            info.gender = gender
        out = filedialog.asksaveasfilename(
            defaultextension=".docx",
            initialdir=str(self.src.parent),
            initialfile=default_output_name(self.ep, cover.language or "Dutch"),
            filetypes=[("Word", "*.docx")],
        )
        if not out:
            return
        try:
            path, warnings = convert(self.src, out, cover, self.ep)
        except Exception as exc:
            messagebox.showerror("Conversion failed", str(exc))
            return
        msg = f"Saved:\n{path}"
        if warnings:
            msg += "\n\nNote:\n- " + "\n- ".join(warnings)
        messagebox.showinfo("Done", msg)
        self.reset()


def _batch(app: "App"):
    folder = filedialog.askdirectory(title="Folder with SyncScript exports (.xlsx)")
    if not folder:
        return
    app.status.configure(text="Converting…")
    app.root.update_idletasks()
    out_dir, report = convert_folder(folder)
    ok = sum(1 for r in report if r.startswith("OK"))
    app.status.configure(text=f"{ok} of {len(report)} converted")
    messagebox.showinfo(
        "Folder converted",
        f"{ok} of {len(report)} files converted.\n\nOutput and rapport.txt are in:\n{out_dir}",
    )


App.batch = _batch


def run():
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except Exception:
        root = tk.Tk()
    App(root)
    root.mainloop()
