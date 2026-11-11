"""Small desktop window: choose (or drop) a SyncScript export, check the cover data and
characters, then save the Iyuno Word document."""
from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from . import __version__
from .convert import gender_code
from .naming import DEFAULT_TEMPLATE as NAME_DEFAULT, LABELS as NAME_LABELS, STYLES
from .pipeline import convert, convert_folder, default_cover, output_name
from .reader import Episode, read_export
from .settings import (DEFAULT_FIELDS, load_defaults, load_filename_template, save_defaults,
                       save_filename_template, settings_path)
from .writer import CoverInfo

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
        self._build_menu()

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

    # ------------------------------------------------------------ settings
    def _build_menu(self):
        menubar = tk.Menu(self.root)
        settings = tk.Menu(menubar, tearoff=0)
        settings.add_command(label="Cover page defaults…", command=lambda: self.edit_settings(0))
        settings.add_command(label="Document name…", command=lambda: self.edit_settings(1))
        menubar.add_cascade(label="Settings", menu=settings)
        self.root.config(menu=menubar)
        try:  # macOS: also under the app menu as Settings… (Cmd+,)
            self.root.createcommand("tk::mac::ShowPreferences", lambda: self.edit_settings(0))
        except tk.TclError:
            pass

    def _form_cover(self) -> CoverInfo:
        """The cover as the form shows it now (or an example when no file is open)."""
        if self.ep:
            cover = default_cover(self.ep)
            for key, _ in COVER_FIELDS:
                setattr(cover, key, self.vars[key].get().strip())
            cover.language = cover.language or "Dutch"
            return cover
        return CoverInfo(original_title="Puppy Pool Party", translated_title="Puppy Pool Party",
                         series_title="Minnie's Bow-Toons", translated_series_title="Minnie's Bow-Toons",
                         episode="102", translated_by="Mark Omvlee")

    def edit_settings(self, tab: int = 0):
        """Settings window: cover page defaults and the document name."""
        win = tk.Toplevel(self.root)
        win.title("Settings")
        win.transient(self.root)
        win.resizable(False, False)
        outer = ttk.Frame(win, padding=10)
        outer.pack(fill="both", expand=True)
        nb = ttk.Notebook(outer)
        nb.pack(fill="both", expand=True)

        # --- tab 1: cover page defaults
        frame = ttk.Frame(nb, padding=12)
        nb.add(frame, text="Cover page defaults")
        ttk.Label(
            frame,
            text="These values are filled in on every cover page. A filled-in default\n"
                 "replaces what the export says; leave a field empty to keep the export's value.\n"
                 "The SDI auth code is kept only on this computer.",
            justify="left",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        current = load_defaults()
        values: dict[str, tk.StringVar] = {}
        for i, (key, label) in enumerate(DEFAULT_FIELDS, start=1):
            ttk.Label(frame, text=label).grid(row=i, column=0, sticky="w", padx=4, pady=2)
            v = tk.StringVar(value=current.get(key, ""))
            ttk.Entry(frame, textvariable=v, width=48).grid(row=i, column=1, sticky="we", padx=4, pady=2)
            values[key] = v

        # --- tab 2: document name
        nframe = ttk.Frame(nb, padding=12)
        nb.add(nframe, text="Document name")
        ttk.Label(
            nframe,
            text="Build the file name from the cover fields. Choose a style, then click a field\n"
                 "to add it where the cursor is. Type anything else (such as _ or -) in between.",
            justify="left",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))
        tpl = tk.StringVar(value=load_filename_template())
        entry = ttk.Entry(nframe, textvariable=tpl, width=70)
        entry.grid(row=1, column=0, columnspan=3, sticky="we", pady=(0, 6))
        ttk.Label(nframe, text="Style").grid(row=2, column=0, sticky="w")
        style = ttk.Combobox(nframe, state="readonly", width=14, values=[name for _, name in STYLES])
        style.current(0)
        style.grid(row=2, column=1, sticky="w", pady=2)
        buttons_box = ttk.Frame(nframe)
        buttons_box.grid(row=3, column=0, columnspan=3, sticky="w", pady=4)

        def add_label(label: str):
            code = STYLES[style.current()][0]
            entry.insert("insert", f"[{label}:{code}]" if code else f"[{label}]")
            entry.focus()

        for n, (label, _) in enumerate(NAME_LABELS):
            ttk.Button(buttons_box, text=label.title(), command=lambda l=label: add_label(l)).grid(
                row=n // 3, column=n % 3, sticky="we", padx=2, pady=2)
        preview = ttk.Label(nframe, text="", foreground="#1d6b2f")
        ttk.Label(nframe, text="Example").grid(row=4, column=0, sticky="w", pady=(8, 0))
        preview.grid(row=4, column=1, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Button(nframe, text="Back to default",
                   command=lambda: tpl.set(NAME_DEFAULT)).grid(row=5, column=0, sticky="w", pady=(8, 0))

        def update_preview(*_):
            cover = self._form_cover()
            if values["auth_code"].get().strip():
                cover.auth_code = values["auth_code"].get().strip()
            preview.configure(text=output_name(cover, tpl.get()))

        tpl.trace_add("write", update_preview)
        update_preview()

        def save():
            data = {k: v.get() for k, v in values.items()}
            try:
                save_defaults(data)
                save_filename_template(tpl.get() or NAME_DEFAULT)
            except OSError as exc:
                messagebox.showerror("Cannot save settings", f"{settings_path()}\n\n{exc}", parent=win)
                return
            if self.ep:  # show the new defaults in the open script right away
                for key, value in data.items():
                    if value.strip():
                        self.vars[key].set(value.strip())
            win.destroy()

        buttons = ttk.Frame(outer)
        buttons.pack(fill="x", pady=(10, 0))
        ttk.Button(buttons, text="Cancel", command=win.destroy).pack(side="right")
        ttk.Button(buttons, text="Save", command=save).pack(side="right", padx=6)
        nb.select(tab)
        win.bind("<Escape>", lambda e: win.destroy())
        win.grab_set()

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
        cover = default_cover(ep, load_defaults())
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
        cover.language = cover.language or "Dutch"
        if not cover.auth_code and not messagebox.askyesno(
            "No SDI auth code", "The SDI auth code is empty. Iyuno's system needs it.\n\n"
            "Tip: enter it once under Settings › Cover page defaults.\n\nContinue anyway?"
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
            initialfile=output_name(cover, load_filename_template()),
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
    out_dir, report = convert_folder(folder, defaults=load_defaults(), template=load_filename_template())
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
