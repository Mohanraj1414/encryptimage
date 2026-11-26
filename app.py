"""Local desktop workbench for reviewing KYC records."""
from datetime import datetime
import json
import os
import sys
from pathlib import Path
from queue import Empty, SimpleQueue
from threading import Thread
from zipfile import BadZipFile
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.exceptions import InvalidFileException
from PIL import Image, ImageTk

from decrypt_batch import save_image
from grid_records import MI_HEADERS, export_grid

from image_reader import configured_key, legacy_key, load_scan

from kyc import FIELDS, MISSING, export, read_sheet


def mi_settings_path():
    root = Path(os.environ.get("APPDATA") or Path.home())
    return root / "KYCWorkbench" / "config.json"


def read_mi_path():
    # A missing, malformed or unreadable config is equivalent to no saved path.
    try:
        config = json.loads(mi_settings_path().read_text(encoding="utf-8"))
        value = config.get("source_workbook") if isinstance(config, dict) else None
        return Path(value) if isinstance(value, str) and value.strip() else None
    except FileNotFoundError:
        # Read the previous release's saved selection until a JSON choice is made.
        legacy_root = Path(os.environ.get("LOCALAPPDATA") or Path.home())
        try:
            value = (legacy_root / "KYC Workbench" / "mi-workbook.txt").read_text(encoding="utf-8").strip()
            return Path(value) if value else None
        except (OSError, ValueError):
            return None
    except (OSError, ValueError):
        return None


def remember_mi_path(path):
    # Keep a working in-session selection even if the config write fails.
    config = mi_settings_path()
    config.parent.mkdir(parents=True, exist_ok=True)
    config.write_text(json.dumps({"source_workbook": str(path)}), encoding="utf-8")


def validate_mi_headers(path):
    if path.suffix.casefold() != ".xlsx" or path.name.startswith("~$"):
        raise ValueError("Choose an original .xlsx workbook. Excel temporary files are excluded.")
    if not path.is_file():
        raise ValueError("The workbook was moved, deleted or is unavailable.")
    # Excel's owner file can outlive the editing session. Use actual read
    # access and the same first-sheet schema as grid_records.load_mi.
    try:
        with path.open("rb") as source:
            workbook = load_workbook(source, read_only=True, data_only=True)
            try:
                if not workbook.worksheets:
                    raise ValueError("The workbook has no worksheets.")
                rows = workbook.worksheets[0].values
                headers = next(rows, ())[:len(MI_HEADERS)]
                normalized = tuple(int(value) if str(value).isdigit() else value for value in headers)
                if normalized != MI_HEADERS:
                    raise ValueError("The first sheet must start with ID and the numbered KYC MI columns.")
                if not any(row and row[0] for row in rows):
                    raise ValueError("The source worksheet contains no record IDs.")
            finally:
                workbook.close()
    except PermissionError as error:
        raise ValueError("The workbook is locked or unreadable. Close it in Excel and check file access.") from error
    except (BadZipFile, InvalidFileException, KeyError, SyntaxError) as error:
        raise ValueError("The workbook is damaged or is not a readable Excel .xlsx file.") from error


def app_mi_workbook_path():
    app_file = sys.executable if getattr(sys, "frozen", False) else __file__
    return Path(app_file).resolve().with_name("KYC_MI_Source.xlsx")


def find_mi_workbook(images, imported_path, session_choice=None):
    def is_valid(path):
        try:
            validate_mi_headers(path)
            return True
        except (OSError, ValueError):
            return False

    # An explicit menu choice is retained for this session if saving failed.
    saved_path = session_choice or read_mi_path()
    for path in (saved_path, imported_path, app_mi_workbook_path()):
        if path is not None and is_valid(path):
            return path

    candidates = find_batch_mi_workbooks(images)
    # Let the exporter compare printed IDs when several sources are available.
    return candidates[0] if len(candidates) == 1 else None


def find_batch_mi_workbooks(images):
    parents = {image.parent.resolve() for image in images}
    downloads = Path.home() / "Downloads"
    folders = parents | {parent.parent for parent in parents} | {downloads}
    candidates = set()
    for folder in folders:
        try:
            for path in folder.glob("*.xlsx"):
                if folder == downloads and "mi" not in path.stem.casefold():
                    continue
                try:
                    validate_mi_headers(path)
                    candidates.add(path.resolve())
                except (OSError, ValueError):
                    continue
        except OSError:
            continue
    return sorted(candidates)


def scan_folder(folder):
    return sorted(
        (path for path in folder.iterdir()
         if path.is_file() and path.suffix.lower() in {".jpg", ".jpeg", ".png"}),
        key=lambda path: (path.name.casefold(), path.name),
    )


def batch_targets(images, destination):
    folders = dict.fromkeys(source.parent for source in images)
    if len(folders) > 1:
        folders = {
            folder: destination / f"{number:03d}_{folder.name or 'images'}"
            for number, folder in enumerate(folders, 1)
        }
    else:
        folders = dict.fromkeys(folders, destination)
    return [folders[source.parent] / source.name for source in images]


class Workbench:
    def __init__(self, root):
        self.root = root
        self.records = []
        self.images = []
        self.index = 0
        self.dirty = False
        self.loading = False
        self.photo = None
        self.folder_keys = {}
        self.manual_key = None
        self.imported_workbook_path = None
        self.source_session_choice = None
        self.values = {field: tk.StringVar() for field in FIELDS}
        self.status = tk.StringVar()
        root.title("KYC Workbench")
        root.geometry("1200x810")
        root.protocol("WM_DELETE_WINDOW", lambda: self.discard() and root.destroy())
        menu = tk.Menu(root)
        file_menu = tk.Menu(menu, tearoff=False)
        file_menu.add_command(label="Decrypt Images...", command=self.decrypt_images)
        file_menu.add_command(label="Export reviewed form rows", command=self.save_reviewed_rows)
        file_menu.add_command(label="Export OCR review...", command=self.export_ocr_review)
        file_menu.add_command(label="Change source workbook...", command=self.choose_mi)
        menu.add_cascade(label="File", menu=file_menu)
        root.configure(menu=menu)
        toolbar = ttk.Frame(root, padding=10)
        toolbar.pack(fill="x")
        for label, action in (
            ("Open image", self.open_image), ("Open folder", self.open_images),
            ("Open multiple folders", self.open_multiple_folders),
            ("Set legacy key", self.set_key),
            ("Import Excel", self.import_excel),
            ("Previous", lambda: self.move(-1)), ("Next", lambda: self.move(1)),
            ("Go to row", self.go_to), ("Download all images", self.download_all_images),
            ("Export Excel", self.save),
        ):
            ttk.Button(toolbar, text=label, command=action).pack(side="left", padx=3)
        ttk.Label(root, textvariable=self.status, padding=(12, 6)).pack(fill="x")
        body = ttk.Panedwindow(root, orient="horizontal")
        body.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        image_panel = ttk.Frame(body, padding=8)
        body.add(image_panel, weight=3)
        navigator = ttk.Frame(body, padding=(4, 8))
        body.insert(0, navigator, weight=0)
        ttk.Label(navigator, text="Loaded scans / rows").pack(anchor="w", pady=(0, 5))
        list_frame = ttk.Frame(navigator)
        list_frame.pack(fill="both", expand=True)
        self.scan_list = tk.Listbox(list_frame, width=26, exportselection=False)
        list_scroll = ttk.Scrollbar(list_frame, command=self.scan_list.yview)
        self.scan_list.configure(yscrollcommand=list_scroll.set)
        list_scroll.pack(side="right", fill="y")
        self.scan_list.pack(side="left", fill="both", expand=True)
        self.scan_list.bind("<<ListboxSelect>>", self.select_scan)
        self.image_canvas = tk.Canvas(image_panel, width=540, height=450, highlightthickness=0)
        vertical = ttk.Scrollbar(image_panel, command=self.image_canvas.yview)
        horizontal = ttk.Scrollbar(image_panel, orient="horizontal", command=self.image_canvas.xview)
        self.image_canvas.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        vertical.pack(side="right", fill="y")
        horizontal.pack(side="bottom", fill="x")
        self.image_canvas.pack(fill="both", expand=True)
        self.image_label = ttk.Label(self.image_canvas)
        self.image_canvas.create_window(0, 0, window=self.image_label, anchor="nw")
        self.image_label.bind("<Configure>", lambda _: self.image_canvas.configure(
            scrollregion=self.image_canvas.bbox("all")
        ))
        for widget in (self.image_canvas, self.image_label):
            widget.bind("<MouseWheel>", self.scroll_preview)
        form = ttk.Frame(body, padding=8)
        body.add(form, weight=2)
        for row, field in enumerate(FIELDS):
            ttk.Label(form, text=field).grid(row=row, column=0, sticky="w", padx=(0, 10), pady=2)
            state = "readonly" if field == "Image Name" else "normal"
            ttk.Entry(form, textvariable=self.values[field], state=state).grid(row=row, column=1, sticky="ew")
            self.values[field].trace_add("write", self.changed)
        form.columnconfigure(1, weight=1)
        self.show()

    def changed(self, *_):
        if self.records and not self.loading:
            self.dirty = True

    def discard(self):
        return not self.dirty or messagebox.askyesno(
            "Unsaved edits", "Continue and discard changes since the last export?"
        )

    def select_scan(self, _event):
        selected = self.scan_list.curselection()
        if selected and selected[0] != self.index:
            self.move(selected[0] - self.index)

    def scroll_preview(self, event):
        units = max(1, abs(event.delta) // 120)
        self.image_canvas.yview_scroll(-units if event.delta > 0 else units, "units")
        return "break"

    def start(self, records, images):
        self.records, self.images = records, images
        self.index = 0
        self.dirty = False
        self.scan_list.delete(0, "end")
        labels = (
            [f"{number}. {path.parent.name}/{path.name}" for number, path in enumerate(images, 1)]
            if images else [f"{number}. {row.get('Record No') or 'Row'}"
                            for number, row in enumerate(records, 1)]
        )
        if labels:
            self.scan_list.insert("end", *labels)
        self.show()

    def open_image(self):
        filename = filedialog.askopenfilename(
            filetypes=[("KYC images", "*.jpg *.jpeg *.png"), ("All files", "*.*")]
        )
        if filename:
            path = Path(filename)
            self.open_scans([path])

    def open_images(self):
        directory = filedialog.askdirectory(parent=self.root)
        if directory:
            self.open_folders([Path(directory)])

    def open_multiple_folders(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Open multiple folders")
        dialog.transient(self.root)
        dialog.grab_set()
        panel = ttk.Frame(dialog, padding=12)
        panel.pack(fill="both", expand=True)
        ttk.Label(panel, text="Add each folder you want to process in one batch.").pack(anchor="w")
        folder_list = tk.Listbox(panel, width=78, height=9, selectmode="extended")
        folder_list.pack(fill="both", expand=True, pady=8)
        folders = []

        def add():
            directory = filedialog.askdirectory(parent=dialog, title="Add an image folder")
            if directory:
                folder = Path(directory).resolve()
                if folder not in folders:
                    folders.append(folder)
                    folder_list.insert("end", str(folder))

        def remove():
            for index in reversed(folder_list.curselection()):
                folder_list.delete(index)
                folders.pop(index)

        def open_selected():
            if not folders:
                return
            dialog.destroy()
            self.open_folders(folders)

        buttons = ttk.Frame(panel)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="Add folder", command=add).pack(side="left")
        ttk.Button(buttons, text="Remove selected", command=remove).pack(side="left", padx=6)
        ttk.Button(buttons, text="Open folders", command=open_selected).pack(side="right")

    def open_folders(self, folders):
        images = []
        issues = []
        for folder in folders:
            try:
                scans = scan_folder(folder)
            except OSError as error:
                issues.append(f"{folder}: {error}")
                continue
            if not scans:
                issues.append(f"{folder}: no JPG or PNG files")
            images.extend(scans)
        if issues:
            messagebox.showwarning("Open folders", "\n".join(issues))
        if images:
            self.open_scans(images)

    def open_scans(self, images):
        if self.discard():
            self.folder_keys = {}
            issues = []
            for directory in dict.fromkeys(path.parent for path in images):
                try:
                    self.folder_keys[directory] = configured_key(directory) if not self.manual_key else None
                except (OSError, ValueError) as error:
                    issues.append(f"{directory}: {error}")
            self.start(
                [{"Record No": path.stem, "Image Name": str(index)} for index, path in enumerate(images, 1)],
                images,
            )
            if issues:
                messagebox.showerror("Image key configuration", "\n".join(issues))

    def key_for(self, path):
        if self.manual_key:
            return self.manual_key
        if self.folder_keys.get(path.parent) is None:
            self.folder_keys[path.parent] = configured_key(path.parent)
        return self.folder_keys[path.parent]

    def decrypt_images(self):
        if not self.images:
            messagebox.showinfo("Decrypt Images", "Open an image or folder first.")
            return
        if len(self.images) == 1:
            filename = filedialog.asksaveasfilename(
                defaultextension=".jpg", initialfile=self.images[0].stem + "_decrypted.jpg",
                filetypes=[("JPEG", "*.jpg")],
            )
            if not filename:
                return
            targets = [Path(filename)]
        else:
            folder = filedialog.askdirectory(title="Choose where to save the decrypted JPEGs")
            if not folder:
                return
            targets = batch_targets(self.images, Path(folder))
        self.save_images(targets, title="Decrypt Images", formats=("JPEG",))

    def download_all_images(self):
        title = "Download all images"
        if not self.images:
            messagebox.showinfo(title, "Open an image or folder first.", parent=self.root)
            return
        folder = filedialog.askdirectory(title="Choose where to save all loaded images", parent=self.root)
        if folder:
            self.save_images(batch_targets(self.images, Path(folder)), title=title, formats=("JPEG", "PNG"))

    def save_images(self, targets, *, title, formats):
        saved = skipped = 0
        errors = []
        for source, target in zip(self.images, targets):
            try:
                target.parent.mkdir(parents=True, exist_ok=True)
                save_image(source, target, self.key_for(source), formats=formats)
                saved += 1
            except FileExistsError:
                skipped += 1
            except (OSError, ValueError, Image.DecompressionBombError) as error:
                errors.append(f"{source.name}: {error}")
        summary = f"Saved {saved}, skipped {skipped} existing files, failed {len(errors)}."
        if errors:
            messagebox.showwarning(title, summary + "\n\n" + "\n".join(errors[:5]), parent=self.root)
        else:
            messagebox.showinfo(title, summary, parent=self.root)

    def set_key(self):
        text = simpledialog.askstring(
            "Legacy image key", "Enter the complete legacy image key. Leave blank to use automatic detection.",
            show="*", parent=self.root,
        )
        if text is None:
            return
        try:
            self.manual_key = legacy_key(text) if text else None
            self.folder_keys.clear()
        except (OSError, ValueError) as error:
            messagebox.showerror("Legacy image key", str(error))
            return
        self.capture()
        self.show()

    def import_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel workbook", "*.xlsx")])
        if not path:
            return
        try:
            workbook = load_workbook(path, read_only=True, data_only=True)
            try:
                sheet_names = workbook.sheetnames
                sheet_name = simpledialog.askstring("Worksheet", "Worksheet name", initialvalue=sheet_names[0])
                if sheet_name is None:
                    return
                sheet = workbook[sheet_name]
                if sheet.max_row is None:
                    sheet.calculate_dimension(force=True)
                header_row = simpledialog.askinteger("Header row", "Row containing the column headers", initialvalue=1, minvalue=1, maxvalue=sheet.max_row)
                if header_row is None:
                    return
                headers = next(sheet.iter_rows(min_row=header_row, max_row=header_row, values_only=True))
            finally:
                workbook.close()
            mapping = self.choose_mapping(headers)
            if mapping is None or not self.discard():
                return
            records = read_sheet(path, sheet_name, mapping, header_row)
            if not records:
                messagebox.showinfo("Import Excel", "The worksheet has no records after the header.")
                return
            self.start(records, [])
            # Only a completed import becomes a candidate for this session.
            self.imported_workbook_path = Path(path).resolve()
        except Exception as error:
            messagebox.showerror("Import Excel", str(error))

    def choose_mapping(self, headers):
        dialog = tk.Toplevel(self.root)
        dialog.title("Map Excel columns")
        dialog.transient(self.root)
        dialog.grab_set()
        ttk.Label(dialog, text="Match each KYC field to a source column. Blank selections are imported as missing.").grid(
            row=0, column=0, columnspan=4, padx=12, pady=12
        )
        choices = [""] + [
            f"{get_column_letter(index)}  |  {str(header or '')[:80]}"
            for index, header in enumerate(headers, 1)
        ]
        widgets = {}
        result = None
        for index, field in enumerate(FIELDS):
            row, pair = index % 11 + 1, (index // 11) * 2
            ttk.Label(dialog, text=field).grid(row=row, column=pair, sticky="w", padx=10)
            combo = ttk.Combobox(dialog, values=choices, state="readonly", width=28)
            combo.grid(row=row, column=pair + 1, padx=8, pady=3)
            combo.current(next(
                (i for i, header in enumerate(headers, 1) if str(header).strip().casefold() == field.casefold()),
                0,
            ))
            widgets[field] = combo

        def accept():
            nonlocal result
            selected = {field: column_index_from_string(combo.get().split()[0]) - 1
                        for field, combo in widgets.items() if combo.get()}
            if not selected or len(set(selected.values())) != len(selected):
                messagebox.showerror("Column mapping", "Select at least one column and assign each column once.", parent=dialog)
                return
            result = selected
            dialog.destroy()

        ttk.Button(dialog, text="Import rows", command=accept).grid(row=12, column=3, pady=12)
        self.root.wait_window(dialog)
        return result

    def capture(self):
        if self.records:
            record = self.records[self.index]
            for field, variable in self.values.items():
                old = record.get(field)
                displayed = "" if old is None else str(old)
                if variable.get() != displayed:
                    record[field] = variable.get()

    def move(self, offset):
        if 0 <= self.index + offset < len(self.records):
            self.capture()
            self.index += offset
            self.show()

    def go_to(self):
        if self.records:
            target = simpledialog.askinteger("Go to row", "Row number", minvalue=1, maxvalue=len(self.records))
            if target is not None:
                self.move(target - 1 - self.index)

    def show(self):
        self.loading = True
        record = self.records[self.index] if self.records else {}
        for field, variable in self.values.items():
            value = record.get(field)
            variable.set("" if value is None else str(value))
        self.loading = False
        if self.records:
            self.scan_list.selection_clear(0, "end")
            self.scan_list.selection_set(self.index)
            self.scan_list.see(self.index)
        self.image_canvas.xview_moveto(0)
        self.image_canvas.yview_moveto(0)
        self.photo = None
        self.image_label.configure(image="", text="Open an image folder or import an Excel workbook.")
        self.status.set(f"Record {self.index + 1:,} of {len(self.records):,}" if self.records else "No records loaded")
        if self.images:
            path = self.images[self.index]
            self.status.set(f"{self.status.get()}  |  {path.parent.name} / {path.name}")
            try:
                with load_scan(path, self.key_for(path)) as image:
                    image.thumbnail((1800, 2400), Image.Resampling.LANCZOS)
                    self.photo = ImageTk.PhotoImage(image)
                self.image_label.configure(image=self.photo, text="")
            except (OSError, ValueError, Image.DecompressionBombError) as error:
                self.image_label.configure(text=f"{path.name}\n{error}", wraplength=480)
        elif self.records:
            self.image_label.configure(text="Excel import")

    def save(self):
        if self.images:
            self.save_grid()
        else:
            self.save_reviewed_rows()

    def save_reviewed_rows(self):
        if not self.records:
            messagebox.showinfo("Export Excel", "Load records before exporting.")
            return
        self.capture()
        if self.images and not any(
            str(record.get("Record No") or "").strip().casefold() not in MISSING | {image.stem.casefold()}
            or any(
                str(value if value is not None else "").strip().casefold() not in MISSING
                for field, value in record.items() if field not in {"Record No", "Image Name"}
            )
            for image, record in zip(self.images, self.records)
        ):
            self.save_grid()
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", initialfile="KYC_Output.xlsx", filetypes=[("Excel workbook", "*.xlsx")]
        )
        if not path:
            return
        try:
            export(path, self.records)
        except Exception as error:
            messagebox.showerror("Export Excel", str(error))
            return
        self.dirty = False
        messagebox.showinfo("Export Excel", f"Exported {len(self.records):,} records.")


    def choose_mi(self):
        downloads = Path.home() / "Downloads"
        initial_directory = downloads if downloads.is_dir() else Path.home()
        while True:
            filename = filedialog.askopenfilename(
                title="Select the original Client 2 KYC MI workbook",
                initialdir=initial_directory,
                filetypes=[("Excel workbook", "*.xlsx")],
                parent=self.root,
            )
            if not filename:
                return None
            path = Path(filename).resolve()
            try:
                validate_mi_headers(path)
            except (OSError, ValueError) as error:
                messagebox.showerror(
                    "KYC MI source workbook",
                    f"{path.name}\n\n{error}",
                    parent=self.root,
                )
                initial_directory = path.parent
            else:
                break

        try:
            remember_mi_path(path)
            self.source_session_choice = None
        except (OSError, ValueError):
            self.source_session_choice = path
            messagebox.showwarning(
                "KYC MI source workbook",
                "The workbook will be used for this session. Its location could not be saved.",
                parent=self.root,
            )
        return path

    def save_grid(self):
        mi_path = find_mi_workbook(self.images, self.imported_workbook_path, self.source_session_choice)
        if mi_path is None:
            self.export_scan_batch(None, recover=True)
        else:
            self.export_scan_batch(mi_path)

    def export_ocr_review(self):
        if not self.images:
            messagebox.showinfo("Export OCR review", "Open an image or folder first.", parent=self.root)
            return
        self.export_scan_batch(None)

    def export_scan_batch(self, mi_path, *, recover=False):
        ocr_only = mi_path is None and not recover
        title = "Export OCR review" if ocr_only else "Export Excel"
        try:
            destination = Path.home() / "Downloads" / "KYC Exports"
            destination.mkdir(parents=True, exist_ok=True)
        except (OSError, ValueError) as error:
            messagebox.showerror(title, str(error))
            return

        folders = {image.parent for image in self.images}
        batch_name = (self.images[0].parent.name or "Batch") if len(folders) == 1 else "Combined"
        prefix = "KYC" if mi_path is not None else "KYC_OCR"
        if recover:
            prefix = "KYC_Recovered"
        filename = destination / f"{prefix}_{batch_name}_{datetime.now():%Y%m%d_%H%M%S_%f}.xlsx"
        scans = []
        for image in self.images:
            try:
                scans.append((image, self.key_for(image)))
            except (OSError, ValueError) as error:
                scans.append((image, ValueError(str(error))))
        events = SimpleQueue()
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.protocol("WM_DELETE_WINDOW", lambda: None)
        batch_status = ttk.Label(dialog, text=f"Preparing all {len(scans):,} loaded scans", padding=12)
        batch_status.pack()
        progress = ttk.Progressbar(dialog, maximum=len(scans), length=340)
        progress.pack(padx=12, pady=(0, 12))

        def run():
            try:
                alternatives = []
                if not ocr_only:
                    possible = dict.fromkeys([
                        self.imported_workbook_path, app_mi_workbook_path(), *find_batch_mi_workbooks(self.images)
                    ])
                    alternatives = [path for path in possible if path is not None and path != mi_path]
                result = export_grid(
                    filename, scans, mi_path, lambda count: events.put(("progress", count)),
                    recover=recover, alternatives=alternatives, require_matches=not ocr_only,
                    page_progress=None if ocr_only else lambda count: events.put(("pages", count)),
                )
                events.put(("done", result))
            except Exception as error:
                events.put(("error", str(error)))

        def poll():
            while True:
                try:
                    kind, value = events.get_nowait()
                except Empty:
                    self.root.after(100, poll)
                    return
                if kind in {"progress", "pages"}:
                    progress.configure(value=value)
                    if kind == "progress":
                        status = f"Read record IDs on {value:,} of {len(scans):,} scans"
                        if value == len(scans) and not ocr_only:
                            status = "Checking MI sources and starting scan recovery"
                    else:
                        status = f"Prepared Excel rows for {value:,} of {len(scans):,} scans"
                    batch_status.configure(text=status)
                    continue
                dialog.destroy()
                if kind == "error":
                    messagebox.showerror(title, value)
                else:
                    rows, flagged = value
                    row_label = "OCR-detected record IDs" if ocr_only else "KYC rows with data"
                    review_note = ("See Scan review for the OCR results." if ocr_only else
                                   "Check Scan review and any KYC Data rows marked REVIEW against the images.")
                    messagebox.showinfo(
                        title,
                        f"Included all {len(scans):,} loaded scans in Scan review and saved {rows:,} {row_label}. "
                        f"{flagged:,} scans need review.\n\nSaved to {filename}\n\n{review_note}",
                    )
                return

        Thread(target=run, daemon=True).start()
        self.root.after(100, poll)


if __name__ == "__main__":
    window = tk.Tk()
    Workbench(window)
    window.mainloop()
