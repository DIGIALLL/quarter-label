"""Quarter Label — print shipping labels (Orlen Paczka, DPD, InPost, DHL...) on a quarter of A4.

Run:  pythonw label_print.py [label.pdf]

- Finds the label on the PDF page by its content, crops it and rotates it text-up.
- Mode A4: places it into one of 4 A6 cells (2x2) of an A4 sheet and remembers used cells.
- Mode A6: prints it centred on an A6 piece (A4 cut into quarters).
- Prints through Windows GDI at exact 1:1 scale, with X/Y offset calibration.
"""
import json
import locale
import sys
import time
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pymupdf
from PIL import Image, ImageGrab, ImageTk, ImageWin
import win32con
import win32gui
import win32print
import win32ui

APP_NAME = "Quarter Label"
APP = Path(__file__).resolve().parent
CFG = APP / "config.json"
STATE = APP / "state.json"
OUT = APP / "output"
DOWNLOADS = Path.home() / "Downloads"

MM = 72 / 25.4
A4_W, A4_H = 210.0, 297.0
COLS, ROWS = 2, 2
CELL_W, CELL_H = A4_W / COLS, A4_H / ROWS     # 105 x 148.5 mm
PAD_MM = 1.5                                   # margin around detected content
FIT_MM = 5.0                                   # margin inside the cell when a label is enlarged
PV_W = 380                                     # preview width, px
PAPER = {"A4": (9, A4_W, A4_H), "A6": (70, 105.0, 148.0)}   # DMPAPER code, width, height (mm)
ARROW_BOX = (102.5, 238.0, 5.0, 50.0)          # mm: x, y, w, h — gap between columns, bottom of sheet
IMG_EXT = (".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".tif", ".tiff")
FILETYPES = [("PDF / images", "*.pdf " + " ".join("*" + e for e in IMG_EXT)), ("PDF", "*.pdf")]
FONT = "C:/Windows/Fonts/arial.ttf"            # Cyrillic-capable font for printed hints

# ---------- translations ----------

T = {
    "ru": {
        "title": "Этикетки на четверть A4",
        "open": "Открыть PDF…",
        "latest": "Последний PDF из Загрузок",
        "no_file": "файл не выбран",
        "paper": "Бумага в принтере:",
        "mode_a4": "Целый лист A4 (4 места)",
        "mode_a6": "Четвертинка A6 (105×148)",
        "position": "Позиция на листе A4:",
        "pos": {1: "1 — верх-лево", 2: "2 — верх-право", 3: "3 — низ-лево", 4: "4 — низ-право"},
        "new_sheet": "Новый лист (всё свободно)",
        "flip": "Повернуть на 180°",
        "printer": "Принтер:",
        "off_x": "Сдвиг X, мм (+вправо)",
        "off_y": "Сдвиг Y, мм (+вниз)",
        "calib": "Калибровка (печать рамок)",
        "arrow_chk": "Стрелка «в лоток» на новом листе A4",
        "print": "ПЕЧАТЬ",
        "print_all": "ПЕЧАТЬ ВСЕХ ({n}) НА ЛИСТ",
        "add": "Добавить PDF в очередь…",
        "paste": "Вставить скрин из буфера (Ctrl+V)",
        "no_clip_t": "Буфер обмена",
        "no_clip": "В буфере обмена нет изображения",
        "crop": "Выбрать область…",
        "crop_t": "Область печати",
        "crop_hint": "Обведите мышью нужную область этикетки",
        "crop_queue": "Нажмите на этикетку в предпросмотре листа — откроется выбор области",
        "ok": "Готово",
        "auto": "Авто (вся этикетка)",
        "cancel": "Отмена",
        "cropped": " (своя область)",
        "rotate90": "Повернуть на 90°",
        "rot_lbl": "Поворот: {r}",
        "rot_auto": "авто",
        "fit": "Растянуть на всю ячейку",
        "result": "Так будет напечатано:",
        "clear": "Очистить очередь",
        "queued": "в очереди: {n}",
        "batch_more": "Свободных мест: {f}, этикеток: {n}. Напечатаются первые {f}.",
        "batch_t": "Пакетная печать",
        "language": "Язык:",
        "info": "Этикетка {w:.0f}×{h:.0f} мм, масштаб {s:.0f}%, поворот {r}°",
        "page": ", стр. {n}/{t}",
        "error": "Ошибка",
        "used": "занято",
        "a6_hint": "кусок A6 — в лоток узкой стороной вперёд",
        "no_pdfs_t": "Нет файлов",
        "no_pdfs": "В Загрузках нет PDF",
        "open_err": "Не открыть PDF:\n{e}",
        "calib_t": "Калибровка",
        "calib_q": "Положите обычную бумагу ({m}). Будут напечатаны рамки в 5 мм от краёв каждой "
                   "наклейки. Наложите на лист с наклейками на просвет и подберите сдвиг X/Y.",
        "print_err": "Ошибка печати",
        "nofile_t": "Нет файла",
        "nofile": "Сначала откройте PDF этикетки",
        "busy_t": "Занято",
        "busy": "Позиция {p} отмечена как использованная. Всё равно печатать?",
        "sheet_done_t": "Лист закончился",
        "sheet_done": "Все 4 места использованы. Следующая печать — на новый лист.",
        "too_big": "Этикетка больше ячейки A6",
        "arrow_txt": "этот край первым в лоток",
        "calib_txt": "рамка 5 мм от краёв наклейки",
    },
    "en": {
        "title": "Shipping labels on a quarter of A4",
        "open": "Open PDF…",
        "latest": "Latest PDF from Downloads",
        "no_file": "no file selected",
        "paper": "Paper in printer:",
        "mode_a4": "Full A4 sheet (4 slots)",
        "mode_a6": "A6 quarter (105×148)",
        "position": "Slot on the A4 sheet:",
        "pos": {1: "1 — top left", 2: "2 — top right", 3: "3 — bottom left", 4: "4 — bottom right"},
        "new_sheet": "New sheet (all free)",
        "flip": "Rotate 180°",
        "printer": "Printer:",
        "off_x": "Offset X, mm (+right)",
        "off_y": "Offset Y, mm (+down)",
        "calib": "Calibration (print frames)",
        "arrow_chk": "“Feed edge” arrow on a new A4 sheet",
        "print": "PRINT",
        "print_all": "PRINT ALL ({n}) ON SHEET",
        "add": "Add PDF to queue…",
        "paste": "Paste screenshot (Ctrl+V)",
        "no_clip_t": "Clipboard",
        "no_clip": "No image in the clipboard",
        "crop": "Select area…",
        "crop_t": "Print area",
        "crop_hint": "Drag the mouse around the label area",
        "crop_queue": "Click a label on the sheet preview to select its area",
        "ok": "Done",
        "auto": "Auto (whole label)",
        "cancel": "Cancel",
        "cropped": " (custom area)",
        "rotate90": "Rotate 90°",
        "rot_lbl": "Rotation: {r}",
        "rot_auto": "auto",
        "fit": "Fill the whole cell",
        "result": "This is what will print:",
        "clear": "Clear queue",
        "queued": "queued: {n}",
        "batch_more": "Free slots: {f}, labels: {n}. The first {f} will be printed.",
        "batch_t": "Batch print",
        "language": "Language:",
        "info": "Label {w:.0f}×{h:.0f} mm, scale {s:.0f}%, rotation {r}°",
        "page": ", page {n}/{t}",
        "error": "Error",
        "used": "used",
        "a6_hint": "A6 piece — feed short edge first",
        "no_pdfs_t": "No files",
        "no_pdfs": "No PDF files in Downloads",
        "open_err": "Cannot open PDF:\n{e}",
        "calib_t": "Calibration",
        "calib_q": "Load plain paper ({m}). Frames will be printed 5 mm inside each label slot. "
                   "Hold it against a label sheet in front of a light and adjust offset X/Y.",
        "print_err": "Print error",
        "nofile_t": "No file",
        "nofile": "Open a label PDF first",
        "busy_t": "Slot used",
        "busy": "Slot {p} is marked as used. Print anyway?",
        "sheet_done_t": "Sheet finished",
        "sheet_done": "All 4 slots are used. Next print goes on a new sheet.",
        "too_big": "Label is larger than an A6 cell",
        "arrow_txt": "feed this edge first",
        "calib_txt": "frame 5 mm inside label edges",
    },
}
LANGS = {"ru": "Русский", "en": "English"}


def open_source(path):
    """PDF as is; an image is converted into a one-page PDF."""
    path = Path(path)
    if path.suffix.lower() in IMG_EXT:
        img = pymupdf.open(path)
        return pymupdf.open("pdf", img.convert_to_pdf())
    return pymupdf.open(path)


def default_lang():
    try:
        loc = (locale.getlocale()[0] or "").lower()
    except Exception:
        loc = ""
    return "ru" if loc.startswith(("ru", "russian", "uk", "ukrainian", "be")) else "en"


def default_printer():
    try:
        return win32print.GetDefaultPrinter()
    except Exception:
        return ""


DEFAULT_CFG = {"printer": "", "offset_x_mm": 0.0, "offset_y_mm": 0.0, "dpi": 600,
               "allow_shrink": True, "mode": "A6", "feed_arrow": True, "lang": ""}


def load_json(p, default):
    try:
        return {**default, **json.loads(p.read_text(encoding="utf-8"))}
    except Exception:
        return dict(default)


def save_json(p, data):
    p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def cell_rect_mm(pos):
    """Cell pos (1..4) in mm from the top-left corner of the A4 sheet."""
    c, r = (pos - 1) % COLS, (pos - 1) // COLS
    return c * CELL_W, r * CELL_H, CELL_W, CELL_H


def page_cell(pos, mode):
    """Target cell: one of 4 on A4, the whole page on an A6 piece."""
    if mode == "A6":
        return 0.0, 0.0, PAPER["A6"][1], PAPER["A6"][2]
    return cell_rect_mm(pos)


# ---------- label detection ----------

def content_bbox(page):
    r = pymupdf.Rect()
    for _, b in page.get_bboxlog():
        bb = pymupdf.Rect(b) & page.rect
        if bb.is_empty or bb.width * bb.height > page.rect.width * page.rect.height * 0.98:
            continue  # ignore full-page background
        r |= bb
    if r.is_empty:
        r = page.rect
    pad = PAD_MM * MM
    return (r + (-pad, -pad, pad, pad)) & page.rect


def text_rotation(page, clip):
    """Dominant text direction: 0, 90 (top-to-bottom), 180 or 270."""
    votes = {0: 0, 90: 0, 180: 0, 270: 0}
    for b in page.get_text("dict", clip=clip)["blocks"]:
        for ln in b.get("lines", []):
            dx, dy = ln["dir"]
            n = sum(len(s["text"].strip()) for s in ln["spans"])
            if abs(dx) >= abs(dy):
                votes[0 if dx > 0 else 180] += n
            else:
                votes[90 if dy > 0 else 270] += n
    best = max(votes, key=votes.get)
    if votes[best] == 0:   # no text (image-only label) — use the shape
        return 90 if clip.width > clip.height else 0
    return best


def build_page(src_page, pos, cfg, rot_override=None, lang="en", out=None, clip=None, fit=False):
    """Returns (single-page doc of the paper size with the label in cell pos, info)."""
    mode = cfg.get("mode", "A4")
    clip = content_bbox(src_page) if clip is None else clip
    rot = text_rotation(src_page, clip) if rot_override is None else rot_override
    w, h = (clip.height, clip.width) if rot in (90, 270) else (clip.width, clip.height)
    cx, cy, cw, ch = page_cell(pos, mode)
    if fit:   # enlarge (or shrink) to the cell minus a margin the printer can reach
        scale = min((cw - 2 * FIT_MM) * MM / w, (ch - 2 * FIT_MM) * MM / h)
    else:
        scale = min(1.0, cw * MM / w, ch * MM / h)
    if scale < 1.0 and not cfg.get("allow_shrink", True):
        raise ValueError(T[lang]["too_big"])
    tw, th = w * scale, h * scale
    x0 = cx * MM + (cw * MM - tw) / 2
    y0 = cy * MM + (ch * MM - th) / 2
    _, pw, ph = PAPER[mode]
    if out is None:
        out = pymupdf.open()
        pg = out.new_page(width=pw * MM, height=ph * MM)
    else:
        pg = out[0]
    # rot is the text direction; show_pdf_page(rotate=rot) makes it upright
    pg.show_pdf_page(pymupdf.Rect(x0, y0, x0 + tw, y0 + th), src_page.parent, src_page.number,
                     clip=clip, rotate=rot)
    return out, {"rot": rot, "scale": scale, "w_mm": tw / MM, "h_mm": th / MM}


# ---------- GDI printing ----------

def paper_devmode(printer, paper=9):
    h = win32print.OpenPrinter(printer)
    try:
        dm = win32print.GetPrinter(h, 2)["pDevMode"]
    finally:
        win32print.ClosePrinter(h)
    dm.PaperSize = paper        # DMPAPER_A4=9, DMPAPER_A6=70
    dm.Orientation = 1          # portrait
    dm.Fields |= win32con.DM_PAPERSIZE | win32con.DM_ORIENTATION
    try:
        dm.Scale = 100
        dm.Fields |= win32con.DM_SCALE
    except Exception:
        pass
    return dm


def print_regions(printer, doc, regions_mm, cfg, title="Label", paper=9):
    """Prints rectangles (mm) of doc[0] at exactly the same place on the paper."""
    dm = paper_devmode(printer, paper)
    hdc = win32ui.CreateDCFromHandle(win32gui.CreateDC("WINSPOOL", printer, dm))
    try:
        dpx, dpy = hdc.GetDeviceCaps(win32con.LOGPIXELSX), hdc.GetDeviceCaps(win32con.LOGPIXELSY)
        offx, offy = hdc.GetDeviceCaps(win32con.PHYSICALOFFSETX), hdc.GetDeviceCaps(win32con.PHYSICALOFFSETY)
        dpi = min(int(cfg.get("dpi", 600)), dpx)
        hdc.StartDoc(title)
        hdc.StartPage()
        page = doc[0]
        for (x, y, w, h) in regions_mm:
            clip = pymupdf.Rect(x * MM, y * MM, (x + w) * MM, (y + h) * MM)
            pix = page.get_pixmap(dpi=dpi, clip=clip, alpha=False)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            X, Y = x + cfg["offset_x_mm"], y + cfg["offset_y_mm"]
            box = (round(X / 25.4 * dpx) - offx, round(Y / 25.4 * dpy) - offy,
                   round((X + w) / 25.4 * dpx) - offx, round((Y + h) / 25.4 * dpy) - offy)
            ImageWin.Dib(img).draw(hdc.GetHandleOutput(), box)
        hdc.EndPage()
        hdc.EndDoc()
    finally:
        hdc.DeleteDC()


def _font(page):
    try:
        page.insert_font(fontname="arial", fontfile=FONT)
        return "arial"
    except Exception:
        return "helv"


def draw_feed_arrow(page, lang="en"):
    """Up arrow in the gap between columns: 'feed this edge first'."""
    x, y, w, h = ARROW_BOX
    cx = (x + w / 2) * MM
    top, bot = y * MM, (y + h) * MM
    page.draw_line((cx, top + 3 * MM), (cx, bot), color=(0, 0, 0), width=0.8)
    page.draw_polyline([(cx - 1.6 * MM, top + 4 * MM), (cx, top), (cx + 1.6 * MM, top + 4 * MM)],
                       color=(0, 0, 0), fill=(0, 0, 0), width=0.5, closePath=True)
    fn = _font(page)
    txt = T[lang]["arrow_txt"] if fn == "arial" else T["en"]["arrow_txt"]
    page.insert_text((cx - 1.4 * MM, bot - 2 * MM), txt, fontname=fn, fontsize=5.5, rotate=90)


def calibration_doc(mode="A4", lang="en"):
    doc = pymupdf.open()
    _, pw, ph = PAPER[mode]
    pg = doc.new_page(width=pw * MM, height=ph * MM)
    fn = _font(pg)
    txt = T[lang]["calib_txt"] if fn == "arial" else T["en"]["calib_txt"]
    for pos in ([1] if mode == "A6" else range(1, 5)):
        x, y, w, h = page_cell(pos, mode)
        r = pymupdf.Rect(x * MM, y * MM, (x + w) * MM, (y + h) * MM)
        inner = r + (5 * MM, 5 * MM, -5 * MM, -5 * MM)       # frame 5 mm inside the cell
        pg.draw_rect(inner, color=(0, 0, 0), width=0.5)
        cxp, cyp = r.x0 + r.width / 2, r.y0 + r.height / 2
        pg.draw_line((cxp - 10 * MM, cyp), (cxp + 10 * MM, cyp), width=0.3)
        pg.draw_line((cxp, cyp - 10 * MM), (cxp, cyp + 10 * MM), width=0.3)
        pg.insert_text((inner.x0 + 4 * MM, inner.y0 + 8 * MM), f"{pos}", fontsize=20)
        pg.insert_text((inner.x0 + 4 * MM, inner.y0 + 14 * MM), txt, fontname=fn, fontsize=7)
    return doc


# ---------- GUI ----------

class App:
    def __init__(self, root, initial=None):
        self.root = root
        self.cfg = load_json(CFG, DEFAULT_CFG)
        if self.cfg.get("lang") not in T:
            self.cfg["lang"] = default_lang()
        if not self.cfg.get("printer"):
            self.cfg["printer"] = default_printer()
        self.state = load_json(STATE, {"used": []})
        self.src = None
        self.src_path = None
        self.page_no = 0
        self.rot_override = None
        self.preview_img = None
        self.main = None
        self.batch = []          # [(pymupdf doc, page_no, path str)] — all labels opened
        self.crops = {}          # (path str, page_no) -> pymupdf.Rect chosen by hand
        self.rots = {}           # (path str, page_no) -> rotation chosen by hand
        self.fits = {}           # (path str, page_no) -> enlarge to the whole cell
        root.resizable(False, False)
        root.bind_all("<Control-v>", lambda e: self.paste_clipboard())
        root.bind_all("<Control-V>", lambda e: self.paste_clipboard())
        self.build_ui()
        if initial:
            self.load(Path(initial))

    @property
    def t(self):
        return T[self.cfg["lang"]]

    def build_ui(self):
        t = self.t
        if self.main is not None:
            self.main.destroy()
        self.root.title(f"{APP_NAME} — {t['title']}")
        main = self.main = ttk.Frame(self.root, padding=10)
        main.grid()

        top = ttk.Frame(main)
        top.grid(row=0, column=0, columnspan=2, sticky="we")
        ttk.Button(top, text=t["open"], command=self.open_dialog).pack(side="left")
        ttk.Button(top, text=t["latest"], command=self.open_latest).pack(side="left", padx=6)
        self.lang_var = tk.StringVar(value=LANGS[self.cfg["lang"]])
        lang_box = ttk.Combobox(top, textvariable=self.lang_var, values=list(LANGS.values()),
                                width=9, state="readonly")
        lang_box.pack(side="right")
        lang_box.bind("<<ComboboxSelected>>", self.on_lang)
        ttk.Label(top, text=t["language"]).pack(side="right", padx=4)
        self.file_lbl = ttk.Label(top, text=self.src_path.name[:45] if self.src_path else t["no_file"],
                                  width=34)
        self.file_lbl.pack(side="left", padx=6)

        self.canvas = tk.Canvas(main, width=PV_W, height=round(PV_W * A4_H / A4_W), bg="white",
                                highlightthickness=1, highlightbackground="#888")
        self.canvas.grid(row=1, column=0, pady=8)
        self.canvas.bind("<Button-1>", self.on_canvas_click)

        side = ttk.Frame(main)
        side.grid(row=1, column=1, sticky="n", padx=10, pady=8)
        ttk.Label(side, text=t["paper"]).pack(anchor="w")
        self.mode = tk.StringVar(value=self.cfg.get("mode", "A6"))
        ttk.Radiobutton(side, text=t["mode_a4"], variable=self.mode, value="A4",
                        command=self.on_mode).pack(anchor="w")
        ttk.Radiobutton(side, text=t["mode_a6"], variable=self.mode, value="A6",
                        command=self.on_mode).pack(anchor="w")
        ttk.Separator(side).pack(fill="x", pady=6)
        ttk.Label(side, text=t["position"]).pack(anchor="w")
        self.pos = tk.IntVar(value=self.next_free() or 1)
        self.pos_btns = []
        for p in range(1, 5):
            b = ttk.Radiobutton(side, text=t["pos"][p], variable=self.pos, value=p, command=self.redraw)
            b.pack(anchor="w")
            self.pos_btns.append(b)
        self.new_btn = ttk.Button(side, text=t["new_sheet"], command=self.new_sheet)
        self.new_btn.pack(fill="x", pady=(8, 2))
        ttk.Button(side, text=t["flip"], command=self.flip).pack(fill="x", pady=2)
        ttk.Button(side, text=t["crop"], command=self.crop_current).pack(fill="x", pady=2)

        ttk.Separator(side).pack(fill="x", pady=8)
        ttk.Label(side, text=t["printer"]).pack(anchor="w")
        printers = [p[2] for p in win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
        self.printer = tk.StringVar(value=self.cfg["printer"])
        ttk.Combobox(side, textvariable=self.printer, values=printers, width=26,
                     state="readonly").pack(anchor="w")
        off = ttk.Frame(side)
        off.pack(anchor="w", pady=6)
        ttk.Label(off, text=t["off_x"]).grid(row=0, column=0, sticky="w")
        ttk.Label(off, text=t["off_y"]).grid(row=1, column=0, sticky="w")
        self.offx = tk.DoubleVar(value=self.cfg["offset_x_mm"])
        self.offy = tk.DoubleVar(value=self.cfg["offset_y_mm"])
        ttk.Spinbox(off, from_=-10, to=10, increment=0.5, textvariable=self.offx, width=6).grid(row=0, column=1)
        ttk.Spinbox(off, from_=-10, to=10, increment=0.5, textvariable=self.offy, width=6).grid(row=1, column=1)
        ttk.Button(side, text=t["calib"], command=self.print_calibration).pack(fill="x", pady=2)
        self.arrow = tk.BooleanVar(value=self.cfg.get("feed_arrow", True))
        ttk.Checkbutton(side, text=t["arrow_chk"], variable=self.arrow,
                        command=self.redraw).pack(anchor="w", pady=(4, 0))

        ttk.Separator(side).pack(fill="x", pady=8)
        self.info = ttk.Label(side, text="", wraplength=210, justify="left")
        self.info.pack(anchor="w")
        ttk.Button(side, text=t["print"], command=self.do_print).pack(fill="x", pady=(10, 0), ipady=8)
        ttk.Separator(side).pack(fill="x", pady=8)
        ttk.Button(side, text=t["paste"], command=self.paste_clipboard).pack(fill="x", pady=(0, 4))
        qf = ttk.Frame(side)
        qf.pack(fill="x")
        ttk.Button(qf, text=t["add"], command=self.add_dialog).pack(side="left", fill="x", expand=True)
        ttk.Button(qf, text=t["clear"], command=self.clear_queue).pack(side="left", padx=(4, 0))
        self.all_btn = ttk.Button(side, command=self.print_all)
        self.all_btn.pack(fill="x", pady=(6, 0), ipady=6)
        self.on_mode()

    def on_lang(self, _ev=None):
        name = self.lang_var.get()
        self.cfg["lang"] = next(k for k, v in LANGS.items() if v == name)
        self.save_cfg()
        self.build_ui()

    # --- sheet state
    def next_free(self):
        for p in range(1, 5):
            if p not in self.state["used"]:
                return p
        return None

    def new_sheet(self):
        self.state["used"] = []
        save_json(STATE, self.state)
        self.pos.set(1)
        self.redraw()

    def flip(self):
        if self.src is None:
            return
        self.rots[(str(self.src_path), self.page_no)] = (self.current_rot() + 180) % 360
        self.redraw()

    def current_rot(self):
        r = self.rots.get((str(self.src_path), self.page_no))
        if r is not None:
            return r
        pg = self.src[self.page_no]
        return text_rotation(pg, self.clip_for(str(self.src_path), self.page_no) or content_bbox(pg))

    def arrow_due(self):
        """The arrow is printed once: in A4 mode, on a fresh sheet (no slot used yet)."""
        return self.cfg.get("mode") == "A4" and self.arrow.get() and not self.state["used"]

    def on_mode(self):
        self.cfg["mode"] = self.mode.get()
        st = "disabled" if self.cfg["mode"] == "A6" else "normal"
        for w in self.pos_btns + [self.new_btn]:
            w.config(state=st)
        self.update_batch_btn()
        self.redraw()

    def queue_preview(self):
        return self.cfg.get("mode") == "A4" and len(self.batch) > 1

    def add_dialog(self):
        self.add_files(filedialog.askopenfilenames(initialdir=DOWNLOADS, filetypes=FILETYPES))

    def add_files(self, fs):
        for f in fs:
            try:
                d = open_source(f)
            except Exception as e:
                messagebox.showerror(self.t["error"], self.t["open_err"].format(e=e))
                continue
            if self.src is None:
                self.load(Path(f))
                self.batch = []
            self.batch += [(d, i, str(Path(f))) for i in range(len(d))]
        self.update_batch_btn()
        self.redraw()

    def paste_clipboard(self):
        """Screenshot (or copied image files) from the clipboard -> load or add to the queue."""
        try:
            clip = ImageGrab.grabclipboard()
        except Exception:
            clip = None
        if isinstance(clip, list):
            files = [f for f in clip if Path(f).suffix.lower() in (".pdf",) + IMG_EXT]
        elif isinstance(clip, Image.Image):
            OUT.mkdir(exist_ok=True)
            f = OUT / f"clip-{time.strftime('%Y%m%d-%H%M%S')}.png"
            clip.convert("RGB").save(f)
            files = [str(f)]
        else:
            files = []
        if not files:
            messagebox.showinfo(self.t["no_clip_t"], self.t["no_clip"])
            return
        self.add_files(files)

    def clip_for(self, key, pn):
        return self.crops.get((key, pn))

    def fit_for(self, key, pn):
        """Hand-cropped labels are enlarged to the cell unless switched off."""
        return self.fits.get((key, pn), (key, pn) in self.crops)

    def free_slots(self):
        return [p for p in range(1, 5) if p not in self.state["used"]] or [1, 2, 3, 4]

    def crop_current(self):
        if self.queue_preview():
            messagebox.showinfo(self.t["crop_t"], self.t["crop_queue"])
            return
        if self.src is None:
            messagebox.showinfo(self.t["nofile_t"], self.t["nofile"])
            return
        self.crop_dialog(self.src, self.page_no, str(self.src_path))

    def crop_dialog(self, doc, pn, key):
        """Whole source page on the left (drag = area to print), live result on the right."""
        t = self.t
        page = doc[pn]
        z = min(760 / page.rect.width, 640 / page.rect.height)
        pix = page.get_pixmap(matrix=pymupdf.Matrix(z, z), alpha=False)
        img = ImageTk.PhotoImage(Image.frombytes("RGB", (pix.width, pix.height), pix.samples))
        win = tk.Toplevel(self.root)
        win.title(f"{t['crop_t']} — {Path(key).name}")
        win.transient(self.root)
        win.resizable(False, False)
        body = ttk.Frame(win, padding=6)
        body.pack()
        left, right = ttk.Frame(body), ttk.Frame(body)
        left.grid(row=0, column=0, sticky="n")
        right.grid(row=0, column=1, sticky="n", padx=(10, 0))
        ttk.Label(left, text=t["crop_hint"]).pack(anchor="w", pady=(0, 2))
        cv = tk.Canvas(left, width=pix.width, height=pix.height, cursor="crosshair",
                       highlightthickness=1, highlightbackground="#888")
        cv.pack()
        cv.create_image(0, 0, anchor="nw", image=img)
        cv.img = img
        cur = self.clip_for(key, pn) or content_bbox(page)
        sel = {"r": [cur.x0 * z, cur.y0 * z, cur.x1 * z, cur.y1 * z]}
        rid = cv.create_rectangle(*sel["r"], outline="#d22", width=2, dash=(5, 3))

        # result preview: one A6-sized cell, the same geometry as on paper
        rw = 280
        rh = round(rw * PAPER["A6"][2] / PAPER["A6"][1])
        ttk.Label(right, text=t["result"]).pack(anchor="w", pady=(0, 2))
        rv_cv = tk.Canvas(right, width=rw, height=rh, bg="white",
                          highlightthickness=1, highlightbackground="#888")
        rv_cv.pack()
        rot = {"v": self.rots.get((key, pn))}
        rot_var = tk.StringVar()
        fit_var = tk.BooleanVar(value=self.fits.get((key, pn), True))
        info_var = tk.StringVar()

        def sel_rect():
            x0, y0, x1, y1 = sel["r"]
            r = pymupdf.Rect(min(x0, x1) / z, min(y0, y1) / z, max(x0, x1) / z, max(y0, y1) / z)
            return r & page.rect if r.width > 10 and r.height > 10 else None

        def render():
            r = sel_rect()
            try:
                cfg = dict(self.cfg, mode="A6")
                built, inf = build_page(page, 1, cfg, rot["v"], self.cfg["lang"],
                                        clip=r, fit=fit_var.get() if r is not None else False)
            except Exception as e:
                info_var.set(f"{t['error']}: {e}")
                return
            zz = rw / (PAPER["A6"][1] * MM)
            px = built[0].get_pixmap(matrix=pymupdf.Matrix(zz, zz), alpha=False)
            rv_cv.img = ImageTk.PhotoImage(Image.frombytes("RGB", (px.width, px.height), px.samples))
            rv_cv.delete("all")
            rv_cv.create_image(0, 0, anchor="nw", image=rv_cv.img)
            k = rw / PAPER["A6"][1]
            rv_cv.create_rectangle(FIT_MM * k, FIT_MM * k, rw - FIT_MM * k, rh - FIT_MM * k,
                                   outline="#bbb", dash=(2, 3))
            auto = t["rot_auto"] + f" ({inf['rot']}°)"
            rot_var.set(t["rot_lbl"].format(r=auto if rot["v"] is None else f"{rot['v']}°"))
            info_var.set(f"{inf['w_mm']:.0f}×{inf['h_mm']:.0f} мм, {inf['scale'] * 100:.0f}%")
            rot["shown"] = inf["rot"]

        def rotate90():
            base = rot["v"] if rot["v"] is not None else rot.get("shown", 0)
            rot["v"] = (base + 90) % 360
            render()

        def press(ev):
            sel["r"] = [ev.x, ev.y, ev.x, ev.y]
            cv.coords(rid, *sel["r"])

        def drag(ev):
            sel["r"][2] = min(max(ev.x, 0), pix.width)
            sel["r"][3] = min(max(ev.y, 0), pix.height)
            cv.coords(rid, *sel["r"])

        def done(rect, rv, fit):
            for store, val in ((self.crops, rect), (self.rots, rv), (self.fits, fit)):
                if val is None:
                    store.pop((key, pn), None)
                else:
                    store[(key, pn)] = val
            win.destroy()
            self.redraw()

        def ok():
            r = sel_rect()
            done(r, rot["v"], bool(fit_var.get()) if r is not None else None)

        cv.bind("<ButtonPress-1>", press)
        cv.bind("<B1-Motion>", drag)
        cv.bind("<ButtonRelease-1>", lambda e: render())
        ttk.Label(right, textvariable=info_var).pack(anchor="w", pady=(4, 0))
        ttk.Checkbutton(right, text=t["fit"], variable=fit_var, command=render).pack(anchor="w", pady=(6, 0))
        ttk.Button(right, text=t["rotate90"], command=rotate90).pack(fill="x", pady=(6, 0))
        ttk.Label(right, textvariable=rot_var).pack(anchor="w", pady=(2, 0))
        ttk.Separator(right).pack(fill="x", pady=8)
        ttk.Button(right, text=t["ok"], command=ok).pack(fill="x", ipady=4)
        ttk.Button(right, text=t["auto"], command=lambda: done(None, None, None)).pack(fill="x", pady=(6, 0))
        ttk.Button(right, text=t["cancel"], command=win.destroy).pack(fill="x", pady=(6, 0))
        win.bind("<Return>", lambda e: ok())
        win.bind("<Escape>", lambda e: win.destroy())
        render()
        win.grab_set()

    def clear_queue(self):
        self.batch = []
        self.update_batch_btn()
        self.redraw()

    def update_batch_btn(self):
        n = len(self.batch)
        self.file_lbl.config(text=(self.src_path.name[:28] if self.src_path else self.t["no_file"])
                             + (f"  ({self.t['queued'].format(n=n)})" if n > 1 else ""))
        self.all_btn.config(text=self.t["print_all"].format(n=n),
                            state="normal" if n > 1 and self.cfg.get("mode") == "A4" else "disabled")

    def on_canvas_click(self, ev):
        if self.cfg.get("mode") == "A6":
            return
        col = 0 if ev.x < PV_W / 2 else 1
        row = 0 if ev.y < PV_W * A4_H / A4_W / 2 else 1
        if self.queue_preview():
            free = self.free_slots()
            pos = row * 2 + col + 1
            if pos in free and free.index(pos) < len(self.batch):
                d, pn, key = self.batch[free.index(pos)]
                self.crop_dialog(d, pn, key)
            return
        self.pos.set(row * 2 + col + 1)
        self.redraw()

    # --- loading
    def open_dialog(self):
        fs = filedialog.askopenfilenames(initialdir=DOWNLOADS, filetypes=FILETYPES)
        if not fs:
            return
        self.load(Path(fs[0]))
        if len(fs) > 1:
            batch = []
            for f in fs:
                try:
                    d = open_source(f)
                except Exception as e:
                    messagebox.showerror(self.t["error"], self.t["open_err"].format(e=e))
                    continue
                batch += [(d, i, str(Path(f))) for i in range(len(d))]
            self.batch = batch
            self.update_batch_btn()

    def open_latest(self):
        pdfs = sorted((p for p in DOWNLOADS.iterdir() if p.suffix.lower() in (".pdf",) + IMG_EXT),
                      key=lambda p: p.stat().st_mtime, reverse=True)
        if not pdfs:
            messagebox.showinfo(self.t["no_pdfs_t"], self.t["no_pdfs"])
            return
        self.load(pdfs[0])

    def load(self, path):
        try:
            self.src = open_source(path)
        except Exception as e:
            messagebox.showerror(self.t["error"], self.t["open_err"].format(e=e))
            return
        self.src_path = path
        self.page_no = 0
        self.rot_override = None
        self.file_lbl.config(text=path.name[:45])
        self.batch = [(self.src, i, str(path)) for i in range(len(self.src))]
        self.update_batch_btn()
        self.redraw()

    # --- preview
    def redraw(self):
        t = self.t
        c = self.canvas
        c.delete("all")
        k = PV_W / A4_W
        if self.src is not None:
            try:
                if self.queue_preview():
                    built = pymupdf.open()
                    built.new_page(width=A4_W * MM, height=A4_H * MM)
                    for pos, (d, pn, key) in zip(self.free_slots(), self.batch):
                        build_page(d[pn], pos, self.cfg, self.rots.get((key, pn)), self.cfg["lang"], out=built,
                                   clip=self.clip_for(key, pn), fit=self.fit_for(key, pn))
                    inf = {"w_mm": 0, "h_mm": 0, "scale": 1, "rot": 0}
                else:
                    built, inf = build_page(self.src[self.page_no], self.pos.get(), self.cfg,
                                            self.rots.get((str(self.src_path), self.page_no)), self.cfg["lang"],
                                            clip=self.clip_for(str(self.src_path), self.page_no),
                                            fit=self.fit_for(str(self.src_path), self.page_no))
                if self.arrow_due():
                    draw_feed_arrow(built[0], self.cfg["lang"])
                z = PV_W / (PAPER[self.cfg["mode"]][1] * MM)
                pix = built[0].get_pixmap(matrix=pymupdf.Matrix(z, z), alpha=False)
                img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                self.preview_img = ImageTk.PhotoImage(img)
                c.create_image(0, 0, anchor="nw", image=self.preview_img)
                pages = t["page"].format(n=self.page_no + 1, t=len(self.src)) if len(self.src) > 1 else ""
                if self.queue_preview():
                    self.info.config(text=t["queued"].format(n=len(self.batch)))
                else:
                    own = t["cropped"] if self.clip_for(str(self.src_path), self.page_no) else ""
                    self.info.config(text=t["info"].format(w=inf["w_mm"], h=inf["h_mm"],
                                                           s=inf["scale"] * 100, r=inf["rot"]) + pages + own)
            except Exception as e:
                self.info.config(text=f"{t['error']}: {e}")
        if self.cfg.get("mode") == "A6":
            c.create_rectangle(2, 2, PV_W - 2, int(c.cget("height")) - 2, outline="#d22", width=3)
            c.create_text(PV_W / 2, 12, text=t["a6_hint"], fill="#d22")
            return
        for p in range(1, 5):
            x, y, w, h = cell_rect_mm(p)
            r = (x * k, y * k, (x + w) * k, (y + h) * k)
            if p in self.state["used"]:
                c.create_rectangle(*r, fill="#bbb", stipple="gray50", outline="")
                c.create_text((r[0] + r[2]) / 2, (r[1] + r[3]) / 2, text=t["used"], fill="#555",
                              font=("Segoe UI", 12))
            c.create_rectangle(*r, outline="#999", dash=(3, 3))
            if p == self.pos.get() and not self.queue_preview():
                c.create_rectangle(r[0] + 2, r[1] + 2, r[2] - 2, r[3] - 2, outline="#d22", width=3)

    # --- printing
    def save_cfg(self):
        self.cfg.update(printer=self.printer.get(), mode=self.mode.get(),
                        feed_arrow=bool(self.arrow.get()),
                        offset_x_mm=float(self.offx.get()), offset_y_mm=float(self.offy.get()))
        save_json(CFG, self.cfg)

    def print_calibration(self):
        self.save_cfg()
        t, mode = self.t, self.cfg["mode"]
        if not messagebox.askokcancel(t["calib_t"], t["calib_q"].format(m=mode)):
            return
        try:
            paper, pw, ph = PAPER[mode]
            print_regions(self.cfg["printer"], calibration_doc(mode, self.cfg["lang"]), [(0, 0, pw, ph)],
                          self.cfg, "Calibration", paper)
        except Exception as e:
            messagebox.showerror(t["print_err"], str(e))

    def print_all(self):
        """Print every opened label into the free slots of one A4 sheet, in one job."""
        t = self.t
        if self.cfg.get("mode") != "A4" or not self.batch:
            return
        self.save_cfg()
        free = [p for p in range(1, 5) if p not in self.state["used"]]
        if not free:
            self.new_sheet()
            free = [1, 2, 3, 4]
        if len(self.batch) > len(free):
            messagebox.showinfo(t["batch_t"], t["batch_more"].format(f=len(free), n=len(self.batch)))
        items = list(zip(free, self.batch))
        try:
            doc = pymupdf.open()
            doc.new_page(width=A4_W * MM, height=A4_H * MM)
            for pos, (d, pn, key) in items:
                build_page(d[pn], pos, self.cfg, self.rots.get((key, pn)), self.cfg["lang"], out=doc,
                           clip=self.clip_for(key, pn), fit=self.fit_for(key, pn))
            regions = [page_cell(pos, "A4") for pos, _ in items]
            if self.arrow_due():
                draw_feed_arrow(doc[0], self.cfg["lang"])
                regions.append(ARROW_BOX)
            OUT.mkdir(exist_ok=True)
            doc.save(OUT / f"{time.strftime('%Y%m%d-%H%M%S')}-batch{len(items)}.pdf")
            print_regions(self.cfg["printer"], doc, regions, self.cfg, "Labels batch", PAPER["A4"][0])
        except Exception as e:
            messagebox.showerror(t["print_err"], str(e))
            return
        for pos, _ in items:
            if pos not in self.state["used"]:
                self.state["used"].append(pos)
        if len(self.state["used"]) >= 4:
            self.state["used"] = []
            messagebox.showinfo(t["sheet_done_t"], t["sheet_done"])
        save_json(STATE, self.state)
        self.batch = self.batch[len(items):]
        self.update_batch_btn()
        self.pos.set(self.next_free() or 1)
        self.redraw()

    def do_print(self):
        t = self.t
        if self.src is None:
            messagebox.showinfo(t["nofile_t"], t["nofile"])
            return
        self.save_cfg()
        pos = self.pos.get()
        a6 = self.cfg["mode"] == "A6"
        if not a6 and pos in self.state["used"] and not messagebox.askyesno(
                t["busy_t"], t["busy"].format(p=pos)):
            return
        try:
            doc, _ = build_page(self.src[self.page_no], pos, self.cfg,
                                self.rots.get((str(self.src_path), self.page_no)), self.cfg["lang"],
                                clip=self.clip_for(str(self.src_path), self.page_no),
                                fit=self.fit_for(str(self.src_path), self.page_no))
            regions = [page_cell(pos, self.cfg["mode"])]
            if self.arrow_due():
                draw_feed_arrow(doc[0], self.cfg["lang"])
                regions.append(ARROW_BOX)
            OUT.mkdir(exist_ok=True)
            tag = "a6" if a6 else f"pos{pos}"
            doc.save(OUT / f"{time.strftime('%Y%m%d-%H%M%S')}-{tag}-{self.src_path.stem[:20]}.pdf")
            print_regions(self.cfg["printer"], doc, regions, self.cfg,
                          self.src_path.name, PAPER[self.cfg["mode"]][0])
        except Exception as e:
            messagebox.showerror(t["print_err"], str(e))
            return
        if not a6:
            if pos not in self.state["used"]:
                self.state["used"].append(pos)
            if len(self.state["used"]) >= 4:
                self.state["used"] = []
                messagebox.showinfo(t["sheet_done_t"], t["sheet_done"])
            save_json(STATE, self.state)
        # next page of a multi-page PDF, next free slot
        if self.page_no + 1 < len(self.src):
            self.page_no += 1
            self.rot_override = None
        self.pos.set(self.next_free() or 1)
        self.redraw()


def main():
    root = tk.Tk()
    App(root, sys.argv[1] if len(sys.argv) > 1 else None)
    root.mainloop()


if __name__ == "__main__":
    main()
