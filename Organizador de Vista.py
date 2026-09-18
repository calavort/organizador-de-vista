# -*- coding: utf-8 -*-
"""
Organizador de Vista 2.1 - Tekla
Programa desenvolvido por Edflávio Calavort - 2026

Versao 2.1 - base estavel da versao 9 com correcao de distribuicao
- Mantém a movimentação que funcionou: View.Origin via Tekla API.
- Corrige a identificação da PEÇA e alinha as linhas pelo rótulo PEÇA.
- Lê o número da peça por três caminhos: título/tag da View, objetos de texto da View/folha e propriedades do modelo.
- Ordena por chave numérica real: 1.1, 1.2, 1.3 ... 10.1, 11.10.
- Não usa mouse para arrastar.
- Correcao principal: quando acaba a altura util, abre uma nova coluna para a direita
  em vez de marcar as proximas vistas como "Sem espaco vertical".
"""

import csv
import ctypes
import json
import math
import os
import queue
import re
import sys
import threading
import time
import traceback
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

IS_WINDOWS = os.name == "nt"
if IS_WINDOWS:
    from ctypes import wintypes

APP_NAME = "Organizador de Vista - Tekla"
BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"
HTML_UI_PATH = BASE_DIR / "interface_organizador_tekla.html"
REPORT_DIR = BASE_DIR / "relatorios"
REPORT_DIR.mkdir(exist_ok=True)
DIMENSION_TOP_GAP_MM = 120.0
STACK_VIEW_SPACING_MM = 10.0
DIVIDER_LABEL_GAP_MM = 8.0
# Meia-altura estimada do rotulo PEÇA (mm na folha), usada so quando a API nao
# entrega o bbox do texto, para estimar a parte inferior do rotulo.
DIVIDER_LABEL_HALF_MM = 2.0
SHEET_FRAME_LEFT_MARGIN_MM = 25.0
SHEET_FRAME_RIGHT_MARGIN_MM = 10.0
SHEET_FRAME_BOTTOM_MARGIN_MM = 10.0
DIMENSION_ATTRIBUTE_NAMES = ("EC_COTAS", "+EC_COTAS", "EC_cotas", "+EC_cotas")

PIECE_LABEL_RE = re.compile(r"\bPECA\s*(?P<item>\d+(?:[\.,]\d+)*)", re.IGNORECASE)
NUMERIC_ITEM_RE = re.compile(r"(?<!\d)(?P<item>\d+(?:[\.,]\d+)*)(?!\d)", re.IGNORECASE)
DRAWING_MARK_RE = re.compile(r"\[(?P<item>\d+(?:(?:\.\.|\.)\d+)*)\]")

if IS_WINDOWS:
    WM_NCCALCSIZE = 0x0083
    WM_NCLBUTTONDOWN = 0x00A1
    WM_HOTKEY = 0x0312
    MOD_NOREPEAT = 0x4000
    HTCAPTION = 2
    HTLEFT = 10
    HTRIGHT = 11
    HTTOP = 12
    HTTOPLEFT = 13
    HTTOPRIGHT = 14
    HTBOTTOM = 15
    HTBOTTOMLEFT = 16
    HTBOTTOMRIGHT = 17
    SM_CXFRAME = 32
    SM_CYFRAME = 33
    SM_CXPADDEDBORDER = 92
    VK_N = 0x4E
    HOTKEY_DIMENSION_N = 0x4E05

    class _NCCALCSIZE_PARAMS(ctypes.Structure):
        _fields_ = [("rgrc", wintypes.RECT * 3), ("lppos", ctypes.c_void_p)]

    WINDOW_RESIZE_HITTEST = {
        "top": HTTOP,
        "bottom": HTBOTTOM,
        "left": HTLEFT,
        "right": HTRIGHT,
        "top-left": HTTOPLEFT,
        "top-right": HTTOPRIGHT,
        "bottom-left": HTBOTTOMLEFT,
        "bottom-right": HTBOTTOMRIGHT,
    }


def prepare_windows_thread_for_tekla():
    if os.name != "nt":
        return
    try:
        ctypes.windll.ole32.CoInitializeEx(None, 2)
    except Exception:
        pass


def norm_spaces(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "").strip())


def remove_diacritics(value: str) -> str:
    return "".join(ch for ch in unicodedata.normalize("NFD", value or "") if unicodedata.category(ch) != "Mn")


def parse_sort_key(text: str):
    out = []
    for part in (text or "").replace(",", ".").split("."):
        try:
            out.append(int(part))
        except Exception:
            out.append(0)
    return out


def try_parse_piece_label(value: str):
    value = norm_spaces(value)
    if not value:
        return None
    searchable = remove_diacritics(value).upper()
    m = PIECE_LABEL_RE.search(searchable)
    if not m:
        return None
    item_text = m.group("item").replace(",", ".")
    key = parse_sort_key(item_text)
    if not key:
        return None
    return value, item_text, key


def try_parse_numeric_item(value: str):
    value = norm_spaces(value)
    if not value:
        return None
    m = NUMERIC_ITEM_RE.search(value)
    if not m:
        return None
    item_text = m.group("item").replace(",", ".")
    key = parse_sort_key(item_text)
    if not key:
        return None
    return item_text, key


def try_parse_drawing_mark(value: str):
    text = norm_spaces(str(value or ""))
    if not text:
        return None
    match = DRAWING_MARK_RE.search(text)
    if match:
        text = match.group("item")
    text = text.replace("..", ".").replace(",", ".")
    text = re.sub(r"\.+", ".", text).strip(".")
    if not re.fullmatch(r"\d+(?:\.\d+)*", text):
        return None
    return text, parse_sort_key(text)


def try_parse_drawing_view_group(value: str):
    text = norm_spaces(str(value or "")).replace(",", ".")
    if not text or ".." not in text:
        return None
    match = re.search(r"(?<!\d)(?P<base>\d+(?:\.\d+)*)\.\.(?P<view>\d+)(?!\d)", text)
    if not match:
        return None
    base_text = match.group("base").strip(".")
    view_text = match.group("view")
    if not base_text or not view_text:
        return None
    return base_text, parse_sort_key(base_text), [int(view_text)]


def try_parse_piece_from_strings(values):
    """
    O Tekla às vezes entrega o título da vista quebrado em vários elementos:
    ["PEÇA", "1.1, (1X)"] ou ["PE", "ÇA 1.1"].
    Esta função junta janelas curtas de texto e procura PEÇA + número.
    """
    clean = [norm_spaces(str(v or "")) for v in values if norm_spaces(str(v or ""))]
    for i in range(len(clean)):
        for j in range(i + 1, min(len(clean), i + 5) + 1):
            joined = norm_spaces(" ".join(clean[i:j]))
            parsed = try_parse_piece_label(joined)
            if parsed:
                return parsed
    # fallback: procura qualquer trecho que contenha PECA e depois um número próximo
    joined_all = norm_spaces(" ".join(clean))
    searchable = remove_diacritics(joined_all).upper()
    m = re.search(r"PECA\s*[^0-9]{0,20}(?P<item>\d+(?:[\.,]\d+)*)", searchable)
    if m:
        item_text = m.group("item").replace(",", ".")
        key = parse_sort_key(item_text)
        return "PEÇA " + item_text, item_text, key
    return None


def numeric_key(key):
    return tuple(key or [])


def parse_scale_denominator(value):
    """Aceita 5, 1:5 ou 1/5 e retorna o denominador 5.0."""
    text = norm_spaces(str(value or "")).replace(",", ".")
    if not text:
        return 1.0
    text = text.replace(" ", "")
    for sep in (":", "/"):
        if sep in text:
            left, right = text.split(sep, 1)
            left_value = float(left or "1")
            right_value = float(right)
            if left_value == 0:
                raise ValueError("Escala invalida.")
            if abs(left_value - 1.0) < 0.0001:
                return max(1.0, right_value)
            return max(1.0, right_value / left_value)
    return max(1.0, float(text))


@dataclass
class TeklaFrame:
    left: float
    bottom: float
    right: float
    top: float

    @property
    def width(self):
        return abs(self.right - self.left)

    @property
    def height(self):
        return abs(self.top - self.bottom)


@dataclass
class PieceLabelInfo:
    label: str
    item_text: str
    sort_key: list
    center_x: float
    center_y: float
    left: float = 0.0
    bottom: float = 0.0
    right: float = 0.0
    top: float = 0.0


@dataclass
class ApiViewItem:
    index: int
    name: str
    item_text: str
    sort_key: list
    view: object
    current: TeklaFrame
    target_left: float = 0.0
    target_bottom: float = 0.0
    target_right: float = 0.0
    target_top: float = 0.0
    has_target: bool = False
    moved: bool = False
    status: str = "Planejada"
    label_center_y: float = None
    label_offset_y: float = None
    label_bottom_y: float = None
    current_scale: float = 1.0
    proposed_scale: float = 0.0
    stack_group_key: str = ""
    stack_group_sort_key: list = None
    stack_order_key: list = None
    # Folga entre a borda do frame da View e o conteudo visivel (desenho/cotas/rotulo).
    # Usada para que "Dist. empilh." seja a distancia real entre os desenhos.
    content_pad_top: float = 0.0
    content_pad_bottom: float = 0.0
    content_pad_left: float = 0.0
    content_pad_right: float = 0.0
    piece_identity_trusted: bool = True

    @property
    def dx(self):
        return self.target_left - self.current.left if self.has_target else 0.0

    @property
    def dy(self):
        return self.target_bottom - self.current.bottom if self.has_target else 0.0


@dataclass
class ViewUndoState:
    index: int
    name: str
    item_text: str
    view: object
    origin_x: float
    origin_y: float
    origin_z: float
    scale: float
    frame_left: float
    frame_bottom: float


class Config:
    def __init__(self):
        self.margin_left = 35.0
        self.margin_top = 35.0
        self.margin_right = 35.0
        self.margin_bottom = 35.0
        self.spacing_x = 25.0
        self.spacing_y = 25.0
        self.stack_spacing_y = STACK_VIEW_SPACING_MM
        self.title_block_width = 185.0
        self.title_block_height = 70.0
        self.max_details = 0
        self.scale_target = 2.0
        self.auto_scale_limit = 0.0
        self.scale_apply_mode = "all"
        self.dimension_top_gap_mm = DIMENSION_TOP_GAP_MM
        self.divider_label_gap_mm = DIVIDER_LABEL_GAP_MM
        self.density_alert = 0.62
        self.topmost = True
        self.target_sheet_width = 0.0
        self.target_sheet_height = 0.0
        self.capacity_factor_x = 1.0
        self.capacity_factor_y = 1.0
        self.capacity_factor_matches = 0
        self.capacity_cursor_item = ""
        self.capacity_cursor_sheet = ""
        self.capacity_cursor_count = 0
        self.capacity_sheet_cursors = {}
        self.capacity_project_key = ""
        self.dimension_hotkey_enabled = True

    def save(self):
        CONFIG_PATH.write_text(json.dumps(self.__dict__, indent=2, ensure_ascii=False), encoding="utf-8")

    @classmethod
    def load(cls):
        cfg = cls()
        if CONFIG_PATH.exists():
            try:
                data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
                for k, v in data.items():
                    if hasattr(cfg, k):
                        setattr(cfg, k, v)
            except Exception:
                pass
        return cfg


UI = {
    "bg": "#eef3fa",
    "panel": "#ffffff",
    "line": "#d9e2ef",
    "line_soft": "#edf2f8",
    "text": "#142033",
    "muted": "#667894",
    "blue": "#2f5fd3",
    "blue_dark": "#244fbd",
    "yellow": "#efd23c",
    "yellow_dark": "#d9ba1e",
}


def draw_line_icon(canvas, name, color=None):
    color = color or UI["blue"]
    canvas.delete("all")
    try:
        size = int(float(canvas.cget("width")))
    except Exception:
        size = 24
    scale = size / 24.0

    def p(value):
        return value * scale

    def line(*coords, **kwargs):
        canvas.create_line(
            [p(v) for v in coords],
            fill=kwargs.get("fill", color),
            width=max(1, int(round(kwargs.get("width", 2.1) * scale))),
            capstyle="round",
            joinstyle="round",
        )

    def rect(x1, y1, x2, y2, **kwargs):
        canvas.create_rectangle(
            p(x1), p(y1), p(x2), p(y2),
            outline=kwargs.get("outline", color),
            fill=kwargs.get("fill", ""),
            width=max(1, int(round(kwargs.get("width", 2.0) * scale))),
        )

    def oval(x1, y1, x2, y2, **kwargs):
        canvas.create_oval(
            p(x1), p(y1), p(x2), p(y2),
            outline=kwargs.get("outline", color),
            fill=kwargs.get("fill", ""),
            width=max(1, int(round(kwargs.get("width", 2.0) * scale))),
        )

    if name == "analyze":
        oval(4.5, 4.5, 14.5, 14.5)
        line(13, 13, 20, 20)
    elif name == "capacity":
        for x in (4, 10, 16):
            for y in (4, 10, 16):
                rect(x, y, x + 4, y + 4, width=1.6)
    elif name == "organize":
        rect(4, 5, 9, 10, width=1.6)
        rect(4, 14, 9, 19, width=1.6)
        line(10, 7.5, 18, 7.5)
        line(15.5, 5, 18, 7.5, 15.5, 10)
        line(10, 16.5, 18, 16.5)
        line(15.5, 14, 18, 16.5, 15.5, 19)
    elif name == "smart":
        canvas.create_polygon(
            p(13), p(2.5), p(5.5), p(13), p(11.5), p(13),
            p(9.5), p(21.5), p(18.5), p(10.5), p(12.5), p(10.5),
            fill=color,
            outline=color,
        )
    elif name == "undo":
        canvas.create_arc(p(5), p(6), p(20), p(20), start=35, extent=285, style="arc", outline=color, width=max(1, int(2.1 * scale)))
        line(7, 7, 4, 7, 4, 4)
    elif name == "horizontal":
        line(4, 12, 20, 12)
        line(8, 8, 4, 12, 8, 16)
        line(16, 8, 20, 12, 16, 16)
    elif name == "vertical":
        line(12, 4, 12, 20)
        line(8, 8, 12, 4, 16, 8)
        line(8, 16, 12, 20, 16, 16)
    elif name == "legend":
        rect(4, 5, 19, 18, width=1.8)
        line(8, 14, 17, 14, width=1.4)
        line(8, 10, 15, 10, width=1.4)
    elif name == "spacing":
        rect(4, 4, 14, 14, width=1.7)
        line(17, 7, 21, 7)
        line(19, 5, 21, 7, 19, 9)
        line(7, 17, 7, 21)
        line(5, 19, 7, 21, 9, 19)
    elif name == "views":
        rect(4, 4, 10, 10, width=1.6)
        rect(14, 4, 20, 10, width=1.6)
        rect(4, 14, 10, 20, width=1.6)
        rect(14, 14, 20, 20, width=1.6)
    elif name == "scale":
        line(4, 18, 20, 18)
        line(7, 18, 17, 6)
        line(14, 6, 17, 6, 17, 9)
    elif name == "log":
        rect(6, 3, 18, 21, width=1.6)
        line(9, 8, 15, 8, width=1.4)
        line(9, 12, 15, 12, width=1.4)
        line(9, 16, 14, 16, width=1.4)
    elif name == "clear":
        rect(7, 8, 17, 20, width=1.6)
        line(5.5, 6, 18.5, 6)
        line(10, 4, 14, 4)
    else:
        oval(5, 5, 19, 19)


class IconButton(tk.Frame):
    def __init__(self, parent, text, icon, command=None, variant="white", width=160):
        self.variant = variant
        self.command = command
        self.state = "normal"
        self.base_width = width
        super().__init__(parent, height=42, width=width, bd=1, relief="solid", cursor="hand2")
        self.pack_propagate(False)
        self.inner = tk.Frame(self)
        self.inner.place(relx=0.5, rely=0.5, anchor="center")
        self.icon_canvas = tk.Canvas(self.inner, width=22, height=22, highlightthickness=0)
        self.icon_canvas.pack(side="left", padx=(0, 9))
        self.text_label = tk.Label(self.inner, text=text, font=("Segoe UI", 10, "bold"))
        self.text_label.pack(side="left")
        self.icon_name = icon
        for widget in (self, self.inner, self.icon_canvas, self.text_label):
            widget.bind("<Button-1>", self._click)
            widget.bind("<Enter>", self._enter)
            widget.bind("<Leave>", self._leave)
        self.apply_style()

    def colors(self, hover=False):
        if self.state == "disabled":
            return "#f4f6fa", "#d8e1ee", "#9aa7ba"
        if self.variant == "blue":
            return ("#244fbd" if hover else UI["blue"]), UI["blue_dark"], "#ffffff"
        if self.variant == "yellow":
            return ("#d9ba1e" if hover else UI["yellow"]), UI["yellow_dark"], "#221d08"
        return ("#f7faff" if hover else "#ffffff"), "#d4deec", UI["text"]

    def apply_style(self, hover=False):
        bg, border, fg = self.colors(hover)
        self.configure(bg=bg, highlightbackground=border)
        self.inner.configure(bg=bg)
        self.icon_canvas.configure(bg=bg)
        self.text_label.configure(bg=bg, fg=fg)
        draw_line_icon(self.icon_canvas, self.icon_name, fg)
        self.configure(cursor="arrow" if self.state == "disabled" else "hand2")

    def _click(self, event=None):
        if self.state != "disabled" and self.command:
            self.command()

    def _enter(self, event=None):
        self.apply_style(hover=True)

    def _leave(self, event=None):
        self.apply_style(hover=False)

    def configure(self, cnf=None, **kwargs):
        state = kwargs.pop("state", None)
        result = super().configure(cnf or {}, **kwargs)
        if state is not None:
            self.state = state
            self.apply_style()
        return result

    config = configure


class TeklaApi:
    def __init__(self, log):
        self.log = log
        self.initialized = False
        self.tekla_root = None
        self.bin_dir = None
        self.net48_dir = None
        self.DrawingHandler = None
        self.View = None
        self.Text = None
        self.TextElement = None
        self.PropertyElement = None
        self.ContainerElement = None
        self.MarkBase = None
        self.Line = None
        self.LineTypes = None
        self.DrawingColors = None
        self.RadiusDimension = None
        self.StraightDimension = None
        self.StraightDimensionSet = None
        self.StraightDimensionSetHandler = None
        self.PointList = None
        self.TeklaPoint = None
        self.TeklaVector = None
        self.Model = None
        self.DrawingModelObject = None
        self.DrawingLink = None
        self.Size = None
        self.dimension_top_gap_mm = DIMENSION_TOP_GAP_MM
        self.last_sheet_w = 0.0
        self.last_sheet_h = 0.0
        self.last_sheet_object_count = None
        self.last_sheet_plan = []
        self.dimension_view_by_key = {}
        # Progresso real: o controller define progress_cb; as operacoes longas
        # chamam self.report_progress(local, texto). A "janela" (inicio/fim) permite
        # que cada fase reporte 0..100 localmente e seja mapeada para uma faixa
        # global (ex.: analisar = 0..35, medir croquis = 35..90).
        self.progress_cb = None
        self._prog_base = 0.0
        self._prog_span = 100.0

    def set_progress_window(self, start, end):
        """Define a faixa global (start..end, em %) que o proximo 0..100 local ocupa."""
        self._prog_base = float(start)
        self._prog_span = max(0.0, float(end) - float(start))

    def report_progress(self, local, text=None):
        cb = getattr(self, "progress_cb", None)
        if not cb:
            return
        local = max(0.0, min(100.0, float(local or 0.0)))
        absolute = self._prog_base + self._prog_span * local / 100.0
        try:
            cb(absolute, text)
        except Exception:
            pass

    def initialize(self):
        prepare_windows_thread_for_tekla()
        if self.initialized:
            return
        self.tekla_root = self.find_tekla_root()
        self.bin_dir = self.tekla_root / "bin"
        direct_net48 = self.bin_dir / "Net48Runtime"
        if direct_net48.exists():
            self.net48_dir = direct_net48
        else:
            found = self.find_file_under(self.tekla_root, "Tekla.Structures.dll")
            self.net48_dir = found.parent if found else direct_net48

        # Necessário para DLLs nativas/dependências do Tekla.
        for p in self.build_probe_dirs():
            try:
                os.add_dll_directory(str(p))
            except Exception:
                pass
            if str(p) not in sys.path:
                sys.path.insert(0, str(p))

        try:
            import pythonnet
            try:
                pythonnet.load("netfx")
            except Exception:
                # Se o runtime já tiver sido carregado, segue normalmente.
                pass
        except Exception:
            pass

        try:
            import clr
        except Exception as ex:
            raise RuntimeError("pythonnet nao esta instalado. Execute 'Instalar dependencias.bat'. Erro: " + str(ex))

        # Carrega primeiro dependências comuns, depois Tekla.
        for dll in [
            self.net48_dir / "Tekla.Structures.dll",
            self.bin_dir / "Tekla.Structures.dll",
            self.bin_dir / "Tekla.Structures.Drawing.dll",
            self.bin_dir / "Tekla.Structures.Model.dll",
            self.bin_dir / "Tekla.Structures.Datatype.dll",
            self.net48_dir / "Tekla.Structures.Plugins.dll",
            self.bin_dir / "Trimble.Remoting.dll",
            self.bin_dir / "DotNetKit.dll",
            self.bin_dir / "Tekla.Technology.Serialization.dll",
        ]:
            if dll.exists():
                try:
                    clr.AddReference(str(dll))
                except Exception:
                    pass

        try:
            clr.AddReference("Tekla.Structures")
            clr.AddReference("Tekla.Structures.Drawing")
            clr.AddReference("Tekla.Structures.Model")
        except Exception:
            pass

        try:
            from Tekla.Structures.Drawing import (
                DrawingHandler, View, Text, TextElement, PropertyElement, ContainerElement, MarkBase,
                Line, RadiusDimension, StraightDimension, StraightDimensionSet, StraightDimensionSetHandler, PointList,
                LineTypes, DrawingColors, DrawingLink, Size,
                ModelObject as DrawingModelObject
            )
            from Tekla.Structures.Geometry3d import Point as TeklaPoint, Vector as TeklaVector
            from Tekla.Structures.Model import Model
        except Exception as ex:
            raise RuntimeError("Falha ao carregar namespaces Tekla. Verifique tekla-root.txt e a versao do Tekla. Erro: " + str(ex))

        self.DrawingHandler = DrawingHandler
        self.View = View
        self.Text = Text
        self.TextElement = TextElement
        self.PropertyElement = PropertyElement
        self.ContainerElement = ContainerElement
        self.MarkBase = MarkBase
        self.Line = Line
        self.LineTypes = LineTypes
        self.DrawingColors = DrawingColors
        self.RadiusDimension = RadiusDimension
        self.StraightDimension = StraightDimension
        self.StraightDimensionSet = StraightDimensionSet
        self.StraightDimensionSetHandler = StraightDimensionSetHandler
        self.PointList = PointList
        self.TeklaPoint = TeklaPoint
        self.TeklaVector = TeklaVector
        self.Model = Model
        self.DrawingModelObject = DrawingModelObject
        self.DrawingLink = DrawingLink
        self.Size = Size
        self.initialized = True
        self.log(f"Tekla API carregada: {self.tekla_root}")

    def build_probe_dirs(self):
        dirs = []
        for p in [self.bin_dir, self.net48_dir, self.tekla_root, self.tekla_root / "nt" / "bin"]:
            if p and p.exists() and p not in dirs:
                dirs.append(p)
        # inclui subpastas mais prováveis, sem varrer demais.
        for parent in [self.bin_dir, self.net48_dir]:
            if parent and parent.exists():
                for child in parent.iterdir():
                    if child.is_dir() and child not in dirs:
                        dirs.append(child)
        return dirs

    def find_file_under(self, root: Path, filename: str):
        try:
            for p in root.rglob(filename):
                return p
        except Exception:
            return None
        return None

    def find_tekla_root(self) -> Path:
        candidates = []
        manual_files = [BASE_DIR / "tekla-root.txt", BASE_DIR / "tekla_path.txt", BASE_DIR / "tekla-install.txt"]
        for mf in manual_files:
            if mf.exists():
                for line in mf.read_text(encoding="utf-8", errors="ignore").splitlines():
                    line = line.strip().strip('"')
                    if line and not line.startswith("#"):
                        candidates.append(Path(line))
        for env in ["TEKLA_STRUCTURES_ROOT", "TEKLA_ROOT", "TEKLA_INSTALLATION_ROOT"]:
            val = os.environ.get(env)
            if val:
                candidates.append(Path(val))
        for base in [r"C:\Program Files\Tekla Structures", r"C:\TeklaStructures", r"D:\TeklaStructures", r"D:\Tekla Structures"]:
            b = Path(base)
            for ver in ["2026.0", "2025.0", "2024.0"]:
                candidates.append(b / ver)
        for c in candidates:
            c = self.normalize_root(c)
            if self.is_valid_root(c):
                return c
        checked = "\n".join(str(self.normalize_root(c)) for c in candidates[:20])
        raise RuntimeError("Nao encontrei a instalacao do Tekla. Informe a pasta em tekla-root.txt. Candidatos verificados:\n" + checked)

    def normalize_root(self, p: Path) -> Path:
        p = Path(str(p).strip().strip('"'))
        if p.name.lower() == "bin":
            return p.parent
        return p

    def is_valid_root(self, p: Path) -> bool:
        if not p or not p.exists():
            return False
        return (p / "bin" / "Tekla.Structures.Drawing.dll").exists() or (p / "bin" / "Net48Runtime" / "Tekla.Structures.dll").exists()

    def get_active_sheet(self):
        prepare_windows_thread_for_tekla()
        self.initialize()
        handler = None
        last_error = ""
        for attempt in range(8):
            if attempt:
                time.sleep(0.35)
            try:
                try:
                    model = self.Model()
                    model.GetConnectionStatus()
                except Exception:
                    pass
                handler = self.DrawingHandler()
                if handler.GetConnectionStatus():
                    break
                last_error = "DrawingHandler.GetConnectionStatus retornou False."
            except Exception as ex:
                handler = None
                last_error = str(ex)
        if handler is None or not handler.GetConnectionStatus():
            detail = f" Detalhe: {last_error}" if last_error else ""
            raise RuntimeError(
                "Nao foi possivel conectar ao Tekla. Confira se o Tekla esta aberto com um modelo carregado "
                "e um desenho ativo. Se o Tekla foi aberto depois do organizador, feche e abra o organizador novamente."
                + detail
            )
        if not handler.IsAnyDrawingOpen():
            raise RuntimeError("Nenhum desenho esta aberto no Tekla.")
        drawing = handler.GetActiveDrawing()
        if drawing is None:
            raise RuntimeError("Nao foi possivel obter o desenho ativo.")
        sheet = drawing.GetSheet()
        if sheet is None:
            raise RuntimeError("Nao foi possivel acessar a folha do desenho ativo.")
        return handler, drawing, sheet

    def sheet_size(self, drawing):
        try:
            return float(drawing.Layout.SheetSize.Width), float(drawing.Layout.SheetSize.Height)
        except Exception:
            return 841.0, 594.0

    def is_view_object(self, obj):
        if obj is None:
            return False
        try:
            if self.View is not None and isinstance(obj, self.View):
                return True
        except Exception:
            pass
        try:
            tname = self.object_type_name(obj)
            if tname == "View":
                return True
            if "View" in tname:
                for attr in ("Origin", "FrameOrigin", "Width", "Height", "GetAllObjects"):
                    if not hasattr(obj, attr):
                        return False
                return True
            return False
        except Exception:
            return False

    def is_drawing_link_object(self, obj):
        if obj is None:
            return False
        try:
            if self.DrawingLink is not None and isinstance(obj, self.DrawingLink):
                return True
        except Exception:
            pass
        try:
            return self.object_type_name(obj) == "DrawingLink"
        except Exception:
            return False

    def collect_views(self, sheet):
        out = []

        def add_view(obj):
            if not self.is_view_object(obj):
                return
            key = self.object_stable_key(obj)
            if key in seen:
                return
            seen.add(key)
            out.append(obj)

        def scan_enum(enum):
            while enum.MoveNext():
                add_view(enum.Current)

        seen = set()
        get_all_views_count = 0
        try:
            enum = sheet.GetAllViews()
            while enum.MoveNext():
                obj = enum.Current
                add_view(obj)
                get_all_views_count += 1
        except Exception:
            pass

        # Algumas situações do Tekla retornam 0 em GetAllViews() mesmo com o
        # multidrawing visível. Neste caso, varre os objetos da folha e coleta
        # diretamente os objetos cujo tipo real é View.
        if not out:
            try:
                scan_enum(sheet.GetAllObjects())
            except Exception:
                pass

        if not out:
            try:
                scan_enum(sheet.GetObjects())
            except Exception:
                pass

        if get_all_views_count == 0 and out:
            self.log(f"GetAllViews retornou 0; fallback encontrou {len(out)} View(s) na folha.")
        return out

    def collect_drawing_links(self, sheet):
        out = []
        seen = set()

        def add_link(obj):
            if not self.is_drawing_link_object(obj):
                return
            key = self.object_stable_key(obj)
            if key in seen:
                return
            seen.add(key)
            out.append(obj)

        for getter_name in ("GetAllObjects", "GetObjects"):
            try:
                enum = getattr(sheet, getter_name)()
                while enum.MoveNext():
                    add_link(enum.Current)
            except Exception:
                pass
        return out

    def count_sheet_objects(self, sheet):
        try:
            enum = sheet.GetAllObjects()
            count = 0
            while enum.MoveNext():
                count += 1
            return count
        except Exception:
            return None

    def get_view_frame(self, view):
        try:
            box = view.GetAxisAlignedBoundingBox()
            if box is not None and box.LowerLeft is not None and box.UpperRight is not None:
                return TeklaFrame(float(box.LowerLeft.X), float(box.LowerLeft.Y), float(box.UpperRight.X), float(box.UpperRight.Y))
        except Exception:
            pass
        left = float(view.Origin.X + view.FrameOrigin.X)
        bottom = float(view.Origin.Y + view.FrameOrigin.Y)
        return TeklaFrame(left, bottom, left + float(view.Width), bottom + float(view.Height))

    def get_view_content_frame(self, view, frame):
        """União dos bboxes dos objetos visíveis da View (desenho, cotas, rótulo).

        O frame da View no Tekla costuma ter folga vazia; este retângulo mede só o
        conteúdo desenhado, para o empilhamento usar a distância visual real.
        Retorna None quando não conseguir medir.
        """
        left = bottom = right = top = None
        try:
            enum = view.GetAllObjects()
            while enum.MoveNext():
                box = self.get_object_bbox(enum.Current)
                if box is None:
                    continue
                # Ignora bboxes fora do frame (objetos com extensão inválida).
                b_left = max(box.left, frame.left)
                b_right = min(box.right, frame.right)
                b_bottom = max(box.bottom, frame.bottom)
                b_top = min(box.top, frame.top)
                if b_right < b_left or b_top < b_bottom:
                    continue
                left = b_left if left is None else min(left, b_left)
                right = b_right if right is None else max(right, b_right)
                bottom = b_bottom if bottom is None else min(bottom, b_bottom)
                top = b_top if top is None else max(top, b_top)
        except Exception:
            return None
        if left is None or right - left <= 0.1 or top - bottom <= 0.1:
            return None
        return TeklaFrame(left, bottom, right, top)

    def get_view_scale(self, view):
        try:
            scale = float(view.Attributes.Scale)
            return scale if scale > 0.0 else 1.0
        except Exception:
            return 1.0

    def set_view_scale(self, view, new_scale):
        new_scale = float(new_scale)
        if new_scale <= 0.0:
            raise RuntimeError("Escala invalida.")
        attr = view.Attributes
        old_scale = self.get_view_scale(view)
        attr.Scale = new_scale
        view.Attributes = attr
        if not view.Modify():
            raise RuntimeError("Tekla recusou a alteracao de escala.")
        return old_scale

    def safe_name(self, view):
        try:
            return str(view.Name or "")
        except Exception:
            return ""

    def add_if_text(self, values, value):
        value = norm_spaces(str(value or ""))
        if value and value.lower() not in {v.lower() for v in values}:
            values.append(value)

    def add_object_string_candidates(self, obj, values, depth=0, seen=None):
        """
        Coleta textos de objetos Tekla de forma tolerante ao pythonnet.
        O rótulo PEÇA pode estar em Text, TextElement, PropertyElement, ContainerElement,
        MarkBase, TagContent ou em propriedades internas expostas por reflexão.
        """
        if obj is None or depth > 6:
            return
        if seen is None:
            seen = set()
        try:
            oid = id(obj)
            if oid in seen:
                return
            seen.add(oid)
        except Exception:
            pass

        # Propriedades com texto direto.
        for prop in [
            "TextString", "Value", "Name", "Content", "FormattedText", "Text",
            "TagContent", "TextContent", "PropertyName", "DisplayName"
        ]:
            try:
                val = getattr(obj, prop)
                if isinstance(val, str):
                    self.add_if_text(values, val)
                elif val is not None and val is not obj:
                    self.add_object_string_candidates(val, values, depth + 1, seen)
            except Exception:
                pass

        # Métodos comuns de ContainerElement/TextElement.
        for meth in ["GetUnformattedString", "GetFormattedString", "ToString"]:
            try:
                fn = getattr(obj, meth)
                val = fn()
                if isinstance(val, str):
                    self.add_if_text(values, val)
            except Exception:
                pass

        # Tipos conhecidos.
        try:
            if self.ContainerElement is not None and isinstance(obj, self.ContainerElement):
                self.iter_container_elements(obj, values, depth, seen)
        except Exception:
            pass
        try:
            if self.MarkBase is not None and isinstance(obj, self.MarkBase):
                enum = obj.GetObjects()
                while enum.MoveNext():
                    self.add_object_string_candidates(enum.Current, values, depth + 1, seen)
        except Exception:
            pass

        # Reflexão por nome do tipo quando isinstance falhar no pythonnet.
        try:
            tname = obj.GetType().Name
            if tname in {"ContainerElement", "TextElement", "PropertyElement", "Text", "Mark", "ViewMark", "MarkBase"}:
                self.iter_container_elements(obj, values, depth, seen)
        except Exception:
            pass

        # Reflexão leve: algumas versões expõem TagContent/Value apenas como PropertyInfo.
        try:
            props = obj.GetType().GetProperties()
            for pr in props:
                try:
                    pname = str(pr.Name)
                    if pname not in {"TextString", "Value", "Name", "Content", "FormattedText", "Text", "TagContent"}:
                        continue
                    val = pr.GetValue(obj, None)
                    if isinstance(val, str):
                        self.add_if_text(values, val)
                    elif val is not None and val is not obj:
                        self.add_object_string_candidates(val, values, depth + 1, seen)
                except Exception:
                    pass
        except Exception:
            pass

    def iter_container_elements(self, obj, values, depth=0, seen=None):
        if obj is None or depth > 6:
            return
        # Python iteration over .NET containers.
        try:
            for el in obj:
                self.add_object_string_candidates(el, values, depth + 1, seen)
            return
        except Exception:
            pass
        # Explicit .NET enumerator.
        try:
            enum = obj.GetEnumerator()
            while enum.MoveNext():
                self.add_object_string_candidates(enum.Current, values, depth + 1, seen)
            return
        except Exception:
            pass
        # Alguns containers têm método GetObjects.
        try:
            enum = obj.GetObjects()
            while enum.MoveNext():
                self.add_object_string_candidates(enum.Current, values, depth + 1, seen)
        except Exception:
            pass

    def add_view_attribute_candidates(self, view, values):
        """Lê tags do título da View. Em muitas versões do Tekla o texto PEÇA fica aqui."""
        try:
            tags = view.Attributes.TagsAttributes
        except Exception:
            tags = None
        if tags is None:
            return

        # Caminho conhecido.
        for tag_name in ["TagA1", "TagA2", "TagA3", "TagA4", "TagA5", "TagB1", "TagB2", "TagB3", "TagB4", "TagB5"]:
            try:
                tag = getattr(tags, tag_name)
                self.add_object_string_candidates(tag, values)
                try:
                    self.add_object_string_candidates(tag.TagContent, values)
                except Exception:
                    pass
            except Exception:
                pass

        # Reflexão para versões/localizações diferentes.
        try:
            props = tags.GetType().GetProperties()
            for pr in props:
                try:
                    pname = str(pr.Name)
                    if "Tag" not in pname:
                        continue
                    tag = pr.GetValue(tags, None)
                    self.add_object_string_candidates(tag, values)
                    try:
                        self.add_object_string_candidates(tag.TagContent, values)
                    except Exception:
                        pass
                except Exception:
                    pass
        except Exception:
            pass

    def add_view_object_text_candidates(self, view, values):
        try:
            enum = view.GetAllObjects()
            while enum.MoveNext():
                self.add_object_string_candidates(enum.Current, values)
        except Exception:
            pass

    def stack_group_from_values(self, values):
        for value in values or []:
            parsed = try_parse_drawing_view_group(value)
            if parsed:
                return parsed
        return None

    def apply_stack_group(self, item, values):
        parsed = self.stack_group_from_values(values)
        if not parsed:
            return item
        group_text, group_sort_key, view_order_key = parsed
        item.stack_group_key = group_text
        item.stack_group_sort_key = group_sort_key
        item.stack_order_key = view_order_key
        return item

    def get_object_bbox(self, obj):
        try:
            box = obj.GetAxisAlignedBoundingBox()
            if box is not None and box.LowerLeft is not None and box.UpperRight is not None:
                left = float(min(box.LowerLeft.X, box.UpperRight.X))
                right = float(max(box.LowerLeft.X, box.UpperRight.X))
                bottom = float(min(box.LowerLeft.Y, box.UpperRight.Y))
                top = float(max(box.LowerLeft.Y, box.UpperRight.Y))
                return TeklaFrame(left, bottom, right, top)
        except Exception:
            pass
        return None

    def get_object_center(self, obj):
        box_frame = self.get_object_bbox(obj)
        if box_frame is not None:
            return ((box_frame.left + box_frame.right) / 2.0, (box_frame.bottom + box_frame.top) / 2.0)
        for prop in ["InsertionPoint"]:
            try:
                pt = getattr(obj, prop)
                if pt is not None:
                    return float(pt.X), float(pt.Y)
            except Exception:
                pass
        return None

    def collect_piece_labels(self, sheet):
        labels = []

        def scan_enum(enum):
            while enum.MoveNext():
                obj = enum.Current
                vals = []
                self.add_object_string_candidates(obj, vals)
                parsed = try_parse_piece_from_strings(vals)
                if not parsed:
                    continue
                center = self.get_object_center(obj)
                if not center:
                    continue
                bbox = self.get_object_bbox(obj)
                label, item_text, key = parsed
                if bbox is not None:
                    labels.append(PieceLabelInfo(label, item_text, key, center[0], center[1], bbox.left, bbox.bottom, bbox.right, bbox.top))
                else:
                    labels.append(PieceLabelInfo(label, item_text, key, center[0], center[1]))

        try:
            scan_enum(sheet.GetAllObjects())
        except Exception:
            pass
        # Muitos rótulos PEÇA pertencem à própria View, não ao Sheet.
        for view in self.collect_views(sheet):
            try:
                scan_enum(view.GetAllObjects())
            except Exception:
                pass
        # remove duplicados pelo item + centro aproximado
        dedup = {}
        for l in labels:
            k = (l.item_text, round(l.center_x, 1), round(l.center_y, 1))
            dedup[k] = l
        return list(dedup.values())

    def find_nearest_piece_label(self, labels, frame):
        if not labels:
            return None
        cx = (frame.left + frame.right) / 2.0
        lower_y = frame.bottom - max(80.0, frame.height * 0.50)
        upper_y = frame.top + 20.0
        left_x = frame.left - 35.0
        right_x = frame.right + 35.0
        candidates = [l for l in labels if left_x <= l.center_x <= right_x and lower_y <= l.center_y <= upper_y]
        if not candidates:
            return None
        return sorted(candidates, key=lambda l: abs(l.center_x - cx) + abs(l.center_y - frame.bottom) * 0.35)[0]

    def try_get_model_string_report_property(self, model_object, property_name):
        """Obtém string de GetReportProperty corretamente via ref no pythonnet."""
        # Forma correta no pythonnet: clr.Reference[String]().
        try:
            import clr
            from System import String
            ref_val = clr.Reference[String]()
            ok = model_object.GetReportProperty(property_name, ref_val)
            if ok and ref_val.Value:
                return str(ref_val.Value)
        except Exception:
            pass
        # Fallback para wrappers que retornam tupla/lista.
        try:
            res = model_object.GetReportProperty(property_name, "")
            if isinstance(res, (tuple, list)) and len(res) >= 2 and res[0] and res[1]:
                return str(res[1])
            if isinstance(res, str) and res:
                return res
        except Exception:
            pass
        return ""

    def model_position_from_view(self, view):
        """
        Lê o número da peça pelo objeto de modelo da View.
        Este é o fallback mais importante quando o título visual aparece como VISTA_AUTO.
        """
        try:
            model = self.Model()
            if not model.GetConnectionStatus():
                return None
            found = {}
            enum = view.GetModelObjects()
            while enum.MoveNext():
                dobj = enum.Current
                try:
                    identifier = dobj.ModelIdentifier
                except Exception:
                    continue
                if identifier is None:
                    continue
                mobj = model.SelectModelObject(identifier)
                if mobj is None:
                    continue
                for prop, score in [
                    ("PART_POS", 120), ("ASSEMBLY_POS", 60), ("CAST_UNIT_POS", 60),
                    ("MAIN_PART_POS", 50), ("PART_POSITION", 40), ("POSITION", 20), ("PREFIX", 5)
                ]:
                    val = self.try_get_model_string_report_property(mobj, prop)
                    if not val:
                        continue
                    parsed = try_parse_numeric_item(str(val))
                    if not parsed:
                        continue
                    item_text, key = parsed
                    # Ignora números soltos grandes demais que não parecem item de lista.
                    if len(key) < 2:
                        continue
                    if item_text not in found:
                        found[item_text] = [item_text, key, 0, []]
                    found[item_text][2] += score
                    found[item_text][3].append(f"{prop}={val}")
            if not found:
                return None
            best = sorted(found.values(), key=lambda x: (-x[2], numeric_key(x[1])))[0]
            return "PEÇA " + best[0], best[0], best[1]
        except Exception:
            return None

    def try_create_piece_item(self, view, seq, labels):
        try:
            if bool(view.IsSheet):
                return None
        except Exception:
            pass
        frame = self.get_view_frame(view)
        if frame.width <= 0.1 or frame.height <= 0.1:
            return None
        candidates = []
        self.add_if_text(candidates, self.safe_name(view))
        self.add_view_attribute_candidates(view, candidates)
        self.add_view_object_text_candidates(view, candidates)
        parsed_joined = try_parse_piece_from_strings(candidates)
        if parsed_joined:
            name, item_text, key = parsed_joined
            return self.apply_stack_group(ApiViewItem(seq, name, item_text, key, view, frame), candidates)
        model_based = self.model_position_from_view(view)
        if model_based:
            name, item_text, key = model_based
            return self.apply_stack_group(ApiViewItem(seq, name, item_text, key, view, frame), candidates)
        nearest = self.find_nearest_piece_label(labels, frame)
        if nearest:
            item = ApiViewItem(seq, nearest.label, nearest.item_text, nearest.sort_key, view, frame)
            item.label_center_y = nearest.center_y
            item.label_offset_y = nearest.center_y - frame.bottom
            item.label_bottom_y = self.label_info_bottom(nearest)
            return self.apply_stack_group(item, candidates + [nearest.label])
        return None

    def analyze(self, cfg):
        self.report_progress(3, "Lendo a folha ativa...")
        # O plano por folhas pertence somente à análise atual.
        self.last_sheet_plan = []
        self.last_sheet_object_count = None
        handler, drawing, sheet = self.get_active_sheet()
        self.ensure_capacity_project_context(cfg, drawing)
        sheet_w, sheet_h = self.sheet_size(drawing)
        self.last_sheet_w = sheet_w
        self.last_sheet_h = sheet_h
        labels = []
        views = []
        links = []
        for attempt in range(1, 7):
            self.report_progress(8, "Lendo rótulos PEÇA...")
            labels = self.collect_piece_labels(sheet)
            self.report_progress(16, "Localizando vistas...")
            views = self.collect_views(sheet)
            links = self.collect_drawing_links(sheet)
            if views or links or labels:
                break
            if attempt < 6:
                self.log(f"A API ainda nao retornou vistas; nova tentativa {attempt + 1}/6...")
                time.sleep(0.55)
                try:
                    handler, drawing, sheet = self.get_active_sheet()
                    sheet_w, sheet_h = self.sheet_size(drawing)
                    self.last_sheet_w = sheet_w
                    self.last_sheet_h = sheet_h
                except Exception:
                    pass
        self.log(f"Folha Tekla: {sheet_w:.0f} x {sheet_h:.0f}.")
        self.log(f"Vistas localizadas pela API: {len(views)}.")
        if links:
            self.log(f"Links de desenho localizados pela API: {len(links)}.")
        self.log(f"Rótulos PEÇA localizados: {len(labels)}.")
        if not views and not links and not labels:
            self.last_sheet_object_count = self.count_sheet_objects(sheet)
            if self.last_sheet_object_count is not None:
                self.log(f"Objetos na folha localizados pela API: {self.last_sheet_object_count}.")
        items = []
        raw_non_sheet_views = []
        seq = 1
        ignored = 0
        total_views = max(1, len(views))
        for view_index, view in enumerate(views, start=1):
            # Faixa 20..75: leitura de cada vista (parte mais custosa da analise).
            self.report_progress(20 + 55.0 * view_index / total_views, "Lendo vistas...")
            try:
                if bool(view.IsSheet):
                    ignored += 1
                    continue
            except Exception:
                pass
            try:
                frame = self.get_view_frame(view)
                if frame.width > 0.1 and frame.height > 0.1:
                    raw_non_sheet_views.append((view, frame))
            except Exception:
                pass
            item = self.try_create_piece_item(view, seq, labels)
            if item:
                items.append(item)
                seq += 1
            else:
                ignored += 1

        if not items and links:
            for link in links:
                item = self.item_from_drawing_link(link, seq)
                if item:
                    items.append(item)
                    seq += 1
                else:
                    ignored += 1

        # Se os objetos da View não entregarem o rótulo, mas os textos PEÇA foram localizados
        # na folha/vistas, associa cada View ao rótulo mais próximo. Assim a organização fica
        # em ordem crescente real: PEÇA 1.1, 1.2, 1.3 ... 5.1, 5.2 ...
        if not items and raw_non_sheet_views and labels:
            assigned = self.assign_labels_to_views(raw_non_sheet_views, labels)
            if assigned:
                self.log(f"Rótulos PEÇA associados por proximidade: {len(assigned)}.")
                items = []
                for idx, (view, frame, label) in enumerate(assigned, start=1):
                    item = ApiViewItem(idx, label.label, label.item_text, label.sort_key, view, frame, status="PEÇA por proximidade")
                    item.label_center_y = label.center_y
                    item.label_offset_y = label.center_y - frame.bottom
                    item.label_bottom_y = self.label_info_bottom(label)
                    self.apply_stack_group(item, [label.label, self.safe_name(view)])
                    items.append(item)
                ignored = max(0, len(views) - len(items))

        # Último fallback: organiza visualmente, mas deixa claro que não há ordenação PEÇA real.
        if not items and raw_non_sheet_views:
            self.log("Fallback ativado: a API encontrou Views, mas não leu rótulo PEÇA. Organizando por posição visual.")
            ordered_raw = self.sort_raw_views_visual(raw_non_sheet_views)
            items = []
            for idx, (view, frame) in enumerate(ordered_raw, start=1):
                name = self.safe_name(view) or f"VISTA_AUTO_{idx:03d}"
                item_text = str(idx)
                key = [idx]
                items.append(ApiViewItem(
                    idx,
                    name,
                    item_text,
                    key,
                    view,
                    frame,
                    status="Fallback visual",
                    piece_identity_trusted=False,
                ))
            ignored = max(0, len(views) - len(items))

        if not items and self.object_type_name(drawing) == "MultiDrawing":
            selected_drawings = self.collect_selected_single_part_drawings(handler)
            if len(selected_drawings) == 1:
                selected_drawings = []
            source_label = "selecionado(s) no Document Manager"
            if not selected_drawings:
                all_drawings = self.collect_all_single_part_drawings(handler)
                cursor_item, cursor_sheet = self.capacity_cursor_for_multidrawing(cfg, drawing)
                if cursor_item:
                    selected_drawings = self.single_part_drawings_after(all_drawings, cursor_item)
                    source_label = f"da fila automatica apos PEÇA {cursor_item}"
                    if cursor_sheet:
                        self.log(f"Fila automatica: folha anterior {cursor_sheet}; continuando apos PEÇA {cursor_item}.")
                else:
                    selected_drawings = all_drawings
                    source_label = "disponiveis no Document Manager"
            selected_drawings = self.sorted_single_part_drawings(selected_drawings)
            total_selected = len(selected_drawings)
            if selected_drawings:
                self.log(
                    f"Multidrawing sem vistas/links; calculando capacidade com {total_selected} "
                    f"croqui(s) {source_label}."
                )
                self.report_progress(20, "Calculando capacidade dos desenhos...")
                all_items = self.build_capacity_items_from_drawings(
                    selected_drawings,
                    handler=handler,
                    original_drawing=drawing,
                )
                start_mark = ""
                try:
                    start_mark = str(drawing.Mark or "").strip()
                except Exception:
                    pass
                factor_matches = int(float(getattr(cfg, "capacity_factor_matches", 0) or 0))
                factor_x = max(1.0, float(getattr(cfg, "capacity_factor_x", 1.0) or 1.0)) if factor_matches >= 3 else 1.0
                factor_y = max(1.0, float(getattr(cfg, "capacity_factor_y", 1.0) or 1.0)) if factor_matches >= 3 else 1.0
                self.last_sheet_plan = self.build_capacity_sheet_plan(
                    all_items,
                    cfg,
                    sheet_w,
                    sheet_h,
                    start_mark,
                    factor_x=factor_x,
                    factor_y=factor_y,
                )
                if factor_matches >= 3:
                    self.log(
                        f"Capacidade calibrada por {factor_matches} vista(s) reais "
                        f"(X={factor_x:.3f}, Y={factor_y:.3f})."
                    )
                fit_count = int(self.last_sheet_plan[0]["count"]) if self.last_sheet_plan else 0
                items = all_items[:fit_count] if fit_count > 0 else []
                if fit_count > 0:
                    self.log(f"Capacidade calculada: cabem {fit_count} de {total_selected} croqui(s).")
                if fit_count < len(all_items):
                    self.log(f"Primeiro croqui da proxima folha: PECA {all_items[fit_count].item_text}.")
                if self.last_sheet_plan:
                    resumo = "; ".join(
                        f"{row['label']}: {row['start']} a {row['end']} ({row['count']})"
                        for row in self.last_sheet_plan[:8]
                    )
                    if len(self.last_sheet_plan) > 8:
                        resumo += " ..."
                    self.log("Plano por folhas: " + resumo)
                if False:
                    self.log(f"Primeiro croqui fora da capacidade: PEÇA {items[fit_count].item_text}.")
                elif False:
                    self.log(f"Primeiro croqui fora da capacidade: PEÇA {first_failed.item_text}.")
                for item in items:
                    item.status = "Planejada para inserir"
                ignored = max(0, total_selected - len(items))
                try:
                    handler, drawing, sheet = self.get_active_sheet()
                except Exception:
                    pass

        items.sort(key=lambda i: (numeric_key(i.sort_key), i.name.lower()))

        if items:
            sample_order = ", ".join(i.item_text for i in items[:12])
            self.log("Ordem PEÇA aplicada: " + sample_order + (" ..." if len(items) > 12 else ""))
            if any((i.name or "").upper().startswith("VISTA_AUTO") for i in items):
                self.log("Aviso: algumas vistas ainda foram nomeadas como VISTA_AUTO; a ordem pode estar visual se a PEÇA não foi lida.")

        self.report_progress(80, "Alinhando rótulos...")
        self.attach_label_offsets(items, labels)
        self.attach_stack_content_padding(items)

        self.report_progress(90, "Planejando layout...")
        for idx, item in enumerate(items, start=1):
            item.index = idx
            if item.view is not None and not self.is_drawing_link_object(item.view):
                item.current_scale = self.get_view_scale(item.view)
        # "Máx. vistas" é uma referência calculada pela capacidade, não um corte rígido.
        # Todas as vistas reais do multidrawing devem ser testadas. Assim nenhuma peça é
        # descartada apenas porque a estimativa anterior foi menor ou imprecisa.
        target_items = items
        if cfg.max_details and int(cfg.max_details) > 0 and len(items) > int(cfg.max_details):
            self.log(
                f"Referência de capacidade: {int(cfg.max_details)} vista(s); "
                f"testando as {len(items)} vistas reais enquanto houver espaço."
            )
        self.plan_layout(target_items, sheet_w, sheet_h, cfg)
        self.report_progress(100, "Análise concluída.")
        planned_count = len([item for item in target_items if item.has_target])
        without_space = max(0, len(target_items) - planned_count)
        self.log(
            f"Vistas PEÇA encontradas: {len(items)}. Planejadas: {planned_count}. "
            f"Sem espaço: {without_space}. Objetos ignorados na leitura: {ignored}."
        )
        if not items:
            sample = []
            for v in views[:8]:
                n = self.safe_name(v)
                if n:
                    sample.append(n)
            if sample:
                self.log("Amostra de vistas ignoradas: " + " | ".join(sample))
        return items, drawing, handler

    def label_info_bottom(self, label):
        """Bottom do bbox do rótulo PEÇA, ou None quando o bbox não foi medido.

        O PieceLabelInfo guarda bottom/top só quando a API entregou o bbox do
        texto; sem bbox, bottom/top ficam em 0.0 (não usáveis). Aqui devolvemos o
        bottom real apenas quando há altura de texto plausível.
        """
        try:
            bottom = float(getattr(label, "bottom", 0.0) or 0.0)
            top = float(getattr(label, "top", 0.0) or 0.0)
        except Exception:
            return None
        if top - bottom > 0.1:
            return bottom
        return None

    def attach_label_offsets(self, items, labels):
        """Associa a posição vertical do rótulo PEÇA à view.

        A organização já é feita pela ordem numérica da peça. Este passo é só para
        melhorar o alinhamento visual: em cada linha, o programa tenta deixar os
        rótulos PEÇA na mesma altura, em vez de alinhar apenas pelo frame/origem.
        """
        if not items or not labels:
            return
        by_text = {}
        for lab in labels:
            by_text.setdefault(lab.item_text, []).append(lab)

        used = set()
        found = 0
        for item in items:
            if item.label_offset_y is not None:
                found += 1
                continue
            candidates = by_text.get(item.item_text) or []
            if not candidates:
                continue
            cx = (item.current.left + item.current.right) / 2.0
            cy = (item.current.bottom + item.current.top) / 2.0
            ranked = []
            for lab in candidates:
                key = (lab.item_text, round(lab.center_x, 2), round(lab.center_y, 2))
                if key in used:
                    continue
                # Dá preferência ao rótulo mais próximo da view atual.
                dist = abs(lab.center_x - cx) + abs(lab.center_y - item.current.bottom) * 0.35 + abs(lab.center_y - cy) * 0.05
                ranked.append((dist, key, lab))
            if ranked:
                _, key, lab = sorted(ranked, key=lambda x: x[0])[0]
                used.add(key)
                item.label_center_y = lab.center_y
                item.label_offset_y = lab.center_y - item.current.bottom
                item.label_bottom_y = self.label_info_bottom(lab)
                found += 1
        if found:
            self.log(f"Alinhamento por rótulo PEÇA ativado: {found} vistas com referência de rótulo.")
        else:
            self.log("Alinhamento por rótulo PEÇA não encontrou referência; usando alinhamento pelo frame.")

    def attach_stack_content_padding(self, items):
        """Mede a folga vazia do frame das vistas que serão empilhadas.

        Só as peças com mais de uma vista precisam disso: no empilhamento, a
        "Dist. empilh." passa a ser medida entre o conteúdo visível das vistas,
        descontando o espaço vazio do frame da View.
        """
        counts = {}
        for item in items:
            key = str(item.item_text or "").strip()
            if key:
                counts[key] = counts.get(key, 0) + 1
        measured = 0
        for item in items:
            item.content_pad_top = 0.0
            item.content_pad_bottom = 0.0
            item.content_pad_left = 0.0
            item.content_pad_right = 0.0
            key = str(item.item_text or "").strip()
            if not key or counts.get(key, 0) < 2 or item.view is None:
                continue
            content = self.get_view_content_frame(item.view, item.current)
            if content is None:
                self.log(f"{item.name}: não consegui medir o conteúdo visível; usando o frame da View.")
                continue
            item.content_pad_top = max(0.0, item.current.top - content.top)
            item.content_pad_bottom = max(0.0, content.bottom - item.current.bottom)
            item.content_pad_left = max(0.0, content.left - item.current.left)
            item.content_pad_right = max(0.0, item.current.right - content.right)
            measured += 1
            self.log(
                f"{item.name}: frame {item.current.height:.1f} mm, conteúdo {content.height:.1f} mm "
                f"(folga sup. {item.content_pad_top:.1f}, inf. {item.content_pad_bottom:.1f})."
            )
        if measured:
            self.log(f"Conteúdo visível medido em {measured} vistas de peças com múltiplas vistas.")

    def plan_layout(self, items, sheet_w, sheet_h, cfg, log_result=True):
        """Distribui em colunas: cima -> baixo, depois esquerda -> direita.

        A versao 9 empacotava em linhas. Quando acabava a altura da folha, as
        proximas vistas recebiam "Sem espaco vertical", mesmo havendo area livre
        para iniciar uma nova coluna. A versao 2.1 troca essa regra: ao atingir
        o limite inferior ou a faixa reservada da legenda, abre uma nova coluna
        para a direita. A vista so e ignorada quando tambem nao existe largura
        util para essa nova coluna.
        """

        margin_left = float(cfg.margin_left)
        margin_top = float(cfg.margin_top)
        margin_right = float(cfg.margin_right)
        margin_bottom = float(cfg.margin_bottom)
        spacing_x = float(cfg.spacing_x)
        spacing_y = float(cfg.spacing_y)
        title_w = max(0.0, float(cfg.title_block_width))
        title_h = max(0.0, float(cfg.title_block_height))

        top_start = sheet_h - margin_top
        right_limit = sheet_w - margin_right
        title_left = sheet_w - margin_right - title_w
        title_right = sheet_w - margin_right

        def bottom_limit_for_item(column_left, item_width):
            normal_limit = margin_bottom
            if title_w <= 0.0 or title_h <= 0.0:
                return normal_limit
            column_right = column_left + item_width
            overlaps_title_x = column_left < title_right and column_right > title_left
            if overlaps_title_x:
                return margin_bottom + title_h + spacing_y
            return normal_limit

        cursor_left = margin_left
        cursor_top = top_start
        column_width = 0.0
        columns_used = 1 if items else 0

        for item in items:
            w = max(1.0, item.current.width)
            h = max(1.0, item.current.height)

            if cursor_left + w > right_limit:
                item.status = "Sem espaco horizontal para nova coluna"
                item.has_target = False
                continue

            bottom_limit = bottom_limit_for_item(cursor_left, w)
            bottom = cursor_top - h

            if bottom < bottom_limit and column_width > 0.01:
                cursor_left += column_width + spacing_x
                cursor_top = top_start
                column_width = 0.0
                columns_used += 1

                if cursor_left + w > right_limit:
                    item.status = "Sem espaco horizontal para nova coluna"
                    item.has_target = False
                    continue

                bottom_limit = bottom_limit_for_item(cursor_left, w)
                bottom = cursor_top - h

            if bottom < bottom_limit:
                item.status = "Sem espaco vertical na coluna"
                item.has_target = False
                continue

            item.target_left = cursor_left
            item.target_bottom = bottom
            item.target_right = cursor_left + w
            item.target_top = bottom + h
            item.has_target = True
            item.status = "Planejada - coluna"

            cursor_top = bottom - spacing_y
            column_width = max(column_width, w)

        if items and log_result:
            planned = len([item for item in items if item.has_target])
            self.log(f"Layout 2.1 por colunas: {planned}/{len(items)} vistas planejadas em {columns_used} coluna(s).")
        return

        def right_limit_for_bottom(row_bottom):
            normal = sheet_w - float(cfg.margin_right)
            title_top = float(cfg.margin_bottom) + float(cfg.title_block_height)
            if row_bottom < title_top:
                return min(normal, sheet_w - float(cfg.margin_right) - float(cfg.title_block_width) - float(cfg.spacing_x))
            return normal

        # 1) Monta as linhas pela largura disponível, mantendo a ordem PEÇA.
        rows = []
        current_row = []
        x = float(cfg.margin_left)
        provisional_top = sheet_h - float(cfg.margin_top)
        row_h = 0.0

        for item in items:
            w = max(1.0, item.current.width)
            h = max(1.0, item.current.height)
            limit = right_limit_for_bottom(provisional_top - h)
            if current_row and x + w > limit:
                rows.append(current_row)
                provisional_top -= row_h + float(cfg.spacing_y)
                current_row = []
                x = float(cfg.margin_left)
                row_h = 0.0
                limit = right_limit_for_bottom(provisional_top - h)

            if x + w > limit:
                item.status = "Maior que a área da linha"
                item.has_target = False
                continue

            current_row.append(item)
            x += w + float(cfg.spacing_x)
            row_h = max(row_h, h)

        if current_row:
            rows.append(current_row)

        # 2) Distribui as linhas. Dentro de cada linha, alinha pela altura do rótulo PEÇA.
        top = sheet_h - float(cfg.margin_top)
        for row in rows:
            usable = [it for it in row if it.status != "Maior que a área da linha"]
            if not usable:
                continue

            # Offset do rótulo em relação ao bottom do frame. Quando não há rótulo lido,
            # usa uma estimativa baseada na mediana da própria linha.
            offsets_known = [float(it.label_offset_y) for it in usable if it.label_offset_y is not None]
            if offsets_known:
                offsets_sorted = sorted(offsets_known)
                default_offset = offsets_sorted[len(offsets_sorted) // 2]
            else:
                default_offset = 0.0

            def label_offset(it):
                return float(it.label_offset_y) if it.label_offset_y is not None else default_offset

            # Para manter todas as vistas abaixo do topo da linha, calcula a linha-base do rótulo.
            max_above_label = max(max(1.0, it.current.height) - label_offset(it) for it in usable)
            label_y = top - max_above_label

            row_bottoms = []
            x = float(cfg.margin_left)
            for item in usable:
                w = max(1.0, item.current.width)
                h = max(1.0, item.current.height)
                bottom = label_y - label_offset(item)
                row_bottoms.append(bottom)

                limit = right_limit_for_bottom(bottom)
                if x + w > limit:
                    item.status = "Sem espaço horizontal"
                    item.has_target = False
                    continue
                if bottom < float(cfg.margin_bottom):
                    item.status = "Sem espaço vertical"
                    item.has_target = False
                    continue

                item.target_left = x
                item.target_bottom = bottom
                item.target_right = x + w
                item.target_top = bottom + h
                item.has_target = True
                item.status = "Planejada - rótulo alinhado" if offsets_known else "Planejada"
                x += w + float(cfg.spacing_x)

            if row_bottoms:
                top = min(row_bottoms) - float(cfg.spacing_y)

    def plan_layout(self, items, sheet_w, sheet_h, cfg, log_result=True):
        """Layout da base enviada: esquerda -> direita, depois linha abaixo.

        Mantem o comportamento visual da versao 9. A diferenca agora e que o
        Organizar inteligente reduz escala antes de organizar quando o conjunto
        nao couber, em vez de trocar o padrao de distribuicao para colunas.
        """

        spacing_y = float(cfg.spacing_y)
        # Aceita valores negativos: o usuario pode aproximar ou sobrepor os
        # frames das vistas empilhadas se quiser.
        stack_spacing_y = min(200.0, max(-200.0, float(getattr(cfg, "stack_spacing_y", STACK_VIEW_SPACING_MM) or 0.0)))

        def right_limit_for_bottom(row_bottom):
            normal = sheet_w - float(cfg.margin_right)
            title_top = float(cfg.margin_bottom) + float(cfg.title_block_height)
            if row_bottom < title_top:
                return min(normal, sheet_w - float(cfg.margin_right) - float(cfg.title_block_width) - float(cfg.spacing_x))
            return normal

        def stack_member_offsets(members):
            """Posicao vertical de cada frame dentro da unidade empilhada."""
            offsets = []
            y = 0.0
            for member in members:
                h = max(1.0, member.current.height)
                offsets.append(y)
                y += h + stack_spacing_y
            return offsets

        def stack_member_x_offsets(members):
            """Deslocamento horizontal de cada frame para alinhar o CONTEUDO.

            Alinha as vistas empilhadas pelo centro do conteudo visivel (desenho +
            cotas), e nao pela borda do frame. Assim, quando um frame e mais largo
            que o outro por causa de folga/cotas em um dos lados, as pecas continuam
            alinhadas entre si. Se as folgas forem simetricas (ou nao tiverem sido
            medidas), o resultado e igual a centralizar os frames, preservando o
            comportamento anterior.
            """
            ccos = []
            for member in members:
                w = max(1.0, member.current.width)
                clo = max(0.0, float(getattr(member, "content_pad_left", 0.0) or 0.0))
                cro = max(0.0, float(getattr(member, "content_pad_right", 0.0) or 0.0))
                # Centro do conteudo medido a partir da borda esquerda do frame.
                ccos.append((w + clo - cro) / 2.0)
            max_cco = max(ccos) if ccos else 0.0
            return [max_cco - cco for cco in ccos]

        def build_layout_units(layout_items):
            grouped = {}
            for it in layout_items:
                key = str(it.item_text or "").strip()
                if key:
                    grouped.setdefault(key, []).append(it)
            stacked_keys = {key for key, group in grouped.items() if len(group) > 1}
            units = []
            used_groups = set()
            stacked_count = 0
            for it in layout_items:
                key = str(it.item_text or "").strip()
                key = key if key in stacked_keys else ""
                if not key:
                    units.append({
                        "items": [it],
                        "width": max(1.0, it.current.width),
                        "height": max(1.0, it.current.height),
                        "label_offset": it.label_offset_y,
                        "stacked": False,
                    })
                    continue
                if key in used_groups:
                    continue
                used_groups.add(key)
                members = sorted(
                    grouped[key],
                    key=lambda member: (member.current.bottom, -member.current.left, member.index),
                )
                offsets = stack_member_offsets(members)
                x_offsets = stack_member_x_offsets(members)
                width = max(
                    x_off + max(1.0, member.current.width)
                    for x_off, member in zip(x_offsets, members)
                )
                height = max(
                    off + max(1.0, member.current.height)
                    for off, member in zip(offsets, members)
                )
                units.append({
                    "items": members,
                    "offsets": offsets,
                    "x_offsets": x_offsets,
                    "width": width,
                    "height": height,
                    "label_offset": None,
                    "stacked": True,
                    "group_key": key,
                })
                stacked_count += 1
            return units, stacked_count

        def set_unit_unplanned(unit, status):
            for member in unit["items"]:
                member.status = status
                member.has_target = False

        def place_unit(unit, left, bottom, status):
            if unit["stacked"]:
                x_offsets = unit.get("x_offsets") or [0.0] * len(unit["items"])
                for member, offset, x_off in zip(unit["items"], unit["offsets"], x_offsets):
                    w = max(1.0, member.current.width)
                    h = max(1.0, member.current.height)
                    member.target_left = left + x_off
                    member.target_bottom = bottom + offset
                    member.target_right = member.target_left + w
                    member.target_top = member.target_bottom + h
                    member.has_target = True
                    member.status = "Planejada - grupo empilhado"
                return
            item = unit["items"][0]
            w = max(1.0, item.current.width)
            h = max(1.0, item.current.height)
            item.target_left = left
            item.target_bottom = bottom
            item.target_right = left + w
            item.target_top = bottom + h
            item.has_target = True
            item.status = status

        for item in items:
            item.has_target = False
            item.target_left = 0.0
            item.target_bottom = 0.0
            item.target_right = 0.0
            item.target_top = 0.0
            if "Ignorada pelo limite" not in item.status:
                item.status = "Planejada"

        units, stacked_count = build_layout_units(items)
        rows = []
        current_row = []
        x = float(cfg.margin_left)
        provisional_top = sheet_h - float(cfg.margin_top)
        row_h = 0.0

        for unit in units:
            w = unit["width"]
            h = unit["height"]
            limit = right_limit_for_bottom(provisional_top - h)
            if current_row and x + w > limit:
                rows.append(current_row)
                provisional_top -= row_h + spacing_y
                current_row = []
                x = float(cfg.margin_left)
                row_h = 0.0
                limit = right_limit_for_bottom(provisional_top - h)

            if x + w > limit:
                set_unit_unplanned(unit, "Maior que a area da linha")
                continue

            current_row.append(unit)
            x += w + float(cfg.spacing_x)
            row_h = max(row_h, h)

        if current_row:
            rows.append(current_row)

        top = sheet_h - float(cfg.margin_top)
        for row_index, row in enumerate(rows):
            usable = [unit for unit in row if all(member.status != "Maior que a area da linha" for member in unit["items"])]
            if not usable:
                continue

            offsets_known = [float(unit["label_offset"]) for unit in usable if unit["label_offset"] is not None]
            if offsets_known:
                offsets_sorted = sorted(offsets_known)
                default_offset = offsets_sorted[len(offsets_sorted) // 2]
            else:
                default_offset = 0.0

            def label_offset(unit):
                return float(unit["label_offset"]) if unit["label_offset"] is not None else default_offset

            max_above_label = max(unit["height"] - label_offset(unit) for unit in usable)
            label_y = top - max_above_label

            row_bottoms = [label_y - label_offset(unit) for unit in usable]

            # Justificacao horizontal: as linhas cheias (todas menos a ultima) sao
            # espalhadas para preencher toda a largura util ate a borda da legenda,
            # em vez de deixarem uma sobra na direita. A ultima linha fica alinhada
            # a esquerda para nao esticar poucas pecas de forma estranha.
            is_last_row = (row_index == len(rows) - 1)
            gap = float(cfg.spacing_x)
            if not is_last_row and len(usable) >= 2 and row_bottoms:
                row_right_limit = min(right_limit_for_bottom(b) for b in row_bottoms)
                total_w = sum(unit["width"] for unit in usable)
                extra = (row_right_limit - float(cfg.margin_left)) - total_w - float(cfg.spacing_x) * (len(usable) - 1)
                if extra > 0.0:
                    gap = float(cfg.spacing_x) + extra / (len(usable) - 1)

            x = float(cfg.margin_left)
            for unit, bottom in zip(usable, row_bottoms):
                w = unit["width"]

                limit = right_limit_for_bottom(bottom)
                if x + w > limit + 0.5:
                    set_unit_unplanned(unit, "Sem espaco horizontal")
                    continue
                if bottom < float(cfg.margin_bottom):
                    set_unit_unplanned(unit, "Sem espaco vertical")
                    continue

                place_unit(unit, x, bottom, "Planejada - rotulo alinhado" if offsets_known else "Planejada")
                x += w + gap

            if row_bottoms:
                top = min(row_bottoms) - spacing_y

        # Segunda passagem: aproveita espaços livres deixados por vistas de tamanhos
        # diferentes. A primeira passagem mantém a organização principal em linhas;
        # esta etapa só tenta encaixar as peças ainda sem alvo em lacunas realmente
        # disponíveis, sem sobrepor vistas nem a legenda.
        occupied = []
        for unit in units:
            placed = [member for member in unit["items"] if member.has_target]
            if not placed:
                continue
            occupied.append((
                min(member.target_left for member in placed),
                min(member.target_bottom for member in placed),
                max(member.target_right for member in placed),
                max(member.target_top for member in placed),
            ))

        area_left = float(cfg.margin_left)
        area_right = sheet_w - float(cfg.margin_right)
        area_bottom = float(cfg.margin_bottom)
        area_top = sheet_h - float(cfg.margin_top)
        title_w = max(0.0, float(cfg.title_block_width))
        title_h = max(0.0, float(cfg.title_block_height))
        title_rect = None
        if title_w > 0.0 and title_h > 0.0:
            title_rect = (
                area_right - title_w - float(cfg.spacing_x),
                area_bottom,
                area_right,
                area_bottom + title_h + spacing_y,
            )

        def rects_conflict(a, b):
            return not (
                a[2] + float(cfg.spacing_x) <= b[0] + 0.01 or
                a[0] >= b[2] + float(cfg.spacing_x) - 0.01 or
                a[3] + spacing_y <= b[1] + 0.01 or
                a[1] >= b[3] + spacing_y - 0.01
            )

        def overlaps_plain(a, b):
            return not (
                a[2] <= b[0] + 0.01 or a[0] >= b[2] - 0.01 or
                a[3] <= b[1] + 0.01 or a[1] >= b[3] - 0.01
            )

        recovered = 0
        for unit in units:
            if any(member.has_target for member in unit["items"]):
                continue
            if any("Ignorada pelo limite" in str(member.status) for member in unit["items"]):
                continue

            w = max(1.0, float(unit["width"]))
            h = max(1.0, float(unit["height"]))

            candidate_x = {area_left}
            candidate_top = {area_top}
            for left, bottom, right, top_edge in occupied:
                candidate_x.add(left)
                candidate_x.add(right + float(cfg.spacing_x))
                candidate_top.add(top_edge)
                candidate_top.add(bottom - spacing_y)
            if title_rect is not None:
                candidate_x.add(max(area_left, title_rect[0] - w - float(cfg.spacing_x)))
                candidate_top.add(min(area_top, title_rect[3] + h))

            placed_here = False
            for candidate_top_value in sorted(candidate_top, reverse=True):
                bottom = candidate_top_value - h
                if bottom < area_bottom - 0.01 or candidate_top_value > area_top + 0.01:
                    continue
                for left in sorted(candidate_x):
                    rect = (left, bottom, left + w, candidate_top_value)
                    if rect[0] < area_left - 0.01 or rect[2] > area_right + 0.01:
                        continue
                    if title_rect is not None and overlaps_plain(rect, title_rect):
                        continue
                    if any(rects_conflict(rect, other) for other in occupied):
                        continue
                    place_unit(unit, left, bottom, "Planejada - aproveitamento de espaço")
                    occupied.append(rect)
                    recovered += len(unit["items"])
                    placed_here = True
                    break
                if placed_here:
                    break

        if recovered and log_result:
            self.log(f"Aproveitamento de espaços livres: {recovered} vista(s) adicional(is) planejada(s).")

        if items and log_result:
            planned = len([item for item in items if item.has_target])
            self.log(f"Layout base por linhas: {planned}/{len(items)} vistas planejadas.")
            if stacked_count:
                self.log(
                    f"Grupos com múltiplas vistas empilhados: {stacked_count} "
                    f"(Dist. empilh. = {stack_spacing_y:.0f} mm entre os desenhos)."
                )

    def assign_labels_to_views(self, raw_views, labels):
        if not raw_views or not labels:
            return []
        pairs = []
        used_labels = set()
        # Primeiro tenta rótulos que ficam dentro ou logo abaixo da View.
        for view, frame in raw_views:
            cx = (frame.left + frame.right) / 2.0
            local = []
            for li, lab in enumerate(labels):
                if li in used_labels:
                    continue
                # Título de vista PEÇA costuma ficar abaixo ou dentro do frame.
                inside_x = frame.left - 60 <= lab.center_x <= frame.right + 60
                near_y = frame.bottom - max(90.0, frame.height * 0.80) <= lab.center_y <= frame.top + 40
                if inside_x and near_y:
                    dist = abs(lab.center_x - cx) + abs(lab.center_y - frame.bottom) * 0.45
                    local.append((dist, li, lab))
            if local:
                _, li, lab = sorted(local, key=lambda x: x[0])[0]
                used_labels.add(li)
                pairs.append((view, frame, lab))
        # Se não conseguiu associar quase nada, não força uma associação ruim.
        if len(pairs) < max(1, min(len(raw_views), len(labels)) * 0.50):
            return []
        return pairs

    def sort_raw_views_visual(self, raw_views):
        # Ordena por posição na folha: de cima para baixo e da esquerda para direita.
        # Tekla usa coordenada Y crescente para cima, então top maior vem primeiro.
        if not raw_views:
            return []
        avg_h = max(1.0, sum(frame.height for _, frame in raw_views) / len(raw_views))
        row_tol = max(10.0, avg_h * 0.55)
        ordered = sorted(raw_views, key=lambda vf: (-vf[1].top, vf[1].left))
        rows = []
        for view, frame in ordered:
            placed = False
            cy = (frame.bottom + frame.top) / 2.0
            for row in rows:
                row_cy = sum((f.bottom + f.top) / 2.0 for _, f in row) / len(row)
                if abs(cy - row_cy) <= row_tol:
                    row.append((view, frame))
                    placed = True
                    break
            if not placed:
                rows.append([(view, frame)])
        rows.sort(key=lambda row: -max(f.top for _, f in row))
        final = []
        for row in rows:
            final.extend(sorted(row, key=lambda vf: vf[1].left))
        return final

    def ensure_view_free_for_moving(self, view):
        try:
            attr = view.Attributes
            if attr is not None and bool(attr.FixedViewPlacing):
                attr.FixedViewPlacing = False
                view.Attributes = attr
        except Exception:
            pass

    def save_backup(self, items):
        path = REPORT_DIR / ("backup_posicoes_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".csv")
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["Sequencia", "Vista", "Ordem", "Escala", "AtualX", "AtualY", "DestinoX", "DestinoY", "Largura", "Altura", "Status"])
            for it in items:
                w.writerow([
                    it.index, it.name, it.item_text, f"1:{it.current_scale:g}",
                    f"{it.current.left:.3f}", f"{it.current.bottom:.3f}",
                    f"{it.target_left:.3f}" if it.has_target else "",
                    f"{it.target_bottom:.3f}" if it.has_target else "",
                    f"{it.current.width:.3f}", f"{it.current.height:.3f}", it.status
                ])
        return path

    def capture_undo_state(self, items):
        states = []
        for item in items:
            if item.view is None:
                continue
            try:
                origin = item.view.Origin
                frame = self.get_view_frame(item.view)
                states.append(ViewUndoState(
                    index=int(item.index),
                    name=str(item.name or ""),
                    item_text=str(item.item_text or ""),
                    view=item.view,
                    origin_x=float(origin.X),
                    origin_y=float(origin.Y),
                    origin_z=float(origin.Z),
                    scale=float(self.get_view_scale(item.view)),
                    frame_left=float(frame.left),
                    frame_bottom=float(frame.bottom),
                ))
            except Exception as ex:
                self.log(f"Nao foi possivel preparar desfazer para {item.name}: {ex}")
        return states

    def resolve_undo_view(self, state, current_items):
        try:
            _ = state.view.Origin
            return state.view
        except Exception:
            pass
        for item in current_items or []:
            if state.item_text and str(item.item_text) == str(state.item_text):
                return item.view
            if state.name and str(item.name) == str(state.name):
                return item.view
        return None

    def restore_undo_state(self, states, current_items, handler, drawing):
        restored = 0
        errors = []
        for state in states:
            view = self.resolve_undo_view(state, current_items)
            if view is None:
                errors.append(f"{state.name}: vista nao localizada")
                continue
            try:
                self.ensure_view_free_for_moving(view)
                current_scale = self.get_view_scale(view)
                if abs(current_scale - float(state.scale)) > 0.001:
                    self.set_view_scale(view, state.scale)
                view.Origin = self.TeklaPoint(state.origin_x, state.origin_y, state.origin_z)
                if not view.Modify():
                    errors.append(f"{state.name}: Tekla recusou restaurar posicao")
                    continue
                restored += 1
                self.log(
                    f"Desfeito: {state.name} -> X={state.frame_left:.1f}, "
                    f"Y={state.frame_bottom:.1f}, escala 1:{state.scale:g}"
                )
            except Exception as ex:
                errors.append(f"{state.name}: {ex}")
                self.log("Erro ao desfazer " + str(state.name) + ": " + str(ex))
        try:
            drawing.CommitChanges("Desfazer organizacao de vistas")
        except Exception:
            pass
        try:
            handler.SaveActiveDrawing()
        except Exception:
            pass
        if errors:
            raise RuntimeError("Algumas vistas nao foram restauradas:\n" + "\n".join(errors[:8]))
        return restored

    def collect_selected_single_part_drawings(self, handler):
        drawings = []
        try:
            selector = handler.GetDrawingSelector()
            enum = selector.GetSelected()
            while enum.MoveNext():
                drawing = enum.Current
                if drawing is not None and drawing.GetType().Name == "SinglePartDrawing":
                    drawings.append(drawing)
        except Exception as ex:
            self.log("Nao consegui ler a selecao do Document Manager: " + str(ex))
        return drawings

    def collect_all_single_part_drawings(self, handler):
        drawings = []
        enum = handler.GetDrawings()
        while enum.MoveNext():
            drawing = enum.Current
            if drawing is not None and drawing.GetType().Name == "SinglePartDrawing":
                drawings.append(drawing)
        return drawings

    def drawing_piece_info(self, drawing):
        if drawing is None:
            return None
        try:
            if drawing.GetType().Name != "SinglePartDrawing":
                return None
        except Exception:
            return None
        mark = ""
        name = ""
        try:
            mark = str(drawing.Mark or "")
        except Exception:
            pass
        try:
            name = str(drawing.Name or "")
        except Exception:
            pass
        parsed = try_parse_drawing_mark(mark) or try_parse_numeric_item(mark)
        if not parsed:
            return None
        item_text, sort_key = parsed
        return item_text, sort_key, drawing, mark, name

    def sheet_mark_key(self, value):
        parsed = try_parse_drawing_mark(value) or try_parse_numeric_item(value)
        if not parsed:
            return None
        return numeric_key(parsed[1])

    def normalized_sheet_mark(self, value):
        text = str(value or "").strip()
        key = self.sheet_mark_key(text)
        if key and len(key) == 1:
            try:
                return f"[{int(key[0])}]"
            except Exception:
                pass
        return text

    def capacity_project_key(self, drawing):
        """Identifica o projeto sem depender do numero da folha."""
        try:
            if drawing is None or drawing.GetType().Name != "MultiDrawing":
                return ""
        except Exception:
            return ""
        try:
            name = norm_spaces(str(drawing.Name or "")).casefold()
        except Exception:
            name = ""
        return f"multidrawing:{name}" if name else ""

    def ensure_capacity_project_context(self, cfg, drawing):
        """Impede que a fila de um projeto seja reutilizada em outro."""
        project_key = self.capacity_project_key(drawing)
        if not project_key:
            return False
        saved_key = str(getattr(cfg, "capacity_project_key", "") or "").strip()
        if saved_key == project_key:
            return False

        cfg.capacity_project_key = project_key
        cfg.capacity_cursor_item = ""
        cfg.capacity_cursor_sheet = ""
        cfg.capacity_cursor_count = 0
        cfg.capacity_sheet_cursors = {}
        cfg.capacity_factor_x = 1.0
        cfg.capacity_factor_y = 1.0
        cfg.capacity_factor_matches = 0
        cfg.save()
        if saved_key:
            self.log("Projeto alterado: fila e calibracao de capacidade reiniciadas.")
        else:
            self.log("Fila de capacidade inicializada para o projeto atual.")
        return True

    def previous_sheet_mark(self, value):
        key = self.sheet_mark_key(value)
        if not key or len(key) != 1:
            return ""
        try:
            number = int(key[0])
        except Exception:
            return ""
        if number <= 1:
            return ""
        return f"[{number - 1}]"

    def capacity_cursor_for_multidrawing(self, cfg, drawing):
        self.ensure_capacity_project_context(cfg, drawing)
        current_mark = ""
        try:
            current_mark = str(drawing.Mark or "").strip()
        except Exception:
            pass

        history = getattr(cfg, "capacity_sheet_cursors", {}) or {}
        if not isinstance(history, dict):
            history = {}

        previous_mark = self.previous_sheet_mark(current_mark)
        if previous_mark:
            entry = history.get(previous_mark) or history.get(previous_mark.strip("[]"))
            if isinstance(entry, dict):
                if str(entry.get("source") or "").strip().lower() != "actual":
                    entry = None
            if isinstance(entry, dict):
                item = str(entry.get("item") or "").strip()
                if item:
                    return item, previous_mark
            elif entry:
                return str(entry).strip(), previous_mark

        cursor_item = str(getattr(cfg, "capacity_cursor_item", "") or "").strip()
        cursor_sheet = str(getattr(cfg, "capacity_cursor_sheet", "") or "").strip()
        if not cursor_item:
            return "", ""

        current_key = self.sheet_mark_key(current_mark)
        cursor_key = self.sheet_mark_key(cursor_sheet)
        if current_key and cursor_key and cursor_key >= current_key:
            return "", ""
        return cursor_item, cursor_sheet

    def sorted_single_part_drawings(self, drawings):
        infos = []
        for drawing in drawings or []:
            info = self.drawing_piece_info(drawing)
            if info:
                infos.append(info)
        infos.sort(key=lambda info: (numeric_key(info[1]), str(info[4] or "").lower()))
        return [info[2] for info in infos]

    def single_part_drawings_after(self, drawings, cursor_item):
        parsed = try_parse_numeric_item(str(cursor_item or ""))
        if not parsed:
            return []
        _, cursor_key = parsed
        cursor_tuple = numeric_key(cursor_key)
        out = []
        for info in self.sorted_single_part_drawings(drawings):
            piece_info = self.drawing_piece_info(info)
            if piece_info and numeric_key(piece_info[1]) > cursor_tuple:
                out.append(info)
        return out

    def single_part_drawings_from(self, drawings, start_item):
        parsed = try_parse_numeric_item(str(start_item or ""))
        if not parsed:
            return []
        _, start_key = parsed
        start_tuple = numeric_key(start_key)
        out = []
        for drawing in self.sorted_single_part_drawings(drawings):
            piece_info = self.drawing_piece_info(drawing)
            if piece_info and numeric_key(piece_info[1]) >= start_tuple:
                out.append(drawing)
        return out

    def measure_drawing_main_views(self, drawing):
        frames = []
        scales = []
        try:
            sheet = drawing.GetSheet()
            enum = sheet.GetAllViews()
            while enum.MoveNext():
                view = enum.Current
                try:
                    if bool(view.IsSheet):
                        continue
                except Exception:
                    pass
                frame = self.get_view_frame(view)
                if frame.width > 0.1 and frame.height > 0.1:
                    frames.append(frame)
                    scales.append(self.get_view_scale(view))
        except Exception:
            return None, 1.0
        if not frames:
            return None, 1.0
        left = min(frame.left for frame in frames)
        bottom = min(frame.bottom for frame in frames)
        right = max(frame.right for frame in frames)
        top = max(frame.top for frame in frames)
        scale = sorted(scales)[len(scales) // 2] if scales else 1.0
        return TeklaFrame(0.0, 0.0, max(1.0, right - left), max(1.0, top - bottom)), scale

    def build_capacity_items_from_drawings(self, drawings, handler=None, original_drawing=None):
        items = []
        drawings = list(drawings or [])
        restore_needed = False
        total = max(1, len(drawings))
        try:
            for pos, drawing in enumerate(drawings, start=1):
                self.report_progress(100.0 * pos / total, f"Medindo croquis {pos}/{len(drawings)}...")
                if handler is not None:
                    # O proxy do Document Manager pode devolver frames incompletos.
                    # A medição do desenho realmente aberto é a referência confiável.
                    restore_needed = True
                    item = self.capacity_item_from_opened_drawing(handler, drawing, len(items) + 1)
                    if item is None:
                        item = self.capacity_item_from_drawing(drawing, len(items) + 1)
                else:
                    item = self.capacity_item_from_drawing(drawing, len(items) + 1)
                if item is None:
                    continue
                items.append(item)
                if len(drawings) >= 20 and (pos % 20 == 0 or pos == len(drawings)):
                    self.log(f"Croquis medidos para capacidade: {pos}/{len(drawings)}.")
        finally:
            if restore_needed and handler is not None and original_drawing is not None:
                self.restore_original_drawing(handler, original_drawing)
        items.sort(key=lambda item: (numeric_key(item.sort_key), item.name.lower()))
        for idx, item in enumerate(items, start=1):
            item.index = idx
        return items

    def build_capacity_until_overflow(self, drawings, cfg, sheet_w, sheet_h, factor_x=1.0, factor_y=1.0, handler=None, original_drawing=None):
        items = []
        drawings = list(drawings or [])
        first_failed = None
        fit_count = 0
        restore_needed = False
        total = max(1, len(drawings))
        try:
            for pos, drawing in enumerate(drawings, start=1):
                self.report_progress(100.0 * pos / total, f"Medindo croquis {pos}/{len(drawings)}...")
                if handler is not None:
                    # O proxy do Document Manager pode devolver frames incompletos.
                    # A medição do desenho realmente aberto é a referência confiável.
                    restore_needed = True
                    item = self.capacity_item_from_opened_drawing(handler, drawing, len(items) + 1)
                    if item is None:
                        item = self.capacity_item_from_drawing(drawing, len(items) + 1)
                else:
                    item = self.capacity_item_from_drawing(drawing, len(items) + 1)
                if item is None:
                    continue
                items.append(item)
                fit_count, first_failed, _ = self.calculate_fit_count(
                    items,
                    sheet_w,
                    sheet_h,
                    cfg,
                    factor_x=factor_x,
                    factor_y=factor_y,
                )
                if len(drawings) >= 20 and (pos % 20 == 0):
                    self.log(f"Croquis medidos na fila automatica: {pos}/{len(drawings)}.")
                if first_failed is not None:
                    break
        finally:
            if restore_needed and handler is not None and original_drawing is not None:
                self.restore_original_drawing(handler, original_drawing)
        for idx, item in enumerate(items, start=1):
            item.index = idx
        if first_failed is None:
            fit_count = len(items)
        return items, fit_count, first_failed

    def capacity_item_from_drawing(self, drawing, index):
        if drawing is None:
            return None
        try:
            if drawing.GetType().Name != "SinglePartDrawing":
                return None
        except Exception:
            return None
        mark = ""
        name = ""
        try:
            mark = str(drawing.Mark or "")
        except Exception:
            pass
        try:
            name = str(drawing.Name or "")
        except Exception:
            pass
        parsed = try_parse_drawing_mark(mark) or try_parse_numeric_item(mark)
        if not parsed:
            return None
        item_text, sort_key = parsed
        frame, scale = self.measure_drawing_main_views(drawing)
        if frame is None:
            return None
        item = ApiViewItem(
            index,
            f"PEÇA {item_text}",
            item_text,
            sort_key,
            None,
            frame,
            status="Croqui de fabricação",
        )
        item.current_scale = scale
        item.source_drawing = drawing
        item.source_mark = mark
        item.source_name = name
        self.apply_stack_group(item, [mark, name])
        return item

    def item_from_drawing_link(self, link, index):
        if link is None:
            return None
        try:
            target = link.Target
        except Exception:
            target = None
        mark = ""
        name = ""
        if target is not None:
            try:
                mark = str(target.Mark or "")
            except Exception:
                pass
            try:
                name = str(target.Name or "")
            except Exception:
                pass
        parsed = try_parse_drawing_mark(mark) or try_parse_numeric_item(mark)
        if not parsed:
            try:
                parsed = try_parse_drawing_mark(str(link.Text or ""))
            except Exception:
                parsed = None
        if not parsed:
            return None
        frame = self.get_object_bbox(link)
        if frame is None or frame.width <= 0.1 or frame.height <= 0.1:
            try:
                pt = link.InsertionPoint
                frame = TeklaFrame(float(pt.X), float(pt.Y), float(pt.X) + 1.0, float(pt.Y) + 1.0)
            except Exception:
                return None
        item_text, sort_key = parsed
        item = ApiViewItem(
            index,
            f"PEÇA {item_text}",
            item_text,
            sort_key,
            link,
            frame,
            status="Link de desenho",
        )
        item.source_drawing = target
        item.source_mark = mark
        item.source_name = name
        self.apply_stack_group(item, [mark, name, getattr(link, "Text", "")])
        return item

    def restore_original_drawing(self, handler, original_drawing):
        """Restaura o desenho que estava aberto antes da medicao de capacidade."""
        try:
            restored = bool(handler.SetActiveDrawing(original_drawing, True, False))
            if restored:
                time.sleep(0.08)
                return True
        except Exception:
            pass
        try:
            # Segunda tentativa em segundo plano e com forceOpen para desenhos
            # marcados como nao atualizados pelo Tekla.
            restored = bool(handler.SetActiveDrawing(original_drawing, False, True))
            if restored:
                time.sleep(0.08)
                try:
                    handler.SetActiveDrawing(original_drawing, True, False)
                except Exception:
                    pass
                return True
        except Exception as ex:
            self.log("Nao consegui restaurar o multidrawing ativo: " + str(ex))
        return False

    def capacity_item_from_opened_drawing(self, handler, drawing, index):
        try:
            # Abre em segundo plano. forceOpen=True permite medir tambem croquis
            # que o Tekla considera desatualizados, sem exibir cada desenho.
            opened = bool(handler.SetActiveDrawing(drawing, False, True))
        except Exception as ex:
            self.log("Nao consegui abrir croqui para medir: " + str(ex))
            return None
        if not opened:
            try:
                mark = str(drawing.Mark or "")
            except Exception:
                mark = ""
            self.log("O Tekla recusou a abertura do croqui" + (f" {mark}" if mark else "") + ".")
            return None
        time.sleep(0.10)
        try:
            active = handler.GetActiveDrawing()
        except Exception:
            active = None
        item = self.capacity_item_from_drawing(active, index) if active is not None else None
        if item is not None:
            return item
        return self.capacity_item_from_drawing(drawing, index)

    def estimate_capacity_factors(self, active_items, source_items):
        if not active_items or not source_items:
            return 1.0, 1.0, 0
        by_text = {}
        for item in source_items:
            by_text.setdefault(str(item.item_text), item)
        ratio_x = []
        ratio_y = []
        matched = 0
        for active in active_items:
            source = by_text.get(str(active.item_text))
            if source is None:
                continue
            active_scale = max(1.0, float(active.current_scale or 1.0))
            source_scale = max(1.0, float(source.current_scale or 1.0))
            scale_factor = source_scale / max(1.0, active_scale)
            source_w = max(1.0, source.current.width * scale_factor)
            source_h = max(1.0, source.current.height * scale_factor)
            if source_w > 0.1 and active.current.width > 0.1:
                ratio_x.append(active.current.width / source_w)
            if source_h > 0.1 and active.current.height > 0.1:
                ratio_y.append(active.current.height / source_h)
            matched += 1
        if not ratio_x or not ratio_y:
            return 1.0, 1.0, 0

        def pct(values, fraction):
            values = sorted(max(1.0, float(value)) for value in values)
            index = int(round((len(values) - 1) * fraction))
            return values[max(0, min(len(values) - 1, index))]

        return pct(ratio_x, 0.85), pct(ratio_y, 0.85), matched

    def copy_layout_metadata(self, clone, item, size_factor=1.0):
        clone.label_center_y = item.label_center_y
        clone.label_offset_y = item.label_offset_y
        clone.stack_group_key = item.stack_group_key
        clone.stack_group_sort_key = list(item.stack_group_sort_key or [])
        clone.stack_order_key = list(item.stack_order_key or [])
        # As folgas do conteúdo escalam junto com o frame do clone.
        size_factor = max(0.0, float(size_factor or 0.0))
        clone.content_pad_top = max(0.0, float(getattr(item, "content_pad_top", 0.0) or 0.0)) * size_factor
        clone.content_pad_bottom = max(0.0, float(getattr(item, "content_pad_bottom", 0.0) or 0.0)) * size_factor
        clone.content_pad_left = max(0.0, float(getattr(item, "content_pad_left", 0.0) or 0.0)) * size_factor
        clone.content_pad_right = max(0.0, float(getattr(item, "content_pad_right", 0.0) or 0.0)) * size_factor
        return clone

    def clone_items_for_capacity(self, items, cfg, factor_x=1.0, factor_y=1.0):
        clones = []
        for item in items:
            old_scale = max(1.0, float(item.current_scale or 1.0))
            width = max(1.0, item.current.width * max(1.0, factor_x))
            height = max(1.0, item.current.height * max(1.0, factor_y))
            frame = TeklaFrame(0.0, 0.0, width, height)
            clone = ApiViewItem(item.index, item.name, item.item_text, item.sort_key, None, frame, status=item.status)
            clone.current_scale = old_scale
            self.copy_layout_metadata(clone, item, size_factor=max(1.0, factor_y))
            for attr in ("source_drawing", "source_mark", "source_name"):
                if hasattr(item, attr):
                    setattr(clone, attr, getattr(item, attr))
            clones.append(clone)
        return clones

    def calculate_fit_count(self, items, sheet_w, sheet_h, cfg, factor_x=1.0, factor_y=1.0):
        clones = self.clone_items_for_capacity(items, cfg, factor_x=factor_x, factor_y=factor_y)
        self.plan_layout(clones, sheet_w, sheet_h, cfg, log_result=False)
        # A capacidade por folha precisa representar uma sequência contínua.
        # Não conta peças menores posteriores quando uma peça anterior já não coube.
        fit_count = 0
        first_failed = None
        for item in clones:
            if not item.has_target:
                first_failed = item
                break
            fit_count += 1
        return fit_count, first_failed, clones

    def max_fitting_prefix(self, items, sheet_w, sheet_h, cfg, factor_x=1.0, factor_y=1.0):
        """Maior quantidade inicial que cabe usando o mesmo layout final."""
        items = list(items or [])
        if not items:
            return 0

        def prefix_fits(count):
            if count <= 0:
                return True
            clones = self.clone_items_for_capacity(items[:count], cfg, factor_x=factor_x, factor_y=factor_y)
            self.plan_layout(clones, sheet_w, sheet_h, cfg, log_result=False)
            return all(item.has_target for item in clones)

        lo, hi = 0, len(items)
        while lo < hi:
            mid = (lo + hi + 1) // 2
            if prefix_fits(mid):
                lo = mid
            else:
                hi = mid - 1
        return lo

    def sheet_mark_with_offset(self, start_mark, offset):
        key = self.sheet_mark_key(start_mark)
        if key and len(key) == 1:
            try:
                return f"[{int(key[0]) + int(offset)}]"
            except Exception:
                pass
        base = str(start_mark or "").strip()
        if offset == 0 and base:
            return base
        return f"Folha {int(offset) + 1}"

    def sheet_label_text(self, sheet_mark):
        key = self.sheet_mark_key(sheet_mark)
        if key and len(key) == 1:
            try:
                return f"FL {int(key[0])}"
            except Exception:
                pass
        return str(sheet_mark or "Folha").strip() or "Folha"

    def build_capacity_sheet_plan(self, items, cfg, sheet_w, sheet_h, start_sheet_mark, factor_x=1.0, factor_y=1.0):
        """Divide a fila completa de croquis em folhas consecutivas."""
        remaining = list(items or [])
        plan = []
        sheet_offset = 0
        while remaining and sheet_offset < 500:
            fit_count = self.max_fitting_prefix(
                remaining,
                sheet_w,
                sheet_h,
                cfg,
                factor_x=factor_x,
                factor_y=factor_y,
            )
            forced = False
            if fit_count <= 0:
                fit_count = 1
                forced = True
            chunk = remaining[:fit_count]
            if not chunk:
                break
            sheet_mark = self.sheet_mark_with_offset(start_sheet_mark, sheet_offset)
            plan.append({
                "sheet": sheet_mark,
                "label": self.sheet_label_text(sheet_mark),
                "start": chunk[0].item_text,
                "end": chunk[-1].item_text,
                "count": len(chunk),
                "overflow": forced,
            })
            remaining = remaining[fit_count:]
            sheet_offset += 1
        return plan

    def sheet_plan_total(self, rows):
        total = 0
        for row in rows or []:
            try:
                total += int(row.get("count", 0) or 0)
            except Exception:
                pass
        return total

    def useful_area(self, sheet_w, sheet_h, cfg):
        width = max(1.0, sheet_w - float(cfg.margin_left) - float(cfg.margin_right))
        height = max(1.0, sheet_h - float(cfg.margin_top) - float(cfg.margin_bottom))
        title_area = max(0.0, float(cfg.title_block_width)) * max(0.0, float(cfg.title_block_height))
        return max(1.0, width * height - title_area)

    def layout_density(self, items, sheet_w, sheet_h, cfg):
        area = sum(max(1.0, it.current.width) * max(1.0, it.current.height) for it in items)
        return area / self.useful_area(sheet_w, sheet_h, cfg)

    def layout_warning(self, items, cfg):
        if not items:
            return ""
        unplanned = [it for it in items if not it.has_target and "Ignorada pelo limite" not in it.status]
        if unplanned:
            return (
                f"{len(unplanned)} vista(s) nao couberam no layout. "
                "Recomendo reduzir escala das maiores vistas e analisar novamente."
            )
        return ""

    def round_scale_up(self, value, limit):
        standards = [1.0, 1.25, 1.5, 2.0, 2.5, 3.0, 4.0, 5.0, 5.5, 6.0, 7.5, 10.0, 12.5, 15.0, 20.0]
        value = max(1.0, float(value))
        limit = max(1.0, float(limit))
        for scale in standards:
            if scale >= value and scale <= limit:
                return scale
        if value <= limit:
            return min(limit, round(value * 2.0 + 0.9999) / 2.0)
        return limit

    def suggest_scale(self, items, cfg):
        if not items:
            return float(cfg.scale_target), []
        current_scales = [max(1.0, float(it.current_scale or 1.0)) for it in items]
        base_scale = sorted(current_scales)[len(current_scales) // 2]
        density = self.layout_density(items, self.last_sheet_w or 841.0, self.last_sheet_h or 594.0, cfg)
        if density > 0.01:
            desired_factor = (max(0.25, float(cfg.density_alert) * 0.85) / density) ** 0.5
            suggested = base_scale / min(1.0, desired_factor)
        else:
            suggested = base_scale
        suggested = max(float(cfg.scale_target), suggested)
        suggested = self.round_scale_up(suggested, float(cfg.auto_scale_limit))

        unplanned = [it for it in items if not it.has_target and "Ignorada pelo limite" not in it.status]
        if unplanned:
            ordered = sorted(items, key=lambda it: it.current.width * it.current.height, reverse=True)
            desired_count = max(len(unplanned), int(round(len(ordered) * 0.35)))
            candidates = list(unplanned)
            for item in ordered:
                if item not in candidates:
                    candidates.append(item)
                if len(candidates) >= desired_count:
                    break
        else:
            ordered = sorted(items, key=lambda it: it.current.width * it.current.height, reverse=True)
            count = max(1, int(round(len(ordered) * 0.35)))
            candidates = ordered[:count]
        return suggested, candidates

    def clone_items_for_scale(self, items, new_scale):
        clones = []
        new_scale = float(new_scale)
        for item in items:
            old_scale = max(1.0, float(item.current_scale or 1.0))
            factor = old_scale / max(1.0, new_scale)
            w = max(1.0, item.current.width * factor)
            h = max(1.0, item.current.height * factor)
            frame = TeklaFrame(item.current.left, item.current.bottom, item.current.left + w, item.current.bottom + h)
            clone = ApiViewItem(item.index, item.name, item.item_text, item.sort_key, item.view, frame)
            clone.current_scale = new_scale
            self.copy_layout_metadata(clone, item, size_factor=factor)
            clones.append(clone)
        return clones

    def standard_scale_candidates(self, items, cfg):
        current_max = max([float(it.current_scale or 1.0) for it in items] + [1.0])
        # A escala preenchida na interface nao deve substituir a escala real das
        # vistas como referencia da simulacao. O Organizar inteligente sempre
        # parte da escala atual das pecas e testa candidatos padrao a partir dela.
        minimum = max(current_max, 2.5)
        standards = [2.5, 5.0, 7.5, 10.0, 12.5, 15.0, 17.5, 20.0, 22.5, 25.0, 27.5, 30.0]
        limit = 30.0
        candidates = [scale for scale in standards if minimum - 0.001 <= scale <= limit + 0.001]
        if not candidates:
            candidates = [minimum]
        return sorted(set(round(scale, 4) for scale in candidates))

    def find_scale_solution(self, items, cfg):
        sheet_w = self.last_sheet_w or 841.0
        sheet_h = self.last_sheet_h or 594.0
        best = None
        for scale in self.standard_scale_candidates(items, cfg):
            clones = self.clone_items_for_scale(items, scale)
            self.plan_layout(clones, sheet_w, sheet_h, cfg, log_result=False)
            unplanned = [it for it in clones if not it.has_target and "Ignorada pelo limite" not in it.status]
            density = self.layout_density(clones, sheet_w, sheet_h, cfg)
            result = {
                "scale": scale,
                "fits": not unplanned,
                "unplanned": len(unplanned),
                "density": density,
                "items": clones,
            }
            best = result
            if result["fits"]:
                return result
        return best or {"scale": float(cfg.scale_target), "fits": False, "unplanned": len(items), "density": 1.0, "items": []}

    def apply_scale_to_items(self, items, new_scale, handler, drawing):
        changed = 0
        errors = []
        for item in items:
            if item.view is None:
                continue
            try:
                old_scale = self.get_view_scale(item.view)
                if abs(float(new_scale) - old_scale) < 0.001:
                    item.status = f"Escala ja era 1:{old_scale:g}"
                    continue
                self.set_view_scale(item.view, new_scale)
                item.current_scale = float(new_scale)
                item.status = f"Escala alterada 1:{old_scale:g} -> 1:{float(new_scale):g}"
                changed += 1
                self.log(f"Escala alterada: {item.name} | 1:{old_scale:g} -> 1:{float(new_scale):g}")
            except Exception as ex:
                item.status = "Erro escala: " + str(ex)
                errors.append(f"{item.name}: {ex}")
                self.log("Erro ao alterar escala de " + item.name + ": " + str(ex))

        if changed:
            try:
                drawing.CommitChanges("Organizador de Vista 2.1 - alterar escala")
            except Exception:
                pass
            try:
                handler.SaveActiveDrawing()
            except Exception:
                pass
        if errors:
            raise RuntimeError("Algumas escalas nao foram alteradas:\n" + "\n".join(errors[:8]))
        return changed

    def object_type_name(self, obj):
        try:
            return str(obj.GetType().Name)
        except Exception:
            return obj.__class__.__name__ if obj is not None else ""

    def object_stable_key(self, obj):
        for attr in ["Identifier", "DrawingObjectIdentifier"]:
            try:
                identifier = getattr(obj, attr)
                if identifier is not None:
                    return f"{self.object_type_name(obj)}:{identifier}"
            except Exception:
                pass
        return f"{self.object_type_name(obj)}:{id(obj)}"

    def is_straight_dimension(self, obj):
        if obj is None:
            return False
        try:
            if self.StraightDimension is not None and isinstance(obj, self.StraightDimension):
                return True
        except Exception:
            pass
        return self.object_type_name(obj) == "StraightDimension"

    def is_straight_dimension_set(self, obj):
        if obj is None:
            return False
        try:
            if self.StraightDimensionSet is not None and isinstance(obj, self.StraightDimensionSet):
                return True
        except Exception:
            pass
        return self.object_type_name(obj) == "StraightDimensionSet"

    def collect_straight_dimensions(self, sheet):
        dimensions = []
        sets = []
        seen_dims = set()
        seen_sets = set()
        self.dimension_view_by_key = {}

        def remember_dim_view(obj, view):
            if obj is None or view is None:
                return
            self.dimension_view_by_key[self.object_stable_key(obj)] = view

        def add_dim(obj, view=None):
            remember_dim_view(obj, view)
            key = self.object_stable_key(obj)
            if key in seen_dims:
                return
            seen_dims.add(key)
            dimensions.append(obj)

        def add_set(obj, view=None):
            key = self.object_stable_key(obj)
            if key in seen_sets:
                return
            seen_sets.add(key)
            sets.append(obj)
            try:
                enum = obj.GetObjects()
                while enum.MoveNext():
                    child = enum.Current
                    if self.is_straight_dimension(child):
                        add_dim(child, view)
            except Exception:
                pass

        def scan_enum(enum, view=None):
            while enum.MoveNext():
                obj = enum.Current
                if self.is_straight_dimension(obj):
                    add_dim(obj, view)
                elif self.is_straight_dimension_set(obj):
                    add_set(obj, view)

        try:
            scan_enum(sheet.GetAllObjects())
        except Exception:
            pass
        try:
            enum_views = sheet.GetAllViews()
            while enum_views.MoveNext():
                view = enum_views.Current
                try:
                    scan_enum(view.GetAllObjects(), view)
                except Exception:
                    pass
        except Exception:
            pass
        return dimensions, sets

    def dimension_points_xy(self, dim):
        start = dim.StartPoint
        end = dim.EndPoint
        return float(start.X), float(start.Y), float(end.X), float(end.Y)

    def dimension_geometry_key(self, dim):
        x1, y1, x2, y2 = self.dimension_points_xy(dim)
        p1 = (round(float(x1), 3), round(float(y1), 3))
        p2 = (round(float(x2), 3), round(float(y2), 3))
        if p2 < p1:
            p1, p2 = p2, p1
        return p1, p2

    def view_group_key(self, view):
        if view is None:
            return ("view", "none")
        try:
            frame = self.get_view_frame(view)
            return (
                "frame",
                round(float(frame.left), 3),
                round(float(frame.bottom), 3),
                round(float(frame.right), 3),
                round(float(frame.top), 3),
            )
        except Exception:
            return ("view", self.object_stable_key(view))

    def dimension_view_group_key(self, dim):
        view = self.dimension_target_view(dim)
        if view is None:
            return ("view", self.object_stable_key(dim))
        return self.view_group_key(view)

    def is_horizontal_dimension(self, dim):
        try:
            x1, y1, x2, y2 = self.dimension_points_xy(dim)
        except Exception:
            return False
        dx = abs(x2 - x1)
        dy = abs(y2 - y1)
        if dx < 1.0:
            return False
        return dy <= max(0.75, dx * 0.06)

    def dimension_effective_direction_y(self, dim):
        try:
            up = dim.UpDirection
            uy = float(up.Y)
        except Exception:
            uy = 0.0
        try:
            distance = float(dim.Distance)
        except Exception:
            distance = 0.0
        return uy * distance, uy, distance

    def dimension_content_frame(self, dim):
        try:
            view = dim.GetView()
        except Exception:
            view = None
        if view is None:
            view = self.dimension_view_by_key.get(self.object_stable_key(dim))
        if view is None:
            return None
        excluded_exact = {
            "StraightDimension", "StraightDimensionSet", "CurvedDimension", "CurvedDimensionSet",
            "AngleDimension", "RadiusDimension", "Text", "TextElement", "PropertyElement",
            "ContainerElement", "Mark", "ViewMark", "MarkBase", "Symbol", "WeldMark",
            "PartMark", "BoltMark", "GridLine", "Grid", "View"
        }
        frames = []
        try:
            enum = view.GetAllObjects()
            while enum.MoveNext():
                obj = enum.Current
                tname = self.object_type_name(obj)
                if tname in excluded_exact or "Dimension" in tname or "Mark" in tname:
                    continue
                frame = self.get_object_bbox(obj)
                if frame is not None and frame.width > 0.1 and frame.height > 0.1:
                    frames.append(frame)
        except Exception:
            pass
        if not frames:
            return None
        return TeklaFrame(
            min(frame.left for frame in frames),
            min(frame.bottom for frame in frames),
            max(frame.right for frame in frames),
            max(frame.top for frame in frames),
        )

    def normalized_axis(self, axis):
        length = math.sqrt(float(axis.X) ** 2 + float(axis.Y) ** 2 + float(axis.Z) ** 2)
        if length < 0.000001:
            return 0.0, 0.0, 0.0
        return float(axis.X) / length, float(axis.Y) / length, float(axis.Z) / length

    def transform_model_point_to_view_xy(self, point, coordinate_system):
        origin = coordinate_system.Origin
        axis_x = self.normalized_axis(coordinate_system.AxisX)
        axis_y = self.normalized_axis(coordinate_system.AxisY)
        rel_x = float(point.X) - float(origin.X)
        rel_y = float(point.Y) - float(origin.Y)
        rel_z = float(point.Z) - float(origin.Z)
        x = rel_x * axis_x[0] + rel_y * axis_x[1] + rel_z * axis_x[2]
        y = rel_x * axis_y[0] + rel_y * axis_y[1] + rel_z * axis_y[2]
        return x, y

    def model_object_view_frame(self, view, drawing_model_object):
        if view is None or drawing_model_object is None:
            return None
        try:
            model_identifier = drawing_model_object.ModelIdentifier
        except Exception:
            return None
        try:
            model = self.Model()
            model_object = model.SelectModelObject(model_identifier)
            if model_object is None:
                return None
            solid = model_object.GetSolid()
            min_point = solid.MinimumPoint
            max_point = solid.MaximumPoint
        except Exception:
            return None
        try:
            coordinate_system = view.DisplayCoordinateSystem
        except Exception:
            try:
                coordinate_system = view.ViewCoordinateSystem
            except Exception:
                return None
        points = []
        for x in [float(min_point.X), float(max_point.X)]:
            for y in [float(min_point.Y), float(max_point.Y)]:
                for z in [float(min_point.Z), float(max_point.Z)]:
                    points.append(self.TeklaPoint(x, y, z))
        view_points = [self.transform_model_point_to_view_xy(point, coordinate_system) for point in points]
        if not view_points:
            return None
        return TeklaFrame(
            min(point[0] for point in view_points),
            min(point[1] for point in view_points),
            max(point[0] for point in view_points),
            max(point[1] for point in view_points),
        )

    def related_part_frame_for_dimension(self, dim):
        view = self.dimension_target_view(dim)
        try:
            enum = dim.GetRelatedObjects()
            while enum.MoveNext():
                obj = enum.Current
                if self.object_type_name(obj) == "Part":
                    frame = self.model_object_view_frame(view, obj)
                    if frame is not None and frame.width > 0.1 and frame.height > 0.1:
                        return frame
        except Exception:
            pass
        if view is None:
            return None
        try:
            enum = view.GetAllObjects()
            while enum.MoveNext():
                obj = enum.Current
                if self.object_type_name(obj) == "Part":
                    frame = self.model_object_view_frame(view, obj)
                    if frame is not None and frame.width > 0.1 and frame.height > 0.1:
                        return frame
        except Exception:
            pass
        return None

    def part_top_horizontal_span(self, view, drawing_model_object):
        if view is None or drawing_model_object is None:
            return None
        try:
            model_identifier = drawing_model_object.ModelIdentifier
        except Exception:
            return None
        try:
            model = self.Model()
            model_object = model.SelectModelObject(model_identifier)
            if model_object is None:
                return None
            solid = model_object.GetSolid()
            edge_enum = solid.GetEdgeEnumerator()
        except Exception:
            return None
        try:
            coordinate_system = view.DisplayCoordinateSystem
        except Exception:
            try:
                coordinate_system = view.ViewCoordinateSystem
            except Exception:
                return None
        candidates = []
        try:
            while edge_enum.MoveNext():
                edge = edge_enum.Current
                start = self.transform_model_point_to_view_xy(edge.StartPoint, coordinate_system)
                end = self.transform_model_point_to_view_xy(edge.EndPoint, coordinate_system)
                dx = abs(end[0] - start[0])
                dy = abs(end[1] - start[1])
                if dx < 1.0:
                    continue
                if dy > max(0.75, dx * 0.02):
                    continue
                left = min(start[0], end[0])
                right = max(start[0], end[0])
                y = (start[1] + end[1]) / 2.0
                candidates.append((y, left, right))
        except Exception:
            return None
        if not candidates:
            return None
        top_y = max(candidate[0] for candidate in candidates)
        top_candidates = [candidate for candidate in candidates if abs(candidate[0] - top_y) <= 1.0]
        if not top_candidates:
            return None
        left = min(candidate[1] for candidate in top_candidates)
        right = max(candidate[2] for candidate in top_candidates)
        if right - left < 1.0:
            return None
        return left, right, top_y

    def unique_xy_points(self, points, tolerance=0.25):
        unique = []
        for x, y in points:
            if not any(abs(x - ux) <= tolerance and abs(y - uy) <= tolerance for ux, uy in unique):
                unique.append((float(x), float(y)))
        return unique

    def circle_from_three_points_xy(self, p1, p2, p3):
        x1, y1 = p1
        x2, y2 = p2
        x3, y3 = p3
        det = 2.0 * (x1 * (y2 - y3) + x2 * (y3 - y1) + x3 * (y1 - y2))
        if abs(det) < 0.000001:
            return None
        q1 = x1 * x1 + y1 * y1
        q2 = x2 * x2 + y2 * y2
        q3 = x3 * x3 + y3 * y3
        cx = (q1 * (y2 - y3) + q2 * (y3 - y1) + q3 * (y1 - y2)) / det
        cy = (q1 * (x3 - x2) + q2 * (x1 - x3) + q3 * (x2 - x1)) / det
        radius = math.hypot(x1 - cx, y1 - cy)
        if radius < 1.0:
            return None
        return float(cx), float(cy), float(radius)

    def part_view_edges_xy(self, view, drawing_model_object):
        if view is None or drawing_model_object is None:
            return []
        try:
            model_object = self.Model().SelectModelObject(drawing_model_object.ModelIdentifier)
            if model_object is None:
                return []
            solid = model_object.GetSolid()
            edge_enum = solid.GetEdgeEnumerator()
        except Exception:
            return []
        try:
            coordinate_system = view.DisplayCoordinateSystem
        except Exception:
            try:
                coordinate_system = view.ViewCoordinateSystem
            except Exception:
                return []
        edges = []
        try:
            while edge_enum.MoveNext():
                edge = edge_enum.Current
                start = self.transform_model_point_to_view_xy(edge.StartPoint, coordinate_system)
                end = self.transform_model_point_to_view_xy(edge.EndPoint, coordinate_system)
                if math.hypot(end[0] - start[0], end[1] - start[1]) > 0.5:
                    edges.append((start, end))
        except Exception:
            return []
        return edges

    def detect_circular_side_cut(self, view, drawing_model_object):
        frame = self.model_object_view_frame(view, drawing_model_object)
        edges = self.part_view_edges_xy(view, drawing_model_object)
        if frame is None or not edges:
            return None
        points = self.unique_xy_points([point for edge in edges for point in edge])
        if len(points) < 3:
            return None
        side_window = max(frame.height * 0.9, frame.width * 0.25, 20.0)
        best = None
        for side in ("left", "right"):
            if side == "left":
                side_points = [point for point in points if point[0] <= frame.left + side_window]
            else:
                side_points = [point for point in points if point[0] >= frame.right - side_window]
            count = len(side_points)
            if count < 3:
                continue
            for i in range(count - 2):
                for j in range(i + 1, count - 1):
                    for k in range(j + 1, count):
                        circle = self.circle_from_three_points_xy(side_points[i], side_points[j], side_points[k])
                        if circle is None:
                            continue
                        cx, cy, radius = circle
                        if radius < frame.height * 0.55 or radius > max(frame.width, frame.height) * 3.0:
                            continue
                        if side == "left" and cx >= frame.left - 0.5:
                            continue
                        if side == "right" and cx <= frame.right + 0.5:
                            continue
                        on_circle = []
                        for point in side_points:
                            residual = abs(math.hypot(point[0] - cx, point[1] - cy) - radius)
                            if residual <= max(1.5, radius * 0.025):
                                on_circle.append(point)
                        if len(on_circle) < 3:
                            continue
                        y_span = max(point[1] for point in on_circle) - min(point[1] for point in on_circle)
                        if y_span < frame.height * 0.45:
                            continue
                        x_span = max(point[0] for point in on_circle) - min(point[0] for point in on_circle)
                        score = len(on_circle) * 1000.0 + y_span * 10.0 - x_span
                        if best is None or score > best["score"]:
                            top = max(on_circle, key=lambda p: p[1])
                            bottom = min(on_circle, key=lambda p: p[1])
                            middle = min(on_circle, key=lambda p: abs(p[1] - cy))
                            best = {
                                "score": score,
                                "side": side,
                                "center": (cx, cy),
                                "radius": radius,
                                "frame": frame,
                                "arc_points": (bottom, middle, top),
                            }
        return best

    def circular_feature_coverage(self, points, center):
        if len(points) < 2:
            return 0.0
        cx, cy = center
        angles = sorted((math.atan2(y - cy, x - cx) % (2.0 * math.pi)) for x, y in points)
        gaps = [angles[index + 1] - angles[index] for index in range(len(angles) - 1)]
        gaps.append((angles[0] + 2.0 * math.pi) - angles[-1])
        return 2.0 * math.pi - max(gaps)

    def detect_part_circular_features(self, view, drawing_model_object):
        side_cut = self.detect_circular_side_cut(view, drawing_model_object)
        if side_cut is not None:
            return [{
                "center": (float(side_cut["center"][0]), float(side_cut["center"][1])),
                "radius": abs(float(side_cut["radius"])),
                "coverage": 0.0,
                "points": 3,
            }]

        edges = self.part_view_edges_xy(view, drawing_model_object)
        points = self.unique_xy_points([point for edge in edges for point in edge])
        if len(points) < 8:
            points = []
        if len(points) > 60:
            step = float(len(points)) / 60.0
            points = [points[min(len(points) - 1, int(index * step))] for index in range(60)]

        frame = self.model_object_view_frame(view, drawing_model_object)
        max_size = max(frame.width, frame.height) if frame is not None else 1000.0
        by_key = {}
        count = len(points)
        for i in range(count - 2):
            for j in range(i + 1, count - 1):
                for k in range(j + 1, count):
                    circle = self.circle_from_three_points_xy(points[i], points[j], points[k])
                    if circle is None:
                        continue
                    cx, cy, radius = circle
                    if radius < 2.0 or radius > max(20.0, max_size * 2.5):
                        continue
                    tolerance = max(0.8, radius * 0.018)
                    on_circle = [
                        point for point in points
                        if abs(math.hypot(point[0] - cx, point[1] - cy) - radius) <= tolerance
                    ]
                    if len(on_circle) < 8:
                        continue
                    coverage = self.circular_feature_coverage(on_circle, (cx, cy))
                    # O caminho generico aceita somente circulos praticamente
                    # completos. Arcos laterais sao tratados acima pelo detector
                    # especifico, evitando circunferencias falsas em perfis retos.
                    if coverage < math.radians(270.0):
                        continue
                    key = (round(cx, 1), round(cy, 1), round(radius, 1))
                    score = len(on_circle) * 1000.0 + coverage * 100.0
                    current = by_key.get(key)
                    if current is None or score > current[0]:
                        by_key[key] = (score, {
                            "center": (float(cx), float(cy)),
                            "radius": float(radius),
                            "coverage": float(coverage),
                            "points": len(on_circle),
                        })

        candidates = [entry[1] for entry in sorted(by_key.values(), key=lambda entry: entry[0], reverse=True)]
        features = []
        for candidate in candidates:
            cx, cy = candidate["center"]
            radius = candidate["radius"]
            duplicate = False
            for feature in features:
                fx, fy = feature["center"]
                center_tolerance = max(1.5, radius * 0.04, feature["radius"] * 0.04)
                radius_tolerance = max(1.5, radius * 0.04, feature["radius"] * 0.04)
                if math.hypot(cx - fx, cy - fy) <= center_tolerance and abs(radius - feature["radius"]) <= radius_tolerance:
                    duplicate = True
                    break
            if not duplicate:
                features.append(candidate)

        return features[:8]

    def is_center_line_type(self, line):
        try:
            attrs = line.Attributes.Line
            return (
                attrs.Type == self.LineTypes.Custom("CENTER")
                or "CENTER" in str(attrs.Type).upper()
            )
        except Exception:
            return False

    def is_generated_center_cross_line(self, line, center, radius):
        if not self.is_center_line_type(line):
            return False
        try:
            x1, y1, x2, y2 = self.line_points_xy(line)
        except Exception:
            return False
        cx, cy = center
        expected_length = abs(radius) * 1.6
        length = math.hypot(x2 - x1, y2 - y1)
        length_tolerance = max(2.0, expected_length * 0.08)
        center_tolerance = max(1.5, abs(radius) * 0.04)
        midpoint_x = (x1 + x2) / 2.0
        midpoint_y = (y1 + y2) / 2.0
        axis_aligned = abs(x2 - x1) <= center_tolerance or abs(y2 - y1) <= center_tolerance
        return (
            axis_aligned
            and abs(length - expected_length) <= length_tolerance
            and math.hypot(midpoint_x - cx, midpoint_y - cy) <= center_tolerance
        )

    def create_center_cross_for_feature(self, view, feature):
        cx, cy = feature["center"]
        radius = abs(feature["radius"])
        half_length = radius * 0.8
        created = 0
        if self.insert_line_xy(view, (cx - half_length, cy), (cx + half_length, cy)) is not None:
            created += 1
        if self.insert_line_xy(view, (cx, cy - half_length), (cx, cy + half_length)) is not None:
            created += 1
        return created

    def create_detected_center_lines(self):
        handler, drawing, sheet = self.get_active_sheet()
        if self.object_type_name(drawing) != "SinglePartDrawing":
            raise RuntimeError("Abra o desenho de peca onde as linhas de centro serao criadas.")

        targets = []
        try:
            enum_views = sheet.GetAllViews()
            while enum_views.MoveNext():
                view = enum_views.Current
                try:
                    objects = view.GetAllObjects()
                    while objects.MoveNext():
                        obj = objects.Current
                        if self.object_type_name(obj) != "Part":
                            continue
                        for feature in self.detect_part_circular_features(view, obj):
                            targets.append((view, feature))
                except Exception:
                    pass
        except Exception:
            pass

        unique_targets = []
        for view, feature in targets:
            cx, cy = feature["center"]
            radius = feature["radius"]
            if any(
                self.view_group_key(view) == self.view_group_key(old_view)
                and math.hypot(cx - old_feature["center"][0], cy - old_feature["center"][1]) <= max(1.5, radius * 0.04)
                and abs(radius - old_feature["radius"]) <= max(1.5, radius * 0.04)
                for old_view, old_feature in unique_targets
            ):
                continue
            unique_targets.append((view, feature))

        removed = 0
        created = 0
        for view, feature in unique_targets:
            previous = []
            try:
                objects = view.GetAllObjects()
                while objects.MoveNext():
                    obj = objects.Current
                    if self.object_type_name(obj) == "Line" and self.is_generated_center_cross_line(
                        obj, feature["center"], feature["radius"]
                    ):
                        previous.append(obj)
            except Exception:
                pass
            for line in previous:
                try:
                    if line.Delete():
                        removed += 1
                except Exception:
                    pass
            created += self.create_center_cross_for_feature(view, feature)

        if created or removed:
            try:
                drawing.CommitChanges("Organizador de Vista - linhas de centro")
            except Exception:
                pass
            try:
                handler.SaveActiveDrawing()
            except Exception:
                pass
        self.log(
            f"L. Centro: {len(unique_targets)} arco(s)/circulo(s) identificado(s); "
            f"{created} linha(s) criada(s); {removed} linha(s) anterior(es) removida(s)."
        )
        return {"features": len(unique_targets), "created": created, "removed": removed}

    def related_part_top_horizontal_span_for_dimension(self, dim):
        view = self.dimension_target_view(dim)
        try:
            enum = dim.GetRelatedObjects()
            while enum.MoveNext():
                obj = enum.Current
                if self.object_type_name(obj) == "Part":
                    span = self.part_top_horizontal_span(view, obj)
                    if span is not None:
                        return span
        except Exception:
            pass
        if view is None:
            return None
        try:
            enum = view.GetAllObjects()
            while enum.MoveNext():
                obj = enum.Current
                if self.object_type_name(obj) == "Part":
                    span = self.part_top_horizontal_span(view, obj)
                    if span is not None:
                        return span
        except Exception:
            pass
        return None

    def current_dimension_top_gap(self):
        try:
            value = float(self.dimension_top_gap_mm)
        except Exception:
            value = DIMENSION_TOP_GAP_MM
        return value if value > 0.0 else DIMENSION_TOP_GAP_MM

    def desired_upper_dimension_distance(self, dim, old_distance, content=None):
        base_distance = abs(float(old_distance or 0.0)) if abs(float(old_distance or 0.0)) > 0.001 else 10.0
        desired_gap = self.current_dimension_top_gap()
        try:
            x1, y1, x2, y2 = self.dimension_points_xy(dim)
        except Exception:
            return max(base_distance, desired_gap)
        if content is None:
            content = self.dimension_content_frame(dim)
        if content is None:
            return max(base_distance, desired_gap)
        anchor_y = max(y1, y2)
        desired_line_y = content.top + desired_gap
        return max(base_distance, desired_line_y - anchor_y, desired_gap)

    def dimension_target_view(self, dim):
        try:
            view = dim.GetView()
        except Exception:
            view = None
        if view is None:
            view = self.dimension_view_by_key.get(self.object_stable_key(dim))
        return view

    def make_point_list(self, points):
        point_list = self.PointList()
        for point in points:
            point_list.Add(point)
        return point_list

    def force_dimension_set_distance(self, dim_set, distance):
        if dim_set is None:
            return None
        distance = float(distance)
        try:
            dim_set.Distance = distance
        except Exception:
            pass
        try:
            enum = dim_set.GetObjects()
            while enum.MoveNext():
                child = enum.Current
                if self.is_straight_dimension(child):
                    try:
                        child.Distance = distance
                    except Exception:
                        pass
                    try:
                        child.Modify()
                    except Exception:
                        pass
        except Exception:
            pass
        try:
            dim_set.Modify()
        except Exception:
            pass
        return dim_set

    def force_upper_dimension_attributes(self, attrs, distance):
        if attrs is None:
            return None
        gap = abs(float(distance))
        try:
            from System import Enum
            placing = attrs.Placing
            placing.Placing = Enum.Parse(placing.Placing.GetType(), "Fixed")
        except Exception:
            placing = None
        if placing is not None:
            try:
                placing.Direction.Positive = False
                placing.Direction.Negative = True
            except Exception:
                pass
            try:
                placing.Distance.SearchMargin = 0.0
                placing.Distance.MinimalDistance = gap
                placing.Distance.MaximalDistance = gap
            except Exception:
                pass
        return attrs

    def load_attributes_from_candidates(self, obj, names):
        try:
            attrs = obj.Attributes
        except Exception:
            return ""
        for name in names:
            try:
                if attrs.LoadAttributes(str(name)):
                    try:
                        obj.Attributes = attrs
                    except Exception:
                        pass
                    return str(name)
            except Exception:
                pass
        return ""

    def apply_dimension_attribute(self, obj, distance=None, force_upper=False):
        if obj is None:
            return ""
        loaded = self.load_attributes_from_candidates(obj, DIMENSION_ATTRIBUTE_NAMES)
        if distance is not None and force_upper:
            try:
                attrs = obj.Attributes
                attrs = self.force_upper_dimension_attributes(attrs, float(distance))
                try:
                    obj.Attributes = attrs
                except Exception:
                    pass
            except Exception:
                pass
        if distance is not None:
            try:
                obj.Distance = float(distance)
            except Exception:
                pass
        try:
            obj.Modify()
        except Exception:
            pass
        return loaded

    def apply_dimension_set_attribute(self, dim_set, distance=None, force_upper=False):
        loaded = self.apply_dimension_attribute(dim_set, distance, force_upper)
        try:
            enum = dim_set.GetObjects()
            while enum.MoveNext():
                child = enum.Current
                if self.is_straight_dimension(child):
                    child_loaded = self.apply_dimension_attribute(child, distance, force_upper)
                    loaded = loaded or child_loaded
        except Exception:
            pass
        if distance is not None:
            self.force_dimension_set_distance(dim_set, float(distance))
        return loaded

    def apply_center_line_attribute(self, line):
        if line is None:
            return False
        try:
            attrs = line.Attributes
        except Exception:
            return False
        try:
            attrs.Line.Color = self.DrawingColors.Yellow
            attrs.Line.Type = self.LineTypes.Custom("CENTER")
            line.Attributes = attrs
            return bool(line.Modify())
        except Exception:
            return False

    def divider_line_points_xy(self, line):
        try:
            start = line.StartPoint
            end = line.EndPoint
            return (float(start.X), float(start.Y)), (float(end.X), float(end.Y))
        except Exception:
            return None

    def is_divider_line_candidate(self, line, sheet_w, cfg):
        points = self.divider_line_points_xy(line)
        if not points:
            return False
        (x1, y1), (x2, y2) = points
        length = abs(x2 - x1)
        if abs(y2 - y1) > 1.0:
            return False
        usable_w = max(1.0, float(sheet_w) - float(cfg.margin_left) - float(cfg.margin_right))
        if length < usable_w * 0.60:
            return False
        try:
            attrs = line.Attributes
            line_attrs = attrs.Line
            center_type = self.LineTypes.Custom("CENTER")
            type_text = str(line_attrs.Type).upper()
            same_type = (
                line_attrs.Type == center_type
                or line_attrs.Type == self.LineTypes.DashDot
                or "CENTER" in type_text
            )
            same_color = (
                line_attrs.Color == self.DrawingColors.Yellow
                or "YELLOW" in str(line_attrs.Color).upper()
            )
            return bool(same_type and same_color)
        except Exception:
            return False

    def delete_existing_divider_lines(self, sheet, sheet_w, cfg):
        removed = []
        try:
            enum = sheet.GetAllObjects()
            while enum.MoveNext():
                obj = enum.Current
                if self.object_type_name(obj) != "Line":
                    continue
                if not self.is_divider_line_candidate(obj, sheet_w, cfg):
                    continue
                try:
                    points = self.divider_line_points_xy(obj)
                    if obj.Delete():
                        if points:
                            (x1, y1), (x2, y2) = points
                            removed.append({"x1": x1, "y1": y1, "x2": x2, "y2": y2})
                except Exception:
                    pass
        except Exception:
            pass
        return removed

    def divider_label_y(self, item):
        if getattr(item, "label_center_y", None) is not None:
            try:
                return float(item.label_center_y)
            except Exception:
                pass
        if getattr(item, "label_offset_y", None) is not None:
            try:
                return float(item.current.bottom) + float(item.label_offset_y)
            except Exception:
                pass
        try:
            return float(item.current.bottom) - 7.0
        except Exception:
            return None

    def divider_label_bottom(self, item):
        """Parte inferior do rotulo PEÇA da vista (coord. Y na folha).

        Quando o bbox do rotulo foi medido, usa o bottom real; caso contrario,
        estima a partir do centro do rotulo, descontando a meia-altura padrao.
        """
        bottom = getattr(item, "label_bottom_y", None)
        if bottom is not None:
            try:
                return float(bottom)
            except Exception:
                pass
        label_y = self.divider_label_y(item)
        if label_y is None:
            return None
        return label_y - DIVIDER_LABEL_HALF_MM

    def divider_piece_anchor(self, group):
        """Membro que representa a peca (mesma PEÇA) para gerar a divisoria.

        Peças com mais de uma vista sao empilhadas e possuem varios frames, mas
        apenas um rotulo PEÇA real (sob a vista inferior). Ao escolher esse membro
        (ou, na falta dele, a vista mais baixa), todas as vistas da peca ficam na
        MESMA fileira. Isso evita a fileira fantasma no topo da pilha, que fazia a
        divisoria cortar a vista de cima das peças empilhadas.
        """
        labeled = [m for m in group if getattr(m, "label_center_y", None) is not None]
        pool = labeled if labeled else group
        return sorted(pool, key=lambda m: (float(m.current.bottom), int(getattr(m, "index", 0))))[0]

    def divider_item_content_top(self, item):
        """Topo do conteudo visivel (desenho + cotas + detalhes) da vista da peca.

        Usado para que a divisoria fique ACIMA das pecas da fileira de baixo,
        mesmo quando uma peca e mais alta que as vizinhas (ex.: PECA 42.2, cujo
        detalhe de angulo sobe bem acima do rotulo). Cai de volta no topo do
        frame da View quando nao consegue medir o conteudo.
        """
        current = getattr(item, "current", None)
        if current is None:
            return None
        frame_top = None
        try:
            frame_top = float(current.top)
        except Exception:
            frame_top = None
        view = getattr(item, "view", None)
        if view is not None:
            try:
                content = self.get_view_content_frame(view, current)
                if content is not None:
                    content_top = float(content.top)
                    if frame_top is None:
                        return content_top
                    # O conteudo nunca deve ultrapassar o frame; usa o menor.
                    return min(frame_top, content_top)
            except Exception:
                pass
        return frame_top

    def divider_row_content_top(self, row):
        """Maior topo de conteudo entre as pecas da fileira (None se nao medir)."""
        tops = []
        for item in row.get("items", []):
            top = self.divider_item_content_top(item)
            if top is not None:
                tops.append(top)
        return max(tops) if tops else None

    def divider_rows_from_items(self, items):
        # Agrupa por PEÇA: cada peca contribui com UM ponto de referencia (o rotulo
        # real), mesmo quando tem varias vistas empilhadas. Assim a pilha inteira
        # fica na mesma fileira e nao gera divisoria fantasma cortando a vista de
        # cima. Vistas sem PEÇA identificada entram individualmente (fallback).
        groups = {}
        order = []
        singles = []
        for item in items or []:
            if item.view is None or item.current is None:
                continue
            key = str(getattr(item, "item_text", "") or "").strip()
            if key:
                if key not in groups:
                    groups[key] = []
                    order.append(key)
                groups[key].append(item)
            else:
                singles.append(item)

        candidates = []

        def add_candidate(anchor):
            label_y = self.divider_label_y(anchor)
            if label_y is None:
                return
            label_bottom = self.divider_label_bottom(anchor)
            if label_bottom is None:
                label_bottom = label_y
            candidates.append((anchor, label_y, label_bottom))

        for key in order:
            add_candidate(self.divider_piece_anchor(groups[key]))
        for item in singles:
            add_candidate(item)

        if not candidates:
            return []

        heights = sorted(max(1.0, float(anchor.current.height)) for anchor, _, _ in candidates)
        median_h = heights[len(heights) // 2]
        tolerance = max(12.0, min(30.0, median_h * 0.45))
        rows = []
        for anchor, label_y, label_bottom in sorted(candidates, key=lambda t: t[1], reverse=True):
            if not rows or abs(label_y - rows[-1]["label_y"]) > tolerance:
                rows.append({"items": [anchor], "ys": [label_y], "bottoms": [label_bottom], "label_y": label_y})
            else:
                rows[-1]["items"].append(anchor)
                rows[-1]["ys"].append(label_y)
                rows[-1]["bottoms"].append(label_bottom)
                rows[-1]["label_y"] = sum(rows[-1]["ys"]) / len(rows[-1]["ys"])
        # Parte inferior do rotulo de cada fileira: a divisoria fica "gap" mm abaixo.
        for row in rows:
            row["label_bottom"] = min(row["bottoms"]) if row["bottoms"] else row["label_y"]
        return rows

    def insert_sheet_divider_line(self, sheet, x1, y, x2):
        line = self.Line(
            sheet,
            self.TeklaPoint(float(x1), float(y), 0.0),
            self.TeklaPoint(float(x2), float(y), 0.0),
        )
        try:
            if not line.Insert():
                return None
            if not self.apply_center_line_attribute(line):
                try:
                    line.Delete()
                except Exception:
                    pass
                return None
            return line
        except Exception:
            return None

    def create_divider_lines(self, items, cfg, label_gap_mm=None):
        handler, drawing, sheet = self.get_active_sheet()
        if self.object_type_name(drawing) != "MultiDrawing":
            raise RuntimeError("Abra o multidrawing para criar as divisorias.")
        sheet_w, sheet_h = self.sheet_size(drawing)
        rows = self.divider_rows_from_items(items)
        if len(rows) < 2:
            raise RuntimeError("Nao ha fileiras suficientes para criar divisorias.")

        try:
            gap = float(str(label_gap_mm if label_gap_mm is not None else getattr(cfg, "divider_label_gap_mm", DIVIDER_LABEL_GAP_MM)).replace(",", "."))
        except Exception:
            gap = DIVIDER_LABEL_GAP_MM
        if gap < 0.0:
            gap = DIVIDER_LABEL_GAP_MM

        x1 = min(SHEET_FRAME_LEFT_MARGIN_MM, max(0.0, float(sheet_w) / 2.0))
        right_margin = min(SHEET_FRAME_RIGHT_MARGIN_MM, max(0.0, float(sheet_w) - x1))
        right_edge = float(sheet_w) - right_margin
        legend_width = max(0.0, float(getattr(cfg, "title_block_width", 0.0) or 0.0))
        legend_height = max(0.0, float(getattr(cfg, "title_block_height", 0.0) or 0.0))
        legend_left = right_edge - legend_width
        legend_top = SHEET_FRAME_BOTTOM_MARGIN_MM + legend_height
        bottom_limit = float(getattr(cfg, "margin_bottom", 35.0) or 35.0)
        top_limit = float(sheet_h) - float(getattr(cfg, "margin_top", 35.0) or 35.0)

        previous_lines = self.delete_existing_divider_lines(sheet, sheet_w, cfg)
        deleted = len(previous_lines)
        created = 0
        for index, row in enumerate(rows):
            # "Folga abaixo do rotulo": distancia da parte inferior do rotulo da
            # fileira ate a linha divisoria. A linha fica sempre "gap" mm abaixo do
            # rotulo da propria fileira (aumentar o valor afasta a linha do rotulo).
            label_bottom = float(row.get("label_bottom", row["label_y"]))
            y = label_bottom - gap
            if y <= bottom_limit or y >= top_limit:
                continue
            x2 = max(x1 + 1.0, legend_left) if legend_width > 0.0 and legend_height > 0.0 and y <= legend_top else right_edge
            if self.insert_sheet_divider_line(sheet, x1, y, x2) is not None:
                created += 1

        if created or deleted:
            try:
                drawing.CommitChanges("Organizador de Vista - divisorias")
            except Exception:
                pass
            try:
                handler.SaveActiveDrawing()
            except Exception:
                pass
        self.log(
            f"Dividir: {created} divisoria(s) criada(s); {deleted} divisoria(s) anterior(es) removida(s). "
            f"Folga abaixo do rotulo: {gap:g} mm na folha."
        )
        return {
            "created": created,
            "deleted": deleted,
            "rows": len(rows),
            "gap": gap,
            "undo_state": {
                "drawing_mark": str(getattr(drawing, "Mark", "") or ""),
                "sheet_width": float(sheet_w),
            },
        }

    def undo_divider_lines(self, undo_state, cfg):
        handler, drawing, sheet = self.get_active_sheet()
        if self.object_type_name(drawing) != "MultiDrawing":
            raise RuntimeError("Abra o multidrawing onde as divisorias foram criadas.")

        saved_mark = str((undo_state or {}).get("drawing_mark", "") or "")
        current_mark = str(getattr(drawing, "Mark", "") or "")
        if saved_mark and current_mark and saved_mark != current_mark:
            raise RuntimeError("Abra a mesma folha do multidrawing para desfazer as divisorias.")

        sheet_w, _ = self.sheet_size(drawing)
        removed = self.delete_existing_divider_lines(sheet, sheet_w, cfg)
        if removed:
            try:
                drawing.CommitChanges("Organizador de Vista - desfazer divisorias")
            except Exception:
                pass
            try:
                handler.SaveActiveDrawing()
            except Exception:
                pass
        self.log(f"Desfazer divisorias: {len(removed)} linha(s) removida(s).")
        return {"removed": len(removed)}

    def create_upper_horizontal_dimension(self, dim, content):
        view = self.dimension_target_view(dim)
        if view is None or content is None:
            return None
        left_x = float(content.left)
        right_x = float(content.right)
        y = float(content.top)
        if right_x - left_x < 1.0:
            return None
        points = self.make_point_list([
            self.TeklaPoint(left_x, y, 0.0),
            self.TeklaPoint(right_x, y, 0.0),
        ])
        handler = self.StraightDimensionSetHandler()
        up = self.TeklaVector(0.0, -1.0, 0.0)
        distance = -float(self.current_dimension_top_gap())
        try:
            old_set = dim.GetDimensionSet()
            attrs = old_set.Attributes
            if attrs is not None:
                attrs = self.force_upper_dimension_attributes(attrs, distance)
                created = handler.CreateDimensionSet(view, points, up, distance, attrs)
                self.apply_dimension_set_attribute(created, distance, force_upper=True)
                return self.force_dimension_set_distance(created, distance)
        except Exception:
            pass
        created = handler.CreateDimensionSet(view, points, up, distance)
        self.apply_dimension_set_attribute(created, distance, force_upper=True)
        return self.force_dimension_set_distance(created, distance)

    def view_straight_dimensions(self, view):
        dimensions = []
        geometry_seen = set()

        def add_dimension(dim):
            try:
                geometry = self.dimension_geometry_key(dim)
            except Exception:
                geometry = self.object_stable_key(dim)
            if geometry in geometry_seen:
                return
            geometry_seen.add(geometry)
            dimensions.append(dim)

        try:
            enum = view.GetAllObjects()
            while enum.MoveNext():
                obj = enum.Current
                if self.is_straight_dimension(obj):
                    add_dimension(obj)
                    continue
                if not self.is_straight_dimension_set(obj):
                    continue
                try:
                    children = obj.GetObjects()
                    while children.MoveNext():
                        child = children.Current
                        if self.is_straight_dimension(child):
                            add_dimension(child)
                except Exception:
                    pass
        except Exception:
            pass
        return dimensions

    def primary_part_and_frame(self, view):
        candidates = []
        try:
            enum = view.GetAllObjects()
            while enum.MoveNext():
                obj = enum.Current
                if self.object_type_name(obj) != "Part":
                    continue
                frame = self.model_object_view_frame(view, obj)
                if frame is None or frame.width < 1.0 or frame.height < 0.1:
                    continue
                candidates.append((frame.width * max(frame.height, 1.0), obj, frame))
        except Exception:
            pass
        if not candidates:
            return None, None
        _, part, frame = max(candidates, key=lambda candidate: candidate[0])
        return part, frame

    def create_missing_upper_horizontal_dimension(self, view):
        dimensions = self.view_straight_dimensions(view)
        if any(self.is_horizontal_dimension(dim) for dim in dimensions):
            return "existing"

        part, frame = self.primary_part_and_frame(view)
        if part is None or frame is None:
            return "ignored"

        span = self.part_top_horizontal_span(view, part)
        if span is not None:
            left_x, right_x, y = span
        else:
            left_x, right_x, y = frame.left, frame.right, frame.top
        if float(right_x) - float(left_x) < 1.0:
            return "ignored"

        points = self.make_point_list([
            self.TeklaPoint(float(left_x), float(y), 0.0),
            self.TeklaPoint(float(right_x), float(y), 0.0),
        ])
        distance = -float(self.current_dimension_top_gap())
        try:
            created = self.StraightDimensionSetHandler().CreateDimensionSet(
                view,
                points,
                self.TeklaVector(0.0, -1.0, 0.0),
                distance,
            )
            if created is None:
                return "error"
            self.apply_dimension_set_attribute(created, distance, force_upper=True)
            self.force_dimension_set_distance(created, distance)
            return "created"
        except Exception as ex:
            self.log("Erro ao criar cota horizontal ausente no multidrawing: " + str(ex))
            return "error"

    def create_missing_multidrawing_horizontal_dimensions(self, sheet):
        created = 0
        ignored = 0
        errors = 0
        try:
            enum = sheet.GetAllViews()
            while enum.MoveNext():
                view = enum.Current
                try:
                    if bool(view.IsSheet):
                        continue
                except Exception:
                    pass
                result = self.create_missing_upper_horizontal_dimension(view)
                if result == "created":
                    created += 1
                elif result == "error":
                    errors += 1
                elif result == "ignored":
                    ignored += 1
        except Exception as ex:
            self.log("Erro ao verificar cotas ausentes no multidrawing: " + str(ex))
            errors += 1
        return created, ignored, errors

    def delete_dimension_group(self, dimensions):
        dimension_sets = []
        deleted = False
        for dim in dimensions:
            try:
                dim_set = dim.GetDimensionSet()
                if dim_set is not None:
                    dimension_sets.append(dim_set)
            except Exception:
                pass
            try:
                if dim.Delete():
                    deleted = True
            except Exception:
                pass
        seen = set()
        for dim_set in dimension_sets:
            key = self.object_stable_key(dim_set)
            if key in seen:
                continue
            seen.add(key)
            try:
                if dim_set.Delete():
                    deleted = True
            except Exception:
                pass
        return deleted

    def dimension_set_child_count(self, dim_set):
        count = 0
        if dim_set is None:
            return count
        try:
            enum = dim_set.GetObjects()
            while enum.MoveNext():
                if self.is_straight_dimension(enum.Current):
                    count += 1
        except Exception:
            return count
        return count

    def delete_old_dimension(self, dim):
        try:
            if dim.Delete():
                return True
        except Exception:
            pass
        try:
            dim_set = dim.GetDimensionSet()
        except Exception:
            dim_set = None
        if dim_set is None:
            return False
        if self.dimension_set_child_count(dim_set) > 1:
            return False
        try:
            return bool(dim_set.Delete())
        except Exception:
            return False

    def delete_created_dimension_set(self, dim_set):
        if dim_set is None:
            return False
        try:
            return bool(dim_set.Delete())
        except Exception:
            return False

    def recreate_horizontal_dimension_group_up(self, dimensions):
        source_dim = None
        for dim in dimensions:
            if self.is_horizontal_dimension(dim):
                source_dim = dim
                break
        if source_dim is None:
            return "ignored"
        content = self.related_part_frame_for_dimension(source_dim)
        if content is None:
            content = self.dimension_content_frame(source_dim)
        if content is None:
            return "error"
        try:
            created = self.create_upper_horizontal_dimension(source_dim, content)
            if created is None:
                return "error"
            if not self.delete_dimension_group(dimensions):
                self.delete_created_dimension_set(created)
                return "error"
            self.force_dimension_set_distance(created, -float(self.current_dimension_top_gap()))
            return "changed"
        except Exception as ex:
            self.log("Erro ao recriar cota horizontal: " + str(ex))
            return "error"

    def adjust_horizontal_dimensions_up(self, top_gap_mm=None):
        if top_gap_mm is not None:
            try:
                top_gap_mm = float(str(top_gap_mm).replace(",", "."))
                if top_gap_mm > 0.0:
                    self.dimension_top_gap_mm = top_gap_mm
            except Exception:
                pass
        handler, drawing, sheet = self.get_active_sheet()
        dimensions, sets = self.collect_straight_dimensions(sheet)
        horizontal_groups = []
        group_by_key = {}
        ignored = 0
        for dim in dimensions:
            if not self.is_horizontal_dimension(dim):
                ignored += 1
                continue
            try:
                key = (self.dimension_view_group_key(dim), self.dimension_geometry_key(dim))
            except Exception:
                key = self.object_stable_key(dim)
            group = group_by_key.get(key)
            if group is None:
                group = []
                group_by_key[key] = group
                horizontal_groups.append(group)
            group.append(dim)

        changed = 0
        moved = 0
        already_up = 0
        errors = 0
        for group in horizontal_groups:
            result = self.recreate_horizontal_dimension_group_up(group)
            if result == "changed":
                changed += 1
            elif result == "ignored":
                ignored += 1
            elif result == "error":
                errors += 1
            else:
                already_up += 1

        created_missing = 0
        if self.object_type_name(drawing) == "MultiDrawing":
            created_missing, missing_ignored, missing_errors = self.create_missing_multidrawing_horizontal_dimensions(sheet)
            ignored += missing_ignored
            errors += missing_errors
            changed += created_missing

        if changed or moved:
            modified_sets = 0
            seen = set()
            for dim in dimensions:
                try:
                    dim_set = dim.GetDimensionSet()
                except Exception:
                    dim_set = None
                if dim_set is None:
                    continue
                key = self.object_stable_key(dim_set)
                if key in seen:
                    continue
                seen.add(key)
                try:
                    if dim_set.Modify():
                        modified_sets += 1
                except Exception:
                    pass
            for dim_set in sets:
                key = self.object_stable_key(dim_set)
                if key in seen:
                    continue
                seen.add(key)
                try:
                    if dim_set.Modify():
                        modified_sets += 1
                except Exception:
                    pass
            try:
                drawing.CommitChanges("Organizador de Vista - ajustar cotas horizontais")
            except Exception:
                pass
            try:
                handler.SaveActiveDrawing()
            except Exception:
                pass
            self.log(
                f"Cotas ajustadas: {changed} horizontal(is) visivel(is) recriada(s) na parte superior."
            )
            if created_missing:
                self.log(
                    f"Multidrawing corrigido: {created_missing} cota(s) horizontal(is) ausente(s) criada(s)."
                )
        else:
            self.log("Ajuste de cotas: nenhuma cota horizontal foi ajustada.")
        self.log(
            f"Cotas analisadas: {len(horizontal_groups)} horizontal(is) visivel(is). "
            f"Objetos internos Tekla: {len(dimensions)}. Ignoradas: {ignored}. Erros: {errors}."
        )
        return {
            "changed": changed,
            "moved": moved,
            "already_up": already_up,
            "ignored": ignored,
            "errors": errors,
            "total": len(horizontal_groups) + created_missing,
            "raw_total": len(dimensions),
            "created_missing": created_missing,
        }

    def insert_line_xy(self, view, p1, p2):
        line = self.Line(
            view,
            self.TeklaPoint(float(p1[0]), float(p1[1]), 0.0),
            self.TeklaPoint(float(p2[0]), float(p2[1]), 0.0),
        )
        try:
            if not line.Insert():
                return None
            self.apply_center_line_attribute(line)
            return line
        except Exception:
            return None

    def create_radius_cut_linear_dimensions(self, view, cut):
        frame = cut["frame"]
        cx, cy = cut["center"]
        y = frame.top
        points = [self.TeklaPoint(cx, y, 0.0)]
        if cut["side"] == "left":
            points.extend([
                self.TeklaPoint(frame.left, y, 0.0),
                self.TeklaPoint(frame.right, y, 0.0),
            ])
        else:
            points.extend([
                self.TeklaPoint(frame.right, y, 0.0),
                self.TeklaPoint(frame.left, y, 0.0),
            ])
        handler = self.StraightDimensionSetHandler()
        distance = -float(self.current_dimension_top_gap())
        dim_set = handler.CreateDimensionSet(
            view,
            self.make_point_list(points),
            self.TeklaVector(0.0, -1.0, 0.0),
            distance,
        )
        self.apply_dimension_set_attribute(dim_set, distance, force_upper=True)
        return dim_set

    def create_radius_cut_vertical_dimension(self, view, cut):
        frame = cut["frame"]
        cx, cy = cut["center"]
        target_y = frame.top if abs(frame.top - cy) >= abs(frame.bottom - cy) else frame.bottom
        points = self.make_point_list([
            self.TeklaPoint(cx, cy, 0.0),
            self.TeklaPoint(cx, target_y, 0.0),
        ])
        handler = self.StraightDimensionSetHandler()
        side = cut["side"]
        outside = -1.0 if side == "left" else 1.0
        vector = self.TeklaVector(outside, 0.0, 0.0)
        distance = max(18.0, min(35.0, abs(cut["radius"]) * 0.45))
        dim_set = handler.CreateDimensionSet(view, points, vector, distance)
        self.apply_dimension_set_attribute(dim_set, distance)
        return dim_set

    def radius_cut_dimension_arc_points(self, cut):
        cx, cy = cut["center"]
        radius = abs(cut["radius"])
        # Usa um trecho curto do arco real voltado para a peca. Pontos no lado
        # externo fazem o Tekla exibir um circulo auxiliar completo.
        if cut["side"] == "left":
            angles = (-37.0, -24.0, -14.0)
        else:
            angles = (217.0, 204.0, 194.0)
        points = []
        for angle in angles:
            radians = math.radians(angle)
            points.append((cx + radius * math.cos(radians), cy + radius * math.sin(radians)))
        return tuple(points)

    def radius_cut_horizontal_center_line_points(self, cut):
        frame = cut["frame"]
        cx, cy = cut["center"]
        radius = abs(cut["radius"])
        x_margin = max(14.0, frame.height * 0.28)
        extension = min(radius * 0.55, 35.0)
        if cut["side"] == "left":
            return (cx - extension, cy), (frame.left + x_margin, cy)
        return (frame.right - x_margin, cy), (cx + extension, cy)

    def radius_dimension_placement_distance(self, cut):
        # Calibrado pela posicao corrigida no Tekla. A proporcao acompanha a
        # variacao do raio e leva a chamada ate a extremidade da linha de centro.
        return max(16.0, abs(cut["radius"]) * 0.4715)

    def force_radius_dimension_placement(self, dim, distance):
        if dim is None:
            return None
        try:
            from System import Enum
            attrs = dim.Attributes
            placing = attrs.Placing
            placing.Placing = Enum.Parse(placing.Placing.GetType(), "Fixed")
            placing.Distance.SearchMargin = 0.0
            placing.Distance.MinimalDistance = 0.0
            placing.Distance.MaximalDistance = 0.0
            dim.Attributes = attrs
        except Exception:
            pass
        try:
            dim.Distance = float(distance)
        except Exception:
            pass
        try:
            dim.Modify()
        except Exception:
            pass
        return dim

    def create_radius_dimension_for_cut(self, view, cut):
        bottom, middle, top = self.radius_cut_dimension_arc_points(cut)
        distance = self.radius_dimension_placement_distance(cut)
        try:
            dim = self.RadiusDimension(
                view,
                self.TeklaPoint(bottom[0], bottom[1], 0.0),
                self.TeklaPoint(middle[0], middle[1], 0.0),
                self.TeklaPoint(top[0], top[1], 0.0),
                distance,
            )
            if not dim.Insert():
                return None
            self.apply_dimension_attribute(dim, distance)
            return self.force_radius_dimension_placement(dim, distance)
        except Exception as ex:
            self.log("Falha ao criar cota radial: " + str(ex))
            return None

    def create_radius_cut_center_lines(self, view, cut):
        frame = cut["frame"]
        cx, cy = cut["center"]
        y_margin = max(8.0, frame.height * 0.18)
        created = 0
        if self.insert_line_xy(view, (cx, frame.bottom - y_margin), (cx, frame.top + y_margin)) is not None:
            created += 1
        line_start, line_end = self.radius_cut_horizontal_center_line_points(cut)
        if self.insert_line_xy(view, line_start, line_end) is not None:
            created += 1
        return created

    def delete_horizontal_dimension_groups(self, dimensions):
        groups = []
        group_by_key = {}
        for dim in dimensions:
            if not self.is_horizontal_dimension(dim):
                continue
            try:
                key = self.dimension_geometry_key(dim)
            except Exception:
                key = self.object_stable_key(dim)
            group = group_by_key.get(key)
            if group is None:
                group = []
                group_by_key[key] = group
                groups.append(group)
            group.append(dim)
        deleted = 0
        for group in groups:
            if self.delete_dimension_group(group):
                deleted += 1
        return deleted

    def is_radius_cut_vertical_dimension(self, dim, cut):
        if self.is_horizontal_dimension(dim):
            return False
        try:
            x1, y1, x2, y2 = self.dimension_points_xy(dim)
        except Exception:
            return False
        frame = cut["frame"]
        cx, cy = cut["center"]
        dx = abs(x2 - x1)
        dy = abs(y2 - y1)
        if dy < 1.0 or dx > max(1.0, frame.width * 0.03):
            return False
        if abs(((x1 + x2) / 2.0) - cx) > max(2.0, frame.width * 0.06):
            return False
        tol_y = max(2.0, frame.height * 0.08)
        endpoints = (y1, y2)
        touches_center = min(abs(y - cy) for y in endpoints) <= tol_y
        touches_edge = min(min(abs(y - frame.top), abs(y - frame.bottom)) for y in endpoints) <= tol_y
        return touches_center and touches_edge

    def delete_radius_cut_vertical_dimension_groups(self, dimensions, cut):
        groups = []
        group_by_key = {}
        for dim in dimensions:
            if not self.is_radius_cut_vertical_dimension(dim, cut):
                continue
            try:
                key = self.dimension_geometry_key(dim)
            except Exception:
                key = self.object_stable_key(dim)
            group = group_by_key.get(key)
            if group is None:
                group = []
                group_by_key[key] = group
                groups.append(group)
            group.append(dim)
        deleted = 0
        for group in groups:
            if self.delete_dimension_group(group):
                deleted += 1
        return deleted

    def line_points_xy(self, line):
        start = line.StartPoint
        end = line.EndPoint
        return float(start.X), float(start.Y), float(end.X), float(end.Y)

    def is_radius_cut_center_line(self, line, cut):
        try:
            x1, y1, x2, y2 = self.line_points_xy(line)
        except Exception:
            return False
        frame = cut["frame"]
        cx, cy = cut["center"]
        tol = max(1.5, frame.height * 0.05)
        dx = abs(x2 - x1)
        dy = abs(y2 - y1)
        if dx < tol and dy > frame.height * 0.45:
            y_low = min(y1, y2)
            y_high = max(y1, y2)
            return abs(((x1 + x2) / 2.0) - cx) <= tol and y_low <= frame.top and y_high >= frame.bottom
        if dy < tol and dx > frame.height * 0.35:
            x_low = min(x1, x2)
            x_high = max(x1, x2)
            expected_start, expected_end = self.radius_cut_horizontal_center_line_points(cut)
            expected_low = min(float(expected_start[0]), float(expected_end[0]))
            expected_high = max(float(expected_start[0]), float(expected_end[0]))
            return (
                abs(((y1 + y2) / 2.0) - cy) <= tol
                and abs(x_low - expected_low) <= tol
                and abs(x_high - expected_high) <= tol
            )
        return False

    def is_radius_cut_radius_dimension(self, radius_dim, cut):
        try:
            points = [
                (float(radius_dim.ArcPoint1.X), float(radius_dim.ArcPoint1.Y)),
                (float(radius_dim.ArcPoint2.X), float(radius_dim.ArcPoint2.Y)),
                (float(radius_dim.ArcPoint3.X), float(radius_dim.ArcPoint3.Y)),
            ]
        except Exception:
            return False
        circle = self.circle_from_three_points_xy(points[0], points[1], points[2])
        if circle is None:
            return False
        cx, cy = cut["center"]
        center_x, center_y, radius = circle
        center_delta = math.hypot(center_x - cx, center_y - cy)
        radius_delta = abs(radius - abs(cut["radius"]))
        return center_delta <= max(8.0, abs(cut["radius"]) * 0.15) and radius_delta <= max(8.0, abs(cut["radius"]) * 0.2)

    def delete_radius_cut_previous_objects(self, view, cut, dimensions):
        deleted = self.delete_radius_cut_vertical_dimension_groups(dimensions, cut)
        targets = []
        try:
            enum = view.GetAllObjects()
            while enum.MoveNext():
                obj = enum.Current
                tname = self.object_type_name(obj)
                if tname == "Line" and self.is_radius_cut_center_line(obj, cut):
                    targets.append(obj)
                elif tname == "RadiusDimension" and self.is_radius_cut_radius_dimension(obj, cut):
                    targets.append(obj)
        except Exception:
            pass
        for obj in targets:
            try:
                if obj.Delete():
                    deleted += 1
            except Exception:
                pass
        return deleted

    def dimension_radius_cut(self, top_gap_mm=None):
        if top_gap_mm is not None:
            try:
                top_gap_mm = float(str(top_gap_mm).replace(",", "."))
                if top_gap_mm > 0.0:
                    self.dimension_top_gap_mm = top_gap_mm
            except Exception:
                pass
        handler, drawing, sheet = self.get_active_sheet()
        dimensions, _ = self.collect_straight_dimensions(sheet)
        target = None
        try:
            enum_views = sheet.GetAllViews()
            while enum_views.MoveNext() and target is None:
                view = enum_views.Current
                enum = view.GetAllObjects()
                while enum.MoveNext():
                    obj = enum.Current
                    if self.object_type_name(obj) != "Part":
                        continue
                    cut = self.detect_circular_side_cut(view, obj)
                    if cut is not None:
                        target = (view, cut)
                        break
        except Exception:
            pass
        if target is None:
            raise RuntimeError("Nao encontrei recorte circular lateral na peca ativa.")
        view, cut = target
        previous = self.delete_radius_cut_previous_objects(view, cut, dimensions)
        deleted = self.delete_horizontal_dimension_groups(dimensions)
        if previous or deleted:
            try:
                drawing.CommitChanges("Organizador de Vista - limpar cotas de raio anteriores")
            except Exception:
                pass
        center_lines = self.create_radius_cut_center_lines(view, cut)
        linear_set = self.create_radius_cut_linear_dimensions(view, cut)
        vertical_set = self.create_radius_cut_vertical_dimension(view, cut)
        radius_dim = self.create_radius_dimension_for_cut(view, cut)
        try:
            drawing.CommitChanges("Organizador de Vista - cotar raio")
        except Exception:
            pass
        try:
            handler.SaveActiveDrawing()
        except Exception:
            pass
        cx, cy = cut["center"]
        self.log(
            f"Cotar raio: centro X={cx:.1f}, Y={cy:.1f}; R={cut['radius']:.1f}; "
            f"linhas centro={center_lines}; cotas lineares={'sim' if linear_set else 'nao'}; "
            f"vertical={'sim' if vertical_set else 'nao'}; radial={'sim' if radius_dim else 'nao'}; "
            f"cotas horizontais removidas={deleted}; objetos anteriores removidos={previous}."
        )
        return {
            "center_x": cx,
            "center_y": cy,
            "radius": cut["radius"],
            "center_lines": center_lines,
            "linear": bool(linear_set),
            "vertical": bool(vertical_set),
            "radial": bool(radius_dim),
            "deleted": deleted,
            "previous_deleted": previous,
        }

    def apply_layout(self, items, handler, drawing):
        backup = self.save_backup(items)
        self.log("Backup salvo: " + str(backup))
        moved = 0
        total_items = max(1, len(items))
        try:
            sheet = drawing.GetSheet()
        except Exception:
            sheet = None
        for item_index, item in enumerate(items, start=1):
            self.report_progress(90.0 * item_index / total_items, "Movendo vistas...")
            if not item.has_target:
                continue
            if item.view is None and getattr(item, "source_drawing", None) is not None:
                if sheet is None:
                    item.status = "Falha: folha indisponivel"
                    continue
                try:
                    point = self.TeklaPoint(float(item.target_left), float(item.target_bottom), 0.0)
                    size = self.Size(max(1.0, item.current.width), max(1.0, item.current.height))
                    link = self.DrawingLink(sheet, point, item.source_drawing, "", size)
                    if not link.Insert():
                        item.status = "Falha ao inserir"
                        self.log("Falha ao inserir link: " + item.name)
                        continue
                    frame = self.get_object_bbox(link)
                    if frame is not None:
                        dx = item.target_left - frame.left
                        dy = item.target_bottom - frame.bottom
                        if abs(dx) > 0.01 or abs(dy) > 0.01:
                            link.MoveObjectRelative(self.TeklaVector(dx, dy, 0.0))
                            link.Modify()
                        item.current = self.get_object_bbox(link) or TeklaFrame(item.target_left, item.target_bottom, item.target_right, item.target_top)
                    item.view = link
                    item.moved = True
                    item.status = "Inserida"
                    moved += 1
                    self.log(f"Inserida: {item.name} -> X={item.target_left:.1f}, Y={item.target_bottom:.1f}")
                except Exception as ex:
                    item.status = "Erro: " + str(ex)
                    self.log("Erro ao inserir " + item.name + ": " + str(ex))
                continue
            if item.view is None:
                continue
            current = self.get_view_frame(item.view)
            dx = item.target_left - current.left
            dy = item.target_bottom - current.bottom
            if abs(dx) < 0.01 and abs(dy) < 0.01:
                item.status = "Já estava na posição"
                continue
            try:
                if self.is_drawing_link_object(item.view):
                    if item.view.MoveObjectRelative(self.TeklaVector(dx, dy, 0.0)):
                        try:
                            item.view.Modify()
                        except Exception:
                            pass
                        item.moved = True
                        item.status = "Movida"
                        moved += 1
                        self.log(f"Movida: {item.name} -> X={item.target_left:.1f}, Y={item.target_bottom:.1f}")
                    else:
                        item.status = "Falha ao mover link"
                        self.log("Falha ao mover link: " + item.name)
                    continue
                origin = item.view.Origin
                self.ensure_view_free_for_moving(item.view)
                item.view.Origin = self.TeklaPoint(origin.X + dx, origin.Y + dy, origin.Z)
                if not item.view.Modify():
                    item.status = "Falha ao modificar"
                    self.log("Falha ao modificar: " + item.name)
                    continue
                item.moved = True
                item.status = "Movida"
                moved += 1
                self.log(f"Movida: {item.name} -> X={item.target_left:.1f}, Y={item.target_bottom:.1f}")
            except Exception as ex:
                item.status = "Erro: " + str(ex)
                self.log("Erro ao mover " + item.name + ": " + str(ex))
        try:
            drawing.CommitChanges("Organizar vistas de pecas")
        except Exception:
            pass
        try:
            handler.SaveActiveDrawing()
        except Exception:
            pass
        return moved


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1080x700")
        self.minsize(1040, 660)
        self.configure(bg=UI["bg"])
        self.cfg = Config.load()
        self.api = TeklaApi(self.log)
        try:
            self.api.dimension_top_gap_mm = float(str(getattr(self.cfg, "dimension_top_gap_mm", DIMENSION_TOP_GAP_MM)).replace(",", "."))
        except Exception:
            self.api.dimension_top_gap_mm = DIMENSION_TOP_GAP_MM
        self.items = []
        self.handler = None
        self.drawing = None
        self.undo_states = []
        self.alert_var = tk.StringVar(value="")
        self.last_multidrawing_sheet = self.configured_target_sheet()
        self.last_capacity_factor_x = max(1.0, float(getattr(self.cfg, "capacity_factor_x", 1.0) or 1.0))
        self.last_capacity_factor_y = max(1.0, float(getattr(self.cfg, "capacity_factor_y", 1.0) or 1.0))
        self.last_capacity_factor_matches = int(float(getattr(self.cfg, "capacity_factor_matches", 0) or 0))
        self._was_iconic = False
        self.attributes("-topmost", bool(self.cfg.topmost))
        self.build_ui_html()
        self.protocol("WM_DELETE_WINDOW", self.confirm_close)
        self.bind("<Unmap>", self.on_window_unmap, add="+")
        self.bind("<Map>", self.on_window_map, add="+")
        self.log("Abra o multidrawing no Tekla e deixe o desenho ativo.")
        self.log("Versao 2.1: usa Tekla API para mover View.Origin e organiza no padrao da base enviada: esquerda para direita, depois para baixo.")

    def configured_target_sheet(self):
        try:
            width = float(getattr(self.cfg, "target_sheet_width", 0.0) or 0.0)
            height = float(getattr(self.cfg, "target_sheet_height", 0.0) or 0.0)
        except Exception:
            return None
        if width > 1.0 and height > 1.0:
            return width, height
        return None

    def drawing_type_name(self, drawing=None):
        drawing = drawing or self.drawing
        try:
            return str(drawing.GetType().Name)
        except Exception:
            return ""

    def remember_multidrawing_sheet(self, sheet_w=None, sheet_h=None):
        if self.drawing_type_name() != "MultiDrawing":
            return False
        try:
            if sheet_w is None or sheet_h is None:
                sheet_w = self.api.last_sheet_w or self.api.sheet_size(self.drawing)[0]
                sheet_h = self.api.last_sheet_h or self.api.sheet_size(self.drawing)[1]
            sheet_w = float(sheet_w)
            sheet_h = float(sheet_h)
        except Exception:
            return False
        if sheet_w <= 1.0 or sheet_h <= 1.0:
            return False
        old = self.last_multidrawing_sheet
        self.last_multidrawing_sheet = (sheet_w, sheet_h)
        self.cfg.target_sheet_width = sheet_w
        self.cfg.target_sheet_height = sheet_h
        self.cfg.save()
        if old is None or abs(old[0] - sheet_w) > 0.1 or abs(old[1] - sheet_h) > 0.1:
            self.log(f"Folha de destino lembrada: multidrawing {sheet_w:.0f} x {sheet_h:.0f}.")
        return True

    def matching_source_drawings(self, active_items, limit=80):
        """Localiza os croquis correspondentes às vistas reais do multidrawing."""
        if self.handler is None or not active_items:
            return []
        wanted = {
            str(item.item_text or "").strip()
            for item in active_items
            if str(item.item_text or "").strip()
        }
        if not wanted:
            return []
        matched = []
        for drawing in self.api.collect_all_single_part_drawings(self.handler):
            info = self.api.drawing_piece_info(drawing)
            if not info or str(info[0] or "").strip() not in wanted:
                continue
            matched.append(drawing)
            if len(matched) >= int(limit):
                break
        return self.api.sorted_single_part_drawings(matched)

    def calibrate_capacity_from_active_items(self, active_items):
        """Aprende a diferença entre o croqui e o DrawingLink realmente inserido."""
        drawings = self.matching_source_drawings(active_items)
        if len(drawings) < 3:
            return 1.0, 1.0, 0
        source_items = self.api.build_capacity_items_from_drawings(
            drawings,
            handler=self.handler,
            original_drawing=self.drawing,
        )
        factor_x, factor_y, matched = self.api.estimate_capacity_factors(active_items, source_items)
        if matched >= 3:
            self.remember_capacity_factors(factor_x, factor_y, matched)
            self.log(
                f"Capacidade calibrada com {matched} vista(s) reais: "
                f"X={factor_x:.3f}, Y={factor_y:.3f}."
            )
        return factor_x, factor_y, matched

    def remember_capacity_factors(self, factor_x, factor_y, matched):
        if matched < 3:
            return
        self.last_capacity_factor_x = max(1.0, float(factor_x or 1.0))
        self.last_capacity_factor_y = max(1.0, float(factor_y or 1.0))
        self.last_capacity_factor_matches = int(matched)
        self.cfg.capacity_factor_x = self.last_capacity_factor_x
        self.cfg.capacity_factor_y = self.last_capacity_factor_y
        self.cfg.capacity_factor_matches = self.last_capacity_factor_matches
        self.cfg.save()

    def remember_sheet_plan(self, rows):
        rows = list(rows or [])
        if not rows:
            return False
        self.sheet_plan_rows = rows
        # O plano e apenas uma sugestao visual. Somente vistas realmente
        # encontradas no multidrawing podem avancar a fila persistente.
        return True

    def capacity_cursor_key(self):
        parsed = try_parse_numeric_item(str(getattr(self.cfg, "capacity_cursor_item", "") or ""))
        if not parsed:
            return None
        return numeric_key(parsed[1])

    def drawing_mark_text(self, drawing=None):
        drawing = drawing or self.drawing
        try:
            return str(drawing.Mark or "")
        except Exception:
            return ""

    def update_capacity_cursor(self, item_text, sort_key=None, count=0, source=""):
        if not item_text:
            return False
        if sort_key is None:
            parsed = try_parse_numeric_item(str(item_text))
            if not parsed:
                return False
            _, sort_key = parsed
        new_key = numeric_key(sort_key)
        current_key = self.capacity_cursor_key()
        if current_key is not None and new_key <= current_key:
            return False
        self.cfg.capacity_cursor_item = str(item_text)
        sheet_mark = self.drawing_mark_text()
        self.cfg.capacity_cursor_sheet = sheet_mark
        self.cfg.capacity_cursor_count = max(int(float(getattr(self.cfg, "capacity_cursor_count", 0) or 0)), int(count or 0))
        history = getattr(self.cfg, "capacity_sheet_cursors", {}) or {}
        if not isinstance(history, dict):
            history = {}
        sheet_key = self.api.normalized_sheet_mark(sheet_mark)
        if sheet_key:
            history[sheet_key] = {
                "item": str(item_text),
                "count": int(count or 0),
            }
            self.cfg.capacity_sheet_cursors = history
        self.cfg.save()
        detail = f"Fila automatica atualizada ate PEÇA {item_text}"
        if source:
            detail += f" ({source})"
        detail += "."
        self.log(detail)
        return True

    def remember_progress_from_multidrawing(self, items):
        if self.drawing_type_name() != "MultiDrawing" or not items:
            return
        # Avança a fila somente pelo trecho contínuo que recebeu posição. Se uma
        # peça intermediária não couber, peças posteriores não podem fazê-la ser
        # pulada na próxima folha.
        ordered = sorted(
            [item for item in items if item.view is not None and item.item_text and item.sort_key],
            key=lambda item: numeric_key(item.sort_key),
        )
        valid = []
        for item in ordered:
            if not item.has_target:
                break
            valid.append(item)
        if not valid:
            return
        last = valid[-1]
        self.update_capacity_cursor(
            last.item_text,
            sort_key=last.sort_key,
            count=len(valid),
            source=f"verificado no multidrawing {self.drawing_mark_text()}",
        )

    def auto_queue_drawings_after_cursor(self):
        cursor_item = str(getattr(self.cfg, "capacity_cursor_item", "") or "").strip()
        if not cursor_item:
            return [], (
                "Nao existe fila automatica registrada ainda. "
                "Abra a folha anterior que ja recebeu as peças e clique em Analisar, "
                "ou selecione os croquis no Document Manager uma vez e clique em Capacidade."
            )
        cursor_sheet = str(getattr(self.cfg, "capacity_cursor_sheet", "") or "").strip()
        current_sheet = self.drawing_mark_text().strip()
        if cursor_sheet and current_sheet and cursor_sheet == current_sheet:
            return [], (
                f"A fila automatica esta registrada nesta mesma folha ({current_sheet}). "
                "Para recalcular esta folha, selecione os croquis no Document Manager. "
                "Para continuar a fila, abra a proxima folha do multidrawing."
            )
        all_drawings = self.api.collect_all_single_part_drawings(self.handler)
        drawings = self.api.single_part_drawings_after(all_drawings, cursor_item)
        if not drawings:
            return [], f"Nao encontrei croquis depois da PEÇA {cursor_item}. A fila automatica parece ter terminado."
        return drawings, ""

    def confirm_close(self):
        if messagebox.askokcancel(APP_NAME, "Fechar o Organizador de Vista?"):
            self.destroy()

    def on_window_unmap(self, event):
        if event.widget is not self:
            return
        try:
            if self.state() == "iconic" and not self._was_iconic:
                self._was_iconic = True
                self.log("Janela minimizada. O programa continua aberto.")
        except Exception:
            pass

    def on_window_map(self, event):
        if event.widget is not self:
            return
        if self._was_iconic:
            self._was_iconic = False
            self.log("Janela restaurada.")

    def build_ui_html(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(
            "Org.Treeview",
            background="#ffffff",
            fieldbackground="#ffffff",
            foreground=UI["text"],
            rowheight=24,
            borderwidth=0,
            font=("Segoe UI", 9),
        )
        style.configure(
            "Org.Treeview.Heading",
            background="#ffffff",
            foreground="#26354d",
            font=("Segoe UI", 8, "bold"),
            relief="flat",
            bordercolor=UI["line"],
        )
        style.map("Org.Treeview", background=[("selected", "#dbe8ff")], foreground=[("selected", UI["text"])])

        root = tk.Frame(self, bg="#ffffff", bd=1, relief="solid")
        root.pack(fill="both", expand=True, padx=12, pady=10)
        main = tk.Frame(root, bg="#fbfdff", padx=16, pady=12)
        main.pack(fill="both", expand=True)

        header = tk.Frame(main, bg="#fbfdff")
        header.pack(fill="x", pady=(0, 10))
        brand = tk.Frame(header, bg="#fbfdff")
        brand.pack(side="left", fill="x")

        cube = tk.Canvas(brand, width=42, height=42, bg="#fbfdff", highlightthickness=0)
        cube.pack(side="left", padx=(0, 12))
        cube.create_polygon(21, 5, 37, 14, 21, 23, 5, 14, fill=UI["blue"], outline=UI["blue"])
        cube.create_polygon(6, 18, 19, 25, 19, 38, 6, 30, fill="#2a56c3", outline="#2a56c3")
        cube.create_polygon(36, 18, 23, 25, 23, 38, 36, 30, fill="#244aa8", outline="#244aa8")
        tk.Label(
            brand,
            text=APP_NAME,
            bg="#fbfdff",
            fg="#111827",
            font=("Segoe UI", 17),
            anchor="w",
        ).pack(side="left")

        self.top_var = tk.BooleanVar(value=bool(self.cfg.topmost))
        tk.Checkbutton(
            header,
            text="Manter à frente",
            variable=self.top_var,
            command=self.toggle_top,
            indicatoron=False,
            bd=1,
            relief="solid",
            bg="#f7faff",
            activebackground="#eef4ff",
            fg=UI["blue"],
            activeforeground=UI["blue"],
            selectcolor="#f7faff",
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=5,
            cursor="hand2",
        ).pack(side="right")

        self.alert_frame = tk.Frame(main, bg="#fff3cd", bd=1, relief="solid")
        self.alert_label = tk.Label(
            self.alert_frame,
            textvariable=self.alert_var,
            bg="#fff3cd",
            fg="#5a4100",
            anchor="w",
            justify="left",
            wraplength=900,
            font=("Segoe UI", 9, "bold"),
            padx=8,
            pady=6,
        )
        self.alert_label.pack(fill="x")
        self.alert_frame.pack_forget()

        opts = tk.Frame(main, bg="#ffffff", bd=1, relief="solid", padx=12, pady=10)
        self.opts_frame = opts
        opts.pack(fill="x", pady=(0, 10))
        tk.Label(
            opts,
            text="Configurar parametros",
            bg="#ffffff",
            fg=UI["blue"],
            font=("Segoe UI", 10, "bold"),
            anchor="w",
        ).pack(fill="x", pady=(0, 8))

        self.vars = {}
        cards = tk.Frame(opts, bg="#ffffff")
        cards.pack(fill="x")
        for col in range(6):
            cards.columnconfigure(col, weight=1, uniform="cards")

        def entry_value(attr):
            value = getattr(self.cfg, attr)
            if attr == "auto_scale_limit" and float(value or 0.0) <= 0.0:
                return ""
            return str(value)

        def add_field(parent, label, attr):
            field = tk.Frame(parent, bg="#ffffff")
            field.pack(side="left", fill="x", expand=True, padx=3)
            tk.Label(
                field,
                text=label.upper(),
                bg="#ffffff",
                fg="#50627d",
                font=("Segoe UI", 7, "bold"),
            ).pack(fill="x")
            var = tk.StringVar(value=entry_value(attr))
            self.vars[attr] = var
            tk.Entry(
                field,
                textvariable=var,
                width=8,
                justify="center",
                bd=1,
                relief="solid",
                fg="#1f56c2",
                bg="#ffffff",
                font=("Segoe UI", 9),
            ).pack(fill="x", pady=(2, 0))

        def add_card(col, icon, title, fields):
            card = tk.Frame(cards, bg="#ffffff", bd=1, relief="solid", height=86)
            card.grid(row=0, column=col, sticky="nsew", padx=(0 if col == 0 else 5, 0 if col == 5 else 5))
            card.grid_propagate(False)
            icon_canvas = tk.Canvas(card, width=28, height=28, bg="#ffffff", highlightthickness=0)
            icon_canvas.grid(row=0, column=0, rowspan=2, padx=(9, 6), pady=12, sticky="n")
            draw_line_icon(icon_canvas, icon, UI["blue"])
            body = tk.Frame(card, bg="#ffffff")
            body.grid(row=0, column=1, sticky="nsew", padx=(0, 8), pady=8)
            card.columnconfigure(1, weight=1)
            tk.Label(
                body,
                text=title,
                bg="#ffffff",
                fg="#162238",
                font=("Segoe UI", 8, "bold"),
            ).pack(fill="x", pady=(0, 5))
            row = tk.Frame(body, bg="#ffffff")
            row.pack(fill="x")
            for field_label, attr in fields:
                add_field(row, field_label, attr)

        add_card(0, "horizontal", "Lateral", [("E", "margin_left"), ("D", "margin_right")])
        add_card(1, "vertical", "Vertical", [("S", "margin_top"), ("I", "margin_bottom")])
        add_card(2, "legend", "Legenda", [("Larg.", "title_block_width"), ("Alt.", "title_block_height")])
        add_card(3, "spacing", "Espaçamento", [("Direção X", "spacing_x"), ("Direção Y", "spacing_y")])
        add_card(4, "views", "Máx. vistas", [("Qtd.", "max_details")])
        add_card(5, "scale", "Escala", [("Sugerida", "auto_scale_limit")])

        scale_mode = tk.Frame(opts, bg="#ffffff")
        scale_mode.pack(fill="x", pady=(10, 0))
        self.scale_apply_mode_var = tk.StringVar(value=getattr(self.cfg, "scale_apply_mode", "all") or "all")
        tk.Label(scale_mode, text="Escala inteligente:", bg="#ffffff", fg="#26354d", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(2, 10))
        for label, value in [("Em todas as peças", "all"), ("Apenas em excedentes", "missing")]:
            tk.Radiobutton(
                scale_mode,
                text=label,
                variable=self.scale_apply_mode_var,
                value=value,
                bg="#ffffff",
                activebackground="#ffffff",
                fg="#26354d",
                selectcolor="#ffffff",
                font=("Segoe UI", 9),
            ).pack(side="left", padx=(0, 16))

        actions = tk.Frame(main, bg="#fbfdff")
        actions.pack(fill="x", pady=(0, 10))
        for col in range(5):
            actions.columnconfigure(col, weight=1, uniform="actions")
        buttons = [
            ("Analisar", "analyze", self.analyze_api, "white", 160),
            ("Capacidade", "capacity", self.calculate_capacity, "white", 160),
            ("Organizar", "organize", self.organize_api, "blue", 160),
            ("Organizar inteligente", "smart", self.auto_scale_and_organize, "yellow", 210),
        ]
        for col, (text, icon, command, variant, width) in enumerate(buttons):
            btn = IconButton(actions, text, icon, command=command, variant=variant, width=width)
            btn.grid(row=0, column=col, sticky="ew", padx=(0 if col == 0 else 5, 5))
        self.undo_button = IconButton(actions, "Desfazer", "undo", command=self.undo_changes, variant="white", width=135)
        self.undo_button.grid(row=0, column=4, sticky="ew", padx=(5, 0))
        self.update_undo_button()

        grid_frame = tk.Frame(main, bg="#ffffff", bd=1, relief="solid")
        grid_frame.pack(fill="both", expand=True, pady=(0, 10))
        columns = ("posicao", "vista", "escala", "status")
        self.tree = ttk.Treeview(grid_frame, columns=columns, show="headings", height=8, selectmode="extended", style="Org.Treeview")
        for col, title, width in [
            ("posicao", "Posição", 120),
            ("vista", "Vista", 260),
            ("escala", "Escala", 140),
            ("status", "Status", 220),
        ]:
            self.tree.heading(col, text=title)
            self.tree.column(col, width=width, minwidth=width, anchor="center", stretch=True)
        scroll = ttk.Scrollbar(grid_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        log_header = tk.Frame(main, bg="#fbfdff")
        log_header.pack(fill="x")
        log_title = tk.Frame(log_header, bg="#fbfdff")
        log_title.pack(side="left")
        log_icon = tk.Canvas(log_title, width=20, height=20, bg="#fbfdff", highlightthickness=0)
        log_icon.pack(side="left", padx=(0, 8))
        draw_line_icon(log_icon, "log", UI["blue"])
        tk.Label(log_title, text="Log", bg="#fbfdff", fg=UI["blue"], font=("Segoe UI", 10, "bold")).pack(side="left")
        clear_btn = tk.Frame(log_header, bg="#fbfdff", cursor="hand2")
        clear_btn.pack(side="right")
        clear_icon = tk.Canvas(clear_btn, width=18, height=18, bg="#fbfdff", highlightthickness=0)
        clear_icon.pack(side="left", padx=(0, 4))
        draw_line_icon(clear_icon, "clear", UI["blue"])
        clear_label = tk.Label(clear_btn, text="Limpar log", bg="#fbfdff", fg=UI["blue"], font=("Segoe UI", 8, "bold"), cursor="hand2")
        clear_label.pack(side="left")
        for widget in (clear_btn, clear_icon, clear_label):
            widget.bind("<Button-1>", lambda event: self.clear_log())

        log_frame = tk.Frame(main, bg="#ffffff", bd=1, relief="solid")
        log_frame.pack(fill="x", pady=(3, 7))
        self.log_text = tk.Text(
            log_frame,
            height=6,
            wrap="word",
            font=("Consolas", 8),
            bd=0,
            bg="#ffffff",
            fg="#111827",
            padx=10,
            pady=7,
        )
        self.log_text.pack(fill="both", expand=True)

        tk.Label(
            main,
            text="Programa desenvolvido por Edflávio Calavort - 2026",
            anchor="center",
            bg="#fbfdff",
            fg="#657895",
            font=("Segoe UI", 9),
        ).pack(fill="x")

    def build_ui(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Title.TLabel", font=("Segoe UI", 13, "bold"))
        style.configure("Accent.TButton", padding=6, font=("Segoe UI", 9, "bold"))
        root = ttk.Frame(self, padding=10)
        root.pack(fill="both", expand=True)
        header = ttk.Frame(root)
        header.pack(fill="x")
        ttk.Label(header, text=APP_NAME, style="Title.TLabel").pack(side="left")
        self.top_var = tk.BooleanVar(value=bool(self.cfg.topmost))
        ttk.Checkbutton(header, text="Manter à frente", variable=self.top_var, command=self.toggle_top).pack(side="right")

        self.alert_frame = tk.Frame(root, bg="#fff3cd", bd=1, relief="solid")
        self.alert_label = tk.Label(
            self.alert_frame,
            textvariable=self.alert_var,
            bg="#fff3cd",
            fg="#5a4100",
            anchor="w",
            justify="left",
            wraplength=660,
            font=("Segoe UI", 9, "bold"),
            padx=8,
            pady=6,
        )
        self.alert_label.pack(fill="x")
        self.alert_frame.pack(fill="x", pady=(0, 6))
        self.alert_frame.pack_forget()

        opts = ttk.LabelFrame(root, text="Configurar parametros")
        self.opts_frame = opts
        opts.pack(fill="x", pady=6)
        self.vars = {}
        fields = [
            ("Margem E", "margin_left"), ("Margem S", "margin_top"), ("Margem D", "margin_right"), ("Margem I", "margin_bottom"),
            ("Esp. X", "spacing_x"), ("Esp. Y", "spacing_y"), ("Larg. legenda", "title_block_width"), ("Alt. legenda", "title_block_height"),
            ("Máx. vistas", "max_details"), ("Escala sugerida", "auto_scale_limit"),
        ]
        for i, (label, attr) in enumerate(fields):
            row = i // 5
            col = (i % 5) * 2
            ttk.Label(opts, text=label).grid(row=row, column=col, padx=(6, 3), pady=4, sticky="w")
            value = getattr(self.cfg, attr)
            if attr == "auto_scale_limit" and float(value or 0.0) <= 0.0:
                value = ""
            var = tk.StringVar(value=str(value))
            self.vars[attr] = var
            ttk.Entry(opts, textvariable=var, width=9).grid(row=row, column=col + 1, padx=(0, 8), pady=4, sticky="w")
        for c in range(10):
            opts.columnconfigure(c, weight=1)

        scale_mode = ttk.Frame(root)
        scale_mode.pack(fill="x", pady=(0, 4))
        self.scale_apply_mode_var = tk.StringVar(value=getattr(self.cfg, "scale_apply_mode", "all") or "all")
        ttk.Label(scale_mode, text="Escala inteligente:").pack(side="left", padx=(2, 8))
        ttk.Radiobutton(scale_mode, text="Em todas as peças", variable=self.scale_apply_mode_var, value="all").pack(side="left", padx=(0, 12))
        ttk.Radiobutton(scale_mode, text="Apenas em excedentes", variable=self.scale_apply_mode_var, value="missing").pack(side="left")

        actions = ttk.Frame(root)
        actions.pack(fill="x", pady=6)
        ttk.Button(actions, text="Analisar", style="Accent.TButton", command=self.analyze_api).pack(side="left", expand=True, fill="x", padx=(0, 5))
        ttk.Button(actions, text="Capacidade", style="Accent.TButton", command=self.calculate_capacity).pack(side="left", expand=True, fill="x", padx=5)
        ttk.Button(actions, text="Organizar", style="Accent.TButton", command=self.organize_api).pack(side="left", expand=True, fill="x", padx=5)
        ttk.Button(actions, text="Organizar inteligente", style="Accent.TButton", command=self.auto_scale_and_organize).pack(side="left", expand=True, fill="x", padx=5)
        self.undo_button = ttk.Button(actions, text="Desfazer", command=self.undo_changes)
        self.undo_button.pack(side="left", expand=True, fill="x", padx=(5, 0))
        self.update_undo_button()

        grid_frame = ttk.Frame(root)
        grid_frame.pack(fill="both", expand=True, pady=6)
        columns = ("posicao", "vista", "escala", "status")
        self.tree = ttk.Treeview(grid_frame, columns=columns, show="headings", height=11, selectmode="extended")
        for col, title, width in [
            ("posicao", "Posição", 110),
            ("vista", "Vista", 250),
            ("escala", "Escala", 130),
            ("status", "Status", 210),
        ]:
            self.tree.heading(col, text=title)
            self.tree.column(col, width=width, minwidth=width, anchor="center", stretch=True)
        scroll = ttk.Scrollbar(grid_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        log_header = ttk.Frame(root)
        log_header.pack(fill="x", pady=(4, 0))
        ttk.Label(log_header, text="Log").pack(side="left")
        ttk.Button(log_header, text="Limpar log", command=self.clear_log).pack(side="right")

        log_frame = ttk.Frame(root)
        log_frame.pack(fill="both", expand=False, pady=(2, 6))
        self.log_text = tk.Text(log_frame, height=6, wrap="word", font=("Consolas", 8))
        self.log_text.pack(fill="both", expand=True)

        footer = ttk.Label(root, text="Programa desenvolvido por Edflávio Calavort - 2026", anchor="center", foreground="#576575")
        footer.pack(fill="x", pady=(2, 0))

    def toggle_top(self):
        self.cfg.topmost = bool(self.top_var.get())
        self.attributes("-topmost", self.cfg.topmost)
        self.cfg.save()

    def update_undo_button(self):
        try:
            state = "normal" if self.undo_states else "disabled"
            self.undo_button.configure(state=state)
        except Exception:
            pass

    def ensure_undo_snapshot(self):
        if self.undo_states:
            return True
        if not self.items:
            if not self.analyze_current(warn_empty=True):
                return False
        states = self.api.capture_undo_state(self.items)
        if not states:
            raise RuntimeError("Nao foi possivel preparar o desfazer. Nenhuma vista valida foi capturada.")
        self.undo_states = states
        self.update_undo_button()
        self.log(f"Desfazer preparado: {len(states)} vista(s) guardadas com posicao e escala originais.")
        return True

    def apply_vars(self):
        for attr, var in self.vars.items():
            raw = var.get().strip().replace(",", ".")
            try:
                if attr == "max_details":
                    setattr(self.cfg, attr, int(float(raw or "0")))
                elif attr == "density_alert":
                    value = float(raw)
                    if value > 1.0:
                        value = value / 100.0
                    setattr(self.cfg, attr, max(0.10, min(0.95, value)))
                elif attr == "auto_scale_limit":
                    setattr(self.cfg, attr, parse_scale_denominator(raw) if raw else 0.0)
                elif attr == "scale_target":
                    setattr(self.cfg, attr, parse_scale_denominator(raw))
                else:
                    setattr(self.cfg, attr, float(raw))
            except Exception:
                raise RuntimeError(f"Valor inválido em {attr}: {var.get()}")
        self.cfg.scale_apply_mode = self.scale_apply_mode_var.get()
        self.cfg.save()

    def log(self, text):
        try:
            self.log_text.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {text}\n")
            self.log_text.see("end")
            self.update_idletasks()
        except Exception:
            print(text)

    def clear_log(self):
        try:
            self.log_text.delete("1.0", "end")
        except Exception:
            pass

    def set_alert(self, text):
        self.alert_var.set(text or "")
        if self.alert_frame.winfo_ismapped():
            self.alert_frame.pack_forget()
        if text:
            self.alert_frame.pack(fill="x", pady=(0, 10), before=self.opts_frame)
            self.log("ALERTA: " + text)

    def selected_items(self):
        selected = []
        by_index = {str(it.index): it for it in self.items}
        for iid in self.tree.selection():
            item = by_index.get(str(iid))
            if item is not None:
                selected.append(item)
        return selected

    def update_layout_alert(self):
        self.set_alert("")

    def update_tree(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for it in self.items:
            atual = f"X={it.current.left:.1f}, Y={it.current.bottom:.1f} | {it.current.width:.1f}x{it.current.height:.1f}"
            dest = "-" if not it.has_target else f"X={it.target_left:.1f}, Y={it.target_bottom:.1f}"
            mov = "-" if not it.has_target else f"{it.dx:+.1f}, {it.dy:+.1f}"
            escala = f"1:{it.current_scale:g}"
            self.tree.insert("", "end", iid=str(it.index), values=(it.index, it.name, escala, it.status))

    def analyze_current(self, warn_empty=True):
        self.apply_vars()
        self.items, self.drawing, self.handler = self.api.analyze(self.cfg)
        self.remember_multidrawing_sheet()
        self.remember_progress_from_multidrawing(self.items)
        self.update_tree()
        if self.drawing_type_name() == "MultiDrawing":
            self.update_layout_alert()
        else:
            self.set_alert("")
        if warn_empty and not self.items:
            messagebox.showwarning(APP_NAME, "Nenhuma vista PEÇA foi encontrada pela API.")
        return bool(self.items)

    def apply_scale_then_reanalyze(self, items, scale):
        if not items:
            messagebox.showinfo(APP_NAME, "Nenhuma vista foi escolhida para alterar escala.")
            return False
        if self.drawing is None or self.handler is None:
            self.analyze_current(warn_empty=True)
        if not self.ensure_undo_snapshot():
            return False
        changed = self.api.apply_scale_to_items(items, scale, self.handler, self.drawing)
        self.log(f"Escala aplicada em {changed} vista(s). Reanalisando...")
        time.sleep(0.4)
        self.analyze_current(warn_empty=False)
        return True

    def auto_scale_and_organize(self):
        try:
            self.apply_vars()
            if not self.analyze_current(warn_empty=True):
                return
            if self.drawing_type_name() != "MultiDrawing":
                messagebox.showinfo(APP_NAME, "Abra o multidrawing de destino para organizar. No croqui use apenas Capacidade.")
                return
            warning = self.api.layout_warning(self.items, self.cfg)
            manual_scale = False
            if not warning and not manual_scale:
                self.log("Layout nao indicou falta de espaco. Organizando sem alterar escala.")
                self.organize_api(skip_space_check=True)
                return

            solution = self.api.find_scale_solution(self.items, self.cfg)
            scale = float(solution["scale"])
            if not solution["fits"]:
                if manual_scale:
                    reason = (
                        f"Nao vou executar: a escala sugerida 1:{scale:g} ainda deixa "
                        f"{solution['unplanned']} vista(s) sem espaco. Aumente a escala sugerida, "
                        "reduza margens/espacamentos ou use uma folha maior."
                    )
                else:
                    reason = (
                        f"Nao vou executar: testei ate a escala 1:30 "
                        f"e ainda sobraram {solution['unplanned']} vista(s) sem espaco. "
                        "Preencha uma Escala sugerida maior, reduza margens/espacamentos ou use uma folha maior."
                    )
                self.set_alert(reason)
                self.log(reason)
                messagebox.showinfo(APP_NAME, reason)
                return

            current_hard_block = [it for it in self.items if not it.has_target and "Ignorada pelo limite" not in it.status]
            if manual_scale:
                all_targets = [item for item in self.items if abs(float(item.current_scale or 1.0) - scale) > 0.001]
            else:
                all_targets = [item for item in self.items if float(item.current_scale or 1.0) < scale - 0.001]
            if self.cfg.scale_apply_mode == "missing" and current_hard_block:
                if manual_scale:
                    targets = [item for item in current_hard_block if abs(float(item.current_scale or 1.0) - scale) > 0.001]
                else:
                    targets = [item for item in current_hard_block if float(item.current_scale or 1.0) < scale - 0.001]
                mode_text = "apenas em excedentes"
            elif self.cfg.scale_apply_mode == "missing" and not manual_scale:
                targets = []
                mode_text = "apenas em excedentes"
            else:
                targets = all_targets
                mode_text = "em todas as peças"
            if not targets:
                if current_hard_block:
                    reason = (
                        f"Nao vou executar: as vistas que nao couberam nao seriam reduzidas pela escala 1:{scale:g}. "
                        "Aumente a Escala sugerida ou mude a opcao para 'Em todas as peças'."
                    )
                    self.set_alert(reason)
                    self.log(reason)
                    messagebox.showinfo(APP_NAME, reason)
                    return
                if self.cfg.scale_apply_mode == "missing" and manual_scale:
                    self.log(f"As vistas ja estao em 1:{scale:g}; organizando sem alterar escala.")
                else:
                    self.log(f"Simulacao indica que 1:{scale:g} cabe, mas as vistas ja estao nessa escala ou menor. Organizando.")
                self.organize_api(skip_space_check=True)
                return

            scale_source = "A escala sugerida preenchida sera aplicada" if manual_scale else "Para tentar resolver"
            ok = messagebox.askokcancel(
                APP_NAME,
                warning
                + f"\n\n{scale_source}: 1:{scale:g} {mode_text} ({len(targets)} vista(s)), "
                + "reanalisar e organizar somente se couber.\n\nConfirmar?",
                icon="warning",
            )
            if not ok:
                self.set_alert("Organizacao cancelada pelo usuario antes de alterar escala.")
                return

            self.apply_scale_then_reanalyze(targets, scale)
            hard_block = [it for it in self.items if not it.has_target and "Ignorada pelo limite" not in it.status]
            if hard_block:
                if self.cfg.scale_apply_mode == "missing":
                    reason = (
                        f"Nao vou executar: aplicar escala 1:{scale:g} apenas em excedentes "
                        f"nao resolveu; ainda sobraram {len(hard_block)} vista(s). "
                        "Mude a opcao para 'Em todas as peças' e rode Organizar inteligente."
                    )
                    self.set_alert(reason)
                    self.log(reason)
                    messagebox.showinfo(APP_NAME, reason)
                    return
                if manual_scale:
                    reason = (
                        f"Nao vou executar: a escala sugerida 1:{scale:g} nao resolveu; "
                        f"ainda sobraram {len(hard_block)} vista(s). Aumente a Escala sugerida, "
                        "reduza margens/espacamentos ou use uma folha maior."
                    )
                    self.set_alert(reason)
                    self.log(reason)
                    messagebox.showinfo(APP_NAME, reason)
                    return
                retry = self.api.find_scale_solution(self.items, self.cfg)
                retry_scale = float(retry["scale"])
                if retry["fits"] and retry_scale > scale + 0.001:
                    more_targets = [item for item in self.items if float(item.current_scale or 1.0) < retry_scale - 0.001]
                    self.log(f"A primeira escala ainda nao coube. Tentando 1:{retry_scale:g}.")
                    self.apply_scale_then_reanalyze(more_targets, retry_scale)
                    hard_block = [it for it in self.items if not it.has_target and "Ignorada pelo limite" not in it.status]
                    scale = retry_scale
                if not hard_block:
                    self.organize_api(skip_space_check=True)
                    return
                reason = (
                    f"Nao vou executar: mesmo depois da escala 1:{scale:g}, "
                    f"{len(hard_block)} vista(s) ainda nao couberam. "
                    "Aumente Escala sugerida, reduza margens/espacamentos, ou use uma folha maior."
                )
                self.set_alert(reason)
                self.log(reason)
                messagebox.showinfo(APP_NAME, reason)
                return

            self.organize_api(skip_space_check=True)
        except Exception as ex:
            self.log("Falha na auto escala: " + str(ex))
            self.log(traceback.format_exc())
            messagebox.showerror(APP_NAME, str(ex))

    def analyze_api(self):
        try:
            self.analyze_current(warn_empty=True)
        except Exception as ex:
            self.log("Falha na análise API: " + str(ex))
            self.log(traceback.format_exc())
            messagebox.showerror(APP_NAME, str(ex))

    def calculate_capacity(self):
        try:
            self.log("Capacidade: iniciando leitura da folha e dos croquis.")
            self.apply_vars()
            self.analyze_current(warn_empty=False)
            if self.drawing is None or self.handler is None:
                raise RuntimeError("Abra um desenho no Tekla antes de calcular a capacidade.")

            active_type = self.drawing_type_name()
            active_items_for_capacity = []
            if active_type == "MultiDrawing":
                sheet_w = self.api.last_sheet_w or self.api.sheet_size(self.drawing)[0]
                sheet_h = self.api.last_sheet_h or self.api.sheet_size(self.drawing)[1]
                active_items_for_capacity = [item for item in self.items if item.view is not None]
                self.remember_multidrawing_sheet(sheet_w, sheet_h)
            elif self.last_multidrawing_sheet:
                sheet_w, sheet_h = self.last_multidrawing_sheet
                self.log(
                    f"Folha de destino usada: ultimo multidrawing lembrado "
                    f"{sheet_w:.0f} x {sheet_h:.0f}. O croqui ativo nao foi usado como destino."
                )
            else:
                raise RuntimeError(
                    "Abra o multidrawing de destino uma vez e clique em Analisar ou Capacidade. "
                    "Depois pode voltar aos croquis/Document Manager para calcular. "
                    "Sem isso eu so vejo a folha do croqui ativo, e a conta fica errada."
                )

            selected_drawings = self.api.collect_selected_single_part_drawings(self.handler)
            auto_queue = False
            if selected_drawings:
                if len(selected_drawings) == 1:
                    start_info = self.api.drawing_piece_info(selected_drawings[0])
                    if not start_info:
                        source_drawings = selected_drawings
                        source_label = "selecionados no Document Manager"
                    else:
                        all_drawings = self.api.collect_all_single_part_drawings(self.handler)
                        source_drawings = self.api.single_part_drawings_from(all_drawings, start_info[0])
                        source_label = f"da sequencia a partir da PEÇA {start_info[0]} selecionada"
                        auto_queue = True
                else:
                    source_drawings = selected_drawings
                    source_label = "selecionados no Document Manager"
            elif active_items_for_capacity:
                source_drawings = []
                source_label = ""
            else:
                source_drawings, reason = self.auto_queue_drawings_after_cursor()
                if not source_drawings:
                    self.log(reason)
                    messagebox.showinfo(APP_NAME, reason)
                    return
                auto_queue = True
                source_label = f"da fila automatica apos PEÇA {self.cfg.capacity_cursor_item}"

            if source_drawings:
                self.log(f"Medindo {len(source_drawings)} croqui(s) {source_label} para calcular capacidade.")
            source_fit = 0
            source_failed = None
            if not source_drawings:
                source_items = []
            elif auto_queue:
                factor_x = self.last_capacity_factor_x if self.last_capacity_factor_matches >= 3 else 1.0
                factor_y = self.last_capacity_factor_y if self.last_capacity_factor_matches >= 3 else 1.0
                source_items, source_fit, source_failed = self.api.build_capacity_until_overflow(
                    source_drawings,
                    self.cfg,
                    sheet_w,
                    sheet_h,
                    factor_x=factor_x,
                    factor_y=factor_y,
                    handler=self.handler,
                    original_drawing=self.drawing,
                )
            else:
                source_items = self.api.build_capacity_items_from_drawings(
                    source_drawings,
                    handler=self.handler,
                    original_drawing=self.drawing,
                )
            if not source_items and not active_items_for_capacity:
                messagebox.showinfo(APP_NAME, "Nao encontrei vistas no multidrawing nem croquis de peça para calcular.")
                return

            applied_capacity = 0
            message_lines = []

            if active_items_for_capacity:
                self.calibrate_capacity_from_active_items(active_items_for_capacity)
                current_fit, current_failed, _ = self.api.calculate_fit_count(active_items_for_capacity, sheet_w, sheet_h, self.cfg)
                applied_capacity = current_fit
                text = f"Capacidade real na folha atual: cabem {current_fit} de {len(active_items_for_capacity)} vista(s) ja inseridas."
                if current_failed:
                    text += f" Primeira excedente: {current_failed.name}."
                self.log(text)
                message_lines.append(text)

            if source_items:
                matched = 0
                if source_fit <= 0:
                    factor_x, factor_y, matched = self.api.estimate_capacity_factors(active_items_for_capacity, source_items)
                    if matched >= 3:
                        self.remember_capacity_factors(factor_x, factor_y, matched)
                    elif self.last_capacity_factor_matches >= 3:
                        factor_x = self.last_capacity_factor_x
                        factor_y = self.last_capacity_factor_y
                    source_fit, source_failed, _ = self.api.calculate_fit_count(
                        source_items,
                        sheet_w,
                        sheet_h,
                        self.cfg,
                        factor_x=factor_x,
                        factor_y=factor_y,
                    )
                if selected_drawings or not applied_capacity:
                    applied_capacity = source_fit
                text = f"Capacidade estimada dos croquis {source_label}: cabem {source_fit} de {len(source_items)} peça(s)."
                if source_failed:
                    text += f" Primeira excedente: {source_failed.name}."
                if matched >= 3:
                    text += f" Calibrado por {matched} vista(s) ja presentes no multidrawing."
                elif self.last_capacity_factor_matches >= 3:
                    text += f" Usando calibracao anterior de {self.last_capacity_factor_matches} vista(s) do multidrawing."
                else:
                    text += " Sem amostra suficiente no multidrawing; estimativa baseada no tamanho dos croquis."
                self.log(text)
                message_lines.append(text)
                message_lines.append("Fila automatica nao foi avancada aqui; ela sera atualizada quando Analisar encontrar as vistas reais na folha.")

            if applied_capacity > 0:
                self.vars["max_details"].set(str(applied_capacity))
                self.cfg.max_details = int(applied_capacity)
                self.cfg.save()
                applied = f"Campo 'Máx. vistas' atualizado para {applied_capacity}."
                self.log(applied)
                message_lines.append(applied)

            messagebox.showinfo(APP_NAME, "\n\n".join(message_lines))
        except Exception as ex:
            self.log("Falha ao calcular capacidade: " + str(ex))
            self.log(traceback.format_exc())
            messagebox.showerror(APP_NAME, str(ex))

    def undo_changes(self):
        try:
            if not self.undo_states:
                messagebox.showinfo(APP_NAME, "Nao ha organizacao para desfazer nesta sessao.")
                return
            ok = messagebox.askokcancel(
                APP_NAME,
                f"Restaurar posicao e escala originais de {len(self.undo_states)} vista(s)?",
                icon="warning",
            )
            if not ok:
                return
            if self.drawing is None or self.handler is None or not self.items:
                self.analyze_current(warn_empty=False)
            restored = self.api.restore_undo_state(self.undo_states, self.items, self.handler, self.drawing)
            self.undo_states = []
            self.update_undo_button()
            time.sleep(0.4)
            self.analyze_current(warn_empty=False)
            self.log(f"Desfazer concluido. Vistas restauradas: {restored}.")
            messagebox.showinfo(APP_NAME, f"Desfazer concluido. Vistas restauradas: {restored}.")
        except Exception as ex:
            self.log("Falha ao desfazer: " + str(ex))
            self.log(traceback.format_exc())
            messagebox.showerror(APP_NAME, str(ex))

    def organize_api(self, skip_space_check=False):
        try:
            if not self.items or self.drawing is None or self.handler is None:
                self.analyze_api()
            if not self.items:
                return
            if self.drawing_type_name() != "MultiDrawing":
                messagebox.showinfo(APP_NAME, "Abra o multidrawing de destino para organizar. No croqui use apenas Capacidade.")
                return
            hard_block = [i for i in self.items if not i.has_target and "Ignorada pelo limite" not in i.status]
            if hard_block and not skip_space_check:
                reason = (
                    f"Nao vou executar: {len(hard_block)} vista(s) ainda nao couberam no layout. "
                    "Use Organizar inteligente."
                )
                self.set_alert(reason)
                messagebox.showinfo(APP_NAME, reason)
                return
            movable = [i for i in self.items if i.has_target]
            if not movable:
                messagebox.showinfo(APP_NAME, "Nenhuma vista planejada para mover.")
                return
            warning = self.api.layout_warning(self.items, self.cfg)
            if warning and not skip_space_check:
                ok_warn = messagebox.askokcancel(
                    APP_NAME,
                    warning + "\n\nDeseja organizar mesmo assim, sem alterar escala?",
                    icon="warning",
                )
                if not ok_warn:
                    self.set_alert("Organizacao pausada. Use Organizar inteligente.")
                    return
            ok = messagebox.askokcancel(APP_NAME, f"Mover {len(movable)} vista(s) pela Tekla API?\n\nO desenho ativo será salvo após aplicar.")
            if not ok:
                return
            if not self.ensure_undo_snapshot():
                return
            moved = self.api.apply_layout(self.items, self.handler, self.drawing)
            self.update_tree()
            self.update_layout_alert()
            self.log(f"Organização concluída. Vistas movidas: {moved}.")
            messagebox.showinfo(APP_NAME, f"Organização concluída. Vistas movidas: {moved}.")
        except Exception as ex:
            self.log("Falha ao organizar via API: " + str(ex))
            self.log(traceback.format_exc())
            messagebox.showerror(APP_NAME, str(ex))


class HtmlController:
    def __init__(self):
        self.cfg = Config.load()
        self.logs = []
        self.alert = ""
        self.api = TeklaApi(self.log)
        self.items = []
        self.handler = None
        self.drawing = None
        self.undo_states = []
        self.divider_undo_state = None
        self.sheet_plan_rows = []
        self._window = None
        self.last_multidrawing_sheet = self.configured_target_sheet()
        self.last_capacity_factor_x = max(1.0, float(getattr(self.cfg, "capacity_factor_x", 1.0) or 1.0))
        self.last_capacity_factor_y = max(1.0, float(getattr(self.cfg, "capacity_factor_y", 1.0) or 1.0))
        self.last_capacity_factor_matches = int(float(getattr(self.cfg, "capacity_factor_matches", 0) or 0))
        self.api.dimension_top_gap_mm = self.safe_dimension_top_gap()
        # Progresso real reportado pela API durante as operacoes longas.
        self._progress = {"value": 0, "text": "Processando..."}
        self.progress_emit_cb = None
        self.api.progress_cb = self.report_progress
        self.log("Abra o multidrawing no Tekla e deixe o desenho ativo.")
        self.log("Versao HTML: usa Tekla API para mover View.Origin e organiza no padrao da base enviada.")

    def reset_progress(self, text="Processando..."):
        self.api.set_progress_window(0.0, 100.0)
        self._progress = {"value": 0, "text": text}
        self._notify_progress()

    def report_progress(self, value, text=None):
        # Mantem no maximo 99 durante o trabalho; a interface fecha em 100.
        v = max(0, min(99, int(round(float(value or 0.0)))))
        cur = dict(self._progress)
        prev = int(cur.get("value", 0) or 0)
        cur["value"] = max(prev, v)  # monotonico: nunca recua
        if text:
            cur["text"] = text
        self._progress = cur
        self._notify_progress()

    def get_progress(self):
        return dict(self._progress)

    def _notify_progress(self):
        cb = getattr(self, "progress_emit_cb", None)
        if not cb:
            return
        try:
            cb(dict(self._progress))
        except Exception:
            pass

    def safe_dimension_top_gap(self):
        try:
            value = float(str(getattr(self.cfg, "dimension_top_gap_mm", DIMENSION_TOP_GAP_MM)).replace(",", "."))
        except Exception:
            value = DIMENSION_TOP_GAP_MM
        return value if value > 0.0 else DIMENSION_TOP_GAP_MM

    def safe_divider_label_gap(self):
        try:
            value = float(str(getattr(self.cfg, "divider_label_gap_mm", DIVIDER_LABEL_GAP_MM)).replace(",", "."))
        except Exception:
            value = DIVIDER_LABEL_GAP_MM
        return value if value >= 0.0 else DIVIDER_LABEL_GAP_MM

    def set_window(self, window):
        self._window = window

    def sync_dimension_hotkey_setting(self):
        try:
            if self._window is not None and hasattr(self._window, "sync_dimension_hotkey_setting"):
                self._window.sync_dimension_hotkey_setting()
        except Exception:
            pass

    def log(self, text):
        line = f"[{datetime.now().strftime('%H:%M:%S')}]  {text}"
        self.logs.append(line)
        self.logs = self.logs[-300:]
        print(text)

    def config_payload(self):
        return {
            "margin_left": self.cfg.margin_left,
            "margin_right": self.cfg.margin_right,
            "margin_top": self.cfg.margin_top,
            "margin_bottom": self.cfg.margin_bottom,
            "title_block_width": self.cfg.title_block_width,
            "title_block_height": self.cfg.title_block_height,
            "spacing_x": self.cfg.spacing_x,
            "spacing_y": self.cfg.spacing_y,
            "stack_spacing_y": float(getattr(self.cfg, "stack_spacing_y", STACK_VIEW_SPACING_MM) or 0.0),
            "max_details": self.cfg.max_details,
            "auto_scale_limit": self.cfg.auto_scale_limit,
            "scale_apply_mode": self.cfg.scale_apply_mode,
            "dimension_top_gap_mm": self.safe_dimension_top_gap(),
            "divider_label_gap_mm": self.safe_divider_label_gap(),
            "topmost": bool(self.cfg.topmost),
            "dimension_hotkey_enabled": bool(getattr(self.cfg, "dimension_hotkey_enabled", True)),
        }

    def item_payload(self, item):
        atual = f"X={item.current.left:.1f}, Y={item.current.bottom:.1f} | {item.current.width:.1f}x{item.current.height:.1f}"
        destino = "-" if not item.has_target else f"X={item.target_left:.1f}, Y={item.target_bottom:.1f}"
        mov = "-" if not item.has_target else f"{item.dx:+.1f}, {item.dy:+.1f}"
        return {
            "posicao": item.index,
            "vista": item.name,
            "escala": f"1:{item.current_scale:g}",
            "atual": atual,
            "destino": destino,
            "mov": mov,
            "status": item.status,
        }

    def state_payload(self, message="", ok=True, confirm_request=None):
        return {
            "ok": bool(ok),
            "message": message or "",
            "config": self.config_payload(),
            "items": [self.item_payload(item) for item in self.items],
            "sheet_plan": list(self.sheet_plan_rows),
            "logs": list(self.logs),
            "alert": self.alert,
            "can_undo": bool(self.undo_states or self.divider_undo_state),
            "confirm_request": confirm_request,
        }

    def apply_payload(self, payload):
        payload = payload or {}
        try:
            self.cfg.margin_left = float(str(payload.get("margin_left", self.cfg.margin_left)).replace(",", ".") or 0)
            self.cfg.margin_right = float(str(payload.get("margin_right", self.cfg.margin_right)).replace(",", ".") or 0)
            self.cfg.margin_top = float(str(payload.get("margin_top", self.cfg.margin_top)).replace(",", ".") or 0)
            self.cfg.margin_bottom = float(str(payload.get("margin_bottom", self.cfg.margin_bottom)).replace(",", ".") or 0)
            self.cfg.title_block_width = float(str(payload.get("title_block_width", self.cfg.title_block_width)).replace(",", ".") or 0)
            self.cfg.title_block_height = float(str(payload.get("title_block_height", self.cfg.title_block_height)).replace(",", ".") or 0)
            self.cfg.spacing_x = float(str(payload.get("spacing_x", self.cfg.spacing_x)).replace(",", ".") or 0)
            self.cfg.spacing_y = float(str(payload.get("spacing_y", self.cfg.spacing_y)).replace(",", ".") or 0)
            self.cfg.stack_spacing_y = min(200.0, max(-200.0, float(str(payload.get("stack_spacing_y", getattr(self.cfg, "stack_spacing_y", STACK_VIEW_SPACING_MM))).replace(",", ".") or 0)))
            self.cfg.max_details = int(float(str(payload.get("max_details", self.cfg.max_details)).replace(",", ".") or 0))
            raw_scale = str(payload.get("auto_scale_limit", self.cfg.auto_scale_limit)).strip()
            self.cfg.auto_scale_limit = parse_scale_denominator(raw_scale) if raw_scale else 0.0
            self.cfg.scale_apply_mode = payload.get("scale_apply_mode", self.cfg.scale_apply_mode) or "all"
            self.cfg.dimension_top_gap_mm = float(str(payload.get("dimension_top_gap_mm", self.cfg.dimension_top_gap_mm)).replace(",", ".") or DIMENSION_TOP_GAP_MM)
            if self.cfg.dimension_top_gap_mm <= 0.0:
                self.cfg.dimension_top_gap_mm = DIMENSION_TOP_GAP_MM
            self.api.dimension_top_gap_mm = self.cfg.dimension_top_gap_mm
            raw_divider_gap = str(
                payload.get(
                    "divider_label_gap_mm",
                    getattr(self.cfg, "divider_label_gap_mm", DIVIDER_LABEL_GAP_MM),
                )
            ).strip()
            self.cfg.divider_label_gap_mm = (
                float(raw_divider_gap.replace(",", "."))
                if raw_divider_gap
                else DIVIDER_LABEL_GAP_MM
            )
            if self.cfg.divider_label_gap_mm < 0.0:
                self.cfg.divider_label_gap_mm = DIVIDER_LABEL_GAP_MM
            self.cfg.topmost = bool(payload.get("topmost", self.cfg.topmost))
            self.cfg.dimension_hotkey_enabled = bool(payload.get("dimension_hotkey_enabled", getattr(self.cfg, "dimension_hotkey_enabled", True)))
            self.cfg.save()
        except Exception as ex:
            raise RuntimeError(f"Valor invalido na interface: {ex}")

    def get_state(self):
        return self.state_payload()

    def clear_log(self):
        self.logs = []
        return self.state_payload()

    def set_topmost(self, value):
        self.cfg.topmost = bool(value)
        self.cfg.save()
        try:
            if self._window is not None:
                if hasattr(self._window, "setWindowFlag"):
                    from PySide6.QtCore import Qt
                    was_maximized = bool(self._window.isMaximized())
                    self._window.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, bool(value))
                    if was_maximized:
                        self._window.showMaximized()
                    else:
                        self._window.show()
                    if bool(value):
                        self._window.raise_()
                        self._window.activateWindow()
                else:
                    self._window.on_top = bool(value)
        except Exception:
            pass
        self.log("Manter a frente ativado." if self.cfg.topmost else "Manter a frente desativado.")
        return self.state_payload()

    def save_config(self, payload=None):
        try:
            self.apply_payload(payload)
            self.sync_dimension_hotkey_setting()
            return self.state_payload()
        except Exception as ex:
            self.log("Falha ao salvar parametros: " + str(ex))
            self.log(traceback.format_exc())
            return self.state_payload(message=str(ex), ok=False)

    def window_minimize(self):
        try:
            if self._window is not None:
                if hasattr(self._window, "showMinimized"):
                    self._window.showMinimized()
                else:
                    self._window.minimize()
        except Exception:
            pass
        return True

    def window_maximize(self):
        try:
            if self._window is not None:
                if hasattr(self._window, "isMaximized"):
                    if self._window.isMaximized():
                        self._window.showNormal()
                    else:
                        self._window.showMaximized()
                else:
                    if self._window.state == "maximized":
                        self._window.restore()
                    else:
                        self._window.maximize()
        except Exception:
            pass
        return True

    def window_move_by(self, dx, dy):
        try:
            if self._window is not None:
                if hasattr(self._window, "x") and callable(getattr(self._window, "x", None)):
                    x = int(float(self._window.x()) + float(dx or 0))
                    y = int(float(self._window.y()) + float(dy or 0))
                    self._window.move(x, y)
                else:
                    x = int(float(self._window.x or 0) + float(dx or 0))
                    y = int(float(self._window.y or 0) + float(dy or 0))
                    self._window.move(x, y)
        except Exception:
            pass
        return True

    def native_window_handle(self):
        try:
            if self._window is None or not hasattr(self._window, "native"):
                return None
            native = self._window.native
            handle = getattr(native, "Handle", None)
            if handle is None:
                return None
            try:
                return int(handle.ToInt64())
            except Exception:
                try:
                    return int(handle.ToInt32())
                except Exception:
                    return int(handle)
        except Exception:
            return None

    def start_native_window_action(self, hit_test):
        if not IS_WINDOWS:
            return False
        hwnd = self.native_window_handle()
        if not hwnd:
            return False
        try:
            ctypes.windll.user32.ReleaseCapture()
            ctypes.windll.user32.SendMessageW(hwnd, WM_NCLBUTTONDOWN, int(hit_test), 0)
            return True
        except Exception:
            return False

    def window_start_drag(self):
        try:
            if self._window is not None and hasattr(self._window, "windowHandle"):
                handle = self._window.windowHandle()
                if handle:
                    handle.startSystemMove()
                    return True
        except Exception:
            pass
        return bool(self.start_native_window_action(HTCAPTION))

    def window_start_resize(self, edge):
        try:
            if self._window is not None and hasattr(self._window, "windowHandle"):
                from PySide6.QtCore import Qt
                resize_edges = {
                    "top": Qt.Edge.TopEdge,
                    "bottom": Qt.Edge.BottomEdge,
                    "left": Qt.Edge.LeftEdge,
                    "right": Qt.Edge.RightEdge,
                    "top-left": Qt.Edge.TopEdge | Qt.Edge.LeftEdge,
                    "top-right": Qt.Edge.TopEdge | Qt.Edge.RightEdge,
                    "bottom-left": Qt.Edge.BottomEdge | Qt.Edge.LeftEdge,
                    "bottom-right": Qt.Edge.BottomEdge | Qt.Edge.RightEdge,
                }
                edges = resize_edges.get(str(edge or ""))
                handle = self._window.windowHandle()
                if handle and edges is not None and not self._window.isMaximized():
                    handle.startSystemResize(edges)
                    return True
        except Exception:
            pass
        hit_test = WINDOW_RESIZE_HITTEST.get(str(edge or ""))
        if hit_test is None:
            return False
        return bool(self.start_native_window_action(hit_test))

    def window_resize_to(self, width, height):
        try:
            if self._window is not None:
                current_width = self._window.width() if callable(getattr(self._window, "width", None)) else getattr(self._window, "width", 960)
                current_height = self._window.height() if callable(getattr(self._window, "height", None)) else getattr(self._window, "height", 720)
                width = max(760, int(float(width or current_width or 800)))
                height = max(680, int(float(height or current_height or 700)))
                if hasattr(self._window, "resize"):
                    self._window.resize(width, height)
        except Exception:
            pass
        return True

    def window_close(self):
        try:
            if self._window is not None:
                if hasattr(self._window, "close"):
                    self._window.close()
                else:
                    self._window.destroy()
        except Exception:
            pass
        return True

    def configured_target_sheet(self):
        try:
            width = float(getattr(self.cfg, "target_sheet_width", 0.0) or 0.0)
            height = float(getattr(self.cfg, "target_sheet_height", 0.0) or 0.0)
        except Exception:
            return None
        if width > 1.0 and height > 1.0:
            return width, height
        return None

    def drawing_type_name(self, drawing=None):
        drawing = drawing or self.drawing
        try:
            return str(drawing.GetType().Name)
        except Exception:
            return ""

    def empty_views_message(self):
        if self.drawing_type_name() == "MultiDrawing" and getattr(self.api, "last_sheet_object_count", None) == 0:
            return (
                "O multidrawing ativo foi aberto, mas a folha esta vazia para a API do Tekla "
                "(0 objetos/vistas). Insira ou carregue as vistas das pecas no multidrawing e tente novamente."
            )
        return "Nenhuma vista PECA foi encontrada pela API."

    def drawing_mark_text(self, drawing=None):
        drawing = drawing or self.drawing
        try:
            return str(drawing.Mark or "")
        except Exception:
            return ""

    def set_alert(self, text):
        self.alert = text or ""
        if text:
            self.log("ALERTA: " + text)

    def remember_multidrawing_sheet(self, sheet_w=None, sheet_h=None):
        if self.drawing_type_name() != "MultiDrawing":
            return False
        try:
            if sheet_w is None or sheet_h is None:
                sheet_w = self.api.last_sheet_w or self.api.sheet_size(self.drawing)[0]
                sheet_h = self.api.last_sheet_h or self.api.sheet_size(self.drawing)[1]
            sheet_w = float(sheet_w)
            sheet_h = float(sheet_h)
        except Exception:
            return False
        if sheet_w <= 1.0 or sheet_h <= 1.0:
            return False
        old = self.last_multidrawing_sheet
        self.last_multidrawing_sheet = (sheet_w, sheet_h)
        self.cfg.target_sheet_width = sheet_w
        self.cfg.target_sheet_height = sheet_h
        self.cfg.save()
        if old is None or abs(old[0] - sheet_w) > 0.1 or abs(old[1] - sheet_h) > 0.1:
            self.log(f"Folha de destino lembrada: multidrawing {sheet_w:.0f} x {sheet_h:.0f}.")
        return True

    def matching_source_drawings(self, active_items, limit=80):
        """Localiza os croquis correspondentes às vistas reais do multidrawing."""
        if self.handler is None or not active_items:
            return []
        wanted = {
            str(item.item_text or "").strip()
            for item in active_items
            if str(item.item_text or "").strip()
        }
        if not wanted:
            return []
        matched = []
        for drawing in self.api.collect_all_single_part_drawings(self.handler):
            info = self.api.drawing_piece_info(drawing)
            if not info or str(info[0] or "").strip() not in wanted:
                continue
            matched.append(drawing)
            if len(matched) >= int(limit):
                break
        return self.api.sorted_single_part_drawings(matched)

    def calibrate_capacity_from_active_items(self, active_items):
        """Aprende a diferença entre o croqui e o DrawingLink realmente inserido."""
        drawings = self.matching_source_drawings(active_items)
        if len(drawings) < 3:
            return 1.0, 1.0, 0
        source_items = self.api.build_capacity_items_from_drawings(
            drawings,
            handler=self.handler,
            original_drawing=self.drawing,
        )
        factor_x, factor_y, matched = self.api.estimate_capacity_factors(active_items, source_items)
        if matched >= 3:
            self.remember_capacity_factors(factor_x, factor_y, matched)
            self.log(
                f"Capacidade calibrada com {matched} vista(s) reais: "
                f"X={factor_x:.3f}, Y={factor_y:.3f}."
            )
        return factor_x, factor_y, matched

    def remember_capacity_factors(self, factor_x, factor_y, matched):
        if matched < 3:
            return
        self.last_capacity_factor_x = max(1.0, float(factor_x or 1.0))
        self.last_capacity_factor_y = max(1.0, float(factor_y or 1.0))
        self.last_capacity_factor_matches = int(matched)
        self.cfg.capacity_factor_x = self.last_capacity_factor_x
        self.cfg.capacity_factor_y = self.last_capacity_factor_y
        self.cfg.capacity_factor_matches = self.last_capacity_factor_matches
        self.cfg.save()

    def remember_sheet_plan(self, rows):
        rows = list(rows or [])
        if not rows:
            return False
        self.sheet_plan_rows = rows
        # O plano e apenas uma sugestao visual. Somente vistas realmente
        # encontradas no multidrawing podem avancar a fila persistente.
        return True

    def capacity_cursor_key(self):
        parsed = try_parse_numeric_item(str(getattr(self.cfg, "capacity_cursor_item", "") or ""))
        if not parsed:
            return None
        return numeric_key(parsed[1])

    def sheet_history_item(self, sheet_mark):
        history = getattr(self.cfg, "capacity_sheet_cursors", {}) or {}
        if not isinstance(history, dict):
            return "", 0, ""

        normalized = self.api.normalized_sheet_mark(sheet_mark)
        candidates = []
        for key in (normalized, str(normalized or "").strip("[]"), str(sheet_mark or "").strip()):
            if key and key not in candidates:
                candidates.append(key)

        for key in candidates:
            entry = history.get(key)
            if isinstance(entry, dict):
                if str(entry.get("source") or "").strip().lower() != "actual":
                    continue
                item = str(entry.get("item") or "").strip()
                if not item:
                    continue
                try:
                    count = int(float(entry.get("count", 0) or 0))
                except Exception:
                    count = 0
                return item, count, key
            if entry:
                return str(entry).strip(), 0, key
        return "", 0, ""

    def source_drawings_from_sheet_history(self):
        current_mark = self.drawing_mark_text().strip()
        if not current_mark:
            return [], "", "", ""

        checks = []
        previous_mark = self.api.previous_sheet_mark(current_mark)
        if previous_mark:
            checks.append((previous_mark, current_mark, "registro da folha anterior"))
        checks.append((current_mark, self.api.sheet_mark_with_offset(current_mark, 1), "registro da folha atual"))

        all_drawings = None
        for cursor_sheet, plan_start_mark, label in checks:
            cursor_item, _count, _key = self.sheet_history_item(cursor_sheet)
            if not cursor_item:
                continue
            if all_drawings is None:
                all_drawings = self.api.collect_all_single_part_drawings(self.handler)
            drawings = self.api.single_part_drawings_after(all_drawings, cursor_item)
            if drawings:
                return (
                    drawings,
                    plan_start_mark,
                    f"{label} apos PECA {cursor_item}",
                    cursor_item,
                )
        return [], "", "", ""

    def update_capacity_cursor(self, item_text, sort_key=None, count=0, source=""):
        if not item_text:
            return False
        if sort_key is None:
            parsed = try_parse_numeric_item(str(item_text))
            if not parsed:
                return False
            _, sort_key = parsed
        new_key = numeric_key(sort_key)
        current_key = self.capacity_cursor_key()
        sheet_mark = self.drawing_mark_text()
        history = getattr(self.cfg, "capacity_sheet_cursors", {}) or {}
        if not isinstance(history, dict):
            history = {}
        sheet_key = self.api.normalized_sheet_mark(sheet_mark)
        if sheet_key:
            history[sheet_key] = {
                "item": str(item_text),
                "count": int(count or 0),
                "source": "actual",
            }
            self.cfg.capacity_sheet_cursors = history

        advanced_global = current_key is None or new_key > current_key
        if advanced_global:
            self.cfg.capacity_cursor_item = str(item_text)
            self.cfg.capacity_cursor_sheet = sheet_mark
            self.cfg.capacity_cursor_count = max(
                int(float(getattr(self.cfg, "capacity_cursor_count", 0) or 0)),
                int(count or 0),
            )
        self.cfg.save()
        if advanced_global:
            detail = f"Fila automatica atualizada ate PECA {item_text}"
        else:
            detail = f"Registro da folha {sheet_mark or sheet_key} atualizado ate PECA {item_text}"
        if source:
            detail += f" ({source})"
        self.log(detail + ".")
        return True

    def trusted_piece_items(self, items):
        trusted = []
        for item in items or []:
            if item is None or item.view is None:
                continue
            if not bool(getattr(item, "piece_identity_trusted", True)):
                continue
            if str(getattr(item, "status", "") or "").strip().lower() == "fallback visual":
                continue
            if not try_parse_numeric_item(str(getattr(item, "item_text", "") or "")):
                continue
            trusted.append(item)
        return trusted

    def contiguous_trusted_items(self, items):
        """Retorna somente o prefixo real presente, sem saltar pecas ausentes."""
        trusted = self.trusted_piece_items(items)
        if not trusted or self.handler is None:
            return []

        by_item = {}
        for item in trusted:
            by_item.setdefault(str(item.item_text).strip(), item)

        drawings = self.api.sorted_single_part_drawings(
            self.api.collect_all_single_part_drawings(self.handler)
        )
        previous_mark = self.api.previous_sheet_mark(self.drawing_mark_text())
        previous_item, _count, _key = self.sheet_history_item(previous_mark)
        if previous_item:
            drawings = self.api.single_part_drawings_after(drawings, previous_item)

        contiguous = []
        for drawing in drawings:
            info = self.api.drawing_piece_info(drawing)
            if not info:
                continue
            item = by_item.get(str(info[0]).strip())
            if item is None:
                break
            contiguous.append(item)
        return contiguous

    def remember_progress_from_multidrawing(self, items):
        if self.drawing_type_name() != "MultiDrawing" or not items:
            return
        # Avança a fila apenas até a última peça do trecho contínuo planejado.
        # Assim uma peça sem espaço nunca é pulada porque outra menor coube depois.
        valid = [item for item in self.contiguous_trusted_items(items) if item.has_target]
        if not valid:
            return
        last = valid[-1]
        self.update_capacity_cursor(last.item_text, sort_key=last.sort_key, count=len(valid), source=f"verificado no multidrawing {self.drawing_mark_text()}")

    def auto_queue_drawings_after_cursor(self):
        cursor_item = str(getattr(self.cfg, "capacity_cursor_item", "") or "").strip()
        if not cursor_item:
            return [], (
                "Nao existe fila automatica registrada ainda. Abra a folha anterior que ja recebeu as pecas e clique em Analisar, "
                "ou selecione os croquis no Document Manager uma vez e clique em Capacidade."
            )
        cursor_sheet = str(getattr(self.cfg, "capacity_cursor_sheet", "") or "").strip()
        current_sheet = self.drawing_mark_text().strip()
        if cursor_sheet and current_sheet and cursor_sheet == current_sheet:
            return [], (
                f"A fila automatica esta registrada nesta mesma folha ({current_sheet}). "
                "Para recalcular esta folha, selecione os croquis no Document Manager. Para continuar a fila, abra a proxima folha do multidrawing."
            )
        drawings = self.api.single_part_drawings_after(self.api.collect_all_single_part_drawings(self.handler), cursor_item)
        if not drawings:
            return [], f"Nao encontrei croquis depois da PECA {cursor_item}. A fila automatica parece ter terminado."
        return drawings, ""

    def update_layout_alert(self):
        # Interface aprovada: nao exibir avisos em caixa azul no layout.
        # Qualquer orientacao extra deve ficar apenas no log ou em confirmacoes.
        self.set_alert("")

    def analyze_current(self, warn_empty=True):
        self.items, self.drawing, self.handler = self.api.analyze(self.cfg)
        self.last_capacity_factor_x = max(1.0, float(getattr(self.cfg, "capacity_factor_x", 1.0) or 1.0))
        self.last_capacity_factor_y = max(1.0, float(getattr(self.cfg, "capacity_factor_y", 1.0) or 1.0))
        self.last_capacity_factor_matches = int(float(getattr(self.cfg, "capacity_factor_matches", 0) or 0))
        self.remember_multidrawing_sheet()
        plan_rows = list(getattr(self.api, "last_sheet_plan", []) or [])
        if plan_rows:
            self.remember_sheet_plan(plan_rows)
        self.remember_progress_from_multidrawing(self.items)
        if self.drawing_type_name() == "MultiDrawing":
            self.update_layout_alert()
        else:
            self.set_alert("")
        return bool(self.items)

    def ensure_undo_snapshot(self):
        self.divider_undo_state = None
        if self.undo_states:
            return True
        if not self.items:
            if not self.analyze_current(warn_empty=True):
                return False
        states = self.api.capture_undo_state(self.items)
        if not states:
            raise RuntimeError("Nao foi possivel preparar o desfazer. Nenhuma vista valida foi capturada.")
        self.undo_states = states
        self.log(f"Desfazer preparado: {len(states)} vista(s) guardadas com posicao e escala originais.")
        return True

    def apply_scale_then_reanalyze(self, items, scale):
        if not items:
            raise RuntimeError("Nenhuma vista foi escolhida para alterar escala.")
        if self.drawing is None or self.handler is None:
            self.analyze_current(warn_empty=True)
        if not self.ensure_undo_snapshot():
            return False
        changed = self.api.apply_scale_to_items(items, scale, self.handler, self.drawing)
        self.log(f"Escala aplicada em {changed} vista(s). Reanalisando...")
        time.sleep(0.4)
        self.analyze_current(warn_empty=False)
        return True

    def analyze(self, payload=None):
        try:
            self.apply_payload(payload)
            found = self.analyze_current(warn_empty=False)
            message = "" if found else self.empty_views_message()
            return self.state_payload(message=message, ok=found)
        except Exception as ex:
            self.log("Falha na analise API: " + str(ex))
            self.log(traceback.format_exc())
            return self.state_payload(message=str(ex), ok=False)

    def capacity(self, payload=None):
        try:
            self.log("Capacidade: iniciando leitura da folha e dos croquis.")
            self.apply_payload(payload)
            self.api.set_progress_window(0.0, 35.0)  # analise da folha ativa
            self.analyze_current(warn_empty=False)
            if self.drawing is None or self.handler is None:
                raise RuntimeError("Abra um desenho no Tekla antes de calcular a capacidade.")

            active_type = self.drawing_type_name()
            active_items_for_capacity = []
            if active_type == "MultiDrawing":
                sheet_w = self.api.last_sheet_w or self.api.sheet_size(self.drawing)[0]
                sheet_h = self.api.last_sheet_h or self.api.sheet_size(self.drawing)[1]
                active_items_for_capacity = [item for item in self.items if item.view is not None]
                self.remember_multidrawing_sheet(sheet_w, sheet_h)
            elif self.last_multidrawing_sheet:
                sheet_w, sheet_h = self.last_multidrawing_sheet
                self.log(f"Folha de destino usada: ultimo multidrawing lembrado {sheet_w:.0f} x {sheet_h:.0f}. O croqui ativo nao foi usado como destino.")
            else:
                raise RuntimeError("Abra o multidrawing de destino uma vez e clique em Analisar ou Capacidade.")

            selected_drawings = self.api.collect_selected_single_part_drawings(self.handler)
            manual_selection = len(selected_drawings) > 1
            trusted_active_items = self.trusted_piece_items(active_items_for_capacity)
            auto_queue = False
            remaining_after_active = False
            plan_start_mark = self.drawing_mark_text().strip()
            if manual_selection:
                source_drawings = selected_drawings
                source_label = "selecionados no Document Manager"
            elif trusted_active_items:
                ordered_real = self.contiguous_trusted_items(trusted_active_items)
                if not ordered_real:
                    source_drawings = self.api.sorted_single_part_drawings(
                        self.api.collect_all_single_part_drawings(self.handler)
                    )
                    source_label = "desde a primeira peca do projeto; a folha atual nao forma uma sequencia continua"
                else:
                    last_real = ordered_real[-1]
                    self.update_capacity_cursor(
                        last_real.item_text,
                        sort_key=last_real.sort_key,
                        count=len(ordered_real),
                        source=f"capacidade recalculada na folha {self.drawing_mark_text()}",
                    )
                    source_drawings = self.api.single_part_drawings_after(
                        self.api.collect_all_single_part_drawings(self.handler),
                        last_real.item_text,
                    )
                    if not source_drawings:
                        self.sheet_plan_rows = []
                        return self.state_payload(
                            message=f"Nao encontrei croquis depois da PECA {last_real.item_text}. A fila restante parece ter terminado.",
                            ok=True,
                        )
                    auto_queue = True
                    remaining_after_active = True
                    current_mark = self.drawing_mark_text().strip()
                    plan_start_mark = self.api.sheet_mark_with_offset(current_mark, 1)
                    source_label = f"restantes apos PECA {last_real.item_text}"
            elif selected_drawings:
                start_info = self.api.drawing_piece_info(selected_drawings[0])
                if start_info:
                    source_drawings = self.api.single_part_drawings_from(
                        self.api.collect_all_single_part_drawings(self.handler),
                        start_info[0],
                    )
                    source_label = f"da sequencia a partir da PECA {start_info[0]} selecionada"
                    auto_queue = True
                else:
                    source_drawings = selected_drawings
                    source_label = "selecionados no Document Manager"
            else:
                source_drawings, history_start_mark, history_label, _history_item = self.source_drawings_from_sheet_history()
                if source_drawings:
                    auto_queue = True
                    plan_start_mark = history_start_mark or plan_start_mark
                    source_label = history_label
                else:
                    source_drawings, reason = self.auto_queue_drawings_after_cursor()
                    if not source_drawings:
                        cursor_item = str(getattr(self.cfg, "capacity_cursor_item", "") or "").strip()
                        if cursor_item:
                            return self.state_payload(message=reason, ok=False)
                        source_drawings = self.api.sorted_single_part_drawings(
                            self.api.collect_all_single_part_drawings(self.handler)
                        )
                        source_label = "desde a primeira peca do projeto"
                    else:
                        source_label = f"da fila automatica apos PECA {self.cfg.capacity_cursor_item}"
                    auto_queue = True

            if source_drawings:
                self.log(f"Medindo {len(source_drawings)} croqui(s) {source_label} para calcular capacidade.")
            source_fit = 0
            source_failed = None
            factor_x = self.last_capacity_factor_x if self.last_capacity_factor_matches >= 3 else 1.0
            factor_y = self.last_capacity_factor_y if self.last_capacity_factor_matches >= 3 else 1.0
            matched = 0
            self.api.set_progress_window(35.0, 90.0)  # medicao dos croquis
            if not source_drawings:
                source_items = []
            else:
                source_items = self.api.build_capacity_items_from_drawings(source_drawings, handler=self.handler, original_drawing=self.drawing)

            if not source_items and not active_items_for_capacity:
                return self.state_payload(message="Nao encontrei vistas no multidrawing nem croquis de peca para calcular.", ok=False)

            applied_capacity = 0
            message_lines = []
            if active_items_for_capacity:
                self.calibrate_capacity_from_active_items(active_items_for_capacity)
                current_fit, current_failed, _ = self.api.calculate_fit_count(active_items_for_capacity, sheet_w, sheet_h, self.cfg)
                applied_capacity = current_fit
                text = f"Capacidade real na folha atual: cabem {current_fit} de {len(active_items_for_capacity)} vista(s) ja inseridas."
                if current_failed:
                    text += f" Primeira excedente: {current_failed.name}."
                self.log(text)
                message_lines.append(text)

            if source_items:
                if source_fit <= 0:
                    factor_x, factor_y, matched = self.api.estimate_capacity_factors(active_items_for_capacity, source_items)
                    if matched >= 3:
                        self.remember_capacity_factors(factor_x, factor_y, matched)
                    elif self.last_capacity_factor_matches >= 3:
                        factor_x = self.last_capacity_factor_x
                        factor_y = self.last_capacity_factor_y
                    source_fit, source_failed, _ = self.api.calculate_fit_count(source_items, sheet_w, sheet_h, self.cfg, factor_x=factor_x, factor_y=factor_y)
                self.sheet_plan_rows = self.api.build_capacity_sheet_plan(
                    source_items,
                    self.cfg,
                    sheet_w,
                    sheet_h,
                    plan_start_mark,
                    factor_x=factor_x,
                    factor_y=factor_y,
                )
                if self.sheet_plan_rows:
                    self.remember_sheet_plan(self.sheet_plan_rows)
                    total_plan = self.api.sheet_plan_total(self.sheet_plan_rows)
                    self.log(f"Distribuicao restante por folhas: {total_plan} peca(s) em {len(self.sheet_plan_rows)} folha(s).")
                if selected_drawings or remaining_after_active or not applied_capacity:
                    applied_capacity = source_fit
                text = f"Capacidade estimada dos croquis {source_label}: cabem {source_fit} de {len(source_items)} peca(s)."
                if source_failed:
                    text += f" Primeira excedente: {source_failed.name}."
                if matched >= 3:
                    text += f" Calibrado por {matched} vista(s) ja presentes no multidrawing."
                elif self.last_capacity_factor_matches >= 3:
                    text += f" Usando calibracao anterior de {self.last_capacity_factor_matches} vista(s) do multidrawing."
                else:
                    text += " Sem amostra suficiente no multidrawing; estimativa baseada no tamanho dos croquis."
                self.log(text)
                message_lines.append(text)
                if remaining_after_active:
                    message_lines.append("Fila restante recalculada a partir das pecas reais da folha atual.")
                else:
                    message_lines.append("Fila automatica nao foi avancada aqui; ela sera atualizada quando Analisar encontrar as vistas reais na folha.")

            if applied_capacity > 0:
                self.cfg.max_details = int(applied_capacity)
                self.cfg.save()
                applied = f"Campo 'Max. vistas' atualizado para {applied_capacity}."
                self.log(applied)
                message_lines.append(applied)
            return self.state_payload(message="\n\n".join(message_lines), ok=True)
        except Exception as ex:
            self.log("Falha ao calcular capacidade: " + str(ex))
            self.log(traceback.format_exc())
            return self.state_payload(message=str(ex), ok=False)

    def organize(self, payload=None, replan=True):
        try:
            payload = payload or {}
            self.apply_payload(payload)
            # Organizar executa o plano atual. Analisar e Capacidade sao acoes separadas.
            # So le a folha automaticamente quando ainda nao ha plano carregado.
            if not self.items or self.drawing is None or self.handler is None:
                self.api.set_progress_window(0.0, 55.0)  # analise antes de mover
                self.analyze_current(warn_empty=False)
            if not self.items:
                return self.state_payload(message=self.empty_views_message(), ok=False)
            if self.drawing_type_name() != "MultiDrawing":
                return self.state_payload(message="Abra o multidrawing de destino para organizar. No croqui use apenas Capacidade.", ok=False)
            hard_block = [item for item in self.items if not item.has_target and "Ignorada pelo limite" not in item.status]
            if hard_block and not bool(payload.get("continue_no_space")):
                reason = (
                    f"Não há espaço para alocar {len(hard_block)} peça(s). "
                    "Deseja continuar e mover apenas as peças que couberam?"
                )
                self.set_alert(reason)
                return self.state_payload(
                    ok=False,
                    confirm_request={
                        "title": "Espaço insuficiente",
                        "message": "Não há espaço para alocar as peças. Deseja continuar?",
                        "buttons": [
                            {"label": "Cancelar", "value": False, "className": ""},
                            {"label": "Continuar", "value": True, "className": "yellow"},
                        ],
                        "payload": {"continue_no_space": True},
                    },
                )
            if hard_block and bool(payload.get("continue_no_space")):
                self.log(f"Usuário optou por continuar: {len(hard_block)} peça(s) sem posição serão ignoradas nesta execução.")
            movable = [item for item in self.items if item.has_target]
            if not movable:
                return self.state_payload(message="Nenhuma vista planejada para mover.", ok=False)
            if not self.ensure_undo_snapshot():
                return self.state_payload(message="Nao foi possivel preparar o desfazer.", ok=False)
            self.api.set_progress_window(55.0, 97.0)  # movimentacao das vistas
            moved = self.api.apply_layout(self.items, self.handler, self.drawing)
            self.update_layout_alert()
            self.log(f"Organizacao concluida. Vistas movidas: {moved}.")
            return self.state_payload(message=f"Organizacao concluida. Vistas movidas: {moved}.", ok=True)
        except Exception as ex:
            self.log("Falha ao organizar via API: " + str(ex))
            self.log(traceback.format_exc())
            return self.state_payload(message=str(ex), ok=False)

    def smart_organize(self, payload=None):
        try:
            payload = payload or {}
            self.apply_payload(payload)
            self.api.set_progress_window(0.0, 45.0)  # analise inicial (deixa folga p/ escala+mover)
            if not self.analyze_current(warn_empty=False):
                return self.state_payload(message=self.empty_views_message(), ok=False)
            if self.drawing_type_name() != "MultiDrawing":
                return self.state_payload(message="Abra o multidrawing de destino para organizar. No croqui use apenas Capacidade.", ok=False)
            warning = self.api.layout_warning(self.items, self.cfg)
            manual_scale = False
            if not warning and not manual_scale:
                self.log("Layout nao indicou falta de espaco. Organizando sem alterar escala.")
                return self.organize(payload, replan=False)
            solution = self.api.find_scale_solution(self.items, self.cfg)
            scale = float(solution["scale"])
            if not solution["fits"]:
                reason = (
                    f"Nao vou executar: a escala sugerida 1:{scale:g} ainda deixa {solution['unplanned']} vista(s) sem espaco."
                    if manual_scale else
                    f"Nao vou executar: testei ate a escala 1:30 e ainda sobraram {solution['unplanned']} vista(s) sem espaco."
                )
                self.set_alert(reason)
                if not bool(payload.get("continue_no_space")):
                    return self.state_payload(
                        ok=False,
                        confirm_request={
                            "title": "Espaço insuficiente",
                            "message": "Não há espaço para alocar as peças. Deseja continuar?",
                            "buttons": [
                                {"label": "Cancelar", "value": False, "className": ""},
                                {"label": "Continuar", "value": True, "className": "yellow"},
                            ],
                            "payload": {"continue_no_space": True},
                        },
                    )
                self.log("Usuário optou por continuar mesmo sem solução completa de escala.")
                return self.organize(payload, replan=False)
            current_hard_block = [item for item in self.items if not item.has_target and "Ignorada pelo limite" not in item.status]
            if manual_scale:
                all_targets = [item for item in self.items if abs(float(item.current_scale or 1.0) - scale) > 0.001]
            else:
                all_targets = [item for item in self.items if float(item.current_scale or 1.0) < scale - 0.001]
            if self.cfg.scale_apply_mode == "missing" and current_hard_block:
                targets = [item for item in current_hard_block if abs(float(item.current_scale or 1.0) - scale) > 0.001] if manual_scale else [item for item in current_hard_block if float(item.current_scale or 1.0) < scale - 0.001]
            elif self.cfg.scale_apply_mode == "missing" and not manual_scale:
                targets = []
            else:
                targets = all_targets
            if targets:
                self.apply_scale_then_reanalyze(targets, scale)
            hard_block = [item for item in self.items if not item.has_target and "Ignorada pelo limite" not in item.status]
            if hard_block and not bool(payload.get("continue_no_space")):
                reason = f"Nao vou executar: depois da escala 1:{scale:g}, ainda sobraram {len(hard_block)} vista(s)."
                self.set_alert(reason)
                return self.state_payload(
                    ok=False,
                    confirm_request={
                        "title": "Espaço insuficiente",
                        "message": "Não há espaço para alocar as peças. Deseja continuar?",
                        "buttons": [
                            {"label": "Cancelar", "value": False, "className": ""},
                            {"label": "Continuar", "value": True, "className": "yellow"},
                        ],
                        "payload": {"continue_no_space": True},
                    },
                )
            if hard_block and bool(payload.get("continue_no_space")):
                self.log(f"Usuário optou por continuar: {len(hard_block)} peça(s) sem posição serão ignoradas nesta execução.")
            return self.organize(payload, replan=False)
        except Exception as ex:
            self.log("Falha na auto escala: " + str(ex))
            self.log(traceback.format_exc())
            return self.state_payload(message=str(ex), ok=False)

    def adjust_dimensions(self, payload=None):
        try:
            self.apply_payload(payload)
            result = self.api.adjust_horizontal_dimensions_up(self.cfg.dimension_top_gap_mm)
            errors = int(result.get("errors", 0) or 0)
            total = int(result.get("total", 0) or 0)
            if total <= 0:
                return self.state_payload(message="Nenhuma cota reta foi encontrada no desenho ativo.", ok=False)
            if errors:
                return self.state_payload(message=f"Atenção: {errors} cota(s) não puderam ser alteradas.", ok=False)
            return self.state_payload(ok=True)
        except Exception as ex:
            self.log("Falha ao ajustar cotas: " + str(ex))
            self.log(traceback.format_exc())
            return self.state_payload(message=str(ex), ok=False)

    def dimension_radius_cut(self, payload=None):
        try:
            self.apply_payload(payload)
            self.api.dimension_radius_cut(self.cfg.dimension_top_gap_mm)
            return self.state_payload(ok=True)
        except Exception as ex:
            self.log("Falha ao cotar raio: " + str(ex))
            self.log(traceback.format_exc())
            return self.state_payload(message=str(ex), ok=False)

    def center_lines(self, payload=None):
        try:
            self.apply_payload(payload)
            self.api.create_detected_center_lines()
            return self.state_payload(ok=True)
        except Exception as ex:
            self.log("Falha ao criar linhas de centro: " + str(ex))
            self.log(traceback.format_exc())
            return self.state_payload(message=str(ex), ok=False)

    def divide(self, payload=None):
        try:
            self.apply_payload(payload)
            self.analyze_current(warn_empty=False)
            if not self.items:
                return self.state_payload(message="Nenhuma vista PECA foi encontrada para criar divisorias.", ok=False)
            result = self.api.create_divider_lines(self.items, self.cfg, self.safe_divider_label_gap())
            if int(result.get("created", 0) or 0) <= 0:
                return self.state_payload(message="Nenhuma divisoria foi criada. Verifique se ha mais de uma fileira na folha.", ok=False)
            self.divider_undo_state = result.get("undo_state")
            return self.state_payload(ok=True)
        except Exception as ex:
            self.log("Falha ao criar divisorias: " + str(ex))
            self.log(traceback.format_exc())
            return self.state_payload(message=str(ex), ok=False)

    def undo(self, payload=None):
        try:
            self.apply_payload(payload)
            if self.divider_undo_state:
                self.api.undo_divider_lines(self.divider_undo_state, self.cfg)
                self.divider_undo_state = None
                return self.state_payload(ok=True)
            if not self.undo_states:
                return self.state_payload(message="Nao ha alteracao para desfazer nesta sessao.", ok=False)
            if self.drawing is None or self.handler is None or not self.items:
                self.analyze_current(warn_empty=False)
            restored = self.api.restore_undo_state(self.undo_states, self.items, self.handler, self.drawing)
            self.undo_states = []
            time.sleep(0.4)
            self.analyze_current(warn_empty=False)
            self.log(f"Desfazer concluido. Vistas restauradas: {restored}.")
            return self.state_payload(message=f"Desfazer concluido. Vistas restauradas: {restored}.", ok=True)
        except Exception as ex:
            self.log("Falha ao desfazer: " + str(ex))
            self.log(traceback.format_exc())
            return self.state_payload(message=str(ex), ok=False)


def install_windows_native_frame_patch():
    if not IS_WINDOWS:
        return
    try:
        import webview.platforms.winforms as winforms
        from System import IntPtr
        from System.Windows import Forms as WinForms
    except Exception:
        return
    try:
        base_form = winforms.BrowserView.BrowserForm
        if getattr(base_form, "_organizador_native_frame_patch", False):
            return

        class NativeBrowserForm(base_form):
            _organizador_native_frame_patch = True

            def WndProc(self, message):
                try:
                    if int(message.Msg) == WM_NCCALCSIZE and int(message.WParam.ToInt64()) != 0:
                        if self.WindowState == WinForms.FormWindowState.Maximized:
                            params = _NCCALCSIZE_PARAMS.from_address(int(message.LParam.ToInt64()))
                            rect = params.rgrc[0]
                            metrics = ctypes.windll.user32.GetSystemMetrics
                            padded = metrics(SM_CXPADDEDBORDER)
                            frame_x = metrics(SM_CXFRAME) + padded
                            frame_y = metrics(SM_CYFRAME) + padded
                            rect.left += frame_x
                            rect.top += frame_y
                            rect.right -= frame_x
                            rect.bottom -= frame_y
                        message.Result = IntPtr.Zero
                        return
                except Exception:
                    pass
                super().WndProc(message)

        winforms.BrowserView.BrowserForm = NativeBrowserForm
    except Exception:
        pass


def run_html_app():
    from PySide6.QtCore import QEventLoop, QObject, QTimer, QUrl, Qt, Signal, Slot
    from PySide6.QtGui import QColor
    from PySide6.QtWebChannel import QWebChannel
    from PySide6.QtWebEngineCore import QWebEngineSettings
    from PySide6.QtWebEngineWidgets import QWebEngineView
    from PySide6.QtWidgets import QApplication, QMainWindow

    class QtBridge(QObject):
        progressChanged = Signal(str)

        def __init__(self, controller):
            super().__init__()
            self.controller = controller
            self.controller.progress_emit_cb = self._emit_progress
            # Todas as chamadas Tekla rodam em uma unica thread dedicada, fora da
            # thread do Qt. Assim a janela continua pintando (barra de carregamento
            # animada) enquanto o Tekla processa.
            self._tasks = queue.Queue()
            self._task_running = False
            self._worker = threading.Thread(target=self._task_loop, daemon=True)
            self._worker.start()

        def _task_loop(self):
            prepare_windows_thread_for_tekla()
            while True:
                fn, box, done = self._tasks.get()
                try:
                    box["result"] = fn()
                except Exception as ex:
                    box["error"] = ex
                    box["trace"] = traceback.format_exc()
                finally:
                    done.set()

        def is_busy(self):
            return bool(self._task_running)

        def _emit_progress(self, progress):
            try:
                self.progressChanged.emit(json.dumps(progress or {}, ensure_ascii=False))
            except Exception:
                pass

        def run_task(self, fn):
            """Executa fn na thread Tekla mantendo a interface responsiva."""
            if self._task_running:
                return self.controller.state_payload(message="Aguarde a operacao atual terminar.", ok=False)
            self._task_running = True
            try:
                self.controller.reset_progress()
                box = {}
                done = threading.Event()
                self._tasks.put((fn, box, done))
                loop = QEventLoop()
                timer = QTimer()
                timer.setInterval(50)
                timer.timeout.connect(lambda: loop.quit() if done.is_set() else None)
                timer.start()
                loop.exec()
                timer.stop()
                if "error" in box:
                    self.controller.log("Falha na operacao: " + str(box["error"]))
                    if box.get("trace"):
                        self.controller.log(str(box["trace"]))
                    return self.controller.state_payload(message=str(box["error"]), ok=False)
                return box.get("result")
            finally:
                self._task_running = False

        def _payload(self, text):
            if isinstance(text, dict):
                return text
            try:
                return json.loads(str(text or "{}"))
            except Exception:
                return {}

        def _result(self, value):
            return json.dumps(value if value is not None else True, ensure_ascii=False)

        @Slot(result=str)
        def get_state(self):
            return self._result(self.controller.get_state())

        @Slot(result=str)
        def get_progress(self):
            # Nao passa por run_task: roda direto na thread do Qt e responde na
            # hora, mesmo com uma operacao em andamento na thread Tekla.
            return self._result(self.controller.get_progress())

        @Slot(result=str)
        def clear_log(self):
            return self._result(self.controller.clear_log())

        @Slot(bool, result=str)
        def set_topmost(self, value):
            return self._result(self.controller.set_topmost(bool(value)))

        @Slot(str, result=str)
        def analyze(self, payload):
            data = self._payload(payload)
            return self._result(self.run_task(lambda: self.controller.analyze(data)))

        @Slot(str, result=str)
        def capacity(self, payload):
            data = self._payload(payload)
            return self._result(self.run_task(lambda: self.controller.capacity(data)))

        @Slot(str, result=str)
        def organize(self, payload):
            data = self._payload(payload)
            return self._result(self.run_task(lambda: self.controller.organize(data)))

        @Slot(str, result=str)
        def smart_organize(self, payload):
            data = self._payload(payload)
            return self._result(self.run_task(lambda: self.controller.smart_organize(data)))

        @Slot(str, result=str)
        def adjust_dimensions(self, payload):
            data = self._payload(payload)
            return self._result(self.run_task(lambda: self.controller.adjust_dimensions(data)))

        @Slot(str, result=str)
        def dimension_radius_cut(self, payload):
            data = self._payload(payload)
            return self._result(self.run_task(lambda: self.controller.dimension_radius_cut(data)))

        @Slot(str, result=str)
        def center_lines(self, payload):
            data = self._payload(payload)
            return self._result(self.run_task(lambda: self.controller.center_lines(data)))

        @Slot(str, result=str)
        def divide(self, payload):
            data = self._payload(payload)
            return self._result(self.run_task(lambda: self.controller.divide(data)))

        @Slot(str, result=str)
        def undo(self, payload):
            data = self._payload(payload)
            return self._result(self.run_task(lambda: self.controller.undo(data)))

        @Slot(str, result=str)
        def save_config(self, payload):
            return self._result(self.controller.save_config(self._payload(payload)))

        @Slot(result=bool)
        def minimizeWindow(self):
            return bool(self.controller.window_minimize())

        @Slot(result=bool)
        def maximizeWindow(self):
            return bool(self.controller.window_maximize())

        @Slot(result=bool)
        def closeWindow(self):
            return bool(self.controller.window_close())

        @Slot(result=bool)
        def startWindowDrag(self):
            return bool(self.controller.window_start_drag())

        @Slot(str, result=bool)
        def startWindowResize(self, edge):
            return bool(self.controller.window_start_resize(edge))

    class OrganizerMainWindow(QMainWindow):
        def __init__(self):
            super().__init__()
            self.controller = HtmlController()
            self.controller.set_window(self)
            self._dimension_hotkeys_registered = False
            self._dimension_hotkeys_attempted = False
            self._dimension_hotkey_ids = []
            self._dimension_hotkey_busy = False
            self.setWindowTitle(APP_NAME)
            if not IS_WINDOWS:
                self.setWindowFlag(Qt.WindowType.FramelessWindowHint, True)
            if bool(self.controller.cfg.topmost):
                self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
            self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
            self.setStyleSheet(f"QMainWindow {{ background: {UI['bg']}; }} QWebEngineView {{ background: #ffffff; }}")
            self.resize(800, 760)
            self.setMinimumSize(760, 760)

            self.web = QWebEngineView(self)
            self.web.setZoomFactor(1.0)
            self.web.settings().setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessRemoteUrls, False)
            self.web.settings().setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessFileUrls, True)
            try:
                self.web.page().setBackgroundColor(QColor("#ffffff"))
            except AttributeError:
                pass

            self.bridge = QtBridge(self.controller)
            self.bridge.progressChanged.connect(self._sync_progress_to_html)
            self.channel = QWebChannel(self.web.page())
            self.channel.registerObject("bridge", self.bridge)
            self.web.page().setWebChannel(self.channel)
            self.setCentralWidget(self.web)
            self.web.setUrl(QUrl.fromLocalFile(str(HTML_UI_PATH)))

        def showEvent(self, event):
            super().showEvent(event)
            self._register_dimension_hotkeys()

        def closeEvent(self, event):
            self._unregister_dimension_hotkeys()
            super().closeEvent(event)

        def _register_dimension_hotkeys(self):
            if not IS_WINDOWS or self._dimension_hotkeys_attempted:
                return
            if not bool(getattr(self.controller.cfg, "dimension_hotkey_enabled", True)):
                self._unregister_dimension_hotkeys()
                return
            self._dimension_hotkeys_attempted = True
            try:
                hwnd = int(self.winId())
                user32 = ctypes.windll.user32
                user32.RegisterHotKey.restype = wintypes.BOOL
                user32.RegisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int, wintypes.UINT, wintypes.UINT]
                if user32.RegisterHotKey(hwnd, int(HOTKEY_DIMENSION_N), MOD_NOREPEAT, int(VK_N)):
                    self._dimension_hotkey_ids.append(int(HOTKEY_DIMENSION_N))
                self._dimension_hotkeys_registered = bool(self._dimension_hotkey_ids)
                if self._dimension_hotkeys_registered:
                    self.controller.log("Atalho global N ativado para Ajustar cotas.")
                else:
                    self.controller.log("Atalho global N nao pode ser registrado pelo Windows.")
            except Exception as ex:
                self.controller.log("Falha ao registrar atalho global N: " + str(ex))

        def _unregister_dimension_hotkeys(self):
            if not IS_WINDOWS or not self._dimension_hotkey_ids:
                return
            try:
                hwnd = int(self.winId())
                user32 = ctypes.windll.user32
                user32.UnregisterHotKey.restype = wintypes.BOOL
                user32.UnregisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int]
                for hotkey_id in list(self._dimension_hotkey_ids):
                    user32.UnregisterHotKey(hwnd, int(hotkey_id))
            except Exception:
                pass
            self._dimension_hotkey_ids = []
            self._dimension_hotkeys_registered = False

        def sync_dimension_hotkey_setting(self):
            if not IS_WINDOWS:
                return
            if bool(getattr(self.controller.cfg, "dimension_hotkey_enabled", True)):
                if not self._dimension_hotkeys_registered:
                    self._dimension_hotkeys_attempted = False
                    self._register_dimension_hotkeys()
            else:
                was_active = bool(self._dimension_hotkeys_registered or self._dimension_hotkey_ids or self._dimension_hotkeys_attempted)
                if self._dimension_hotkeys_registered or self._dimension_hotkey_ids:
                    self._unregister_dimension_hotkeys()
                self._dimension_hotkeys_attempted = False
                if was_active:
                    self.controller.log("Atalho global desativado para Ajustar cotas.")

        def _sync_state_to_html(self, state):
            try:
                payload = json.dumps(state or self.controller.get_state(), ensure_ascii=False)
                self.web.page().runJavaScript(f"if (typeof applyState === 'function') applyState({payload});")
            except Exception:
                pass

        def _sync_progress_to_html(self, payload):
            try:
                self.web.page().runJavaScript(f"if (typeof applyProgress === 'function') applyProgress({payload});")
            except Exception:
                pass

        def _trigger_adjust_dimensions_hotkey(self):
            if self._dimension_hotkey_busy or self.bridge.is_busy():
                return
            self._dimension_hotkey_busy = True
            try:
                state = self.bridge.run_task(lambda: self.controller.adjust_dimensions({}))
                message = str((state or {}).get("message", "") or "").strip()
                if message:
                    self.controller.log("Atalho B: " + message)
                    state = self.controller.get_state()
                self._sync_state_to_html(state)
            finally:
                self._dimension_hotkey_busy = False

        def nativeEvent(self, event_type, message):
            if IS_WINDOWS and event_type == b"windows_generic_MSG":
                msg = wintypes.MSG.from_address(int(message))
                if (
                    msg.message == WM_HOTKEY
                    and bool(getattr(self.controller.cfg, "dimension_hotkey_enabled", True))
                    and int(msg.wParam) in self._dimension_hotkey_ids
                ):
                    QTimer.singleShot(0, self._trigger_adjust_dimensions_hotkey)
                    return True, 0
                if msg.message == WM_NCCALCSIZE and msg.wParam:
                    if self.isMaximized():
                        params = _NCCALCSIZE_PARAMS.from_address(msg.lParam)
                        rect = params.rgrc[0]
                        metrics = ctypes.windll.user32.GetSystemMetrics
                        padded = metrics(SM_CXPADDEDBORDER)
                        frame_x = metrics(SM_CXFRAME) + padded
                        frame_y = metrics(SM_CYFRAME) + padded
                        rect.left += frame_x
                        rect.top += frame_y
                        rect.right -= frame_x
                        rect.bottom -= frame_y
                    return True, 0
            return super().nativeEvent(event_type, message)

    app = QApplication.instance() or QApplication(sys.argv)
    window = OrganizerMainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    try:
        run_html_app()
    except Exception as ex:
        print("Falha ao abrir interface HTML. Abrindo interface Tkinter antiga:", ex)
        app = App()
        app.mainloop()
