import tkinter as tk
from tkinter import CENTER, filedialog
import customtkinter as ctk
import openpyxl
import pyautogui
import time
import threading
import sys
import os
from pynput import keyboard as kb

# ─── Utilities ────────────────────────────────────────────────────────────────

def resource_path(relative_path):
    """Resolve asset paths for both dev and PyInstaller."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ─── Theme ────────────────────────────────────────────────────────────────────

BG         = "#0f1117"
SURFACE    = "#1a1d27"
BORDER     = "#2a2d3e"
TEXT       = "#ffffff"
SUBTEXT    = "#94a3b8"
ACCENT     = "#22c55e"
ACCENT_HOV = "#16a34a"
ERROR      = "#ef4444"
HEADER_BG  = "#16213e"

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

# ─── Global State ─────────────────────────────────────────────────────────────

stop_event = threading.Event()
excel_lock = threading.Lock()
validExcel = []
entryVar   = None   # initialized in ClearPage
pathVar    = None   # initialized in KeyPage
cellVar    = None   # initialized in KeyPage
sheetVar   = None   # initialized in KeyPage


# ─── ClearPage ────────────────────────────────────────────────────────────────

class ClearPage(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.configure(fg_color=BG)

        global entryVar
        entryVar = tk.StringVar()

        # Header
        header = ctk.CTkFrame(self, fg_color=HEADER_BG, height=60, corner_radius=0)
        header.place(relx=0, rely=0, relwidth=1)

        ctk.CTkButton(
            header, text="← Back",
            font=("Inter", 13), width=80, height=36,
            fg_color="transparent", text_color=SUBTEXT, hover_color=SURFACE,
            command=self.backhome
        ).place(relx=0.03, rely=0.5, anchor="w")

        ctk.CTkLabel(
            header, text="C L E A R",
            font=("Inter", 20, "bold"), text_color=TEXT, fg_color="transparent"
        ).place(relx=0.5, rely=0.5, anchor=CENTER)

        # Card
        card = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=14)
        card.place(relx=0.5, rely=0.57, anchor=CENTER, relwidth=0.78, relheight=0.62)

        ctk.CTkLabel(
            card,
            text="A 5-second countdown starts after Confirm.\nSwitch to the PI Count box before time runs out.",
            font=("Inter", 12, "italic"), text_color=SUBTEXT, justify="center"
        ).place(relx=0.5, rely=0.17, anchor=CENTER)

        ctk.CTkLabel(
            card, text="How many boxes do you need to clear?",
            font=("Inter", 15, "bold"), text_color=TEXT
        ).place(relx=0.5, rely=0.36, anchor=CENTER)

        self.clear_entry = ctk.CTkEntry(
            card, placeholder_text="Enter number of boxes...",
            width=300, height=40, border_color=BORDER, fg_color=BG, text_color=TEXT,
            textvariable=entryVar
        )
        self.clear_entry.place(relx=0.5, rely=0.54, anchor=CENTER)

        self.error_label = ctk.CTkLabel(
            card, text="", font=("Inter", 12), text_color=ERROR
        )
        self.error_label.place(relx=0.5, rely=0.67, anchor=CENTER)

        ctk.CTkButton(
            card, text="Confirm",
            font=("Inter", 15, "bold"), fg_color=ACCENT, hover_color=ACCENT_HOV,
            text_color="#000000", width=200, height=40,
            command=self.validate_input
        ).place(relx=0.5, rely=0.82, anchor=CENTER)

    def validate_input(self):
        value = entryVar.get()
        try:
            if int(value) > 0:
                self.error_label.configure(text="")
                self.clear_entry.configure(border_color=BORDER)
                self.controller.show_page("countDown")
            else:
                raise ValueError
        except ValueError:
            self.error_label.configure(text="Please enter a positive whole number.")
            self.clear_entry.configure(border_color=ERROR)

    def backhome(self):
        entryVar.set("")
        self.error_label.configure(text="")
        self.clear_entry.configure(border_color=BORDER)
        self.controller.show_page("Home")


# ─── KeyPage ──────────────────────────────────────────────────────────────────

class KeyPage(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.configure(fg_color=BG)

        global pathVar, cellVar, sheetVar
        pathVar  = tk.StringVar()
        cellVar  = tk.StringVar()
        sheetVar = tk.StringVar()

        # Header
        header = ctk.CTkFrame(self, fg_color=HEADER_BG, height=60, corner_radius=0)
        header.place(relx=0, rely=0, relwidth=1)

        ctk.CTkButton(
            header, text="← Back",
            font=("Inter", 13), width=80, height=36,
            fg_color="transparent", text_color=SUBTEXT, hover_color=SURFACE,
            command=self.backhome
        ).place(relx=0.03, rely=0.5, anchor="w")

        ctk.CTkLabel(
            header, text="K E Y",
            font=("Inter", 20, "bold"), text_color=TEXT, fg_color="transparent"
        ).place(relx=0.5, rely=0.5, anchor=CENTER)

        # Card
        card = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=14)
        card.place(relx=0.5, rely=0.58, anchor=CENTER, relwidth=0.88, relheight=0.72)

        ctk.CTkLabel(
            card,
            text="A 5-second countdown starts after Confirm.\nSwitch to the PI Count box before time runs out.",
            font=("Inter", 12, "italic"), text_color=SUBTEXT, justify="center"
        ).place(relx=0.5, rely=0.09, anchor=CENTER)

        # File path
        ctk.CTkLabel(
            card, text="Excel File Path:",
            font=("Inter", 13, "bold"), text_color=TEXT
        ).place(relx=0.05, rely=0.22, anchor="w")

        self.file_path_entry = ctk.CTkEntry(
            card, placeholder_text="Enter or browse for Excel file path...",
            width=468, height=38, border_color=BORDER, fg_color=BG, text_color=TEXT,
            textvariable=pathVar
        )
        self.file_path_entry.place(relx=0.05, rely=0.34, anchor="w")

        ctk.CTkButton(
            card, text="Browse",
            font=("Inter", 12, "bold"), fg_color=ACCENT, hover_color=ACCENT_HOV,
            text_color="#000000", width=90, height=38,
            command=self.browse_file
        ).place(relx=0.94, rely=0.34, anchor="e")

        # Starting cell
        ctk.CTkLabel(
            card, text="Starting Cell:",
            font=("Inter", 13, "bold"), text_color=TEXT
        ).place(relx=0.05, rely=0.50, anchor="w")

        self.cell_entry = ctk.CTkEntry(
            card, placeholder_text="e.g. C2",
            width=120, height=38, border_color=BORDER, fg_color=BG, text_color=TEXT,
            textvariable=cellVar
        )
        self.cell_entry.place(relx=0.05, rely=0.62, anchor="w")

        # Sheet name
        ctk.CTkLabel(
            card, text="Sheet Name:",
            font=("Inter", 13, "bold"), text_color=TEXT
        ).place(relx=0.53, rely=0.50, anchor="w")

        self.sheet_entry = ctk.CTkEntry(
            card, placeholder_text="e.g. Sheet1",
            width=165, height=38, border_color=BORDER, fg_color=BG, text_color=TEXT,
            textvariable=sheetVar
        )
        self.sheet_entry.place(relx=0.53, rely=0.62, anchor="w")

        self.error_text = ctk.CTkLabel(
            card, text="", font=("Inter", 12), text_color=ERROR
        )
        self.error_text.place(relx=0.5, rely=0.76, anchor=CENTER)

        ctk.CTkButton(
            card, text="Confirm",
            font=("Inter", 15, "bold"), fg_color=ACCENT, hover_color=ACCENT_HOV,
            text_color="#000000", width=200, height=40,
            command=self.validate_input
        ).place(relx=0.5, rely=0.88, anchor=CENTER)

    def backhome(self):
        cellVar.set("")
        pathVar.set("")
        sheetVar.set("")
        self.error_text.configure(text="")
        self.file_path_entry.configure(border_color=BORDER)
        self.cell_entry.configure(border_color=BORDER)
        self.sheet_entry.configure(border_color=BORDER)
        self.controller.show_page("Home")

    def validate_input(self):
        global validExcel
        self.error_text.configure(text="Validating...")
        self.update()

        excelfile = pathVar.get().replace('"', '')
        pathVar.set(excelfile)
        cell  = cellVar.get()
        sheet = sheetVar.get()

        try:
            values = self._read_excel_column(excelfile, sheet, cell)

            if not values:
                self.error_text.configure(text="No data found at that location.")
                return

            bad = [v for v in values if not isinstance(v, int) or v < 0]
            if bad:
                self.error_text.configure(text=f"Invalid data in column: {bad[:3]}...")
                return

            self.error_text.configure(text="")
            with excel_lock:
                validExcel = values

            cellVar.set("")
            pathVar.set("")
            sheetVar.set("")
            self.controller.show_page("countDown2")

        except Exception as e:
            print("Validation Error:", e)
            self.error_text.configure(text=str(e))

    def _read_excel_column(self, file_path, sheet_name, start_cell):
        """Read non-empty values from a column starting at start_cell."""
        wb = openpyxl.load_workbook(file_path, data_only=True)
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found.")
        ws = wb[sheet_name]
        from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
        col_letter, start_row = coordinate_from_string(start_cell)
        start_col = column_index_from_string(col_letter)
        values = []
        for r in range(start_row, ws.max_row + 1):
            v = ws.cell(row=r, column=start_col).value
            if v is None:
                break
            values.append(v)
        return values

    def browse_file(self):
        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if path:
            pathVar.set(path)


# ─── CountDownPage (unified for both modes) ───────────────────────────────────

class CountDownPage(ctk.CTkFrame):
    """Unified countdown + execution page for both 'key' and 'clear' modes."""

    def __init__(self, parent, controller, mode: str, countdown_time: int = 5):
        super().__init__(parent)
        self.controller     = controller
        self.mode           = mode           # "key" | "clear"
        self.countdown_time = countdown_time
        self._cancelled     = False
        self.remaining      = 0

        self.configure(fg_color=BG)

        # Header
        header = ctk.CTkFrame(self, fg_color=HEADER_BG, height=60, corner_radius=0)
        header.place(relx=0, rely=0, relwidth=1)

        ctk.CTkLabel(
            header,
            text="K E Y I N G" if mode == "key" else "C L E A R I N G",
            font=("Inter", 20, "bold"), text_color=TEXT, fg_color="transparent"
        ).place(relx=0.5, rely=0.5, anchor=CENTER)

        # Large count display
        self.count = ctk.CTkLabel(
            self, text="",
            font=("Inter", 72, "bold"), text_color=TEXT
        )
        self.count.place(relx=0.5, rely=0.42, anchor=CENTER)

        # Sub-status label
        self.sub_label = ctk.CTkLabel(
            self, text="",
            font=("Inter", 14), text_color=SUBTEXT
        )
        self.sub_label.place(relx=0.5, rely=0.60, anchor=CENTER)

        # Hotkey hint
        ctk.CTkLabel(
            self,
            text="Ctrl+Shift+Q to stop at any time",
            font=("Inter", 11), text_color=BORDER
        ).place(relx=0.5, rely=0.72, anchor=CENTER)

        # Cancel button
        self.cancel_button = ctk.CTkButton(
            self, text="Cancel",
            font=("Inter", 14, "bold"),
            fg_color="transparent", border_width=2, border_color=ERROR,
            text_color=ERROR, hover_color=SURFACE,
            width=160, height=38,
            command=self.cancel_countdown
        )
        self.cancel_button.place(relx=0.5, rely=0.83, anchor=CENTER)

        # Home button (shown after done/cancel)
        self.home_button = ctk.CTkButton(
            self, text="Return to Home",
            font=("Inter", 14, "bold"),
            fg_color=ACCENT, hover_color=ACCENT_HOV, text_color="#000000",
            width=180, height=40,
            command=self.return_home
        )
        self.home_button.place(relx=0.5, rely=0.83, anchor=CENTER)
        self.home_button.lower()

    # ── Lifecycle ──────────────────────────────────────────────────────────────

    def start_countdown(self):
        """Reset state and begin countdown. Called by App.show_page."""
        stop_event.clear()
        self._cancelled = False
        self.remaining  = self.countdown_time
        self.cancel_button.lift()
        self.home_button.lower()
        self.count.configure(text="", text_color=TEXT)
        self.sub_label.configure(text="Starting in...", text_color=SUBTEXT)
        self._tick()

    def _tick(self):
        if stop_event.is_set():
            self._cancelled = True

        if not self._cancelled and self.remaining >= 0:
            self.count.configure(text=str(self.remaining), text_color=TEXT)
            self.remaining -= 1
            self.after(1000, self._tick)
        elif not self._cancelled:
            self.count.configure(text="GO", text_color=ACCENT)
            self.sub_label.configure(text="Starting...", text_color=TEXT)
            self.after(600, self._start_worker)
        else:
            self._show_cancelled()

    def _start_worker(self):
        if stop_event.is_set() or self._cancelled:
            self._show_cancelled()
            return
        t = threading.Thread(target=self._run_process, daemon=True)
        t.start()

    def _show_cancelled(self):
        self.count.configure(text="✕", text_color=ERROR)
        self.sub_label.configure(text="Cancelled", text_color=ERROR)
        self.home_button.lift()
        self.cancel_button.lower()

    # ── Cancel / Navigation ────────────────────────────────────────────────────

    def cancel_countdown(self):
        self._cancelled = True
        stop_event.set()
        self._clear_inputs()
        self._show_cancelled()

    def return_home(self):
        stop_event.clear()
        self._cancelled = False
        self.controller.show_page("Home")

    def _clear_inputs(self):
        if self.mode == "clear":
            if entryVar:
                entryVar.set("")
        else:
            for var in (cellVar, pathVar, sheetVar):
                if var:
                    var.set("")

    # ── Worker ─────────────────────────────────────────────────────────────────

    def _run_process(self):
        if self.mode == "clear":
            self._run_clear()
        else:
            self._run_key()

    def _run_clear(self):
        try:
            boxes = int(entryVar.get())
        except (ValueError, TypeError):
            self.after(0, lambda: (
                self.count.configure(text="Err", text_color=ERROR),
                self.sub_label.configure(text="Invalid box count", text_color=ERROR),
                self.home_button.lift(),
                self.cancel_button.lower()
            ))
            return

        self.after(0, lambda: entryVar.set(""))

        for i in range(boxes):
            if stop_event.is_set():
                break
            label = f"{i + 1} / {boxes}"
            self.after(0, lambda l=label: (
                self.count.configure(text=l, text_color=TEXT),
                self.sub_label.configure(text="Clearing...", text_color=SUBTEXT)
            ))
            pyautogui.typewrite("-1")
            time.sleep(0.1)
            pyautogui.press("enter")
            time.sleep(0.1)
            pyautogui.press("esc")
            time.sleep(0.1)

        if not stop_event.is_set():
            self.after(500, self._update_done)
        else:
            self.after(500, self._update_cancelled)

    def _run_key(self):
        global validExcel
        with excel_lock:
            values = list(validExcel)
            validExcel = []

        total = len(values)
        for i, value in enumerate(values):
            if stop_event.is_set():
                break
            label = f"{i + 1} / {total}"
            self.after(0, lambda l=label: (
                self.count.configure(text=l, text_color=TEXT),
                self.sub_label.configure(text="Keying...", text_color=SUBTEXT)
            ))
            pyautogui.typewrite(str(value))
            pyautogui.press("enter")
            pyautogui.press("esc")
            time.sleep(0.6)

        if not stop_event.is_set():
            self.after(500, self._update_done)
        else:
            self.after(500, self._update_cancelled)

    # ── UI Updates (called from main thread via after()) ───────────────────────

    def _update_done(self):
        self.count.configure(text="✓", text_color=ACCENT)
        self.sub_label.configure(text="Done!", text_color=ACCENT)
        self.cancel_button.lower()
        self.home_button.lift()

    def _update_cancelled(self):
        self._show_cancelled()


# ─── HomePage ─────────────────────────────────────────────────────────────────

class HomePage(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.configure(fg_color=BG)

        # Header banner
        header = ctk.CTkFrame(self, fg_color=HEADER_BG, height=110, corner_radius=0)
        header.place(relx=0, rely=0, relwidth=1)

        ctk.CTkLabel(
            header, text="PI COUNT AUTOKEY",
            font=("Inter", 28, "bold"), text_color=TEXT, fg_color="transparent"
        ).place(relx=0.5, rely=0.42, anchor=CENTER)

        ctk.CTkLabel(
            header, text="Automated keyboard entry for Physical Inventory",
            font=("Inter", 12), text_color=SUBTEXT, fg_color="transparent"
        ).place(relx=0.5, rely=0.73, anchor=CENTER)

        # Prompt
        ctk.CTkLabel(
            self, text="What would you like to do?",
            font=("Inter", 15), text_color=SUBTEXT
        ).place(relx=0.5, rely=0.46, anchor=CENTER)

        # Card buttons
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.place(relx=0.5, rely=0.65, anchor=CENTER)

        ctk.CTkButton(
            btn_frame, text="⌨  KEY",
            font=("Inter", 18, "bold"),
            fg_color=SURFACE, border_color=ACCENT, border_width=2,
            hover_color="#1a2b1a", text_color=ACCENT,
            width=165, height=85, corner_radius=14,
            command=lambda: controller.show_page("Key")
        ).pack(side="left", padx=14)

        ctk.CTkButton(
            btn_frame, text="✕  CLEAR",
            font=("Inter", 18, "bold"),
            fg_color=SURFACE, border_color=BORDER, border_width=2,
            hover_color="#1e2130", text_color=TEXT,
            width=165, height=85, corner_radius=14,
            command=lambda: controller.show_page("Clear")
        ).pack(side="left", padx=14)

        # Footer
        ctk.CTkLabel(
            self, text="v2.0  ·  PI Count AutoKey  ·  Ctrl+Shift+Q to stop",
            font=("Inter", 10), text_color=SUBTEXT
        ).place(relx=0.5, rely=0.92, anchor=CENTER)


# ─── App ──────────────────────────────────────────────────────────────────────

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("PI Count AutoKey")
        self.geometry("720x480")
        self.configure(fg_color=BG)
        self.resizable(False, False)

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        container = ctk.CTkFrame(self, fg_color=BG)
        container.pack(fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.pages = {}
        for PageClass, name, kwargs in [
            (HomePage,      "Home",       {}),
            (KeyPage,       "Key",        {}),
            (ClearPage,     "Clear",      {}),
            (CountDownPage, "countDown",  {"mode": "clear"}),
            (CountDownPage, "countDown2", {"mode": "key"}),
        ]:
            page = PageClass(container, self, **kwargs)
            self.pages[name] = page
            page.grid(row=0, column=0, sticky="nsew")
            page.lower()

        # Global emergency stop hotkey
        self._hotkey = kb.GlobalHotKeys({"<ctrl>+<shift>+q": self._emergency_stop})
        self._hotkey.start()

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.show_page("Home")

    def _emergency_stop(self):
        stop_event.set()

    def _on_close(self):
        self._hotkey.stop()
        self.destroy()

    def show_page(self, page_name: str):
        for page in self.pages.values():
            page.lower()
        self.pages[page_name].tkraise()
        if page_name in ("countDown", "countDown2"):
            self.pages[page_name].start_countdown()


if __name__ == "__main__":
    app = App()
    app.mainloop()
