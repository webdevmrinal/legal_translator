import threading
import time
import os
import sys
import json
import base64
import tkinter as tk
from tkinter import ttk, messagebox
import requests
import keyboard
import pyperclip
import pystray
from PIL import Image, ImageDraw, ImageFont
import queue
import ctypes

APP_VERSION = 1
APP_NAME = "Legal Translator"
CONFIG_DIR = os.path.join(os.environ.get("USERPROFILE", os.path.expanduser("~")), "Documents", "LegalTranslator")
CONFIG_FILE = os.path.join(CONFIG_DIR, "local_config.json")
UPDATER_FLAG = os.path.join(CONFIG_DIR, "update_ready.txt")
REMOTE_CONFIG_URL = "https://gist.githubusercontent.com/webdevmrinal/b71f32a1b102549ea5011433605d1b6b/raw/9a6ac91f5a86ca2256c2f0d48dfd748ab679adb3/config.json"

DEFAULT_CONFIG = {
    "model": "gemini-2.5-flash",
    "api_key": "",
    "version": 1,
    "provider": "gemini",
    "prompt_to_english": (
        "You are an expert legal translator for Indian courts. "
        "Input is ROMAN HINDI (Hindi in English letters like 'mera muawakkil nirdosh hai') or any language. "
        "Translate to formal Legal English for Indian High Courts.\n"
        "RULES:\n"
        "- Use proper Indian legal terminology\n"
        "- Formal and court-appropriate\n"
        "- Do NOT add or remove facts\n"
        "- Output ONLY the translation\n\n"
        "Translate:\n"
    ),
    "prompt_to_hindi": (
        "You are an expert translator. "
        "Translate the following text into Hindi (Devanagari script). "
        "RULES:\n"
        "- Use proper Hindi grammar\n"
        "- If input is legal English, use appropriate Hindi legal terminology\n"
        "- Do NOT add or remove any facts\n"
        "- Output ONLY the Hindi translation in Devanagari script, nothing else\n\n"
        "Translate:\n"
    ),
    "gemini_url": "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}",
    "huggingface_url": "https://api-inference.huggingface.co/models/{model}",
    "huggingface_key": ""
}

# ── Design tokens ──────────────────────────────────────────────────────────────
C = {
    # Backgrounds
    "bg":           "#F2F2F7",   # Apple system bg
    "surface":      "#FFFFFF",
    "fill":         "#F9F9FB",
    "orig_bg":      "#F4F4F8",
    # Accents
    "primary":      "#007AFF",
    "primary_dn":   "#005FCC",
    "primary_tint": "#E8F2FF",
    "success":      "#34C759",
    "success_dn":   "#24A242",
    "indigo":       "#5856D6",
    "indigo_dn":    "#3E3BB5",
    "amber":        "#FF9F0A",
    "amber_dn":     "#D4820A",
    "red":          "#FF3B30",
    "red_dn":       "#C02820",
    # Typography
    "label":        "#1C1C1E",
    "label2":       "#3C3C43",
    "label3":       "#6E6E73",
    "label4":       "#AEAEB2",
    # Structural
    "sep":          "#E5E5EA",
    "sep2":         "#D1D1D6",
    # Toolbar
    "tbar":         "#1C1C1E",
    "tbar2":        "#2C2C2E",
    "tbar_sep":     "#3A3A3C",
    # Transparent key for pill window
    "TRANS":        "#020202",
}

FONT = "Segoe UI"

# ── UI Helpers ─────────────────────────────────────────────────────────────────

def _btn_anim(btn, normal, hover, press=None):
    """Attach lightweight hover + press colour animations to a tk.Button."""
    if press is None:
        press = hover
    btn.configure(bg=normal, activebackground=press, activeforeground="white")
    btn.bind("<Enter>",           lambda e: btn.configure(bg=hover))
    btn.bind("<Leave>",           lambda e: btn.configure(bg=normal))
    btn.bind("<ButtonPress-1>",   lambda e: btn.configure(bg=press))
    btn.bind("<ButtonRelease-1>", lambda e: btn.configure(bg=hover))

def _pill(canvas, x1, y1, x2, y2, fill, outline=None):
    """Draw a filled pill shape on a canvas (radius = half height)."""
    if outline is None:
        outline = fill
    r = (y2 - y1) // 2
    canvas.create_arc(x1, y1, x1 + 2*r, y2, start=90,  extent=180, fill=fill, outline=outline)
    canvas.create_arc(x2 - 2*r, y1, x2, y2, start=270, extent=180, fill=fill, outline=outline)
    canvas.create_rectangle(x1 + r, y1, x2 - r, y2, fill=fill, outline=fill)

def _rounded_rect(canvas, x1, y1, x2, y2, r, fill, outline=None):
    """Draw a rounded rectangle on a canvas."""
    if outline is None:
        outline = fill
    canvas.create_arc(x1,       y1,       x1+2*r, y1+2*r, start=90,  extent=90, fill=fill, outline=outline)
    canvas.create_arc(x2-2*r,   y1,       x2,     y1+2*r, start=0,   extent=90, fill=fill, outline=outline)
    canvas.create_arc(x1,       y2-2*r,   x1+2*r, y2,     start=180, extent=90, fill=fill, outline=outline)
    canvas.create_arc(x2-2*r,   y2-2*r,   x2,     y2,     start=270, extent=90, fill=fill, outline=outline)
    canvas.create_rectangle(x1+r, y1, x2-r, y2, fill=fill, outline=fill)
    canvas.create_rectangle(x1, y1+r, x2, y2-r, fill=fill, outline=fill)

def _separator(parent, bg=None, pady=0):
    """Thin 1px horizontal separator line."""
    color = bg or C["sep"]
    tk.Frame(parent, bg=color, height=1).pack(fill="x", pady=pady)

def _label(parent, text, size=10, weight="normal", color=None, bg=None, **kw):
    """Shorthand for a styled label."""
    return tk.Label(
        parent, text=text,
        font=(FONT, size, weight),
        fg=color or C["label"],
        bg=bg or C["surface"],
        **kw
    )


# ── Config & Providers (unchanged logic) ───────────────────────────────────────

class ConfigManager:
    def __init__(self):
        os.makedirs(CONFIG_DIR, exist_ok=True)
        self.config = dict(DEFAULT_CONFIG)
        self.load_local()
        threading.Thread(target=self.fetch_remote, daemon=True).start()

    def load_local(self):
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    saved = json.load(f)
                self.config.update(saved)
        except:
            pass

    def save_local(self):
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=2)
        except:
            pass

    def fetch_remote(self):
        try:
            r = requests.get(REMOTE_CONFIG_URL, timeout=10)
            if r.status_code == 200:
                remote = r.json()
                local_key = self.config.get("api_key", "")
                encoded_key = remote.get("api_key_encoded", "")
                if encoded_key:
                    try:
                        remote["api_key"] = base64.b64decode(encoded_key).decode("utf-8")
                    except Exception:
                        pass
                for key in remote:
                    self.config[key] = remote[key]
                if local_key:
                    self.config["api_key"] = local_key
                self.save_local()
                remote_ver = remote.get("version", APP_VERSION)
                if remote_ver > APP_VERSION:
                    self.signal_update(remote)
        except:
            pass

    def signal_update(self, remote):
        update_url = remote.get("update_url", "")
        if update_url:
            try:
                with open(UPDATER_FLAG, "w") as f:
                    f.write(update_url)
            except:
                pass

    def get(self, key):
        return self.config.get(key, DEFAULT_CONFIG.get(key, ""))


class GeminiProvider:
    def __init__(self, config):
        self.config = config

    def translate(self, text, prompt_key="prompt_to_english"):
        model = self.config.get("model")
        key = self.config.get("api_key")
        prompt = self.config.get(prompt_key)
        url_template = self.config.get("gemini_url")
        if not key:
            return "ERROR: No API key configured. Please set it from the system tray menu."
        url = url_template.replace("{model}", model).replace("{key}", key)
        full_prompt = prompt + text
        body = {"contents": [{"parts": [{"text": full_prompt}]}]}
        try:
            r = requests.post(url, json=body, timeout=60)
            data = r.json()
            if "error" in data:
                err_msg = data["error"].get("message", "Unknown error")
                status = data["error"].get("code", 0)
                if status == 429:
                    return "ERROR: Daily limit reached. Try tomorrow."
                if status in (401, 403):
                    return "ERROR: API key problem. Please check your key."
                if status == 400:
                    return f"ERROR: Bad request - {err_msg}"
                if status == 404:
                    return "ERROR: Model not found."
                return f"ERROR: {err_msg}"
            candidates = data.get("candidates", [])
            if candidates:
                parts = candidates[0].get("content", {}).get("parts", [])
                if parts:
                    return parts[0].get("text", "").strip()
            return "ERROR: Empty response from Google."
        except requests.exceptions.ConnectionError:
            return "ERROR: No internet connection."
        except requests.exceptions.Timeout:
            return "ERROR: Request timed out. Try again."
        except Exception as e:
            return f"ERROR: {str(e)}"


class HuggingFaceProvider:
    def __init__(self, config):
        self.config = config

    def translate(self, text, prompt_key="prompt_to_english"):
        model = self.config.get("model")
        key = self.config.get("huggingface_key")
        prompt = self.config.get(prompt_key)
        url_template = self.config.get("huggingface_url")
        if not key:
            return "ERROR: No HuggingFace key configured."
        url = url_template.replace("{model}", model)
        headers = {"Authorization": f"Bearer {key}"}
        body = {"inputs": prompt + text}
        try:
            r = requests.post(url, json=body, headers=headers, timeout=60)
            data = r.json()
            if isinstance(data, list) and len(data) > 0:
                return data[0].get("generated_text", "").replace(prompt + text, "").strip()
            if isinstance(data, dict) and "error" in data:
                return f"ERROR: {data['error']}"
            return "ERROR: Unexpected response."
        except Exception as e:
            return f"ERROR: {str(e)}"


def get_provider(config):
    if config.get("provider") == "huggingface":
        return HuggingFaceProvider(config)
    return GeminiProvider(config)


# ── Main App ───────────────────────────────────────────────────────────────────

class TranslatorApp:
    def __init__(self):
        self.config = ConfigManager()
        self.tray_icon = None
        self.is_translating = False
        self.is_toolbar_visible = True
        self.ui_queue = queue.Queue()
        self.root = tk.Tk()
        self.root.withdraw()
        self._setup_styles()
        self.progress_win = None
        self.toolbar_win = None
        self.spinner_arc = None
        self.spinner_angle = 0
        self.spinner_task = None

    def _setup_styles(self):
        """Configure ttk styles once, shared across all windows."""
        style = ttk.Style(self.root)
        style.theme_use("default")
        # Slim, clean scrollbar
        style.configure(
            "Slim.Vertical.TScrollbar",
            gripcount=0,
            background=C["sep2"],
            darkcolor=C["sep2"],
            lightcolor=C["sep2"],
            troughcolor=C["fill"],
            bordercolor=C["fill"],
            arrowcolor=C["label4"],
            relief="flat",
            width=6,
            arrowsize=6,
        )
        style.map(
            "Slim.Vertical.TScrollbar",
            background=[("active", C["label3"]), ("pressed", C["label2"]), ("!active", C["sep2"])],
        )

    def start(self):
        self.check_update_on_start()
        keyboard.add_hotkey("ctrl+shift+e", lambda: self.trigger_translation("prompt_to_english"), suppress=True)
        keyboard.add_hotkey("ctrl+shift+d", lambda: self.trigger_translation("prompt_to_hindi"), suppress=True)
        threading.Thread(target=self.run_tray, daemon=True).start()
        self.create_floating_toolbar()
        self.root.after(100, self.process_queue)
        self.root.mainloop()

    def process_queue(self):
        try:
            while True:
                msg = self.ui_queue.get_nowait()
                cmd = msg[0]
                if cmd == "SHOW_PROGRESS":
                    self._show_progress_ui()
                elif cmd == "HIDE_PROGRESS":
                    self._hide_progress_ui()
                elif cmd == "SHOW_RESULT":
                    self._show_result_ui(msg[1], msg[2], msg[3])
                elif cmd == "SHOW_ERROR":
                    self._show_error_ui(msg[1])
                elif cmd == "SHOW_KEY_DIALOG":
                    self._show_key_dialog_ui()
                elif cmd == "SET_TOOLBAR_STATE":
                    if msg[1]:
                        self.toolbar_win.deiconify()
                    else:
                        self.toolbar_win.withdraw()
                elif cmd == "QUIT":
                    self._quit_app()
                    return
        except queue.Empty:
            pass
        self.root.after(100, self.process_queue)

    def check_update_on_start(self):
        try:
            if os.path.exists(UPDATER_FLAG):
                with open(UPDATER_FLAG, "r") as f:
                    update_url = f.read().strip()
                if update_url:
                    new_path = os.path.join(CONFIG_DIR, "LegalTranslator_new.exe")
                    r = requests.get(update_url, timeout=30)
                    if r.status_code == 200:
                        with open(new_path, "wb") as f:
                            f.write(r.content)
                        bat = os.path.join(CONFIG_DIR, "update.bat")
                        current = sys.executable if getattr(sys, 'frozen', False) else __file__
                        with open(bat, "w") as f:
                            f.write(f'@echo off\ntimeout /t 2 >nul\n')
                            f.write(f'copy /y "{new_path}" "{current}"\n')
                            f.write(f'start "" "{current}"\n')
                            f.write(f'del "{new_path}"\n')
                            f.write(f'del "{UPDATER_FLAG}"\n')
                            f.write(f'del "%~f0"\n')
                os.remove(UPDATER_FLAG)
        except:
            pass

    # ── Tray ──────────────────────────────────────────────────────────────────

    def create_tray_image(self):
        img = Image.new("RGB", (64, 64), (0, 80, 160))
        d = ImageDraw.Draw(img)
        d.rectangle([2, 2, 62, 62], outline="white", width=2)
        try:
            font = ImageFont.truetype("arial.ttf", 22)
            d.text((6, 18), "H>E", fill="white", font=font)
        except:
            d.text((10, 20), "H>E", fill="white")
        return img

    def toggle_toolbar_action(self, icon, item):
        self.is_toolbar_visible = not self.is_toolbar_visible
        self.ui_queue.put(("SET_TOOLBAR_STATE", self.is_toolbar_visible))

    def run_tray(self):
        menu = pystray.Menu(
            pystray.MenuItem("Legal English (Ctrl+Shift+E)", lambda: self.trigger_translation("prompt_to_english")),
            pystray.MenuItem("Hindi Devanagari (Ctrl+Shift+D)", lambda: self.trigger_translation("prompt_to_hindi")),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Show Floating Toolbar", self.toggle_toolbar_action, checked=lambda item: self.is_toolbar_visible),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Provider", pystray.Menu(
                pystray.MenuItem("Gemini", lambda: self.set_provider("gemini"),
                                 checked=lambda item: self.config.get("provider") == "gemini"),
                pystray.MenuItem("HuggingFace", lambda: self.set_provider("huggingface"),
                                 checked=lambda item: self.config.get("provider") == "huggingface"),
            )),
            pystray.MenuItem("Set API Key", lambda: self.ui_queue.put(("SHOW_KEY_DIALOG",))),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Quit", lambda: self.ui_queue.put(("QUIT",))),
        )
        self.tray_icon = pystray.Icon(APP_NAME, self.create_tray_image(), APP_NAME, menu)
        self.tray_icon.run()

    def set_provider(self, name):
        self.config.config["provider"] = name
        self.config.save_local()

    # ── Translation trigger ───────────────────────────────────────────────────

    def trigger_translation(self, direction):
        if self.is_translating:
            return
        threading.Thread(target=self._capture_and_translate, args=(direction,), daemon=True).start()

    def _capture_and_translate(self, direction):
        self.is_translating = True
        try:
            user32 = ctypes.windll.user32
            active_hwnd = user32.GetForegroundWindow()
            pyperclip.copy("")
            time.sleep(0.1)
            keyboard.send("ctrl+c")
            text = ""
            for _ in range(15):
                time.sleep(0.1)
                try:
                    text = pyperclip.paste()
                    if text:
                        break
                except:
                    pass
            if not text or len(text.strip()) < 2:
                self.ui_queue.put(("SHOW_ERROR", "No text selected!\n\n1. Select text in Word\n2. Click a button or use a hotkey (Ctrl+Shift+E / Ctrl+Shift+D)."))
                return
            text = text.strip()
            if len(text) > 15000:
                self.ui_queue.put(("SHOW_ERROR", "Text is very long (over 15,000 characters).\n\nPlease select a smaller portion for better results and to avoid API limits."))
                return
            self.ui_queue.put(("SHOW_PROGRESS",))
            provider = get_provider(self.config)
            result = provider.translate(text, prompt_key=direction)
            self.ui_queue.put(("HIDE_PROGRESS",))
            if result.startswith("ERROR:"):
                self.ui_queue.put(("SHOW_ERROR", result))
                return
            self.ui_queue.put(("SHOW_RESULT", text, result, active_hwnd))
        except Exception as e:
            self.ui_queue.put(("HIDE_PROGRESS",))
            self.ui_queue.put(("SHOW_ERROR", f"Unexpected error: {str(e)}"))
        finally:
            self.is_translating = False

    # ── Floating Toolbar (pill-shaped, transparent corners) ───────────────────

    def create_floating_toolbar(self):
        TRANS = C["TRANS"]
        BG    = C["tbar"]
        W, H  = 252, 46

        self.toolbar_win = tk.Toplevel(self.root)
        self.toolbar_win.overrideredirect(True)
        self.toolbar_win.attributes("-topmost", True)
        self.toolbar_win.attributes("-transparentcolor", TRANS)
        self.toolbar_win.configure(bg=TRANS)

        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.toolbar_win.geometry(f"{W}x{H}+{sw - W - 44}+{sh - H - 90}")

        # Canvas for pill background
        canvas = tk.Canvas(self.toolbar_win, width=W, height=H, bg=TRANS, highlightthickness=0)
        canvas.place(x=0, y=0)
        _rounded_rect(canvas, 0, 0, W, H, 14, BG)

        # Drag support (bind to canvas AND window)
        self._drag_data = {"x": 0, "y": 0}

        def start_drag(e):
            self._drag_data["x"] = e.x_root
            self._drag_data["y"] = e.y_root

        def do_drag(e):
            dx = e.x_root - self._drag_data["x"]
            dy = e.y_root - self._drag_data["y"]
            self.toolbar_win.geometry(
                f"+{self.toolbar_win.winfo_x() + dx}+{self.toolbar_win.winfo_y() + dy}"
            )
            self._drag_data["x"] = e.x_root
            self._drag_data["y"] = e.y_root

        canvas.bind("<ButtonPress-1>",   start_drag)
        canvas.bind("<B1-Motion>",       do_drag)

        # ── Action buttons ───────────────────────────────────────────────────
        BTN_H = 30
        BTN_Y = (H - BTN_H) // 2  # vertically centered = 8

        # Legal English button
        btn1 = tk.Button(
            self.toolbar_win, text="⚖  Legal Eng",
            font=(FONT, 9, "bold"), fg="white", bg=C["primary"],
            relief="flat", bd=0, cursor="hand2",
            command=lambda: self.trigger_translation("prompt_to_english"),
        )
        btn1.place(x=12, y=BTN_Y, width=104, height=BTN_H)
        _btn_anim(btn1, C["primary"], C["primary_dn"])

        # Hindi button
        btn2 = tk.Button(
            self.toolbar_win, text="अ  Hindi",
            font=(FONT, 9, "bold"), fg="white", bg=C["success"],
            relief="flat", bd=0, cursor="hand2",
            command=lambda: self.trigger_translation("prompt_to_hindi"),
        )
        btn2.place(x=122, y=BTN_Y, width=84, height=BTN_H)
        _btn_anim(btn2, C["success"], C["success_dn"])

        # Close ✕
        close_btn = tk.Button(
            self.toolbar_win, text="✕",
            font=(FONT, 9), fg=C["label3"], bg=BG,
            relief="flat", bd=0, cursor="hand2",
            command=lambda: self.ui_queue.put(("SET_TOOLBAR_STATE", False)),
        )
        close_btn.place(x=218, y=BTN_Y, width=24, height=BTN_H)
        _btn_anim(close_btn, BG, C["red"], C["red_dn"])

        # Allow dragging from buttons too
        for b in (btn1, btn2, close_btn):
            b.bind("<ButtonPress-1>",  lambda e, ob=b: (start_drag(e), ob.configure(bg=ob.cget("activebackground"))), add="+")
            b.bind("<B1-Motion>",      lambda e: do_drag(e), add="+")

        def _apply_noactivate():
            try:
                hwnd = int(self.toolbar_win.frame(), 16)
                GWL_EXSTYLE = -20
                ex = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
                ctypes.windll.user32.SetWindowLongW(
                    hwnd, GWL_EXSTYLE,
                    ex | 0x08000000 | 0x00000080 | 0x00000008  # NOACTIVATE | TOOLWINDOW | TOPMOST
                )
            except Exception:
                pass

        self.toolbar_win.after(100, _apply_noactivate)

    # ── Progress Window ───────────────────────────────────────────────────────

    def _show_progress_ui(self):
        if self.progress_win is not None:
            return

        TRANS = C["TRANS"]
        BG    = "#1C1C1E"
        W, H  = 300, 86

        self.progress_win = tk.Toplevel(self.root)
        self.progress_win.overrideredirect(True)
        self.progress_win.attributes("-topmost", True)
        self.progress_win.attributes("-transparentcolor", TRANS)
        self.progress_win.configure(bg=TRANS)

        x = (self.root.winfo_screenwidth()  - W) // 2
        y = (self.root.winfo_screenheight() - H) // 2
        self.progress_win.geometry(f"{W}x{H}+{x}+{y}")

        canvas = tk.Canvas(self.progress_win, width=W, height=H, bg=TRANS, highlightthickness=0)
        canvas.pack()

        # Card background
        _rounded_rect(canvas, 0, 0, W, H, 16, BG)

        # Thin accent border ring
        _rounded_rect(canvas, 1, 1, W-1, H-1, 15, "", C["primary"])

        # Spinner arc
        self.spinner_angle = 0
        self.spinner_arc = canvas.create_arc(
            16, 20, 52, 56,
            start=0, extent=260,
            outline=C["primary"], width=3, style="arc"
        )

        # Text
        canvas.create_text(
            68, H // 2,
            text="Translating…  please wait",
            fill="#FFFFFF",
            font=(FONT, 11, "bold"),
            anchor="w",
        )

        def _apply_noactivate_p():
            try:
                hwnd = int(self.progress_win.frame(), 16)
                GWL_EXSTYLE = -20
                ex = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
                ctypes.windll.user32.SetWindowLongW(
                    hwnd, GWL_EXSTYLE,
                    ex | 0x08000000 | 0x00000080 | 0x00000008
                )
            except Exception:
                pass

        self.progress_win.after(50, _apply_noactivate_p)
        self._animate_spinner(canvas)

    def _animate_spinner(self, canvas):
        try:
            if not self.progress_win or not canvas.winfo_exists():
                return
            self.spinner_angle = (self.spinner_angle + 12) % 360
            canvas.itemconfig(self.spinner_arc, start=self.spinner_angle)
            self.spinner_task = self.root.after(35, lambda: self._animate_spinner(canvas))
        except Exception:
            pass

    def _hide_progress_ui(self):
        if self.spinner_task:
            self.root.after_cancel(self.spinner_task)
            self.spinner_task = None
        if self.progress_win:
            self.progress_win.destroy()
            self.progress_win = None

    # ── Result Window ─────────────────────────────────────────────────────────

    def _show_result_ui(self, original, translated, active_hwnd):
        W, H = 760, 560

        win = tk.Toplevel(self.root)
        win.title(f"{APP_NAME} — Review Translation")
        win.attributes("-topmost", True)
        win.configure(bg=C["bg"])
        win.resizable(True, True)
        win.minsize(600, 460)

        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        win.geometry(f"{W}x{H}+{(sw - W)//2}+{(sh - H)//2}")

        # ── Header ────────────────────────────────────────────────────────────
        header = tk.Frame(win, bg=C["surface"], pady=0)
        header.pack(fill="x")

        hinner = tk.Frame(header, bg=C["surface"], padx=20, pady=14)
        hinner.pack(fill="x")

        tk.Label(
            hinner, text="⚖",
            font=(FONT, 18), fg=C["primary"], bg=C["surface"]
        ).pack(side="left", padx=(0, 10))

        title_col = tk.Frame(hinner, bg=C["surface"])
        title_col.pack(side="left")
        tk.Label(
            title_col, text="Translation Ready",
            font=(FONT, 14, "bold"), fg=C["label"], bg=C["surface"]
        ).pack(anchor="w")
        tk.Label(
            title_col, text="Review, edit if needed, then choose an action below",
            font=(FONT, 9), fg=C["label3"], bg=C["surface"]
        ).pack(anchor="w")

        _separator(header, C["sep"])

        # ── Content ───────────────────────────────────────────────────────────
        content = tk.Frame(win, bg=C["bg"], padx=18, pady=12)
        content.pack(fill="both", expand=True)

        # Original text label row
        orig_row = tk.Frame(content, bg=C["bg"])
        orig_row.pack(fill="x", pady=(0, 4))
        tk.Label(
            orig_row, text="ORIGINAL",
            font=(FONT, 8, "bold"), fg=C["label3"], bg=C["bg"]
        ).pack(side="left")

        # Original text box — fixed height, read-only
        orig_outer = tk.Frame(content, bg=C["sep"], padx=1, pady=1)
        orig_outer.pack(fill="x", pady=(0, 12))
        orig_text = tk.Text(
            orig_outer, height=4,
            font=(FONT, 10), fg=C["label2"], bg=C["orig_bg"],
            wrap="word", relief="flat", bd=0,
            padx=10, pady=8,
            selectbackground=C["primary_tint"],
        )
        orig_text.pack(fill="x")
        orig_text.insert("1.0", original)
        orig_text.config(state="disabled")

        # Translation label row
        trans_row = tk.Frame(content, bg=C["bg"])
        trans_row.pack(fill="x", pady=(0, 4))
        tk.Label(
            trans_row, text="TRANSLATION",
            font=(FONT, 8, "bold"), fg=C["primary"], bg=C["bg"]
        ).pack(side="left")
        tk.Label(
            trans_row, text="  — you can edit this",
            font=(FONT, 8), fg=C["label4"], bg=C["bg"]
        ).pack(side="left")

        # Translation text + external scrollbar in a bordered container
        trans_outer = tk.Frame(content, bg=C["sep"], padx=1, pady=1)
        trans_outer.pack(fill="both", expand=True, pady=(0, 6))

        trans_inner = tk.Frame(trans_outer, bg=C["surface"])
        trans_inner.pack(fill="both", expand=True)

        trans_text = tk.Text(
            trans_inner,
            font=(FONT, 12), fg=C["label"], bg=C["surface"],
            wrap="word", relief="flat", bd=0,
            padx=12, pady=10,
            insertbackground=C["primary"],
            selectbackground=C["primary_tint"],
            selectforeground=C["label"],
        )
        # Scrollbar — external, with breathing room on both sides
        scrollbar = ttk.Scrollbar(
            trans_inner,
            orient="vertical",
            style="Slim.Vertical.TScrollbar",
            command=trans_text.yview,
        )
        scrollbar.pack(side="right", fill="y", padx=(0, 4), pady=6)
        trans_text.pack(side="left", fill="both", expand=True)
        trans_text.config(yscrollcommand=scrollbar.set)
        trans_text.insert("1.0", translated)

        # ── Action bar ────────────────────────────────────────────────────────
        _separator(win, C["sep"])

        bar = tk.Frame(win, bg=C["surface"], padx=16, pady=12)
        bar.pack(fill="x")

        BTN  = {"font": (FONT, 10, "bold"), "relief": "flat", "bd": 0,
                "cursor": "hand2", "padx": 0, "pady": 0}
        BH   = 36
        BW_L = 148   # wide buttons
        BW_S = 90    # narrow
        BW_X = 36    # icon-only close

        def _execute_action(action_type):
            final_text = trans_text.get("1.0", "end-1c").strip()

            if action_type == "copy":
                pyperclip.copy(final_text)
                copy_btn.config(text="✔  Copied!", bg=C["success"])
                win.after(1100, win.destroy)
                return

            win.destroy()
            self.root.update()

            def _thread():
                if action_type == "replace":
                    pyperclip.copy(final_text)
                elif action_type == "insert":
                    pyperclip.copy("\n\n" + final_text + "\n")

                user32   = ctypes.windll.user32
                kernel32 = ctypes.windll.kernel32
                target_tid  = user32.GetWindowThreadProcessId(active_hwnd, None)
                current_tid = kernel32.GetCurrentThreadId()
                try:
                    if target_tid != current_tid and target_tid != 0:
                        user32.AttachThreadInput(current_tid, target_tid, True)
                        user32.SetForegroundWindow(active_hwnd)
                        user32.BringWindowToTop(active_hwnd)
                        user32.AttachThreadInput(current_tid, target_tid, False)
                    else:
                        user32.SetForegroundWindow(active_hwnd)
                except Exception:
                    pass
                time.sleep(0.3)
                if action_type == "insert":
                    keyboard.send("right")
                    time.sleep(0.1)
                keyboard.send("ctrl+v")

            threading.Thread(target=_thread, daemon=True).start()

        # Replace
        rep_btn = tk.Button(
            bar, text="↩  Replace Selection",
            fg="white", bg=C["success"],
            width=0, height=0, **BTN,
            command=lambda: _execute_action("replace"),
        )
        rep_btn.pack(side="left", padx=(0, 8), ipadx=14, ipady=6)
        _btn_anim(rep_btn, C["success"], C["success_dn"])

        # Insert Below
        ins_btn = tk.Button(
            bar, text="⬇  Insert Below",
            fg="white", bg=C["primary"],
            **BTN,
            command=lambda: _execute_action("insert"),
        )
        ins_btn.pack(side="left", padx=(0, 8), ipadx=14, ipady=6)
        _btn_anim(ins_btn, C["primary"], C["primary_dn"])

        # Copy Only
        copy_btn = tk.Button(
            bar, text="⎘  Copy Only",
            fg="white", bg=C["amber"],
            **BTN,
            command=lambda: _execute_action("copy"),
        )
        copy_btn.pack(side="left", ipadx=14, ipady=6)
        _btn_anim(copy_btn, C["amber"], C["amber_dn"])

        # Cancel — right-aligned
        cancel_btn = tk.Button(
            bar, text="Cancel",
            fg=C["label3"], bg=C["surface"],
            **BTN,
            command=win.destroy,
        )
        cancel_btn.pack(side="right", ipadx=10, ipady=6)
        _btn_anim(cancel_btn, C["surface"], C["bg"], C["sep"])

    # ── Error dialog ─────────────────────────────────────────────────────────

    def _show_error_ui(self, msg):
        messagebox.showerror(APP_NAME, msg, parent=self.root)

    # ── API Key Dialog ────────────────────────────────────────────────────────

    def _show_key_dialog_ui(self):
        win = tk.Toplevel(self.root)
        win.title(f"{APP_NAME} — API Key")
        win.attributes("-topmost", True)
        win.configure(bg=C["bg"])
        win.resizable(False, False)

        W, H = 460, 230
        win.geometry(f"{W}x{H}+{(self.root.winfo_screenwidth()-W)//2}+{(self.root.winfo_screenheight()-H)//2}")

        # Header
        hdr = tk.Frame(win, bg=C["surface"], padx=20, pady=14)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🔑  Set API Key",
                 font=(FONT, 13, "bold"), fg=C["label"], bg=C["surface"]).pack(anchor="w")
        tk.Label(hdr, text="Get a free key at  aistudio.google.com",
                 font=(FONT, 9), fg=C["label3"], bg=C["surface"]).pack(anchor="w")
        _separator(hdr, C["sep"])

        # Body
        body = tk.Frame(win, bg=C["bg"], padx=20, pady=16)
        body.pack(fill="both", expand=True)

        tk.Label(body, text="Gemini API Key",
                 font=(FONT, 10, "bold"), fg=C["label2"], bg=C["bg"]).pack(anchor="w", pady=(0, 5))

        entry_outer = tk.Frame(body, bg=C["sep"], padx=1, pady=1)
        entry_outer.pack(fill="x")

        entry = tk.Entry(
            entry_outer,
            font=("Consolas", 11), fg=C["label"], bg=C["surface"],
            relief="flat", bd=0,
            insertbackground=C["primary"],
        )
        entry.pack(fill="x", ipady=8, padx=8)
        entry.insert(0, self.config.get("api_key"))

        def save():
            key = entry.get().strip()
            if key:
                self.config.config["api_key"] = key
                self.config.save_local()
                messagebox.showinfo(APP_NAME, "API Key saved successfully.", parent=win)
                win.destroy()

        # Footer
        _separator(win, C["sep"])
        foot = tk.Frame(win, bg=C["surface"], padx=16, pady=10)
        foot.pack(fill="x")

        save_btn = tk.Button(
            foot, text="Save Key",
            font=(FONT, 10, "bold"), fg="white", bg=C["primary"],
            relief="flat", bd=0, cursor="hand2",
            command=save,
        )
        save_btn.pack(side="left", ipadx=20, ipady=7)
        _btn_anim(save_btn, C["primary"], C["primary_dn"])

        cancel_btn = tk.Button(
            foot, text="Cancel",
            font=(FONT, 10), fg=C["label3"], bg=C["surface"],
            relief="flat", bd=0, cursor="hand2",
            command=win.destroy,
        )
        cancel_btn.pack(side="right", ipadx=12, ipady=7)
        _btn_anim(cancel_btn, C["surface"], C["bg"], C["sep"])

        entry.focus_set()
        win.bind("<Return>", lambda e: save())
        win.bind("<Escape>", lambda e: win.destroy())

    # ── Quit ─────────────────────────────────────────────────────────────────

    def _quit_app(self):
        keyboard.unhook_all()
        if self.tray_icon:
            self.tray_icon.stop()
        self.root.quit()
        os._exit(0)


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    mutex_name = "Global\\LegalTranslator_App_Mutex"
    mutex = ctypes.windll.kernel32.CreateMutexW(None, False, mutex_name)
    last_error = ctypes.windll.kernel32.GetLastError()
    if last_error == 183:
        sys.exit(0)

    try:
        app = TranslatorApp()
        app.start()
    finally:
        if mutex:
            ctypes.windll.kernel32.ReleaseMutex(mutex)
            ctypes.windll.kernel32.CloseHandle(mutex)