import threading
import time
import os
import sys
import json
import base64
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
import requests
import keyboard
import pyperclip
import pystray
from PIL import Image, ImageDraw, ImageFont
import queue
import ctypes

APP_VERSION = 2
APP_NAME = "Legal Translator"
CONFIG_DIR = os.path.join(os.environ.get("USERPROFILE", os.path.expanduser("~")), "Documents", "LegalTranslator")
CONFIG_FILE = os.path.join(CONFIG_DIR, "local_config.json")
UPDATER_FLAG = os.path.join(CONFIG_DIR, "update_ready.txt")


def _load_remote_url():
    """Load config URL from multiple sources with fallbacks."""
    # 1. Try build_config.py (bundled or local)
    try:
        from build_config import REMOTE_CONFIG_URL as url
        if url:
            return url
    except ImportError:
        pass

    # 2. Try config_url.txt file (bundled or local)
    for base in [
        getattr(sys, '_MEIPASS', None),
        os.path.dirname(os.path.abspath(__file__)),
        CONFIG_DIR,
    ]:
        if base is None:
            continue
        txt_path = os.path.join(base, "config_url.txt")
        try:
            if os.path.exists(txt_path):
                with open(txt_path, "r", encoding="utf-8") as f:
                    url = f.read().strip()
                if url:
                    return url
        except:
            pass

    # 3. Try environment variable
    url = os.environ.get("LEGAL_TRANSLATOR_CONFIG_URL", "")
    if url:
        return url

    # 4. Try previously saved URL in local config
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
            url = saved.get("_remote_url", "")
            if url:
                return url
    except:
        pass

    return ""


REMOTE_CONFIG_URL = _load_remote_url()

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


def get_icon_path(filename):
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)


class ConfigManager:
    def __init__(self):
        os.makedirs(CONFIG_DIR, exist_ok=True)
        self.config = dict(DEFAULT_CONFIG)
        self.load_local()
        if REMOTE_CONFIG_URL:
            self.config["_remote_url"] = REMOTE_CONFIG_URL
            self.save_local()
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
        url = REMOTE_CONFIG_URL or self.config.get("_remote_url", "")
        if not url:
            return False
        try:
            r = requests.get(url, timeout=10)
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
                if url:
                    self.config["_remote_url"] = url
                self.save_local()
                remote_ver = remote.get("version", APP_VERSION)
                if remote_ver > APP_VERSION:
                    self.signal_update(remote)
                return True
            return False
        except:
            return False

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

    def get_display_config(self):
        display = dict(self.config)
        for k in ("api_key", "huggingface_key"):
            val = display.get(k, "")
            if val and len(val) > 8:
                display[k] = val[:4] + "*" * (len(val) - 8) + val[-4:]
            elif val:
                display[k] = "****"
        url = display.get("_remote_url", "")
        if url and len(url) > 40:
            display["_remote_url"] = url[:35] + "..."
        display["_app_version"] = APP_VERSION
        display["_config_url_source"] = (
            "build_config.py" if REMOTE_CONFIG_URL else
            "saved in local config" if self.config.get("_remote_url") else
            "(not set)"
        )
        return display


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


class TranslatorApp:
    def __init__(self):
        self.config = ConfigManager()
        self.tray_icon = None
        self.is_translating = False
        self.is_toolbar_visible = True
        self.ui_queue = queue.Queue()
        self.root = tk.Tk()
        self.root.withdraw()
        self.progress_win = None
        self.toolbar_win = None
        self.dropdown_win = None
        self.spinner_arc = None
        self.spinner_angle = 0
        self.spinner_task = None
        self._last_direction = "prompt_to_english"
        self._upd_win = None
        self._upd_widgets = {}
        self._cfg_win = None
        self._cfg_text = None
        self._cfg_refresh_btn = None

    def _get_effective_url(self):
        return REMOTE_CONFIG_URL or self.config.config.get("_remote_url", "")

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
                    self._show_result_ui(msg[1], msg[2], msg[3], msg[4])
                elif cmd == "SHOW_ERROR":
                    self._show_error_ui(msg[1])
                elif cmd == "SHOW_KEY_DIALOG":
                    self._show_key_dialog_ui()
                elif cmd == "SHOW_CONFIG_VIEWER":
                    self._show_config_viewer_ui()
                elif cmd == "CONFIG_VIEWER_REFRESH":
                    self._handle_config_viewer_refresh(msg[1])
                elif cmd == "SET_TOOLBAR_STATE":
                    if msg[1]:
                        self.toolbar_win.deiconify()
                        self.is_toolbar_visible = True
                    else:
                        self.toolbar_win.withdraw()
                        self.is_toolbar_visible = False
                elif cmd == "FETCH_CONFIG_RESULT":
                    self._show_fetch_config_result(msg[1])
                elif cmd == "SHOW_UPDATE_DIALOG":
                    self._show_update_dialog_ui()
                elif cmd == "UPD":
                    self._handle_update_msg(msg[1], msg[2:])
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

    def create_tray_image(self):
        try:
            icon_path = get_icon_path("icon.png")
            if os.path.exists(icon_path):
                img = Image.open(icon_path)
                img = img.resize((64, 64), Image.LANCZOS)
                return img
        except:
            pass
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
            pystray.MenuItem("Legal English (Ctrl+Shift+E)",
                             lambda: self.trigger_translation("prompt_to_english")),
            pystray.MenuItem("Hindi Devanagari (Ctrl+Shift+D)",
                             lambda: self.trigger_translation("prompt_to_hindi")),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Show Floating Toolbar", self.toggle_toolbar_action,
                             checked=lambda item: self.is_toolbar_visible),
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

    def trigger_translation(self, direction):
        if self.is_translating:
            return
        self._last_direction = direction
        threading.Thread(target=self._capture_and_translate, args=(direction,), daemon=True).start()

    def _retranslate(self, original_text, direction, active_hwnd):
        if self.is_translating:
            return
        self._last_direction = direction
        threading.Thread(target=self._do_retranslate, args=(original_text, direction, active_hwnd),
                         daemon=True).start()

    def _do_retranslate(self, text, direction, active_hwnd):
        self.is_translating = True
        try:
            self.ui_queue.put(("SHOW_PROGRESS",))
            provider = get_provider(self.config)
            result = provider.translate(text, prompt_key=direction)
            self.ui_queue.put(("HIDE_PROGRESS",))
            if result.startswith("ERROR:"):
                self.ui_queue.put(("SHOW_ERROR", result))
                return
            self.ui_queue.put(("SHOW_RESULT", text, result, active_hwnd, direction))
        except Exception as e:
            self.ui_queue.put(("HIDE_PROGRESS",))
            self.ui_queue.put(("SHOW_ERROR", f"Unexpected error: {str(e)}"))
        finally:
            self.is_translating = False

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
                self.ui_queue.put(("SHOW_ERROR",
                                   "No text selected!\n\n"
                                   "1. Select text in Word\n"
                                   "2. Click a button or use a hotkey\n"
                                   "   (Ctrl+Shift+E / Ctrl+Shift+D)."))
                return
            text = text.strip()
            if len(text) > 15000:
                self.ui_queue.put(("SHOW_ERROR",
                                   "Text is very long (over 15,000 characters).\n\n"
                                   "Please select a smaller portion for better results."))
                return
            self.ui_queue.put(("SHOW_PROGRESS",))
            provider = get_provider(self.config)
            result = provider.translate(text, prompt_key=direction)
            self.ui_queue.put(("HIDE_PROGRESS",))
            if result.startswith("ERROR:"):
                self.ui_queue.put(("SHOW_ERROR", result))
                return
            self.ui_queue.put(("SHOW_RESULT", text, result, active_hwnd, direction))
        except Exception as e:
            self.ui_queue.put(("HIDE_PROGRESS",))
            self.ui_queue.put(("SHOW_ERROR", f"Unexpected error: {str(e)}"))
        finally:
            self.is_translating = False

    def _apply_noactivate(self, win):
        try:
            hwnd = int(win.frame(), 16)
            GWL_EXSTYLE = -20
            WS_EX_NOACTIVATE = 0x08000000
            WS_EX_TOOLWINDOW = 0x00000080
            WS_EX_TOPMOST = 0x00000008
            ex_style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            ctypes.windll.user32.SetWindowLongW(
                hwnd, GWL_EXSTYLE,
                ex_style | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW | WS_EX_TOPMOST
            )
        except Exception:
            pass

    # ── Dropdown menu ──

    def _dismiss_dropdown(self, event=None):
        if self.dropdown_win and self.dropdown_win.winfo_exists():
            self.dropdown_win.destroy()
            self.dropdown_win = None

    def _toggle_dropdown(self):
        if self.dropdown_win and self.dropdown_win.winfo_exists():
            self._dismiss_dropdown()
            return

        self.dropdown_win = tk.Toplevel(self.root)
        self.dropdown_win.overrideredirect(True)
        self.dropdown_win.attributes("-topmost", True)
        self.dropdown_win.configure(bg="#2C2C2E")

        tx = self.toolbar_win.winfo_x()
        ty = self.toolbar_win.winfo_y() + self.toolbar_win.winfo_height() + 2
        dw, dh = 155, 94
        self.dropdown_win.geometry(f"{dw}x{dh}+{tx}+{ty}")

        def make_dd_btn(parent, text, command):
            btn = tk.Button(parent, text=text, bg="#2C2C2E", fg="#D0D0D0",
                            font=("Segoe UI", 8), relief="flat",
                            activebackground="#3A3A3C", activeforeground="white",
                            cursor="hand2", command=command, anchor="w",
                            bd=0, padx=10, pady=4)
            btn.pack(fill="x")
            btn.bind("<Enter>", lambda e: btn.config(bg="#3A3A3C", fg="white"))
            btn.bind("<Leave>", lambda e: btn.config(bg="#2C2C2E", fg="#D0D0D0"))
            return btn

        def do_fetch():
            self._dismiss_dropdown()
            url = self._get_effective_url()
            if not url:
                messagebox.showwarning(APP_NAME,
                                       "No remote config URL found.\n\n"
                                       "Create one of these:\n"
                                       "1. build_config.py with REMOTE_CONFIG_URL\n"
                                       "2. config_url.txt with the URL\n"
                                       "3. Set LEGAL_TRANSLATOR_CONFIG_URL env var",
                                       parent=self.root)
                return
            threading.Thread(target=self._fetch_config_threaded, daemon=True).start()

        def do_show_cfg():
            self._dismiss_dropdown()
            self.ui_queue.put(("SHOW_CONFIG_VIEWER",))

        def do_update():
            self._dismiss_dropdown()
            self.ui_queue.put(("SHOW_UPDATE_DIALOG",))

        make_dd_btn(self.dropdown_win, "  Refresh Config", do_fetch)
        tk.Frame(self.dropdown_win, bg="#444", height=1).pack(fill="x", padx=6)
        make_dd_btn(self.dropdown_win, "  Show Config", do_show_cfg)
        tk.Frame(self.dropdown_win, bg="#444", height=1).pack(fill="x", padx=6)
        make_dd_btn(self.dropdown_win, "  Update App", do_update)

        self.dropdown_win.after(100, lambda: self._apply_noactivate(self.dropdown_win))
        self.dropdown_win.after(8000, self._dismiss_dropdown)

    # ── Fetch config ──

    def _fetch_config_threaded(self):
        success = self.config.fetch_remote()
        self.ui_queue.put(("FETCH_CONFIG_RESULT", success))

    def _show_fetch_config_result(self, success):
        if success:
            messagebox.showinfo(APP_NAME, "Configuration refreshed successfully.", parent=self.root)
        else:
            messagebox.showwarning(APP_NAME,
                                   "Could not fetch remote config.\nCheck your internet connection.",
                                   parent=self.root)

    # ── Config viewer ──

    def _show_config_viewer_ui(self):
        if self._cfg_win and self._cfg_win.winfo_exists():
            self._cfg_win.lift()
            return

        win = tk.Toplevel(self.root)
        win.title(f"{APP_NAME} \u2014 Active Configuration")
        win.attributes("-topmost", True)
        win.configure(bg="#1E1E2E")

        w, h = 580, 460
        x = (self.root.winfo_screenwidth() - w) // 2
        y = (self.root.winfo_screenheight() - h) // 2
        win.geometry(f"{w}x{h}+{x}+{y}")
        win.minsize(450, 350)

        try:
            icon_path = get_icon_path("icon.png")
            if os.path.exists(icon_path):
                icon_img = tk.PhotoImage(file=icon_path)
                win.iconphoto(False, icon_img)
                win._icon_ref = icon_img
        except:
            pass

        self._cfg_win = win

        btn_frame = tk.Frame(win, bg="#1E1E2E", pady=10)
        btn_frame.pack(side="bottom", fill="x")

        def do_refresh():
            url = self._get_effective_url()
            if not url:
                messagebox.showwarning(APP_NAME, "No remote config URL configured.", parent=win)
                return
            self._cfg_refresh_btn.config(text="Refreshing...", state="disabled")
            threading.Thread(target=self._cfg_refresh_thread, daemon=True).start()

        refresh_btn = tk.Button(btn_frame, text="Refresh Config", bg="#2196f3", fg="white",
                                font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2",
                                padx=14, pady=4, command=do_refresh)
        refresh_btn.pack(side="left", padx=(20, 6))
        refresh_btn.bind("<Enter>",
                         lambda e: refresh_btn.config(
                             bg="#1E88E5") if refresh_btn["state"] != "disabled" else None)
        refresh_btn.bind("<Leave>",
                         lambda e: refresh_btn.config(
                             bg="#2196f3") if refresh_btn["state"] != "disabled" else None)
        self._cfg_refresh_btn = refresh_btn

        close_btn = tk.Button(btn_frame, text="Close", bg="#555", fg="white",
                              font=("Segoe UI", 9), relief="flat", cursor="hand2",
                              padx=14, pady=4, command=win.destroy)
        close_btn.pack(side="right", padx=(6, 20))
        close_btn.bind("<Enter>", lambda e: close_btn.config(bg="#777"))
        close_btn.bind("<Leave>", lambda e: close_btn.config(bg="#555"))

        title_frame = tk.Frame(win, bg="#283593", pady=10)
        title_frame.pack(side="top", fill="x")
        tk.Label(title_frame, text="Active Configuration",
                 font=("Segoe UI", 13, "bold"), fg="white", bg="#283593").pack()
        tk.Label(title_frame, text="API keys are partially masked for security",
                 font=("Segoe UI", 8), fg="#B0BEC5", bg="#283593").pack()

        text_frame = tk.Frame(win, bg="#1E1E2E")
        text_frame.pack(side="top", fill="both", expand=True, padx=16, pady=12)

        scroll = tk.Scrollbar(text_frame, orient="vertical")
        scroll.pack(side="right", fill="y")

        text_widget = tk.Text(text_frame, font=("Consolas", 10), bg="#282A36", fg="#F8F8F2",
                              wrap="word", relief="flat", bd=0, padx=14, pady=12,
                              yscrollcommand=scroll.set, insertbackground="#F8F8F2",
                              selectbackground="#44475A", selectforeground="#F8F8F2")
        text_widget.pack(fill="both", expand=True)
        scroll.config(command=text_widget.yview)
        self._cfg_text = text_widget

        display_cfg = self.config.get_display_config()
        pretty = json.dumps(display_cfg, indent=2, ensure_ascii=False)
        text_widget.insert("1.0", pretty)
        text_widget.config(state="disabled")

        def on_close():
            self._cfg_win = None
            self._cfg_text = None
            self._cfg_refresh_btn = None
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", on_close)

    def _cfg_refresh_thread(self):
        success = self.config.fetch_remote()
        self.ui_queue.put(("CONFIG_VIEWER_REFRESH", success))

    def _handle_config_viewer_refresh(self, success):
        if not self._cfg_win or not self._cfg_win.winfo_exists():
            return
        try:
            self._cfg_text.config(state="normal")
            self._cfg_text.delete("1.0", "end")
            new_cfg = self.config.get_display_config()
            self._cfg_text.insert("1.0", json.dumps(new_cfg, indent=2, ensure_ascii=False))
            self._cfg_text.config(state="disabled")
            if success:
                self._cfg_refresh_btn.config(text="\u2713 Refreshed!", bg="#43A047", state="normal")
                self._cfg_win.after(1500, lambda: (
                    self._cfg_refresh_btn.config(text="Refresh Config", bg="#2196f3")
                    if self._cfg_win and self._cfg_win.winfo_exists() else None
                ))
            else:
                self._cfg_refresh_btn.config(text="Failed", bg="#E53935", state="normal")
                self._cfg_win.after(1500, lambda: (
                    self._cfg_refresh_btn.config(text="Refresh Config", bg="#2196f3")
                    if self._cfg_win and self._cfg_win.winfo_exists() else None
                ))
        except:
            pass

    # ── Update app ──

    def _show_update_dialog_ui(self):
        if self._upd_win and self._upd_win.winfo_exists():
            self._upd_win.lift()
            return

        win = tk.Toplevel(self.root)
        win.title(f"{APP_NAME} \u2014 Update")
        win.attributes("-topmost", True)
        win.configure(bg="#FAFAFA")
        win.resizable(False, False)

        w, h = 470, 310
        x = (self.root.winfo_screenwidth() - w) // 2
        y = (self.root.winfo_screenheight() - h) // 2
        win.geometry(f"{w}x{h}+{x}+{y}")

        try:
            icon_path = get_icon_path("icon.png")
            if os.path.exists(icon_path):
                icon_img = tk.PhotoImage(file=icon_path)
                win.iconphoto(False, icon_img)
                win._icon_ref = icon_img
        except:
            pass

        self._upd_win = win

        btn_frame = tk.Frame(win, bg="#EEEEEE", pady=10, padx=20)
        btn_frame.pack(side="bottom", fill="x")

        close_btn = tk.Button(btn_frame, text="Close", bg="#777", fg="white",
                              font=("Segoe UI", 9), relief="flat", cursor="hand2",
                              padx=14, pady=4, command=win.destroy)
        close_btn.pack(side="right")
        close_btn.bind("<Enter>", lambda e: close_btn.config(bg="#999"))
        close_btn.bind("<Leave>", lambda e: close_btn.config(bg="#777"))

        action_btn = tk.Button(btn_frame, text="Checking...", bg="#BDBDBD", fg="#888",
                               font=("Segoe UI", 10, "bold"), relief="flat",
                               state="disabled", padx=16, pady=4, bd=0)
        action_btn.pack(side="left")

        title_frame = tk.Frame(win, bg="#283593", pady=10)
        title_frame.pack(side="top", fill="x")
        tk.Label(title_frame, text="App Update",
                 font=("Segoe UI", 13, "bold"), fg="white", bg="#283593").pack()

        content = tk.Frame(win, bg="#FAFAFA", padx=30, pady=15)
        content.pack(side="top", fill="both", expand=True)

        status_label = tk.Label(content, text="Checking for updates...",
                                font=("Segoe UI", 12), fg="#333", bg="#FAFAFA")
        status_label.pack(pady=(20, 5))

        detail_label = tk.Label(content, text=f"Current version: v{APP_VERSION}",
                                font=("Segoe UI", 9), fg="#888", bg="#FAFAFA")
        detail_label.pack(pady=(0, 15))

        progress_bar = ttk.Progressbar(content, length=390, mode='determinate')
        progress_label = tk.Label(content, text="",
                                  font=("Segoe UI", 8), fg="#888", bg="#FAFAFA")

        self._upd_widgets = {
            "status": status_label,
            "detail": detail_label,
            "progress": progress_bar,
            "progress_label": progress_label,
            "action": action_btn,
            "close": close_btn,
        }

        def on_close():
            self._upd_win = None
            self._upd_widgets = {}
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", on_close)

        url = self._get_effective_url()
        if not url:
            self.ui_queue.put(("UPD", "error", "No remote config URL configured."))
        else:
            threading.Thread(target=self._upd_check_thread, daemon=True).start()

    def _upd_check_thread(self):
        try:
            url = self._get_effective_url()
            if not url:
                self.ui_queue.put(("UPD", "error", "No remote config URL."))
                return

            r = requests.get(url, timeout=10)
            if r.status_code != 200:
                self.ui_queue.put(("UPD", "error",
                                   f"Could not reach update server (HTTP {r.status_code})."))
                return

            remote = r.json()
            remote_ver = remote.get("version", APP_VERSION)
            update_url = remote.get("update_url", "")

            if remote_ver <= APP_VERSION:
                self.ui_queue.put(("UPD", "latest"))
                return

            if not update_url:
                self.ui_queue.put(("UPD", "error",
                                   f"v{remote_ver} is available but no download URL configured."))
                return

            self.ui_queue.put(("UPD", "available", remote_ver, update_url))

        except requests.exceptions.ConnectionError:
            self.ui_queue.put(("UPD", "error", "No internet connection."))
        except Exception as e:
            self.ui_queue.put(("UPD", "error", f"Check failed: {str(e)}"))

    def _handle_update_msg(self, state, args):
        w = self._upd_widgets
        if not w or not self._upd_win or not self._upd_win.winfo_exists():
            return

        if state == "latest":
            w["status"].config(text="\u2713  You're up to date!", fg="#2E7D32")
            w["detail"].config(text=f"Version v{APP_VERSION} is the latest")
            w["action"].config(text="Up to date", bg="#43A047", fg="white",
                               state="disabled", disabledforeground="white")

        elif state == "available":
            remote_ver, url = args[0], args[1]
            w["status"].config(text=f"Update available: v{remote_ver}", fg="#E65100")
            w["detail"].config(text=f"Current: v{APP_VERSION}  \u2192  New: v{remote_ver}")

            def start_download():
                w["action"].config(text="Downloading...", state="disabled",
                                   bg="#BDBDBD", fg="#888", disabledforeground="#888")
                threading.Thread(target=self._upd_download_thread,
                                 args=(url,), daemon=True).start()

            w["action"].config(text="Download && Install", bg="#4CAF50", fg="white",
                               state="normal", cursor="hand2", command=start_download)
            w["action"].bind("<Enter>", lambda e: w["action"].config(bg="#43A047"))
            w["action"].bind("<Leave>", lambda e: w["action"].config(bg="#4CAF50"))

        elif state == "downloading":
            pct = args[0]
            downloaded = args[1] if len(args) > 1 else 0
            total = args[2] if len(args) > 2 else 0

            if not w["progress"].winfo_ismapped():
                w["progress"].pack(pady=(5, 3))
                w["progress_label"].pack()

            w["progress"]["value"] = pct
            w["status"].config(text="Downloading update...", fg="#1565C0")

            if total > 0:
                dl_mb = downloaded / (1024 * 1024)
                tot_mb = total / (1024 * 1024)
                w["progress_label"].config(text=f"{dl_mb:.1f} MB / {tot_mb:.1f} MB  ({pct}%)")
            else:
                dl_mb = downloaded / (1024 * 1024)
                w["progress_label"].config(text=f"{dl_mb:.1f} MB downloaded...")

        elif state == "downloaded":
            new_path = args[0]
            w["progress"]["value"] = 100
            w["progress_label"].config(text="Download complete")
            w["status"].config(text="\u2713  Download complete!", fg="#2E7D32")
            w["detail"].config(text="Click below to install and restart the app.")

            def do_install():
                self._apply_update(new_path)

            w["action"].config(text="Install && Restart", bg="#4CAF50", fg="white",
                               state="normal", cursor="hand2", command=do_install)
            w["action"].bind("<Enter>", lambda e: w["action"].config(bg="#43A047"))
            w["action"].bind("<Leave>", lambda e: w["action"].config(bg="#4CAF50"))

        elif state == "error":
            err_msg = args[0] if args else "Unknown error"
            w["status"].config(text="Update failed", fg="#D32F2F")
            w["detail"].config(text=err_msg)

            if w["progress"].winfo_ismapped():
                w["progress"].pack_forget()
                w["progress_label"].pack_forget()

            def do_retry():
                w["status"].config(text="Checking for updates...", fg="#333")
                w["detail"].config(text=f"Current version: v{APP_VERSION}")
                w["action"].config(text="Checking...", bg="#BDBDBD", fg="#888",
                                   state="disabled", disabledforeground="#888")
                threading.Thread(target=self._upd_check_thread, daemon=True).start()

            w["action"].config(text="Retry", bg="#2196F3", fg="white",
                               state="normal", cursor="hand2", command=do_retry)
            w["action"].bind("<Enter>", lambda e: w["action"].config(bg="#1E88E5"))
            w["action"].bind("<Leave>", lambda e: w["action"].config(bg="#2196F3"))

    def _upd_download_thread(self, url):
        try:
            new_path = os.path.join(CONFIG_DIR, "LegalTranslator_update.exe")

            r = requests.get(url, stream=True, timeout=120, allow_redirects=True)
            if r.status_code != 200:
                self.ui_queue.put(("UPD", "error",
                                   f"Download failed (HTTP {r.status_code})"))
                return

            total = int(r.headers.get('content-length', 0))
            downloaded = 0
            last_pct = -1

            with open(new_path, 'wb') as f:
                for chunk in r.iter_content(chunk_size=65536):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if total > 0:
                            pct = int(downloaded * 100 / total)
                        else:
                            pct = min(int(downloaded / 10000), 99)
                        if pct != last_pct:
                            self.ui_queue.put(("UPD", "downloading",
                                               pct, downloaded, total))
                            last_pct = pct

            if os.path.exists(new_path) and os.path.getsize(new_path) > 50000:
                self.ui_queue.put(("UPD", "downloaded", new_path))
            else:
                self.ui_queue.put(("UPD", "error",
                                   "Downloaded file is too small or corrupt.\n"
                                   "Please check the update URL in the config."))
        except requests.exceptions.ConnectionError:
            self.ui_queue.put(("UPD", "error", "Lost internet connection during download."))
        except Exception as e:
            self.ui_queue.put(("UPD", "error", f"Download error: {str(e)}"))

    def _apply_update(self, new_path):
        w = self._upd_widgets
        if w and self._upd_win and self._upd_win.winfo_exists():
            w["status"].config(text="Installing update...", fg="#1565C0")
            w["detail"].config(text="The app will close and restart automatically.")
            w["action"].config(state="disabled", text="Installing...",
                               bg="#BDBDBD", fg="#888", disabledforeground="#888")
            w["close"].config(state="disabled")
            self._upd_win.update()

        if not getattr(sys, 'frozen', False):
            messagebox.showinfo(
                APP_NAME,
                "Update downloaded successfully!\n\n"
                "Auto-install only works for the packaged .exe.\n"
                f"New file saved at:\n{new_path}",
                parent=self._upd_win if self._upd_win else self.root
            )
            if w:
                w["status"].config(text="Download saved", fg="#2E7D32")
                w["action"].config(text="Done", state="disabled")
                w["close"].config(state="normal")
            return

        current_exe = sys.executable
        pid = os.getpid()
        bat_path = os.path.join(CONFIG_DIR, "do_update.bat")

        bat_content = (
            '@echo off\n'
            'setlocal\n'
            '\n'
            'REM Wait for old process to exit\n'
            ':waitloop\n'
            f'tasklist /FI "PID eq {pid}" 2>NUL | find /I "{pid}" >NUL\n'
            'if not errorlevel 1 (\n'
            '    timeout /t 1 /nobreak >nul\n'
            '    goto waitloop\n'
            ')\n'
            '\n'
            'REM Extra wait for file handles to release\n'
            'timeout /t 3 /nobreak >nul\n'
            '\n'
            'REM Clean up PyInstaller temp folders\n'
            'for /d %%i in ("%TEMP%\\_MEI*") do rd /s /q "%%i" 2>nul\n'
            '\n'
            'REM Kill any lingering processes using the exe\n'
            f'taskkill /f /im "{os.path.basename(current_exe)}" 2>nul\n'
            'timeout /t 1 /nobreak >nul\n'
            '\n'
            'REM Try to copy new exe over old\n'
            f'copy /y "{new_path}" "{current_exe}"\n'
            'if errorlevel 1 (\n'
            '    echo First copy attempt failed, retrying...\n'
            f'    del /f "{current_exe}" 2>nul\n'
            '    timeout /t 2 /nobreak >nul\n'
            f'    copy /y "{new_path}" "{current_exe}"\n'
            '    if errorlevel 1 (\n'
            '        echo.\n'
            '        echo ============================================\n'
            '        echo   UPDATE FAILED - Could not replace file\n'
            '        echo ============================================\n'
            '        echo.\n'
            f'        echo New version saved at: {new_path}\n'
            '        echo You can manually copy it to replace the old exe.\n'
            '        echo.\n'
            '        pause\n'
            '        exit /b 1\n'
            '    )\n'
            ')\n'
            '\n'
            'REM Start new version\n'
            f'start "" "{current_exe}"\n'
            '\n'
            'REM Cleanup temp files\n'
            f'del "{new_path}" >nul 2>&1\n'
            'del "%~f0" >nul 2>&1\n'
        )

        try:
            with open(bat_path, 'w') as f:
                f.write(bat_content)

            subprocess.Popen(
                ['cmd', '/c', bat_path],
                creationflags=0x08000000,
                close_fds=True
            )

            self._quit_app()

        except Exception as e:
            messagebox.showerror(APP_NAME, f"Failed to start updater:\n{str(e)}",
                                 parent=self._upd_win if self._upd_win else self.root)
            if w:
                w["status"].config(text="Install failed", fg="#D32F2F")
                w["close"].config(state="normal")

    # ── Floating toolbar ──

    def create_floating_toolbar(self):
        self.toolbar_win = tk.Toplevel(self.root)
        self.toolbar_win.overrideredirect(True)
        self.toolbar_win.attributes("-topmost", True)
        self.toolbar_win.configure(bg="#202124", bd=0)

        w, h = 255, 36
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.toolbar_win.geometry(f"{w}x{h}+{sw - w - 50}+{sh - h - 100}")

        self._drag_data = {"x": 0, "y": 0}

        def start_drag(event):
            self._drag_data["x"] = event.x
            self._drag_data["y"] = event.y

        def do_drag(event):
            x = self.toolbar_win.winfo_x() - self._drag_data["x"] + event.x
            y = self.toolbar_win.winfo_y() - self._drag_data["y"] + event.y
            self.toolbar_win.geometry(f"+{x}+{y}")

        drag_frame = tk.Frame(self.toolbar_win, bg="#202124", width=16, cursor="fleur")
        drag_frame.pack(side="left", fill="y")
        drag_frame.bind("<ButtonPress-1>", start_drag)
        drag_frame.bind("<B1-Motion>", do_drag)

        grip_canvas = tk.Canvas(drag_frame, width=16, height=36, bg="#202124", highlightthickness=0)
        grip_canvas.pack()
        for y_pos in range(8, 28, 4):
            grip_canvas.create_rectangle(5, y_pos, 7, y_pos + 2, fill="#555", outline="")
            grip_canvas.create_rectangle(9, y_pos, 11, y_pos + 2, fill="#555", outline="")
        grip_canvas.bind("<ButtonPress-1>", start_drag)
        grip_canvas.bind("<B1-Motion>", do_drag)

        def make_hover_button(parent, text, bg_color, hover_color, command):
            btn = tk.Button(parent, text=text, bg=bg_color, fg="white",
                            font=("Segoe UI", 8, "bold"), relief="flat",
                            activebackground=hover_color, activeforeground="white",
                            cursor="hand2", command=command, bd=0, padx=6, pady=2)
            btn.bind("<Enter>", lambda e: btn.config(bg=hover_color))
            btn.bind("<Leave>", lambda e: btn.config(bg=bg_color))
            btn.pack(side="left", padx=(1, 0), fill="both", expand=True)
            return btn

        make_hover_button(self.toolbar_win, "Legal Eng", "#1976D2", "#1565C0",
                          lambda: self.trigger_translation("prompt_to_english"))
        make_hover_button(self.toolbar_win, "Hindi", "#388E3C", "#2E7D32",
                          lambda: self.trigger_translation("prompt_to_hindi"))

        caret_btn = tk.Button(self.toolbar_win, text="\u25BE", bg="#202124", fg="#757575",
                              font=("Segoe UI", 9), relief="flat",
                              activebackground="#333", activeforeground="#aaa",
                              cursor="hand2", bd=0, padx=3, pady=0,
                              command=self._toggle_dropdown)
        caret_btn.bind("<Enter>", lambda e: caret_btn.config(bg="#333", fg="#aaa"))
        caret_btn.bind("<Leave>", lambda e: caret_btn.config(bg="#202124", fg="#757575"))
        caret_btn.pack(side="left", fill="y")

        close_btn = tk.Button(self.toolbar_win, text="\u00D7", bg="#202124", fg="#666",
                              font=("Segoe UI", 11), relief="flat",
                              activebackground="#D32F2F", activeforeground="white",
                              cursor="hand2", bd=0, padx=4,
                              command=lambda: self.ui_queue.put(("SET_TOOLBAR_STATE", False)))
        close_btn.bind("<Enter>", lambda e: close_btn.config(bg="#D32F2F", fg="white"))
        close_btn.bind("<Leave>", lambda e: close_btn.config(bg="#202124", fg="#666"))
        close_btn.pack(side="right", fill="y")

        self.toolbar_win.after(100, lambda: self._apply_noactivate(self.toolbar_win))

    # ── Progress spinner ──

    def _show_progress_ui(self):
        if self.progress_win is not None:
            return
        self.progress_win = tk.Toplevel(self.root)
        self.progress_win.overrideredirect(True)
        self.progress_win.attributes("-topmost", True)
        self.progress_win.configure(bg="#1E272E")

        w, h = 300, 100
        x = (self.root.winfo_screenwidth() - w) // 2
        y = (self.root.winfo_screenheight() - h) // 2
        self.progress_win.geometry(f"{w}x{h}+{x}+{y}")

        canvas = tk.Canvas(self.progress_win, width=w, height=h, bg="#1E272E",
                           highlightthickness=2, highlightbackground="#00A8FF")
        canvas.pack(fill="both", expand=True)

        self.spinner_angle = 0
        self.spinner_arc = canvas.create_arc(
            20, 25, 70, 75,
            start=0, extent=100, outline="#00A8FF", width=5, style="arc"
        )
        canvas.create_text(90, 50, text="Translating... Please wait",
                           fill="white", font=("Segoe UI", 12, "bold"), anchor="w")

        self.progress_win.after(50, lambda: self._apply_noactivate(self.progress_win))
        self._animate_spinner(canvas)

    def _animate_spinner(self, canvas):
        try:
            if not self.progress_win or not canvas.winfo_exists():
                return
            self.spinner_angle = (self.spinner_angle + 15) % 360
            canvas.itemconfig(self.spinner_arc, start=self.spinner_angle)
            self.spinner_task = self.root.after(40, lambda: self._animate_spinner(canvas))
        except Exception:
            pass

    def _hide_progress_ui(self):
        if self.spinner_task:
            self.root.after_cancel(self.spinner_task)
            self.spinner_task = None
        if self.progress_win:
            self.progress_win.destroy()
            self.progress_win = None

    # ── Result dialog ──

    def _show_result_ui(self, original, translated, active_hwnd, direction):
        win = tk.Toplevel(self.root)
        win.title(f"{APP_NAME} - Review")
        win.attributes("-topmost", True)
        win.configure(bg="#FAFAFA")

        try:
            icon_path = get_icon_path("icon.png")
            if os.path.exists(icon_path):
                icon_img = tk.PhotoImage(file=icon_path)
                win.iconphoto(False, icon_img)
                win._icon_ref = icon_img
        except:
            pass

        w, h = 780, 620
        x = (self.root.winfo_screenwidth() - w) // 2
        y = (self.root.winfo_screenheight() - h) // 2
        win.geometry(f"{w}x{h}+{x}+{y}")
        win.minsize(620, 480)

        hint_frame = tk.Frame(win, bg="#EEEEEE")
        hint_frame.pack(side="bottom", fill="x")
        tk.Label(hint_frame,
                 text="Tip: Edit the translation above, then choose an action  "
                      "\u2022  Drag the divider to resize panels",
                 font=("Segoe UI", 8), fg="#999", bg="#EEEEEE").pack(pady=5)

        btn_bar = tk.Frame(win, bg="#EEEEEE", pady=10, padx=18)
        btn_bar.pack(side="bottom", fill="x")

        direction_label = "Legal English" if direction == "prompt_to_english" else "Hindi (Devanagari)"
        title_frame = tk.Frame(win, bg="#1a237e", pady=12)
        title_frame.pack(side="top", fill="x")
        tk.Label(title_frame, text=f"Translation Ready  \u2192  {direction_label}",
                 font=("Segoe UI", 15, "bold"), fg="white", bg="#1a237e").pack()

        paned = tk.PanedWindow(win, orient=tk.VERTICAL, bg="#CCCCCC",
                               sashwidth=6, sashrelief="flat", bd=0,
                               opaqueresize=True)
        paned.pack(side="top", fill="both", expand=True, padx=18, pady=(12, 6))

        orig_pane = tk.Frame(paned, bg="#FAFAFA")
        tk.Label(orig_pane, text="ORIGINAL TEXT",
                 font=("Segoe UI", 9, "bold"), fg="#888", bg="#FAFAFA").pack(anchor="w", pady=(0, 3))

        orig_inner = tk.Frame(orig_pane, bg="#E8EAF6", bd=1, relief="solid")
        orig_inner.pack(fill="both", expand=True)
        orig_inner.grid_columnconfigure(0, weight=1)
        orig_inner.grid_rowconfigure(0, weight=1)

        orig_scroll = tk.Scrollbar(orig_inner, orient="vertical")
        orig_scroll.grid(row=0, column=1, sticky="ns")
        orig_text = tk.Text(orig_inner, font=("Consolas", 10),
                            bg="#E8EAF6", fg="#333", wrap="word", relief="flat",
                            bd=0, padx=12, pady=10, yscrollcommand=orig_scroll.set,
                            selectbackground="#C5CAE9", selectforeground="#1a237e")
        orig_text.grid(row=0, column=0, sticky="nsew")
        orig_scroll.config(command=orig_text.yview)
        orig_text.insert("1.0", original)
        orig_text.config(state="disabled")
        paned.add(orig_pane, stretch="always")

        trans_pane = tk.Frame(paned, bg="#FAFAFA")
        tk.Label(trans_pane, text="TRANSLATION  (editable)",
                 font=("Segoe UI", 10, "bold"), fg="#1a237e", bg="#FAFAFA").pack(anchor="w", pady=(4, 3))

        trans_inner = tk.Frame(trans_pane, bg="white", bd=1, relief="solid",
                               highlightbackground="#C5CAE9", highlightthickness=1)
        trans_inner.pack(fill="both", expand=True)
        trans_inner.grid_columnconfigure(0, weight=1)
        trans_inner.grid_rowconfigure(0, weight=1)

        trans_scroll = tk.Scrollbar(trans_inner, orient="vertical")
        trans_scroll.grid(row=0, column=1, sticky="ns")
        trans_text = tk.Text(trans_inner, font=("Georgia", 13), bg="white", fg="#222",
                             wrap="word", relief="flat", bd=0, padx=14, pady=12,
                             yscrollcommand=trans_scroll.set, insertbackground="#1a237e",
                             selectbackground="#BBDEFB", selectforeground="#0D47A1")
        trans_text.grid(row=0, column=0, sticky="nsew")
        trans_scroll.config(command=trans_text.yview)
        trans_text.insert("1.0", translated)
        paned.add(trans_pane, stretch="always")

        def set_equal_sash():
            try:
                total = paned.winfo_height()
                if total > 50:
                    paned.sash_place(0, 0, total // 2)
            except:
                pass

        win.after(150, set_equal_sash)

        def _execute_action_thread(action_type, final_text, hwnd):
            if action_type == "replace":
                pyperclip.copy(final_text)
            elif action_type == "insert":
                pyperclip.copy("\n\n" + final_text + "\n")
            user32 = ctypes.windll.user32
            kernel32 = ctypes.windll.kernel32
            target_tid = user32.GetWindowThreadProcessId(hwnd, None)
            current_tid = kernel32.GetCurrentThreadId()
            try:
                if target_tid != current_tid and target_tid != 0:
                    user32.AttachThreadInput(current_tid, target_tid, True)
                    user32.SetForegroundWindow(hwnd)
                    user32.BringWindowToTop(hwnd)
                    user32.AttachThreadInput(current_tid, target_tid, False)
                else:
                    user32.SetForegroundWindow(hwnd)
            except Exception:
                pass
            time.sleep(0.3)
            if action_type == "insert":
                keyboard.send("right")
                time.sleep(0.1)
            keyboard.send("ctrl+v")

        def _execute_action(action_type):
            final_text = trans_text.get("1.0", "end-1c").strip()
            if action_type == "copy":
                pyperclip.copy(final_text)
                copy_btn.config(text="\u2713 Copied", bg="#43A047")
                win.after(800, win.destroy)
                return
            win.destroy()
            self.root.update()
            threading.Thread(target=_execute_action_thread,
                             args=(action_type, final_text, active_hwnd),
                             daemon=True).start()

        def _do_retranslate():
            cur_original = original
            win.destroy()
            self._retranslate(cur_original, direction, active_hwnd)

        btn_style = {"font": ("Segoe UI", 10, "bold"), "padx": 12, "pady": 5,
                     "relief": "flat", "cursor": "hand2", "bd": 0}

        replace_btn = tk.Button(btn_bar, text="Replace Selection", bg="#4caf50", fg="white",
                                command=lambda: _execute_action("replace"), **btn_style)
        replace_btn.pack(side="left", padx=(0, 6))
        replace_btn.bind("<Enter>", lambda e: replace_btn.config(bg="#43A047"))
        replace_btn.bind("<Leave>", lambda e: replace_btn.config(bg="#4caf50"))

        insert_btn = tk.Button(btn_bar, text="Insert Below", bg="#2196f3", fg="white",
                               command=lambda: _execute_action("insert"), **btn_style)
        insert_btn.pack(side="left", padx=(0, 6))
        insert_btn.bind("<Enter>", lambda e: insert_btn.config(bg="#1E88E5"))
        insert_btn.bind("<Leave>", lambda e: insert_btn.config(bg="#2196f3"))

        copy_btn = tk.Button(btn_bar, text="Copy", bg="#ff9800", fg="white",
                             command=lambda: _execute_action("copy"), **btn_style)
        copy_btn.pack(side="left", padx=(0, 6))
        copy_btn.bind("<Enter>", lambda e: copy_btn.config(bg="#FB8C00"))
        copy_btn.bind("<Leave>", lambda e: copy_btn.config(bg="#ff9800"))

        retranslate_btn = tk.Button(btn_bar, text="Retranslate", bg="#7B1FA2", fg="white",
                                    command=_do_retranslate, **btn_style)
        retranslate_btn.pack(side="left", padx=(0, 6))
        retranslate_btn.bind("<Enter>", lambda e: retranslate_btn.config(bg="#6A1B9A"))
        retranslate_btn.bind("<Leave>", lambda e: retranslate_btn.config(bg="#7B1FA2"))

        cancel_btn = tk.Button(btn_bar, text="Cancel", bg="#BDBDBD", fg="#333",
                               command=win.destroy, **btn_style)
        cancel_btn.pack(side="right")
        cancel_btn.bind("<Enter>", lambda e: cancel_btn.config(bg="#E53935", fg="white"))
        cancel_btn.bind("<Leave>", lambda e: cancel_btn.config(bg="#BDBDBD", fg="#333"))

    # ── Error / Key dialogs ──

    def _show_error_ui(self, msg):
        messagebox.showerror(APP_NAME, msg, parent=self.root)

    def _show_key_dialog_ui(self):
        win = tk.Toplevel(self.root)
        win.title(f"{APP_NAME} - API Key")
        win.attributes("-topmost", True)
        win.geometry("480x220")
        win.configure(bg="#f5f5f5")
        win.resizable(False, False)

        try:
            icon_path = get_icon_path("icon.png")
            if os.path.exists(icon_path):
                icon_img = tk.PhotoImage(file=icon_path)
                win.iconphoto(False, icon_img)
                win._icon_ref = icon_img
        except:
            pass

        win.update_idletasks()
        x = (self.root.winfo_screenwidth() - 480) // 2
        y = (self.root.winfo_screenheight() - 220) // 2
        win.geometry(f"+{x}+{y}")

        tk.Label(win, text="Enter Gemini API Key:",
                 font=("Segoe UI", 12, "bold"), bg="#f5f5f5").pack(pady=(15, 5))
        tk.Label(win, text="Get a free key from: aistudio.google.com",
                 font=("Segoe UI", 9), fg="#666", bg="#f5f5f5").pack()

        entry = tk.Entry(win, font=("Consolas", 11), width=45, show="*")
        entry.pack(pady=10, padx=20)
        entry.insert(0, self.config.get("api_key"))

        btn_frame = tk.Frame(win, bg="#f5f5f5")
        btn_frame.pack(pady=5)

        show_var = tk.BooleanVar(value=False)

        def toggle_show():
            if show_var.get():
                entry.config(show="")
                show_btn.config(text="Hide")
            else:
                entry.config(show="*")
                show_btn.config(text="Show")

        show_btn = tk.Button(btn_frame, text="Show", font=("Segoe UI", 9),
                             bg="#e0e0e0", relief="flat", cursor="hand2", padx=10,
                             command=lambda: (show_var.set(not show_var.get()), toggle_show()))
        show_btn.pack(side="left", padx=5)

        def save():
            key = entry.get().strip()
            if key:
                self.config.config["api_key"] = key
                self.config.save_local()
                messagebox.showinfo(APP_NAME, "API Key Saved!", parent=win)
                win.destroy()
            else:
                messagebox.showwarning(APP_NAME, "Please enter a key.", parent=win)

        save_btn = tk.Button(btn_frame, text="Save Key", font=("Segoe UI", 11, "bold"),
                             bg="#4caf50", fg="white", relief="flat", cursor="hand2",
                             command=save, padx=20, pady=3)
        save_btn.pack(side="left", padx=5)
        save_btn.bind("<Enter>", lambda e: save_btn.config(bg="#43A047"))
        save_btn.bind("<Leave>", lambda e: save_btn.config(bg="#4caf50"))

        cancel_btn = tk.Button(btn_frame, text="Cancel", font=("Segoe UI", 9),
                               bg="#e0e0e0", relief="flat", cursor="hand2", padx=10,
                               command=win.destroy)
        cancel_btn.pack(side="left", padx=5)

    def _quit_app(self):
        keyboard.unhook_all()
        if self.tray_icon:
            self.tray_icon.stop()
        self.root.quit()
        os._exit(0)


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