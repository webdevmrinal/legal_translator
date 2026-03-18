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


def get_icon_path(filename):
    """Get path to bundled resource file, works both in dev and PyInstaller exe."""
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)


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
        self.spinner_arc = None
        self.spinner_angle = 0
        self.spinner_task = None

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
                        self.is_toolbar_visible = True
                    else:
                        self.toolbar_win.withdraw()
                        self.is_toolbar_visible = False
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
        # Try custom icon first
        try:
            icon_path = get_icon_path("icon.png")
            if os.path.exists(icon_path):
                img = Image.open(icon_path)
                img = img.resize((64, 64), Image.LANCZOS)
                return img
        except:
            pass

        # Fallback: draw H>E
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

            self.ui_queue.put(("SHOW_RESULT", text, result, active_hwnd))

        except Exception as e:
            self.ui_queue.put(("HIDE_PROGRESS",))
            self.ui_queue.put(("SHOW_ERROR", f"Unexpected error: {str(e)}"))
        finally:
            self.is_translating = False

    def _apply_noactivate(self, win):
        """Apply WS_EX_NOACTIVATE so window never steals focus from Word."""
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

    def create_floating_toolbar(self):
        self.toolbar_win = tk.Toplevel(self.root)
        self.toolbar_win.overrideredirect(True)
        self.toolbar_win.attributes("-topmost", True)
        self.toolbar_win.configure(bg="#202124", bd=1, relief="solid")

        w, h = 220, 40
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

        # Grip
        drag_frame = tk.Frame(self.toolbar_win, bg="#202124", width=18, cursor="fleur")
        drag_frame.pack(side="left", fill="y")
        drag_frame.bind("<ButtonPress-1>", start_drag)
        drag_frame.bind("<B1-Motion>", do_drag)

        grip_canvas = tk.Canvas(drag_frame, width=18, height=40, bg="#202124", highlightthickness=0)
        grip_canvas.pack()
        for y_pos in range(10, 31, 4):
            grip_canvas.create_rectangle(6, y_pos, 8, y_pos + 2, fill="#757575", outline="")
            grip_canvas.create_rectangle(10, y_pos, 12, y_pos + 2, fill="#757575", outline="")
        grip_canvas.bind("<ButtonPress-1>", start_drag)
        grip_canvas.bind("<B1-Motion>", do_drag)

        def make_hover_button(parent, text, bg_color, hover_color, command):
            btn = tk.Button(parent, text=text, bg=bg_color, fg="white",
                            font=("Segoe UI", 9, "bold"), relief="flat",
                            activebackground=hover_color, activeforeground="white",
                            cursor="hand2", command=command, bd=0, padx=8, pady=4)
            btn.bind("<Enter>", lambda e: btn.config(bg=hover_color))
            btn.bind("<Leave>", lambda e: btn.config(bg=bg_color))
            btn.pack(side="left", padx=2, fill="both", expand=True)
            return btn

        make_hover_button(self.toolbar_win, "Legal Eng", "#1976D2", "#1565C0",
                          lambda: self.trigger_translation("prompt_to_english"))
        make_hover_button(self.toolbar_win, "Hindi", "#388E3C", "#2E7D32",
                          lambda: self.trigger_translation("prompt_to_hindi"))

        # Close button
        close_btn = tk.Button(self.toolbar_win, text="x", bg="#202124", fg="#9AA0A6",
                              font=("Segoe UI", 10), relief="flat",
                              activebackground="#D32F2F", activeforeground="white",
                              cursor="hand2", bd=0, padx=5,
                              command=lambda: self.ui_queue.put(("SET_TOOLBAR_STATE", False)))
        close_btn.bind("<Enter>", lambda e: close_btn.config(bg="#D32F2F", fg="white"))
        close_btn.bind("<Leave>", lambda e: close_btn.config(bg="#202124", fg="#9AA0A6"))
        close_btn.pack(side="right", fill="y")

        self.toolbar_win.after(100, lambda: self._apply_noactivate(self.toolbar_win))

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

    def _show_result_ui(self, original, translated, active_hwnd):
        win = tk.Toplevel(self.root)
        win.title(f"{APP_NAME} - Review")
        win.attributes("-topmost", True)
        win.configure(bg="#f5f5f5")

        # Set window icon
        try:
            icon_path = get_icon_path("icon.png")
            if os.path.exists(icon_path):
                icon_img = tk.PhotoImage(file=icon_path)
                win.iconphoto(False, icon_img)
                win._icon_ref = icon_img  # prevent garbage collection
        except:
            pass

        win.update_idletasks()
        w, h = 750, 550
        x = (self.root.winfo_screenwidth() - w) // 2
        y = (self.root.winfo_screenheight() - h) // 2
        win.geometry(f"{w}x{h}+{x}+{y}")
        win.minsize(600, 450)

        # Title bar
        title_frame = tk.Frame(win, bg="#1a237e", pady=10)
        title_frame.pack(fill="x")
        tk.Label(title_frame, text="Translation Ready",
                 font=("Segoe UI", 16, "bold"), fg="white", bg="#1a237e").pack()

        # Content
        content = tk.Frame(win, bg="#f5f5f5", padx=15, pady=10)
        content.pack(fill="both", expand=True)

        tk.Label(content, text="ORIGINAL TEXT:",
                 font=("Segoe UI", 9, "bold"), fg="#666", bg="#f5f5f5").pack(anchor="w")
        orig_text = tk.Text(content, height=4, font=("Consolas", 10),
                            bg="#fff3e0", wrap="word", relief="solid", bd=1)
        orig_text.pack(fill="x", pady=(2, 10))
        orig_text.insert("1.0", original)
        orig_text.config(state="disabled")

        tk.Label(content, text="TRANSLATION (You can edit this):",
                 font=("Segoe UI", 10, "bold"), fg="#1a237e", bg="#f5f5f5").pack(anchor="w")
        trans_text = tk.Text(content, font=("Georgia", 13), bg="white",
                             wrap="word", relief="solid", bd=1, padx=10, pady=10)
        trans_text.pack(fill="both", expand=True, pady=(2, 10))
        trans_text.insert("1.0", translated)

        scrollbar = ttk.Scrollbar(trans_text, command=trans_text.yview)
        scrollbar.pack(side="right", fill="y")
        trans_text.config(yscrollcommand=scrollbar.set)

        # Buttons
        btn_frame = tk.Frame(win, bg="#e0e0e0", pady=12, padx=15)
        btn_frame.pack(fill="x")

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
                copy_btn.config(text="Copied!", bg="#43A047")
                win.after(1000, win.destroy)
                return

            win.destroy()
            self.root.update()
            threading.Thread(target=_execute_action_thread,
                             args=(action_type, final_text, active_hwnd),
                             daemon=True).start()

        btn_style = {"font": ("Segoe UI", 11, "bold"), "padx": 15, "pady": 6,
                     "relief": "flat", "cursor": "hand2"}

        replace_btn = tk.Button(btn_frame, text="Replace Selection", bg="#4caf50", fg="white",
                                command=lambda: _execute_action("replace"), **btn_style)
        replace_btn.pack(side="left", padx=5)
        replace_btn.bind("<Enter>", lambda e: replace_btn.config(bg="#43A047"))
        replace_btn.bind("<Leave>", lambda e: replace_btn.config(bg="#4caf50"))

        insert_btn = tk.Button(btn_frame, text="Insert Below", bg="#2196f3", fg="white",
                               command=lambda: _execute_action("insert"), **btn_style)
        insert_btn.pack(side="left", padx=5)
        insert_btn.bind("<Enter>", lambda e: insert_btn.config(bg="#1E88E5"))
        insert_btn.bind("<Leave>", lambda e: insert_btn.config(bg="#2196f3"))

        copy_btn = tk.Button(btn_frame, text="Copy Only", bg="#ff9800", fg="white",
                             command=lambda: _execute_action("copy"), **btn_style)
        copy_btn.pack(side="left", padx=5)
        copy_btn.bind("<Enter>", lambda e: copy_btn.config(bg="#FB8C00"))
        copy_btn.bind("<Leave>", lambda e: copy_btn.config(bg="#ff9800"))

        cancel_btn = tk.Button(btn_frame, text="Cancel", bg="#f44336", fg="white",
                               command=win.destroy, **btn_style)
        cancel_btn.pack(side="right", padx=5)
        cancel_btn.bind("<Enter>", lambda e: cancel_btn.config(bg="#E53935"))
        cancel_btn.bind("<Leave>", lambda e: cancel_btn.config(bg="#f44336"))

        # Hint
        tk.Label(btn_frame, text="You can edit the translation before clicking a button",
                 font=("Segoe UI", 8), fg="#666", bg="#e0e0e0").pack(side="bottom", pady=(5, 0))

    def _show_error_ui(self, msg):
        messagebox.showerror(APP_NAME, msg, parent=self.root)

    def _show_key_dialog_ui(self):
        win = tk.Toplevel(self.root)
        win.title(f"{APP_NAME} - API Key")
        win.attributes("-topmost", True)
        win.geometry("480x220")
        win.configure(bg="#f5f5f5")
        win.resizable(False, False)

        # Set window icon
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

        # Toggle show/hide key
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