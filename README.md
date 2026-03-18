````md
# ⚖️ Legal Translator

**Instant Hindi → Legal English translation for Indian Courts**

Translate selected text directly from Microsoft Word using a single hotkey.

---

## ✨ Features

- One-click translation from Word using hotkeys
- Floating toolbar with quick access buttons
- Gemini and HuggingFace AI model support
- Auto-updates via GitHub Releases
- Configurable via system tray
- Secure - API keys stored locally and masked in UI

---

## 📥 Installation

### Option A - Download and Run (Recommended)

1. Go to **Releases**  
   https://github.com/webdevmrinal/LegalTranslator/releases/latest
2. Download `LegalTranslator.exe`
3. Double-click to run - no installation needed
4. Set your Gemini API key from the system tray icon

Get a free API key: https://aistudio.google.com/apikey

### Option B - Run from Source

```bash
git clone https://github.com/webdevmrinal/LegalTranslator.git
cd LegalTranslator

python -m venv venv
venv\Scripts\activate

pip install -r requirements.txt

python translator.py
````

---

## 🔧 Requirements

### For the executable

* Windows 10 or later

### For running from source

* Python 3.9 or later
* Dependencies listed in `requirements.txt`

---

## 🚀 Usage

### Hotkeys

| Shortcut         | Action                                       |
| ---------------- | -------------------------------------------- |
| Ctrl + Shift + E | Translate selected text → Legal English      |
| Ctrl + Shift + D | Translate selected text → Hindi (Devanagari) |

### Floating Toolbar

A small toolbar appears at the bottom-right of your screen with buttons for Legal English, Hindi, a dropdown menu for configuration and updates, and a close button.

### Workflow

1. Open Microsoft Word
2. Select the text you want to translate
3. Press Ctrl + Shift + E or click a toolbar button
4. Review the translation in the popup
5. Click **Replace Selection**, **Insert Below**, or **Copy**

---

## ⚙️ Configuration

| Setting        | How to Change                                  |
| -------------- | ---------------------------------------------- |
| API Key        | System tray → Set API Key                      |
| AI Provider    | System tray → Provider → Gemini or HuggingFace |
| View Config    | Toolbar dropdown → Show Config                 |
| Refresh Config | Toolbar dropdown → Refresh Config              |
| Check Updates  | Toolbar dropdown → Update App                  |

Settings are stored in:

```
Documents\LegalTranslator\local_config.json
```

---

## 🏗️ Building the Executable

```bash
pip install pyinstaller

pyinstaller LegalTranslator.spec
```

The packaged executable will be available in the `dist/` folder.

---

## 📁 Project Structure

```
LegalTranslator/
├── translator.py
├── icon.png
├── icon.ico
├── requirements.txt
├── .gitignore
├── LICENSE
└── README.md
```

---

## 🔒 Security

* API keys are stored locally on your machine
* Keys are masked in the configuration viewer
* The app connects only to Google AI and HuggingFace APIs
* Updates are downloaded only from GitHub Releases

---

## 📋 License

MIT License - see `LICENSE` for details.

---

Made for Indian legal professionals ⚖️

Report Bug - [https://github.com/webdevmrinal/LegalTranslator/issues](https://github.com/webdevmrinal/LegalTranslator/issues)
Request Feature - [https://github.com/webdevmrinal/LegalTranslator/issues](https://github.com/webdevmrinal/LegalTranslator/issues)

```
```
