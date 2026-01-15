import os
import json
import logging
import sys

# Ścieżka bazowa - obsługa PyInstaller
if getattr(sys, 'frozen', False):
    # Uruchomione jako EXE
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # Uruchomione jako skrypt
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Ścieżka do katalogu kartotek
KATALOG_KARTOTEK = os.path.join(BASE_DIR, "kartotek")

# Ścieżka do pliku konfiguracji
CONFIG_FILE = os.path.join(BASE_DIR, "settings.json")

# Domyślne wartości
DEFAULT_AGE_FROM = 0
DEFAULT_AGE_TO = 120
DEFAULT_JUBILEE_DAYS = 30
DEFAULT_JSON_FILE = "imiona.json"

# Stałe kolory dla interfejsu
COLORS = {
    "primary": "#3498DB",
    "secondary": "#34495E",
    "accent": "#3498DB",
    "info": "#3498DB",
    "error": "#E74C3C",
    "warning": "#F39C12",
    "success": "#27AE60",
    "background": "#ECF0F1",
    "panel": "#FFFFFF",
    "text": "#2C3E50",
    "white": "#FFFFFF",
    "text_white": "#FFFFFF",
    "button_primary": "#3498DB",
    "sidebar": "#2C3E50",
    "border": "#BDC3C7",
}

def load_settings():
    """Wczytuje zapisane ustawienia z pliku."""
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except json.JSONDecodeError as e:
        logging.error(f"Błąd parsowania settings.json: {e}")
    except Exception as e:
        logging.error(f"Błąd wczytywania ustawień: {e}")
    return {}

def save_settings(settings):
    """Zapisuje ustawienia do pliku."""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        logging.error(f"Błąd zapisywania ustawień: {e}")
        return False

def set_window_icon(window):
    """Ustawia niestandardową ikonę okna (zamiast domyślnej ikony Pythona)."""
    try:
        # Debug: pokaż gdzie szukamy ikon
        print(f"[DEBUG] set_window_icon: BASE_DIR={BASE_DIR}")
        icon_files = ["logo.ico", "icon.ico"]
        for icon_name in icon_files:
            icon_path = os.path.join(BASE_DIR, icon_name)
            print(f"[DEBUG] set_window_icon: sprawdzam {icon_path}")
            if os.path.exists(icon_path):
                try:
                    window.iconbitmap(icon_path)
                    print(f"[DEBUG] set_window_icon: ustawiono {icon_path}")
                    return
                except Exception as e:
                    print(f"[DEBUG] set_window_icon: błąd ustawiania {icon_path}: {e}")
        # Jeśli nie ma pliku ico, użyj pustej ikony
        try:
            import tkinter as tk
            img = tk.PhotoImage(width=1, height=1)
            img.blank()
            window.iconphoto(True, img)
            print(f"[DEBUG] set_window_icon: ustawiono pustą ikonę")
        except Exception as e:
            print(f"[DEBUG] set_window_icon: błąd pustej ikony: {e}")
    except Exception as e:
        print(f"[DEBUG] set_window_icon: wyjątek główny: {e}")

