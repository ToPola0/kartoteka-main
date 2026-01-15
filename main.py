import sys
import os
import tkinter as tk
import logging
import time

# Sprawdź czy Pillow jest dostępny
try:
    from PIL import Image, ImageTk
except ImportError:
    # Sprawdź czy jesteśmy w venv
    venv_python = os.path.join(os.path.dirname(__file__), ".venv", "Scripts", "python.exe")
    if os.path.exists(venv_python):
        print("\n" + "="*60)
        print("BŁĄD: Pillow nie jest zainstalowany!")
        print("="*60)
        print("\nAby uruchomić program poprawnie, użyj jednej z metod:\n")
        print(f"1. Aktywuj venv: .venv\\Scripts\\Activate.ps1")
        print(f"   Potem uruchom: python main.py\n")
        print(f"2. Lub bezpośrednio: .venv\\Scripts\\python.exe main.py\n")
        print("="*60)
        input("\nNaciśnij Enter aby zakończyć...")
        sys.exit(1)
    else:
        print("\nBŁĄD: Brak modułu Pillow. Zainstaluj: pip install Pillow")
        sys.exit(1)

from splash_screen import SplashScreen
from gui_main import MainWindow

# Konfiguracja logowania
logging.basicConfig(
    filename='kartoteka_errors.log',
    level=logging.WARNING,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)

def load_application_with_splash():
    """Ładuje aplikację z profesjonalnym ekranem powitalnym."""
    # Utwórz główne okno root
    root = tk.Tk()
    root.withdraw()  # Ukryj główne okno na czas splash
    from config import set_window_icon
    set_window_icon(root)

    # Ścieżka do logo
    import sys
    if getattr(sys, 'frozen', False):
        base_dirs = [getattr(sys, '_MEIPASS', None), os.path.dirname(sys.executable)]
    else:
        base_dirs = [os.path.dirname(os.path.abspath(__file__))]
    logo_path = None
    for d in base_dirs:
        if d:
            p = os.path.join(d, "logo.png")
            if os.path.exists(p):
                logo_path = p
                break
    # Utwórz i wyświetl splash screen
    splash = SplashScreen(
        parent=root,
        logo_path=logo_path,
        title="Kartoteka Parafialna",
        subtitle="System Zarządzania i Analizy",
        version="v3.2"
    )
    splash.show()
    
    # Symuluj proces ładowania z prawdziwymi krokami
    loading_steps = [
        ("Ładowanie konfiguracji...", 0.3),
        ("Ładowanie słownika imion...", 0.5),
        ("Przygotowanie interfejsu graficznego...", 0.5),
        ("Finalizacja uruchamiania...", 0.2),
    ]
    
    for i, (step_text, duration) in enumerate(loading_steps):
        progress = int((i / len(loading_steps)) * 100)
        splash.update_progress(progress, step_text)
        time.sleep(duration)
    

    # Przywróć napisy statusu ładowania
    loading_steps = [
        ("Ładowanie konfiguracji...", 0.7),
        ("Ładowanie słownika imion...", 1.0),
        ("Przygotowanie interfejsu graficznego...", 1.0),
        ("Finalizacja uruchamiania...", 0.7),
    ]
    for i, (step_text, duration) in enumerate(loading_steps):
        progress = int((i / len(loading_steps)) * 100)
        splash.update_progress(progress, step_text)
        splash.root.update()  # Wymuś odświeżenie GUI
        time.sleep(duration)
    splash.update_progress(100, "Gotowe!")
    splash.root.update()
    time.sleep(0.7)

    # Zamknij splash screen
    splash.close()
    root.deiconify()  # Pokaż główne okno
    app = MainWindow(root)
    return root, app

def main():
    """Główna funkcja uruchamiająca aplikację."""
    try:
        root, app = load_application_with_splash()
        root.mainloop()
    except Exception as e:
        logging.error(f"Błąd krytyczny podczas uruchamiania: {e}", exc_info=True)
        import traceback
        print(f"\n{'='*60}")
        print("BŁĄD KRYTYCZNY")
        print(f"{'='*60}")
        print(f"\n{traceback.format_exc()}")
        print(f"{'='*60}")
        input("\nNaciśnij Enter aby zakończyć...")
        sys.exit(1)

if __name__ == "__main__":
    main()
