"""
Moduł splash screen - wyświetla profesjonalny ekran powitalny podczas uruchamiania aplikacji.
"""
import tkinter as tk
from tkinter import ttk
import os
import time

class SplashScreen:
    """Profesjonalny ekran powitalny z paskiem postępu."""
    
    def __init__(self, parent, logo_path=None, title="Kartoteka Parafialna", 
                 subtitle="System Zarządzania i Analizy", version="v3.2"):
        """
        Inicjalizuje splash screen.
        
        Args:
            parent: Rodzic okna (musi być root lub Toplevel)
            logo_path: Ścieżka do pliku logo
            title: Tytuł aplikacji
            subtitle: Podtytuł
            version: Wersja aplikacji
        """
        # ZAWSZE korzystaj z przekazanego parent (root lub Toplevel)
        if parent is None:
            raise ValueError("SplashScreen wymaga przekazania parent (root)")
        self.root = tk.Toplevel(parent)
        
        # Ustaw niestandardową ikonę
        try:
            from config import set_window_icon, BASE_DIR
            print(f"[DEBUG] SplashScreen BASE_DIR={BASE_DIR}")
            set_window_icon(self.root)
        except Exception as e:
            print(f"[DEBUG] SplashScreen: błąd set_window_icon: {e}")
        
        self.root.overrideredirect(True)  # Usuń ramkę okna
        # Wymuś logo_path na logo.png z BASE_DIR jeśli nie podano
        if logo_path is None:
            try:
                from config import BASE_DIR
                logo_path = os.path.join(BASE_DIR, "logo.png")
                print(f"[DEBUG] SplashScreen: wymuszam logo_path={logo_path}")
            except Exception as e:
                print(f"[DEBUG] SplashScreen: błąd ustalania logo_path: {e}")
        self.logo_path = logo_path
        self.title = title
        self.subtitle = subtitle
        self.version = version
        
        # Wymiary okna
        self.width = 600
        self.height = 440  # +40px na dolne napisy
        
        # Wycentruj okno
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - self.width) // 2
        y = (screen_height - self.height) // 2
        self.root.geometry(f"{self.width}x{self.height}+{x}+{y}")
        
        # Ustawienia koloru
        self.bg_color = "#2C3E50"  # Ciemny niebieski
        self.accent_color = "#3498DB"  # Jasny niebieski
        self.text_color = "#ECF0F1"  # Jasnoszary
        self.progress_color = "#27AE60"  # Zielony
        
        self.root.configure(bg=self.bg_color)
        
        # Ramka z efektem cienia
        self.create_shadow_effect()
        
        # Główny kontener
        self.main_frame = tk.Frame(self.root, bg=self.bg_color, 
                                   highlightbackground=self.accent_color, 
                                   highlightthickness=3)
        self.main_frame.place(relx=0.5, rely=0.5, anchor="center", 
                             relwidth=0.95, relheight=0.95)
        
        self.setup_ui()
        
    def create_shadow_effect(self):
        """Tworzy efekt cienia wokół okna."""
        # Gradient border effect
        border_frame = tk.Frame(self.root, bg=self.accent_color)
        border_frame.place(x=0, y=0, width=self.width, height=self.height)
        
    def setup_ui(self):
        """Tworzy interfejs splash screen."""
        # Logo
        if self.logo_path and os.path.exists(self.logo_path):
            try:
                from PIL import Image, ImageTk
                img = Image.open(self.logo_path)
                
                # Resize do odpowiedniego rozmiaru
                max_size = 150
                img.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)
                
                self.logo_image = ImageTk.PhotoImage(img)
                logo_label = tk.Label(self.main_frame, image=self.logo_image, 
                                     bg=self.bg_color, bd=0)
                logo_label.pack(pady=(40, 20))
            except Exception as e:
                print(f"Nie można załadować logo: {e}")
                self._create_text_logo()
        else:
            self._create_text_logo()
        
        # Tytuł
        title_label = tk.Label(self.main_frame, text=self.title, 
                              font=("Segoe UI", 28, "bold"),
                              bg=self.bg_color, fg=self.text_color)
        title_label.pack(pady=(0, 5))
        
        # Podtytuł
        subtitle_label = tk.Label(self.main_frame, text=self.subtitle,
                                 font=("Segoe UI", 12),
                                 bg=self.bg_color, fg=self.text_color)
        subtitle_label.pack(pady=(0, 30))
        
        # Separator
        separator = tk.Frame(self.main_frame, height=2, bg=self.accent_color)
        separator.pack(fill="x", padx=50, pady=(0, 30))
        

        # Pasek postępu - modern style
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Splash.Horizontal.TProgressbar",
                   troughcolor=self.bg_color,
                   background=self.progress_color,
                   darkcolor=self.progress_color,
                   lightcolor=self.progress_color,
                   bordercolor=self.accent_color,
                   thickness=25)


        # Najpierw elementy na samym dole (bez duplikatów)
        self.status_label = tk.Label(self.main_frame, text="Inicjalizacja...",
                                     font=("Segoe UI", 14, "bold"),
                                     bg=self.bg_color, fg=self.accent_color)
        self.status_label.pack(pady=(10, 10))

        self.progress = ttk.Progressbar(self.main_frame, 
                                       length=400,
                                       mode='determinate',
                                       style="Splash.Horizontal.TProgressbar")
        self.progress.pack(side="bottom", pady=(0, 6))

        copyright_label = tk.Label(self.main_frame, 
                                  text="© 2026 Parafia Przyborów",
                                  font=("Segoe UI", 7),
                                  bg=self.bg_color, fg=self.text_color)
        copyright_label.pack(side="bottom", pady=(0, 2))

        version_label = tk.Label(self.main_frame, text=self.version,
                                font=("Segoe UI", 9),
                                bg=self.bg_color, fg=self.text_color)
        version_label.pack(side="bottom", pady=(0, 8))
        

        # (usunięto duplikaty version_label i copyright_label)
    
    def _create_text_logo(self):
        """Tworzy logo tekstowe jeśli brak pliku graficznego."""
        logo_text = tk.Label(self.main_frame, text="⛪",
                            font=("Segoe UI Emoji", 60),
                            bg=self.bg_color, fg=self.accent_color)
        logo_text.pack(pady=(40, 20))
    
    def update_progress(self, value, status_text=""):
        print(f"[DEBUG] update_progress: value={value}, status_text='{status_text}'")
        """
        Aktualizuje pasek postępu.
        
        Args:
            value: Wartość postępu (0-100)
            status_text: Tekst statusu do wyświetlenia
        """
        self.progress['value'] = value
        if status_text:
            self.status_label.config(text=status_text)
        self.root.update()
    
    def close(self):
        """Zamyka splash screen z efektem zanikania."""
        # Efekt zanikania
        for i in range(10, -1, -1):
            alpha = i / 10
            try:
                self.root.attributes('-alpha', alpha)
                self.root.update()
                time.sleep(0.02)
            except:
                pass
        
        self.root.destroy()
    
    def show(self):
        """Wyświetla splash screen."""
        self.root.attributes('-alpha', 0.0)
        self.root.update()
        
        # Efekt pojawiania się
        for i in range(11):
            alpha = i / 10
            try:
                self.root.attributes('-alpha', alpha)
                self.root.update()
                time.sleep(0.02)
            except:
                pass


def simulate_loading_with_splash(parent=None, logo_path=None, loading_steps=None):
    """
    Funkcja demonstracyjna - pokazuje splash screen z symulowanym ładowaniem.
    
    Args:
        parent: Rodzic okna
        logo_path: Ścieżka do logo
        loading_steps: Lista kroków ładowania [(opis, czas_trwania), ...]
    
    Returns:
        SplashScreen: Instancja splash screen
    """
    if loading_steps is None:
        loading_steps = [
            ("Ładowanie konfiguracji...", 0.3),
            ("Inicjalizacja bazy danych...", 0.4),
            ("Ładowanie słownika imion...", 0.3),
            ("Przygotowanie interfejsu...", 0.3),
            ("Uruchamianie aplikacji...", 0.2)
        ]
    
    splash = SplashScreen(parent, logo_path)
    splash.show()
    
    total_steps = len(loading_steps)
    for i, (step_text, duration) in enumerate(loading_steps):
        progress = int((i / total_steps) * 100)
        splash.update_progress(progress, step_text)
        time.sleep(duration)
    
    splash.update_progress(100, "Gotowe!")
    time.sleep(0.3)
    
    return splash


if __name__ == "__main__":
    # Test splash screen
    logo_path = os.path.join(os.path.dirname(__file__), "logo przeżroczyste.png")
    splash = simulate_loading_with_splash(logo_path=logo_path if os.path.exists(logo_path) else None)
    splash.close()
