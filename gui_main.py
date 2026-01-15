import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import threading
import os
import re
import pandas as pd
import logging
from datetime import datetime
from config import COLORS, DEFAULT_AGE_FROM, DEFAULT_AGE_TO, DEFAULT_JUBILEE_DAYS, KATALOG_KARTOTEK, load_settings, save_settings
from file_operations import load_names
from gui_dialogs import show_results_dialog, edit_unknown_name
from analysis import analyze_marriage_jubilees, analyze_grandparents_jubilees, extract_marriage_info, extract_grandparents_marriage_info
from data_processing import extract_words, extract_birth_date, calculate_age, remove_diacritics, format_person_name, extract_number_from_text
from statistics import Statistics
from export_statistics import export_statistics_to_excel, export_all_results_to_excel
import openpyxl

def save_found_people_to_xlsx(found_people, sort_key=None):
    """Zapisuje znalezione osoby do pliku XLSX."""
    if not found_people:
        messagebox.showinfo("Brak danych", "Brak osób do zapisania.")
        return

    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Plik Excel", "*.xlsx")])
    if not save_path:
        return

    list_to_save = found_people.copy()

    if sort_key:
        try:
            if sort_key == "wiek":
                list_to_save.sort(key=lambda p: (p.get("wiek") is None, p.get("wiek")))
            elif sort_key == "adres":
                list_to_save.sort(key=lambda p: ((p.get("adres") or "").lower()))
            elif sort_key == "stary_adres":
                list_to_save.sort(key=lambda p: ((p.get("old_address") or "").lower()))
            elif sort_key == "stary_numer":
                list_to_save.sort(key=lambda p: (
                    (p.get("old_address") is None),
                    (extract_number_from_text(p.get("old_address")) if extract_number_from_text(p.get("old_address")) is not None else float('inf'))
                ))
            elif sort_key == "alfabetycznie":
                list_to_save.sort(key=lambda p: (
                    (format_person_name(p.get("nazwisko","")).lower()),
                    (format_person_name(p.get("imie","")).lower())
                ))
        except (KeyError, AttributeError, TypeError) as e:
            logging.warning(f"Błąd podczas sortowania: {e}")
            messagebox.showwarning("Ostrzeżenie", "Nie udało się posortować listy - niektóre dane mogą być niepełne.")

    data = []
    for p in list_to_save:
        imie = format_person_name(p.get("imie", ""))
        nazwisko = format_person_name(p.get("nazwisko", ""))
        adres = p.get("adres", "") or ""
        if p.get("old_address"):
            adres += f" (stary: {p['old_address']})"
        wiek = p.get("wiek", "")
        plec = p.get("plec", "")
        data.append([imie, nazwisko, adres, wiek, plec])

    df = pd.DataFrame(data, columns=["Imię", "Nazwisko", "Adres", "Wiek", "Płeć"])

    try:
        with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Wyniki")
            worksheet = writer.sheets["Wyniki"]
            for col_num, column_title in enumerate(df.columns, 1):
                column_len = max(df[column_title].astype(str).map(len).max(), len(column_title))
                col_letter = openpyxl.utils.get_column_letter(col_num)
                worksheet.column_dimensions[col_letter].width = column_len + 2
        messagebox.showinfo("Zapisano", f"Zapisano {len(found_people)} osób do pliku XLSX.")
    except Exception as e:
        messagebox.showerror("Błąd", f"Nie udało się zapisać pliku: {e}")

class MainWindow:
    def reset_search_state(self):
        """Resetuje stan wyszukiwania w wynikach (po zmianie wyników)."""
        self._last_search_term = None
        self.search_index = None
    def __init__(self, root):
        self.root = root
        self.root.title("Kartoteka Parafialna Przyborów - System Zarządzania v3.2")
        self.root.geometry("1100x700")  # Domyślny rozmiar - pozycja zostanie ustawiona później
        
        # Ustaw minimalny rozmiar okna
        self.root.minsize(950, 600)
        
        # Kolor tła
        self.root.configure(bg=COLORS["background"])
        
        # Ustaw ikonę okna (Windows: .ico na pasku, Tkinter: logo.png w GUI)
        debug_msgs = []
        try:
            import sys
            from PIL import Image, ImageTk

            if getattr(sys, 'frozen', False):
                exe_dir = os.path.dirname(sys.executable)
                meipass = getattr(sys, '_MEIPASS', None)
                # Najpierw katalog EXE, potem _MEIPASS/_internal
                base_dirs = [exe_dir]
                if meipass and meipass != exe_dir:
                    base_dirs.append(meipass)
                # Dodatkowo _internal jeśli istnieje
                internal_dir = os.path.join(exe_dir, '_internal')
                if os.path.isdir(internal_dir):
                    base_dirs.append(internal_dir)
            else:
                base_dirs = [os.path.dirname(os.path.abspath(__file__))]

            debug_msgs.append(f"[DEBUG] base_dirs: {base_dirs}")

            # Ikona na pasku Windows (.ico)
            ico_path = None
            for d in base_dirs:
                if d:
                    p = os.path.join(d, "logo.ico")
                    debug_msgs.append(f"[DEBUG] Szukam logo.ico: {p}")
                    if os.path.exists(p):
                        ico_path = p
                        debug_msgs.append(f"[DEBUG] Znalazłem logo.ico: {p}")
                        break
            if ico_path:
                try:
                    self.root.iconbitmap(ico_path)
                    debug_msgs.append(f"[DEBUG] Ikona .ico ustawiona: {ico_path}")
                except Exception as e:
                    debug_msgs.append(f"[DEBUG] Błąd ustawiania .ico: {e}")
            else:
                debug_msgs.append(f"[DEBUG] Nie znaleziono logo.ico w {base_dirs}")

            # Ikona w GUI (logo.png)
            logo_path = None
            for d in base_dirs:
                if d:
                    p = os.path.join(d, "logo.png")
                    debug_msgs.append(f"[DEBUG] Szukam logo.png: {p}")
                    if os.path.exists(p):
                        logo_path = p
                        debug_msgs.append(f"[DEBUG] Znalazłem logo.png: {p}")
                        break
            if logo_path:
                try:
                    icon_image = Image.open(logo_path)
                    self.icon_photo = ImageTk.PhotoImage(icon_image)
                    self.root.iconphoto(True, self.icon_photo)
                    debug_msgs.append(f"[DEBUG] Logo.png ustawione jako iconphoto: {logo_path}")
                    # Usunięto testowy label, logo.png przypisane do self.icon_photo
                except Exception as e:
                    debug_msgs.append(f"[DEBUG] Błąd ustawiania logo.png: {e}")
            else:
                debug_msgs.append(f"[DEBUG] Nie znaleziono logo.png w {base_dirs}")
        except Exception as e:
            debug_msgs.append(f"[DEBUG] Wyjątek podczas ustawiania ikon: {e}")
            import traceback
            debug_msgs.append(traceback.format_exc())

        # Usunięto okno debug

        self.found_people = []
        self.names_dict = {}
        self.folder_path = ""
        self.statistics = Statistics()  # Obiekt do zbierania statystyk
        self.json_file_path = None  # Będzie ustawiony dynamicznie
        self.search_index = None
        self.refresh_after_id = None
        self.btn_analyze = None
        self.edit_unknown_btn = None  # Przycisk edycji nieznanych imion
        self.loading_settings = True  # Flaga aby nie triggerować reanalysis podczas ładowania
        
        # Dane z ostatniej analizy
        self.jubilees_found = []  # Lista jubileuszy
        self.marriages_in_range = []  # Lista ślubów w zakresie
        self.all_unknown = {}  # Słownik nieznanych imion
        self.analysis_details = []  # Szczegóły analizy (błędy, ostrzeżenia)
        
        # Zmienne do filtrowania ślubów
        self.marriage_year_from_var = tk.IntVar(value=1900)
        self.marriage_year_to_var = tk.IntVar(value=datetime.now().year)

        self.setup_ui()
        self.load_saved_settings()
        
        # Podłącz obsługę zamknięcia okna
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Przywróć pozycję i rozmiar okna (musi być po setup_ui)
        self.restore_window_geometry()
        
        self.loading_settings = False  # Teraz można już uruchamiać reanalysis
        
        # Inicjalizuj imiona i uruchom automatyczną analizę jeśli jest folder
        self.initialize_names()

    def setup_ui(self):
        """Konfiguruje interfejs użytkownika."""
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

        # Kontener dla lewego panelu ze scrollbarem
        left_container = tk.Frame(self.root, bg=COLORS.get("sidebar", COLORS["panel"]))
        left_container.grid(row=0, column=0, sticky="ns")
        
        # Canvas dla przewijania
        self.left_canvas = tk.Canvas(left_container, width=280, 
                                     bg=COLORS.get("sidebar", COLORS["panel"]),
                                     highlightthickness=0)
        self.left_scrollbar = tk.Scrollbar(left_container, orient="vertical", 
                                           command=self.left_canvas.yview)
        
        # Frame wewnętrzny dla zawartości
        self.left_frame = tk.Frame(self.left_canvas, bg=COLORS.get("sidebar", COLORS["panel"]))
        
        # Konfiguracja canvas
        self.left_canvas.create_window((0, 0), window=self.left_frame, anchor="nw")
        self.left_canvas.configure(yscrollcommand=self.left_scrollbar.set)
        
        # Pakowanie
        self.left_canvas.pack(side="left", fill="both", expand=True)
        self.left_scrollbar.pack(side="right", fill="y")
        
        # Aktualizacja regionu przewijania
        def configure_scroll_region(event=None):
            self.left_canvas.configure(scrollregion=self.left_canvas.bbox("all"))
        
        self.configure_scroll_region = configure_scroll_region  # Zapisz jako metodę
        self.left_frame.bind("<Configure>", configure_scroll_region)
        
        # Przewijanie kółkiem myszy
        def on_mousewheel(event):
            self.left_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        self.left_canvas.bind_all("<MouseWheel>", on_mousewheel)

        self.setup_left_panel()

        # Prawy panel
        self.right_frame = tk.Frame(self.root, bg=COLORS["background"])
        self.right_frame.grid(row=0, column=1, sticky="nsew")

        self.setup_right_panel()

    def setup_left_panel(self):
        """Konfiguruje lewy panel kontrolny z profesjonalnym wyglądem."""
        # Nagłówek z logo
        header_frame = tk.Frame(self.left_frame, bg=COLORS.get("sidebar", COLORS["panel"]))
        header_frame.pack(fill="x", pady=(8, 5))
        
        # Dodaj logo na górze
        try:
            from PIL import Image, ImageTk
            logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")
            
            if os.path.exists(logo_path):
                logo_image = Image.open(logo_path)
                
                # Zmień rozmiar jeśli potrzeba (zachowaj proporcje)
                max_width = 200
                ratio = max_width / logo_image.width
                new_height = int(logo_image.height * ratio)
                logo_image = logo_image.resize((max_width, new_height), Image.Resampling.LANCZOS)
                
                # KRYTYCZNE: zachowaj referencję jako atrybut self, żeby garbage collector nie usunął obrazu!
                self.logo_photo = ImageTk.PhotoImage(logo_image)
                self.logo_label = tk.Label(header_frame, image=self.logo_photo, 
                                          bg=COLORS.get("sidebar", COLORS["panel"]))
                self.logo_label.pack(pady=(0, 5))
        except Exception as e:
            # Jeśli nie można załadować logo, pokaż emoji
            logging.error(f"Błąd ładowania logo: {type(e).__name__}: {e}")
            emoji_label = tk.Label(header_frame, text="⛪", font=("Segoe UI Emoji", 32),
                                  bg=COLORS.get("sidebar", COLORS["panel"]),
                                  fg=COLORS.get("accent", "#3498DB"))
            emoji_label.pack(pady=(0, 3))
        
        # Tytuł aplikacji
        title_label = tk.Label(header_frame, text="Kartoteka", 
                              font=("Segoe UI", 13, "bold"),
                              bg=COLORS.get("sidebar", COLORS["panel"]),
                              fg=COLORS.get("text_white", "#FFFFFF"))
        title_label.pack()
        
        subtitle_label = tk.Label(header_frame, text="Parafia Przyborów",
                                 font=("Segoe UI", 8),
                                 bg=COLORS.get("sidebar", COLORS["panel"]),
                                 fg=COLORS.get("text_light", "#B0B0B0"))
        subtitle_label.pack(pady=(0, 5))
        
        # Separator
        sep1 = tk.Frame(self.left_frame, height=2, bg=COLORS.get("accent", "#3498DB"))
        sep1.pack(fill="x", padx=15, pady=(0, 8))
        
        # Sekcja Pliki
        files_lf = tk.LabelFrame(self.left_frame, text="📁 Pliki", padx=8, pady=6,
                                bg=COLORS.get("sidebar", COLORS["panel"]),
                                fg=COLORS.get("text_white", "#FFFFFF"),
                                font=("Segoe UI", 9, "bold"),
                                relief="flat")
        files_lf.pack(fill="x", padx=12, pady=(0, 6))
        
        # Styl przycisków - nowoczesny, płaski
        btn_style = {
            "relief": "flat",
            "bd": 0,
            "padx": 12,
            "pady": 6,
            "font": ("Segoe UI", 8),
            "cursor": "hand2",
            "activebackground": COLORS.get("button_primary_hover", "#2980B9"),
            "activeforeground": "white"
        }
        
        btn1 = tk.Button(files_lf, text="📂 Wybierz folder Excel", 
                        command=self.select_folder, 
                        bg=COLORS.get("button_primary", "#3498DB"), 
                        fg="white", **btn_style)
        btn1.pack(fill="x", pady=(0, 4))
        
        btn2 = tk.Button(files_lf, text="📋 Wybierz plik JSON", 
                        command=self.select_json_file, 
                        bg=COLORS.get("button_secondary", "#95A5A6"), 
                        fg="white", **btn_style)
        btn2.pack(fill="x", pady=(0, 4))
        
        btn3 = tk.Button(files_lf, text="➕ Nowa kartoteka", 
                        command=self.open_sample_kartoteka, 
                        bg=COLORS.get("info", "#3498DB"), 
                        fg="white", **btn_style)
        btn3.pack(fill="x")

        # Sekcja Ustawienia
        settings_lf = tk.LabelFrame(self.left_frame, text="⚙️ Ustawienia", padx=8, pady=6,
                                    bg=COLORS.get("sidebar", COLORS["panel"]),
                                    fg=COLORS.get("text_white", "#FFFFFF"),
                                    font=("Segoe UI", 9, "bold"),
                                    relief="flat")
        settings_lf.pack(fill="x", padx=12, pady=(0, 6))
        
        # Wyszukiwanie
        search_label = tk.Label(settings_lf, text="🔍 Wyszukaj w wynikach:",
                               bg=COLORS.get("sidebar", COLORS["panel"]),
                               fg=COLORS.get("text_white", "#FFFFFF"),
                               font=("Segoe UI", 9))
        search_label.pack(anchor="w", pady=(0, 4))
        
        self.search_entry = tk.Entry(settings_lf, font=("Segoe UI", 10),
                         relief="flat", bd=2)
        self.search_entry.pack(fill="x", pady=(0, 6))
        self.search_entry.bind("<Return>", lambda e: self.search_in_results())
        search_btn = tk.Button(settings_lf, text="🔎 Szukaj", 
                      command=self.search_in_results, 
                      bg=COLORS.get("accent", "#3498DB"), 
                      fg="white", **btn_style)
        search_btn.pack(fill="x", pady=(0, 8))

        # Wiek
        age_label = tk.Label(settings_lf, text="👤 Zakres wieku:",
                            bg=COLORS.get("sidebar", COLORS["panel"]),
                            fg=COLORS.get("text_white", "#FFFFFF"),
                            font=("Segoe UI", 8))
        age_label.pack(anchor="w", pady=(0, 3))
        
        age_row = tk.Frame(settings_lf, bg=COLORS.get("sidebar", COLORS["panel"]))
        age_row.pack(fill="x", pady=(0, 8))
        
        tk.Label(age_row, text="od:", bg=COLORS.get("sidebar", COLORS["panel"]),
                fg=COLORS.get("text_white", "#FFFFFF"),
                font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w")
        self.age_from_var = tk.IntVar(value=DEFAULT_AGE_FROM)
        age_from_entry = tk.Entry(age_row, textvariable=self.age_from_var, width=6,
                font=("Segoe UI", 10), relief="flat", bd=2)
        age_from_entry.grid(row=0, column=1, padx=(5, 10))
        age_from_entry.bind("<Return>", lambda e: self.apply_settings())
        
        tk.Label(age_row, text="do:", bg=COLORS.get("sidebar", COLORS["panel"]),
                fg=COLORS.get("text_white", "#FFFFFF"),
                font=("Segoe UI", 9)).grid(row=0, column=2)
        self.age_to_var = tk.IntVar(value=DEFAULT_AGE_TO)
        age_to_entry = tk.Entry(age_row, textvariable=self.age_to_var, width=6,
                font=("Segoe UI", 9), relief="flat", bd=2)
        age_to_entry.grid(row=0, column=3, padx=(5, 0))
        age_to_entry.bind("<Return>", lambda e: self.apply_settings())

        # Jubileusze
        jub_label = tk.Label(settings_lf, text="🎉 Jubileusze:",
                            bg=COLORS.get("sidebar", COLORS["panel"]),
                            fg=COLORS.get("text_white", "#FFFFFF"),
                            font=("Segoe UI", 8))
        jub_label.pack(anchor="w", pady=(0, 3))
        
        jub_row = tk.Frame(settings_lf, bg=COLORS.get("sidebar", COLORS["panel"]))
        jub_row.pack(fill="x")
        
        tk.Label(jub_row, text="Dni do jubileuszu:", 
                bg=COLORS.get("sidebar", COLORS["panel"]),
                fg=COLORS.get("text_white", "#FFFFFF"),
                font=("Segoe UI", 8)).pack(side="left")
        self.jubilee_days_var = tk.IntVar(value=DEFAULT_JUBILEE_DAYS)
        jubilee_entry = tk.Entry(jub_row, textvariable=self.jubilee_days_var, width=5,
                font=("Segoe UI", 9), relief="flat", bd=2)
        jubilee_entry.pack(side="left", padx=(5, 0))
        jubilee_entry.bind("<Return>", lambda e: self.apply_settings())

        # Przycisk Śluby
        marriages_btn = tk.Button(settings_lf, text="💍 Śluby w latach...", 
                                 command=self.show_marriages_dialog,
                                 bg=COLORS.get("info", "#3498DB"), 
                                 fg="white", **btn_style)
        marriages_btn.pack(fill="x", pady=(8, 0))

        # Sekcja Analiza
        analysis_lf = tk.LabelFrame(self.left_frame, text="📊 Analiza", padx=10, pady=10,
                                    bg=COLORS.get("sidebar", COLORS["panel"]),
                                    fg=COLORS.get("text_white", "#FFFFFF"),
                                    font=("Segoe UI", 10, "bold"),
                                    relief="flat")
        analysis_lf.pack(fill="x", padx=12, pady=(0, 10))
        
        # ...usunięto logo nad przyciskiem analizy...
        
        analyze_btn_style = btn_style.copy()
        analyze_btn_style["pady"] = 8
        analyze_btn_style["font"] = ("Segoe UI", 9, "bold")
        
        self.btn_analyze = tk.Button(analysis_lf, text="▶️ Analizuj teraz", 
                                     command=lambda: self.analyze_current_settings(show_dialog=True), 
                                     bg=COLORS.get("success", "#27AE60"), 
                                     fg="white", **analyze_btn_style)
        self.btn_analyze.pack(fill="x")        
        # Miejsce na przycisk edycji nieznanych imion (będzie dodany dynamicznie)
        self.edit_unknown_btn = None
        self.analysis_section = analysis_lf  # Zapisz referencję do sekcji analizy

    def setup_right_panel(self):
        """Konfiguruje prawy panel z wynikami - nowoczesny design."""
        # Nagłówek z tytułem
        header = tk.Frame(self.right_frame, bg=COLORS["background"])
        header.pack(fill="x", padx=15, pady=(15, 10))
        
        title = tk.Label(header, text="📋 Wyniki Analizy", 
                        font=("Segoe UI", 18, "bold"),
                        bg=COLORS["background"],
                        fg=COLORS["text"])
        title.pack(side="left")
        
        # Kontener z wynikami - z cieniem
        result_container = tk.Frame(self.right_frame, bg=COLORS["border"], relief="flat")
        result_container.pack(fill="both", expand=True, padx=15, pady=(0, 10))



        result_frame = tk.Frame(result_container, bg=COLORS["panel"], relief="flat")
        result_frame.pack(fill="both", expand=True, padx=2, pady=2)

        self.result_text = scrolledtext.ScrolledText(
            result_frame, 
            bg=COLORS["panel"], 
            fg=COLORS["text"], 
            font=("Consolas", 10),
            relief="flat",
            bd=0,
            padx=10,
            pady=10,
            wrap="word"
        )
        self.result_text.pack(fill="both", expand=True)
        
        # Tagi dla formatowania
        self.result_text.tag_configure("info", foreground=COLORS["success"], 
                                      font=("Consolas", 10, "bold"))
        self.result_text.tag_configure("error", foreground=COLORS["error"], 
                                      font=("Consolas", 10, "bold"))
        self.result_text.tag_configure("warning", foreground=COLORS["warning"], 
                                      font=("Consolas", 10, "bold"))
        self.result_text.tag_configure("bold", font=("Consolas", 11, "bold"))
        self.result_text.tag_configure("link", foreground=COLORS.get("accent", "#3498DB"), 
                                      underline=True, font=("Consolas", 10, "bold"))
        self.result_text.tag_configure("analyzing", foreground=COLORS.get("accent", "#3498DB"), 
                                      background="#E3F2FD", font=("Consolas", 14, "bold"))

        # Panel akcji na dole
        action_frame = tk.Frame(self.right_frame, bg=COLORS["background"])
        action_frame.pack(fill="x", padx=15, pady=(0, 15))
        
        # Styl przycisków akcji
        action_btn_style = {
            "relief": "flat",
            "bd": 0,
            "padx": 20,
            "pady": 10,
            "font": ("Segoe UI", 10, "bold"),
            "cursor": "hand2"
        }
        
        save_btn = tk.Button(action_frame, text="💾 Lista osób", 
                            command=lambda: save_found_people_to_xlsx(self.found_people), 
                            bg=COLORS.get("button_success", "#27AE60"), 
                            fg="white",
                            activebackground=COLORS.get("button_success_hover", "#229954"),
                            activeforeground="white",
                            **action_btn_style)
        save_btn.pack(side="left", padx=(0, 10))
        
        stats_btn = tk.Button(action_frame, text="📊 Pokaż statystyki", 
                            command=self.show_statistics, 
                            bg=COLORS.get("info", "#3498DB"), 
                            fg="white",
                            activebackground=COLORS.get("button_info_hover", "#2980B9"),
                            activeforeground="white",
                            **action_btn_style)
        stats_btn.pack(side="left", padx=(0, 10))
        
        export_all_btn = tk.Button(action_frame, text="📊 Pełny raport", 
                            command=self.export_all_results, 
                            bg=COLORS.get("warning", "#F39C12"), 
                            fg="white",
                            activebackground="#E67E22",
                            activeforeground="white",
                            **action_btn_style)
        export_all_btn.pack(side="left", padx=(0, 10))

    def restore_window_geometry(self):
        """Przywraca rozmiar i pozycję okna z zapisanych ustawień."""
        settings = load_settings()
        if settings and "window_geometry" in settings:
            try:
                geometry = settings["window_geometry"]
                # Update okna musi być najpierw wykonany
                self.root.update_idletasks()
                self.root.geometry(geometry)
            except Exception as e:
                logging.warning(f"Nie można przywrócić pozycji okna: {e}")

    def load_saved_settings(self):
        """Wczytuje zapisane ustawienia z poprzedniej sesji."""
        settings = load_settings()
        if settings:
            # Przywróć folder
            if "folder_path" in settings and os.path.exists(settings["folder_path"]):
                self.folder_path = settings["folder_path"]
            
            # Przywróć plik JSON
            if "json_file" in settings and os.path.exists(settings["json_file"]):
                self.json_file_path = settings["json_file"]
            
            # Przywróć parametry wieku
            if "age_from" in settings:
                try:
                    age_val = int(settings["age_from"])
                    if 0 <= age_val <= 150:
                        self.age_from_var.set(age_val)
                    else:
                        logging.warning(f"age_from poza zakresem: {age_val}, użyto domyślnej")
                        self.age_from_var.set(DEFAULT_AGE_FROM)
                except (ValueError, TypeError) as e:
                    logging.warning(f"Błąd wczytania age_from: {e}")
                    self.age_from_var.set(DEFAULT_AGE_FROM)
            
            if "age_to" in settings:
                try:
                    age_val = int(settings["age_to"])
                    if 0 <= age_val <= 150:
                        self.age_to_var.set(age_val)
                    else:
                        logging.warning(f"age_to poza zakresem: {age_val}, użyto domyślnej")
                        self.age_to_var.set(DEFAULT_AGE_TO)
                except (ValueError, TypeError) as e:
                    logging.warning(f"Błąd wczytania age_to: {e}")
                    self.age_to_var.set(DEFAULT_AGE_TO)
            
            # Przywróć dni jubileuszowe
            if "jubilee_days" in settings:
                try:
                    days_val = int(settings["jubilee_days"])
                    if 1 <= days_val <= 365:
                        self.jubilee_days_var.set(days_val)
                    else:
                        logging.warning(f"jubilee_days poza zakresem: {days_val}, użyto domyślnej")
                        self.jubilee_days_var.set(DEFAULT_JUBILEE_DAYS)
                except (ValueError, TypeError) as e:
                    logging.warning(f"Błąd wczytania jubilee_days: {e}")
                    self.jubilee_days_var.set(DEFAULT_JUBILEE_DAYS)
            
            # Przywróć zakres lat ślubów
            if "marriage_year_from" in settings:
                try:
                    year_val = int(settings["marriage_year_from"])
                    if 1800 <= year_val <= 2100:
                        self.marriage_year_from_var.set(year_val)
                except (ValueError, TypeError) as e:
                    logging.warning(f"Błąd wczytania marriage_year_from: {e}")
            
            if "marriage_year_to" in settings:
                try:
                    year_val = int(settings["marriage_year_to"])
                    if 1800 <= year_val <= 2100:
                        self.marriage_year_to_var.set(year_val)
                except (ValueError, TypeError) as e:
                    logging.warning(f"Błąd wczytania marriage_year_to: {e}")

    def initialize_names(self):
        """Wczytuje plik JSON z imionami z katalogu Excel."""
        # Sprawdź czy mamy folder z plikami Excel
        if self.folder_path and os.path.exists(self.folder_path):
            # Spróbuj wczytać imiona.json z katalogu z plikami Excel
            self.json_file_path = os.path.join(self.folder_path, "imiona.json")
            if os.path.exists(self.json_file_path):
                self.names_dict = load_names(self.json_file_path)
                self.result_text.insert(tk.END, f"[INFO] Wczytano plik JSON z imionami: {self.json_file_path}\n", "info")
            else:
                # Jeśli nie ma w katalogu Excel, utwórz pusty słownik
                self.names_dict = {}
                self.result_text.insert(tk.END, f"[INFO] Nie znaleziono imiona.json w katalogu {self.folder_path}\n", "warning")
                self.result_text.insert(tk.END, f"[INFO] Plik zostanie utworzony po dodaniu pierwszego imienia.\n", "info")
        else:
            # Jeśli nie ma jeszcze folderu, wczytaj z katalogu programu (fallback)
            import sys
            if getattr(sys, 'frozen', False):
                base_dir = os.path.dirname(sys.executable)
            else:
                base_dir = os.path.dirname(os.path.abspath(__file__))
            self.json_file_path = os.path.join(base_dir, "imiona.json")
            if os.path.exists(self.json_file_path):
                self.names_dict = load_names(self.json_file_path)
                self.result_text.insert(tk.END, f"[INFO] Wczytano domyślny plik JSON: {self.json_file_path}\n", "info")
            else:
                self.names_dict = {}
                self.result_text.insert(tk.END, f"[INFO] Nie znaleziono imiona.json. Wybierz folder z plikami Excel.\n", "warning")
        
        # Jeśli jest zapisany folder, automatycznie uruchom analizę
        if self.folder_path and os.path.exists(self.folder_path):
            self.result_text.insert(tk.END, f"[INFO] Wykryto poprzednio używany folder: {self.folder_path}\n", "info")
            self.result_text.insert(tk.END, f"[INFO] Rozpoczynam automatyczną analizę...\n", "info")
            # Uruchom analizę po pełnej inicjalizacji interfejsu
            self.root.after(500, lambda: self.analyze_current_settings(show_dialog=False))

    def select_folder(self):
        """Wybiera folder do analizy."""
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path = folder
            # Wczytaj imiona.json z wybranego folderu
            self.json_file_path = os.path.join(folder, "imiona.json")
            if os.path.exists(self.json_file_path):
                self.names_dict = load_names(self.json_file_path)
                self.result_text.insert(tk.END, f"[INFO] Załadowano imiona.json z folderu: {self.json_file_path}\n", "info")
            else:
                self.result_text.insert(tk.END, f"[INFO] Nie znaleziono imiona.json w folderze. Plik zostanie utworzony po dodaniu imion.\n", "warning")
            self.analyze_folder(self.folder_path, self.names_dict, jubilee_days=self.jubilee_days_var.get())

    def select_json_file(self):
        """Wybiera plik JSON z imionami."""
        json_path = filedialog.askopenfilename(filetypes=[("Pliki JSON", "*.json")])
        if json_path:
            self.json_file_path = json_path
            self.names_dict = load_names(json_path)
            self.result_text.insert(tk.END, f"[INFO] Załadowano nowy plik JSON: {json_path}\n", "info")
    
    def show_marriages_dialog(self):
        """Wyświetla dialog z wyszukiwaniem ślubów w zakresie lat."""
        from gui_dialogs import show_marriages_dialog
        show_marriages_dialog(self.folder_path)

    
    def on_closing(self):
        """Obsługuje zamknięcie okna - zapisuje ustawienia."""
        try:
            # Anuluj zaplanowane zadania
            if self.refresh_after_id is not None:
                try:
                    self.root.after_cancel(self.refresh_after_id)
                except (ValueError, tk.TclError):
                    pass
            
            # Zapisz rozmiar i pozycję okna
            window_geometry = self.root.geometry()
            
            settings = {
                "folder_path": self.folder_path,
                "json_file": self.json_file_path,
                "age_from": self.age_from_var.get(),
                "age_to": self.age_to_var.get(),
                "jubilee_days": self.jubilee_days_var.get(),
                "marriage_year_from": self.marriage_year_from_var.get(),
                "marriage_year_to": self.marriage_year_to_var.get(),
                "window_geometry": window_geometry
            }
            save_settings(settings)
        except Exception as e:
            logging.error(f"Błąd podczas zapisywania ustawień: {e}")
        finally:
            self.root.destroy()

    def _show_analysis_error(self, error_msg):
        """Wyświetla błąd analizy (wywoływane z wątku)."""
        try:
            messagebox.showerror("Błąd", f"Wystąpił błąd podczas analizy: {error_msg}")
        except Exception:
            pass

    def _restore_ui_after_analysis(self):
        """Przywraca UI po zakończeniu analizy (wywoływane z wątku)."""
        try:
            if self.btn_analyze:
                self.btn_analyze.config(state=tk.NORMAL)
            self.root.config(cursor="")
        except Exception:
            pass

    def open_sample_kartoteka(self):
        """Otwiera plik wzór.xlsx."""
        sample_name = "wzór.xlsx"
        if not self.folder_path:
            messagebox.showwarning("Brak wybranego folderu", "Wybierz najpierw folder z kartotek.")
            return
        sample_path = os.path.join(self.folder_path, sample_name)
        try:
            if not os.path.exists(sample_path):
                messagebox.showerror("Błąd", f"Nie znaleziono pliku wzór: {sample_path}")
                return
            os.startfile(sample_path)
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się otworzyć pliku: {e}")

    def search_in_results(self):
        """Wyszukuje tekst w wynikach. Enter przechodzi do następnego wystąpienia."""
        search_term = self.search_entry.get().strip()
        # Inicjalizacja atrybutów pomocniczych
        if not hasattr(self, '_last_search_term'):
            self._last_search_term = None
        if not hasattr(self, 'search_index'):
            self.search_index = None

        # Jeśli pole puste, usuń komunikat i resetuj indeksy
        if not search_term:
            self.result_text.delete("1.0", "2.0")
            self.result_text.tag_remove("highlight", "1.0", tk.END)
            self._last_search_term = None
            self.search_index = None
            return

        import re
        pattern = rf'\m{re.escape(search_term)}\M'  # \m i \M to granice słowa w Tkinter

        # Jeśli zmienił się tekst wyszukiwania, licz od nowa i podświetl pierwsze
        if self._last_search_term != search_term:
            # Zlicz wystąpienia
            count = 0
            temp_index = "1.0"
            while True:
                temp_index = self.result_text.search(pattern, temp_index, stopindex=tk.END, nocase=True, regexp=True)
                if not temp_index:
                    break
                count += 1
                temp_index = f"{temp_index}+{len(search_term)}c"
            self.result_text.delete("1.0", "2.0")
            if count == 0:
                self.result_text.insert("1.0", f"[WYSZUKIWANIE] Nie znaleziono: '{search_term}'\n", "searchinfo")
                self.result_text.tag_remove("highlight", "1.0", tk.END)
                self._last_search_term = search_term
                self.search_index = None
                return
            else:
                self.result_text.insert("1.0", f"[WYSZUKIWANIE] Znaleziono {count} wystąpień dla: '{search_term}'\n", "searchinfo")
            self.search_index = "1.0"
            self._last_search_term = search_term
        # Szukaj następnego wystąpienia od aktualnego indeksu
        idx = self.result_text.search(pattern, self.search_index, stopindex=tk.END, nocase=True, regexp=True)
        if not idx:
            # Jeśli nie znaleziono dalej, zacznij od początku
            idx = self.result_text.search(pattern, "1.0", stopindex=tk.END, nocase=True, regexp=True)
            if not idx:
                self.result_text.tag_remove("highlight", "1.0", tk.END)
                self.search_index = None
                return
        end_index = f"{idx}+{len(search_term)}c"
        self.result_text.tag_remove("highlight", "1.0", tk.END)
        self.result_text.tag_add("highlight", idx, end_index)
        self.result_text.tag_configure("highlight", background="yellow", foreground="black")
        self.result_text.see(idx)
        # Ustaw następny indeks na pozycję po znalezionym
        self.search_index = end_index
        # Styl komunikatu na górze
        self.result_text.tag_configure("searchinfo", foreground=COLORS.get("info", "#27AE60"), font=("Consolas", 10, "bold"))

    def apply_settings(self):
        """Stosuje ustawienia wieku i jubileuszy po naciśnięciu Enter."""
        if not self.folder_path or not self.names_dict:
            return
        
        try:
            age_from = self.age_from_var.get()
            age_to = self.age_to_var.get()
            jubilee_days = self.jubilee_days_var.get()
            marriage_year_from = self.marriage_year_from_var.get()
            marriage_year_to = self.marriage_year_to_var.get()
            
            if age_from < 0 or age_to < 0 or age_from > age_to:
                messagebox.showwarning("Błąd", "Podano nieprawidłowy zakres wieku (od musi być mniejsze lub równe do).")
                return
            
            if jubilee_days < 0:
                messagebox.showwarning("Błąd", "Liczba dni do jubileuszu musi być nieujemna.")
                return
            
            if marriage_year_from < 1800 or marriage_year_to > 2100 or marriage_year_from > marriage_year_to:
                messagebox.showwarning("Błąd", "Podano nieprawidłowy zakres lat ślubów (od musi być mniejsze lub równe do).")
                return
            
            # Wykonaj analizę z nowymi ustawieniami
            self.analyze_current_settings(show_dialog=False)
            
        except tk.TclError:
            messagebox.showerror("Błąd", "Wprowadź poprawne liczby całkowite w polach.")

    def analyze_current_settings(self, show_dialog=False):
        """Analizuje z bieżącymi ustawieniami."""
        print(f"[DEBUG] analyze_current_settings: wywołano (show_dialog={show_dialog}) o {datetime.now()}")
        if not self.folder_path:
            folder = filedialog.askdirectory(title="Wybierz folder z plikami Excel")
            if not folder:
                print("[DEBUG] analyze_current_settings: przerwano, brak folderu")
                return
            self.folder_path = folder

        try:
            age_from = self.age_from_var.get()
            age_to = self.age_to_var.get()
            jubilee_days = self.jubilee_days_var.get()

            if age_from < 0 or age_to < 0 or age_from > age_to:
                messagebox.showwarning("Błąd", "Podano nieprawidłowy zakres wieku.")
                return

            def worker():
                try:
                    self.analyze_folder(self.folder_path, self.names_dict, age_from, age_to, jubilee_days=jubilee_days, show_dialog=show_dialog)
                except Exception as e:
                    logging.error(f"Błąd w wątku analizy: {e}")
                    self.root.after(0, self._show_analysis_error, str(e))
                finally:
                    self.root.after(0, self._restore_ui_after_analysis)

            if self.btn_analyze:
                self.btn_analyze.config(state=tk.DISABLED)
            self.root.config(cursor="watch")
            self.root.update_idletasks()

            t = threading.Thread(target=worker, daemon=True)
            t.start()

        except Exception as e:
            messagebox.showerror("Błąd", f"Wystąpił błąd podczas uruchamiania analizy: {e}")

    def schedule_reanalysis(self, delay=500):
        """Harmonogramuje ponowną analizę."""
        # DEBUG: logowanie wywołania i flagi loading_settings
        print(f"[DEBUG] Wywołano schedule_reanalysis (delay={delay}), loading_settings={self.loading_settings}, czas={datetime.now()}")
        # Nie uruchamiaj podczas ładowania ustawień
        if self.loading_settings:
            print("[DEBUG] schedule_reanalysis przerwane: loading_settings=True")
            return
        if self.refresh_after_id is not None:
            try:
                self.root.after_cancel(self.refresh_after_id)
            except (ValueError, tk.TclError):
                pass  # ID już nieważne
        print("[DEBUG] Zaplanowano analyze_current_settings przez after")
        self.refresh_after_id = self.root.after(delay, lambda: self.analyze_current_settings(show_dialog=False))

    def show_statistics(self):
        """Wyświetla okno ze statystykami z kolorowymi elementami i filtrowaniem po imieniu/nazwisku."""
        stats_text = self.statistics.format_statistics()

        # Utwórz nowe okno dla statystyk
        stats_window = tk.Toplevel(self.root)
        stats_window.title("📊 Statystyki Analizy")
        stats_window.geometry("800x700")
        stats_window.configure(bg="#1E1E1E")  # Ciemne tło

        # Nagłówek z gradientem
        header = tk.Frame(stats_window, bg="#2C3E50", height=80)
        header.pack(fill="x", padx=0, pady=0)
        header.pack_propagate(False)

        title = tk.Label(header, text="📊 STATYSTYKI ANALIZY KARTOTEKI", 
                        font=("Segoe UI", 18, "bold"),
                        bg="#2C3E50",
                        fg="#ECF0F1")
        title.pack(pady=25)


        # Obszar tekstowy ze statystykami - ciemny motyw
        text_frame = tk.Frame(stats_window, bg="#34495E", relief="flat")
        text_frame.pack(fill="both", expand=True, padx=15, pady=(15, 10))

        text_inner = tk.Frame(text_frame, bg="#2C3E50", relief="flat")
        text_inner.pack(fill="both", expand=True, padx=3, pady=3)

        stats_text_widget = scrolledtext.ScrolledText(
            text_inner,
            bg="#1E1E1E",
            fg="#ECF0F1",
            font=("Consolas", 10),
            relief="flat",
            bd=0,
            padx=15,
            pady=15,
            wrap="word",
            insertbackground="#ECF0F1"
        )
        stats_text_widget.pack(fill="both", expand=True)
        # Przyciski akcji
        bottom_frame = tk.Frame(stats_window, bg="#1E1E1E")
        bottom_frame.pack(fill="x", side="bottom", pady=(0, 15))
        buttons_frame = tk.Frame(bottom_frame, bg="#1E1E1E")
        buttons_frame.pack()
        # Przycisk zapisz pełny raport (osoby + statystyki + jubileusze + śluby)
        save_all_btn = tk.Button(buttons_frame, text="💾 Eksportuj do Excel",
                            command=lambda: export_all_results_to_excel(
                                self.found_people,
                                self.statistics,
                                self.jubilees_found,
                                self.marriages_in_range,
                                self.all_unknown,
                                default_name="pelny_raport.xlsx"
                            ),
                            bg="#27AE60",
                            fg="white",
                            font=("Segoe UI", 11, "bold"),
                            relief="flat",
                            bd=0,
                            padx=20,
                            pady=12,
                            cursor="hand2",
                            activebackground="#229954",
                            activeforeground="white")
        save_all_btn.pack(side="left", padx=(0, 10))

        # Przycisk zamknij - bardziej widoczny
        close_btn = tk.Button(buttons_frame, text="✖ Zamknij",
                            command=stats_window.destroy,
                            bg="#E74C3C",
                            fg="white",
                            font=("Segoe UI", 11, "bold"),
                            relief="flat",
                            bd=0,
                            padx=30,
                            pady=12,
                            cursor="hand2",
                            activebackground="#C0392B",
                            activeforeground="white")
        close_btn.pack(side="left")

        # Wstaw tekst z kolorowaniem
        self._insert_colored_statistics(stats_text_widget, stats_text)
        stats_text_widget.config(state=tk.DISABLED)

        # (Usunięto duplikat przycisków akcji - wszystko jest już w bottom_frame powyżej)
    
    def _insert_colored_statistics(self, text_widget, stats_text):
        """Wstawia tekst statystyk z kolorowaniem."""
        lines = stats_text.split('\n')
        
        # Konfiguruj tagi kolorów
        text_widget.tag_configure("header", foreground="#3498DB", font=("Consolas", 12, "bold"))
        text_widget.tag_configure("section", foreground="#E67E22", font=("Consolas", 11, "bold"))
        text_widget.tag_configure("number", foreground="#2ECC71", font=("Consolas", 10, "bold"))
        text_widget.tag_configure("bar_filled", foreground="#27AE60")
        text_widget.tag_configure("bar_empty", foreground="#34495E")
        text_widget.tag_configure("emoji", foreground="#F39C12")
        text_widget.tag_configure("border", foreground="#7F8C8D")
        text_widget.tag_configure("percentage", foreground="#9B59B6")
        text_widget.tag_configure("highlight", foreground="#E74C3C", font=("Consolas", 10, "bold"))
        
        for line in lines:
            # Nagłówki sekcji
            if "STATYSTYKI ANALIZY" in line:
                text_widget.insert(tk.END, line + '\n', "header")
            elif any(emoji in line for emoji in ["👥", "📂", "🎂", "💍", "📊", "📍", "🎉", "⚠️", "⏱️"]):
                # Linia z emoji - sekcja
                text_widget.insert(tk.END, line + '\n', "section")
            elif "─" in line or "│" in line or "┌" in line or "└" in line or "├" in line or "╔" in line or "╚" in line or "║" in line or "═" in line:
                # Ramki
                text_widget.insert(tk.END, line + '\n', "border")
            elif "[" in line and "]" in line and "█" in line:
                # Linia z paskiem
                parts = line.split('[')
                text_widget.insert(tk.END, parts[0])
                
                if len(parts) > 1:
                    bar_and_rest = parts[1].split(']')
                    bar = bar_and_rest[0]
                    
                    # Koloruj wypełniony pasek
                    filled = bar.count('█')
                    text_widget.insert(tk.END, '[')
                    text_widget.insert(tk.END, '█' * filled, "bar_filled")
                    text_widget.insert(tk.END, '░' * (len(bar) - filled), "bar_empty")
                    text_widget.insert(tk.END, ']')
                    
                    if len(bar_and_rest) > 1:
                        rest = bar_and_rest[1]
                        # Koloruj procenty
                        if '%' in rest:
                            before_percent = rest.split('%')[0]
                            after_percent = '%' + rest.split('%')[1] if len(rest.split('%')) > 1 else '%'
                            text_widget.insert(tk.END, before_percent, "percentage")
                            text_widget.insert(tk.END, after_percent, "percentage")
                        else:
                            text_widget.insert(tk.END, rest)
                text_widget.insert(tk.END, '\n')
            elif any(keyword in line for keyword in ["Średnia", "Mediana", "Najmłodszy", "Najstarszy"]):
                # Statystyki wiekowe - podświetl
                text_widget.insert(tk.END, line + '\n', "highlight")
            else:
                # Zwykły tekst
                text_widget.insert(tk.END, line + '\n')
    
    def export_all_results(self):
        """Eksportuje wszystkie wyniki (osoby + statystyki) do Excel."""
        if not self.found_people:
            messagebox.showwarning("Brak danych", "Brak osób do zapisania. Wykonaj najpierw analizę.")
            return
        export_all_results_to_excel(
            self.found_people, 
            self.statistics, 
            self.jubilees_found, 
            self.marriages_in_range,
            self.all_unknown
        )
    
    def analyze_folder(self, selected_folder, names_dict_local, age_from=None, age_to=None, jubilee_days=None, show_dialog=False):
        """Główna funkcja analizy folderów."""
        self.folder_path = selected_folder
        self.found_people.clear()
        self.statistics.start_analysis()  # Rozpocznij zbieranie statystyk

        if not names_dict_local:
            self.result_text.insert(tk.END, "[ERROR] Lista imion jest pusta. Sprawdź plik JSON.\n", "error")
            return

        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, "\n\n   ⌛ TRWA ANALIZA... PROSZĘ CZEKAĆ ⌛\n\n", "analyzing")
        self.result_text.update_idletasks()
        # Bufor na szczegóły analizy
        analysis_details = []
        total_m = 0
        total_k = 0
        error_count = 0
        warning_count = 0
        all_unknown = {}
        unknown_names = set()  # Zbiór nieznanych imion
        scanned_files_count = 0
        jubilees_found = []
        marriages_in_range = []  # Lista ślubów w zakresie lat

        if age_from is None:
            age_from = self.age_from_var.get()
        if age_to is None:
            age_to = self.age_to_var.get()
        if jubilee_days is None:
            jubilee_days = self.jubilee_days_var.get()


        for filename in os.listdir(self.folder_path):
            if filename.startswith("~$") or not filename.lower().endswith((".xls", ".xlsx")):
                continue
            file_path = os.path.join(self.folder_path, filename)
            # Pomijaj plik wzór.xlsx/wzor.xlsx w statystykach, ale wypisz w wynikach
            is_template = filename.lower() in ["wzór.xlsx", "wzor.xlsx"]
            analysis_details.append((f"[INFO] Analiza pliku: {filename}\n", "link", file_path))
            if is_template:
                # Wypisz, ale nie licz w statystykach
                analysis_details.append((f"  [INFO] Plik {filename} jest wzorcem i nie jest liczony w statystykach.\n", "info", None))
                # Dodaj podsumowanie pliku do bufora z zerami
                analysis_details.append((f"  [INFO] Wynik dla pliku: Kobiety=0, Mężczyźni=0\n", "bold", None))
                analysis_details.append(("-"*50 + "\n", None, None))
                continue
            scanned_files_count += 1
            self.statistics.add_file()  # Zlicz plik
            # ...existing code...

            try:
                xl = pd.ExcelFile(file_path)
            except Exception as e:
                analysis_details.append((f"[ERROR] Nie można wczytać pliku: {e}\n", "error", None))
                error_count += 1
                continue

            file_m = 0
            file_k = 0

            for sheet_name in xl.sheet_names:
                try:
                    df = xl.parse(sheet_name, header=None)
                except Exception as e:
                    analysis_details.append((f"[ERROR] Nie można wczytać arkusza {sheet_name}: {e}\n", "error", None))
                    error_count += 1
                    self.statistics.add_error()  # Zlicz błąd
                    continue

                self.statistics.add_sheet()  # Zlicz arkusz
                analysis_details.append((f"  [INFO] Analiza arkusza: {sheet_name}\n", "info", None))

                # Wyodrębnij nazwisko, adres i inne dane
                surname = ""
                try:
                    if df.shape[0] >= 7 and df.shape[1] >= 2:
                        surname_parts = []
                        for col in range(0, 2):
                            column_data = df.iloc[1:7, col].dropna()
                            if not column_data.empty:
                                surname_parts.append(str(column_data.iloc[0]))
                        surname = " ".join(surname_parts)
                        if surname:
                            analysis_details.append((f"    [INFO] Nazwisko: {surname}\n", "info", None))
                except Exception as e:
                    logging.debug(f"Nie można wyodrębnić nazwiska z {filename}: {e}")
                    surname = ""

                address = ""
                try:
                    if df.shape[0] >= 4 and df.shape[1] >= 5:
                        address_parts = []
                        for col in range(2, 5):
                            column_data = df.iloc[1:4, col].dropna()
                            if not column_data.empty:
                                address_parts.append(str(column_data.iloc[0]))
                        address = " ".join(address_parts)
                        if address:
                            analysis_details.append((f"    [INFO] Adres: {address}\n", "info", None))
                except Exception as e:
                    logging.debug(f"Nie można wyodrębnić adresu z {filename}: {e}")

                old_address = ""
                try:
                    if df.shape[0] >= 4 and df.shape[1] >= 7:
                        old_address_parts = []
                        for col in range(5, 7):
                            column_data = df.iloc[1:4, col].dropna()
                            if not column_data.empty:
                                old_address_parts.append(str(column_data.iloc[0]))
                        old_address = " ".join(old_address_parts)
                        if old_address:
                            analysis_details.append((f"    [INFO] Adres stary: {old_address}\n", "info", None))
                except Exception as e:
                    logging.debug(f"Nie można wyodrębnić starego adresu z {filename}: {e}")

                # Wyodrębnij i wyświetl informacje o ślubie małżonków
                try:
                    marriage_info = extract_marriage_info(df)
                    if marriage_info.get("husband") and marriage_info.get("wife"):
                        analysis_details.append((f"    [INFO] Mąż: {marriage_info['husband']}, Żona: {marriage_info['wife']}\n", "info", None))
                    if marriage_info.get("marriage_date"):
                        analysis_details.append((f"    [INFO] Data ślubu: {marriage_info['marriage_date']}\n", "info", None))
                        
                        # Dodaj rok ślubu do statystyk
                        try:
                            from datetime import datetime
                            marriage_year = datetime.fromisoformat(marriage_info['marriage_date']).year
                            self.statistics.add_marriage_year(marriage_year)
                        except Exception:
                            pass
                        
                        # Sprawdź czy ślub jest w zakresie lat
                        try:
                            from datetime import datetime
                            marriage_year = datetime.fromisoformat(marriage_info['marriage_date']).year
                            if self.marriage_year_from_var.get() <= marriage_year <= self.marriage_year_to_var.get():
                                marriages_in_range.append({
                                    "surname": surname,
                                    "husband": marriage_info['husband'],
                                    "wife": marriage_info['wife'],
                                    "date": marriage_info['marriage_date'],
                                    "year": marriage_year,
                                    "address": address,
                                    "old_address": old_address,
                                    "file_path": file_path
                                })
                        except Exception:
                            pass
                except Exception as e:
                    logging.debug(f"Nie można wyodrębnić danych małżonków z {filename}: {e}")

                # Wyodrębnij i wyświetl informacje o ślubie dziadków
                try:
                    gp_marriage_info = extract_grandparents_marriage_info(df)
                    gp_marriage_date = gp_marriage_info.get("marriage_date") if isinstance(gp_marriage_info, dict) else gp_marriage_info
                    if gp_marriage_date:
                        analysis_details.append((f"    [INFO] Data ślubu dziadków: {gp_marriage_date}\n", "info", None))
                        
                        # Dodaj rok ślubu dziadków do statystyk
                        try:
                            from datetime import datetime
                            gp_marriage_year = datetime.fromisoformat(gp_marriage_date).year
                            self.statistics.add_marriage_year(gp_marriage_year)
                        except Exception:
                            pass
                        
                        # Sprawdź czy ślub dziadków jest w zakresie lat
                        try:
                            from datetime import datetime
                            gp_marriage_year = datetime.fromisoformat(gp_marriage_date).year
                            if self.marriage_year_from_var.get() <= gp_marriage_year <= self.marriage_year_to_var.get():
                                marriages_in_range.append({
                                    "surname": surname,
                                    "husband": "Dziadek",
                                    "wife": "Babcia",
                                    "date": gp_marriage_date,
                                    "year": gp_marriage_year,
                                    "address": address,
                                    "old_address": old_address,
                                    "file_path": file_path,
                                    "type": "DZIADKOWIE"
                                })
                        except Exception:
                            pass
                except Exception as e:
                    logging.debug(f"Nie można wyodrębnić danych dziadków z {filename}: {e}")

                try:
                    mar_list = analyze_marriage_jubilees(df, file_path, surname, window_days=jubilee_days, marriage_year_from=self.marriage_year_from_var.get(), marriage_year_to=self.marriage_year_to_var.get())
                    for j in mar_list:
                        j["old_address"] = old_address
                        self.statistics.add_jubilee()  # Zlicz jubileusz
                    jubilees_found.extend(mar_list)

                    gp_list = analyze_grandparents_jubilees(df, file_path, surname, window_days=jubilee_days, marriage_year_from=self.marriage_year_from_var.get(), marriage_year_to=self.marriage_year_to_var.get())
                    for j in gp_list:
                        j["old_address"] = old_address
                        self.statistics.add_jubilee()  # Zlicz jubileusz
                    jubilees_found.extend(gp_list)
                except Exception as e:
                    logging.error(f"BŁĄD analizy jubileuszy w {filename}: {e}")
                    analysis_details.append((f"    [ERROR] Nie można przeanalizować jubileuszy: {str(e)}\n", "error", None))
                    error_count += 1
                    self.statistics.add_error()  # Zlicz błąd

                sheet_people = []
                for row_idx, row in df.iterrows():
                    if len(row) < 3:
                        continue
                    name_cell = row.iloc[1]
                    birth_cell = row.iloc[2]

                    tokens = extract_words(name_cell)
                    if not tokens:
                        continue

                    given_name = tokens[0]
                    second_member = tokens[1] if len(tokens) > 1 else None

                    # Sprawdź czy brak daty urodzenia
                    if pd.isna(birth_cell) or (isinstance(birth_cell, str) and birth_cell.strip() == ""):
                        analysis_details.append((f"    [WARNING] Brak daty urodzenia dla '{given_name}'\n", "warning", None))
                        warning_count += 1
                        self.statistics.add_warning()  # Zlicz ostrzeżenie
                        continue

                    # Sprawdź czy w komórce jest wzorzec daty, ale nie jest poprawny
                    if pd.notna(birth_cell) and isinstance(birth_cell, str):
                        date_pattern = re.search(r"(\d{1,2})[./-](\d{1,2})[./-](\d{4})", birth_cell)
                        if date_pattern:
                            from data_processing import validate_date_components
                            d, m, y = date_pattern.groups()
                            is_valid, error_msg = validate_date_components(d, m, y)
                            if not is_valid:
                                # Specjalna obsługa daty umownej 99/99/9999: warning zamiast error
                                if str(d) == '99' and str(m) == '99' and str(y) == '9999':
                                    analysis_details.append((f"    [WARNING] Błędna data dla '{given_name}': {birth_cell} - {error_msg}\n", "warning", None))
                                    warning_count += 1
                                    self.statistics.add_warning()  # Zlicz ostrzeżenie
                                else:
                                    analysis_details.append((f"    [ERROR] Błędna data dla '{given_name}': {birth_cell} - {error_msg}\n", "error", None))
                                    error_count += 1
                                    self.statistics.add_error()  # Zlicz błąd

                    birth_date = extract_birth_date(birth_cell)
                    if not birth_date:
                        analysis_details.append((f"    [WARNING] Nie można odczytać daty urodzenia dla '{given_name}': {birth_cell}\n", "warning", None))
                        warning_count += 1
                        self.statistics.add_warning()  # Zlicz ostrzeżenie
                        continue
                    
                    # Dodaj rok urodzenia do statystyk
                    try:
                        birth_year = birth_date.year
                        self.statistics.add_birth_year(birth_year)
                    except Exception:
                        pass

                    # Specjalna obsługa daty umownej: przypisz medianę wieku lub 40
                    if birth_date == "MEDIANA_WIEKU":
                        # Wylicz medianę z dotychczasowych osób
                        ages_so_far = self.statistics.ages
                        if ages_so_far:
                            sorted_ages = sorted(ages_so_far)
                            count = len(sorted_ages)
                            if count % 2 == 0:
                                med = (sorted_ages[count // 2 - 1] + sorted_ages[count // 2]) / 2
                            else:
                                med = sorted_ages[count // 2]
                            age = int(round(med))
                        else:
                            age = 40
                    else:
                        age = calculate_age(birth_date)
                    if age is None or not (age_from <= age <= age_to):
                        continue

                    normalized_word = remove_diacritics(given_name.strip().lower())
                    if normalized_word in names_dict_local:
                        plec = names_dict_local[normalized_word]

                        if plec == "K":
                            file_k += 1
                        elif plec == "M":
                            file_m += 1

                        final_surname = format_person_name(second_member) if second_member else surname

                        person_entry = {
                            "imie": given_name,
                            "nazwisko": final_surname,
                            "adres": address,
                            "old_address": old_address,
                            "wiek": age,
                            "plec": plec,
                            "file": filename,
                            "file_path": file_path
                        }
                        self.found_people.append(person_entry)
                        sheet_people.append(person_entry)
                        self.statistics.add_person(person_entry)  # Dodaj osobę do statystyk
                    else:
                        # Nieznane imię - dodaj do zbioru i zapisz ostrzeżenie
                        if normalized_word not in unknown_names:
                            unknown_names.add(normalized_word)
                            analysis_details.append((f"    [WARNING] Nieznane imię '{given_name}' - dodaj do słownika imiona.json\n", "warning", None))
                            warning_count += 1
                            self.statistics.add_warning()  # Zlicz ostrzeżenie
                            self.statistics.add_unknown_name()  # Zlicz nieznane imię
                        
                        # Dodaj lokalizację nieznanegoienia do słownika all_unknown
                        location_key = f"{file_path} -> {sheet_name}"
                        if normalized_word not in all_unknown:
                            all_unknown[normalized_word] = []
                        if location_key not in all_unknown[normalized_word]:
                            all_unknown[normalized_word].append(location_key)

                # Wyświetl znalezione osoby z arkusza
                if sheet_people:
                    analysis_details.append((f"    Znalezione osoby w arkuszu {sheet_name}:\n", "info", None))
                    for person in sheet_people:
                        imie = format_person_name(person.get("imie", ""))
                        nazwisko = format_person_name(person.get("nazwisko", ""))
                        wiek = person.get("wiek", "")
                        plec = person.get("plec", "")
                        plec_str = "K" if plec == "K" else "M"
                        analysis_details.append((f"      - {imie} {nazwisko}, wiek: {wiek}, płeć: {plec_str}\n", None, None))

            total_k += file_k
            total_m += file_m
            
            # Dodaj podsumowanie pliku do bufora z podświetleniem zer
            analysis_details.append((f"  [INFO] Wynik dla pliku: Kobiety=", "bold", None))
            analysis_details.append((f"{file_k}", "error" if file_k == 0 else "bold", None))
            analysis_details.append((f", Mężczyźni=", "bold", None))
            analysis_details.append((f"{file_m}", "error" if file_m == 0 else "bold", None))
            analysis_details.append(("\n", "bold", None))
            analysis_details.append(("-"*50 + "\n", None, None))

        # Przygotuj podsumowanie końcowe
        final_summary_lines = []
        final_summary_lines.append(f"[RESULT] Przeskanowano plików: {scanned_files_count}\n")
        
        if all_unknown:
            final_summary_lines.append(f"[RESULT] Wszystkie nieznane imiona wyświetlono w wynikach.\n")

        # Zakończ zbieranie statystyk
        self.statistics.update_family_stats(self.found_people)
        self.statistics.end_analysis()
        
        # Usuń komunikat analizy i wstaw wyniki
        self.result_text.delete(1.0, tk.END)
        
        # Wstaw początek podsumowania
        for line in final_summary_lines:
            self.result_text.insert(tk.END, line, "bold")
        
        # Wstaw sumę z podświetleniem zer
        self.result_text.insert(tk.END, "[RESULT] Suma wszystkich plików: ", "bold")
        self.result_text.insert(tk.END, "Kobiety=", "bold")
        self.result_text.insert(tk.END, f"{total_k}", "error" if total_k == 0 else "bold")
        self.result_text.insert(tk.END, ", Mężczyźni=", "bold")
        self.result_text.insert(tk.END, f"{total_m}", "error" if total_m == 0 else "bold")
        self.result_text.insert(tk.END, ", Razem=", "bold")
        self.result_text.insert(tk.END, f"{total_k + total_m}", "bold")
        self.result_text.insert(tk.END, "\n", "bold")
        
        # Wstaw informacje o błędach i ostrzeżeniach NA CZERWONO
        if error_count > 0 or warning_count > 0:
            self.result_text.insert(tk.END, "[WARNING] Znaleziono: ", "warning")
            if error_count > 0:
                self.result_text.insert(tk.END, f"{error_count} błędów", "error")
            if error_count > 0 and warning_count > 0:
                self.result_text.insert(tk.END, ", ", "bold")
            if warning_count > 0:
                self.result_text.insert(tk.END, f"{warning_count} ostrzeżeń", "warning")
            self.result_text.insert(tk.END, "\n", "bold")
        
        self.result_text.insert(tk.END, "=" * 50 + "\n\n", "bold")
        
        # Wstaw sekcję jubileuszy
        self.result_text.insert(tk.END, "-" * 60 + "\n", "bold")
        if jubilees_found:
            self.result_text.insert(tk.END, f"[JUBILEUSZE ŚLUBÓW – NAJBLIŻSZE {jubilee_days} DNI]\n", "bold")
            for j in sorted(jubilees_found, key=lambda x: x["days"]):
                type_str = j.get("type", "MAŁŻONKOWIE").upper()
                old = j.get("old_address", "")
                name_paren = f"{j.get('surname','')}" + (f", {old}" if old else "")
                self.result_text.insert(tk.END, f"{j['date']} – {j['years']} lat – {j['husband']} i {j['wife']} ({name_paren}) [{type_str}] za {j['days']} dni\n")
        else:
            self.result_text.insert(tk.END, f"[INFO] Nie znaleziono nadchodzących jubileuszy w ciągu {jubilee_days} dni.\n", "info")
        self.result_text.insert(tk.END, "-" * 60 + "\n\n", "bold")
        
        # Wstaw szczegóły analizy z bufora
        file_link_counter = 0
        for detail_text, tag, link_path in analysis_details:
            if tag:
                self.result_text.insert(tk.END, detail_text, tag)
            else:
                self.result_text.insert(tk.END, detail_text)
            
            # Jeśli to link do pliku, dodaj obsługę kliknięcia
            if link_path:
                file_link_counter += 1
                start_index = self.result_text.index(f"{tk.END} - 2 lines")
                end_index = self.result_text.index(f"{tk.END} - 1 lines")
                self.result_text.tag_add(f"file-link-{file_link_counter}", start_index, end_index)
                self.result_text.tag_configure(f"file-link-{file_link_counter}", foreground="blue", underline=True)
                self.result_text.tag_bind(f"file-link-{file_link_counter}", "<Button-1>", lambda e, path=link_path: os.startfile(path))

        # Usuwanie starego przycisku edycji jeśli istnieje
        if hasattr(self, 'edit_unknown_btn') and self.edit_unknown_btn and self.edit_unknown_btn.winfo_exists():
            self.edit_unknown_btn.destroy()
            self.edit_unknown_btn = None
        
        if all_unknown:
            # Dodaj separator przed przyciskiem w wynikach
            self.result_text.insert(tk.END, "\n" + "="*60 + "\n", "bold")
            self.result_text.insert(tk.END, "⚠️ UWAGA: Znaleziono nieznane imiona!\n", "warning")
            self.result_text.insert(tk.END, f"Liczba nieznanych imion: {len(all_unknown)}\n\n", "info")
            
            for name, locations in all_unknown.items():
                self.result_text.insert(tk.END, f"  • {name} ({len(locations)} wystąpień)\n")
            
            self.result_text.insert(tk.END, "\n👉 Przewiń panel boczny w dół i kliknij przycisk 'Edytuj nieznane imiona'\n", "info")
            self.result_text.insert(tk.END, "="*60 + "\n\n", "bold")
            
            if hasattr(self, 'analysis_section'):
                # Styl przycisku edycji
                edit_btn_style = {
                    "relief": "flat",
                    "bd": 0,
                    "padx": 12,
                    "pady": 8,
                    "font": ("Segoe UI", 9, "bold"),
                    "cursor": "hand2",
                    "bg": COLORS.get("warning", "#F39C12"),
                    "fg": "white",
                    "activebackground": COLORS.get("button_warning_hover", "#D68910"),
                    "activeforeground": "white"
                }
                
                self.edit_unknown_btn = tk.Button(
                    self.analysis_section, 
                    text=f"✏️ Edytuj nieznane ({len(all_unknown)})", 
                    command=lambda: edit_unknown_name(all_unknown, self.names_dict, self.folder_path, self.result_text, self),
                    **edit_btn_style
                )
                self.edit_unknown_btn.pack(fill="x", pady=(6, 0))
                
                # Aktualizuj region przewijania canvas po dodaniu przycisku
                if hasattr(self, 'configure_scroll_region'):
                    self.left_frame.update_idletasks()  # Wymusza aktualizację geometrii
                    self.configure_scroll_region()
                
                # Przewiń w dół, aby przycisk był widoczny
                self.root.after(100, lambda: self.left_canvas.yview_moveto(1.0))  # Przewiń na sam dół z opóźnieniem
        
        # Zapisz wszystkie dane z analizy do atrybutów klasy
        self.jubilees_found = jubilees_found
        self.marriages_in_range = marriages_in_range
        self.all_unknown = all_unknown
        self.analysis_details = analysis_details


        # Resetuj stan wyszukiwania po każdej analizie
        self.reset_search_state()
        if show_dialog:
            show_results_dialog(self.found_people, self.root)
            self.reset_search_state()

        if self.btn_analyze:
            self.btn_analyze.config(state=tk.NORMAL)
        self.root.config(cursor="")
        self.root.update_idletasks()
