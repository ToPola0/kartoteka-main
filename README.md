# Kartoteka Parafialna - System Zarządzania v3.1

![Status](https://img.shields.io/badge/status-active-success.svg)
![Version](https://img.shields.io/badge/version-2.0-blue.svg)
![Python](https://img.shields.io/badge/python-3.13+-blue.svg)

Profesjonalny program do zarządzania i analizy kartoteki parafialnej z nowoczesnym interfejsem graficznym.

## ✨ Nowe funkcje w wersji 2.0

### 🎨 Profesjonalny wygląd
- **Bootlogo** - elegancki ekran powitalny z paskiem ładowania podczas startu
- **Nowoczesny interfejs** - przeprojektowany UI z ikonami i emoji
- **Lepsze fonty** - Segoe UI dla lepszej czytelności
- **Niestandardowe ikony** - logo.ico zamiast domyślnej ikony Python

### 🔧 Nowe funkcje
- **Dialog ślubów** - osobne okno do wyszukiwania ślubów w zakresie lat
- **Export do Excel** - zapis wyników ślubów z auto-dopasowanymi kolumnami
- **Modułowa architektura** - łatwa do rozbudowy struktura
- **Lepsze logowanie** - szczegółowe logi błędów
- **📊 Statystyki** - szczegółowe statystyki analizy z:
  - Dekadami urodzin (od najstarszych)
  - Dekadami ślubów (od najstarszych)
  - Statystykami wieku (średnia, mediana, najmłodszy, najstarszy)
  - Wizualizacją graficzną (kolorowe paski)
  - Czasem analizy i wydajnością
  - Ciemnym motywem z kolorowym tekstem
  - **Eksport statystyk do Excel** - wieloarkuszowy plik z wszystkimi danymi
- **imiona.json** - automatyczne zapisywanie i wczytywanie z katalogu z plikami Excel
- **Export wszystkich wyników** - jeden przycisk zapisuje osoby + statystyki do Excel

## 🎯 Główne funkcje

- ✅ Wyszukiwanie osób według wieku
- ✅ Analiza jubileuszy małżeńskich (50, 60, 65 lat)
- ✅ Analiza jubileuszy dziadków (90, 95, 100 lat)
- ✅ Walidacja dat i wykrywanie błędów
- ✅ **Export wyników do Excel** - z auto-dopasowaniem szerokości kolumn
- ✅ **Export statystyk do Excel** - wieloarkuszowy raport ze WSZYSTKIMI danymi:
  - Podsumowanie (osoby, pliki, adresy, czasy)
  - Statystyki wieku (średnia, mediana, min, max, rozstęp)
  - Grupy wiekowe z procentami
  - Urodziny w dekadach
  - Śluby w dekadach
- ✅ **Export wszystkich wyników** - kompletny raport w jednym pliku:
  - Znalezione osoby (wszystkie pola: imię, nazwisko, adresy, wiek, płeć, plik)
  - Jubileusze (data, lata małżeństwa, małżonkowie, typ, dni do jubileuszu)
  - Śluby w zakresie lat (rok, data, małżonkowie, adres)
  - Nieznane imiona (nazwa, lokalizacja, liczba wystąpień)
  - Wszystkie statystyki
- ✅ Graficzny interfejs użytkownika (Tkinter)
- ✅ Wyszukiwanie w wynikach
- ✅ Edycja nieznanych imion
- ✅ Automatyczne zapisywanie ustawień
- ✅ **📊 Zaawansowane statystyki**:
  - Rozkład wieku i płci
  - **Statystyki wiekowe**: średnia, mediana, najmłodszy, najstarszy
  - Urodziny w dekadach (chronologicznie)
  - Śluby w dekadach (chronologicznie)
  - Liczba plików i arkuszy
  - Błędy i ostrzeżenia
  - Czas analizy i wydajność
  - **Kolorowy interfejs** z ciemnym motywem
  - **Export do Excel** z formatowaniem i wieloma arkuszami

## 📋 Wymagania

- Python 3.13+
- pandas
- openpyxl
- Pillow

## 🚀 Instalacja

1. **Sklonuj repozytorium:**
```bash
git clone https://github.com/[TWOJ_USERNAME]/kartoteka.git
cd kartoteka
```

2. **Utwórz środowisko wirtualne:**
```bash
python -m venv .venv
.venv\Scripts\activate  # Windows
```

3. **Zainstaluj zależności:**
```bash
pip install pandas openpyxl Pillow
```

## ▶️ Uruchomienie

### Metoda 1: Python
```bash
python main.py
```

### Metoda 2: Plik wsadowy (Windows)
```bash
Uruchom.bat
```

### Metoda 3: Środowisko wirtualne
```bash
.venv\Scripts\python.exe main.py
```

## 📦 Kompilacja do EXE

Aby utworzyć standalone aplikację (.exe):

```bash
pip install pyinstaller
pyinstaller --name="Kartoteka" --windowed --icon="logo.ico" --add-data="imiona.json;." --add-data="logo przeżroczyste.png;." main.py
```

Skompilowany program znajdziesz w folderze `dist/Kartoteka/`

## 📁 Struktura projektu

```
Kartoteka/
├── main.py                  # Punkt wejścia aplikacji
├── splash_screen.py         # Ekran powitalny (bootlogo)
├── themes.py               # System motywów kolorystycznych
├── gui_main.py             # Główne okno aplikacji
├── gui_dialogs.py          # Okna dialogowe
├── analysis.py             # Logika analizy jubileuszy
├── data_processing.py      # Przetwarzanie i walidacja danych
├── file_operations.py      # Operacje na plikach
├── config.py               # Konfiguracja i ustawienia
├── imiona.json             # Słownik imion i płci
├── settings.json           # Zapisane ustawienia użytkownika
└── README.md               # Ten plik
```

## 🔍 Funkcje szczegółowe

### Wyszukiwanie osób
- Filtrowanie według przedziału wiekowego
- Walidacja dat urodzenia
- Wykrywanie błędnych dat (np. 33.1.1970)
- Wyszukiwanie w wynikach po imieniu, nazwisku lub adresie

### Analiza jubileuszy
- Jubileusze małżeńskie: 50, 60, 65 lat
- Jubileusze dziadków: 90, 95, 100 lat
- Automatyczne wyliczanie nadchodzących jubileuszy (konfigurowalne dni)
- Wykrywanie błędów w datach ślubu

### Wyniki
- Szczegółowy raport z każdego pliku Excel
- Podsumowanie błędów i ostrzeżeń
- Kolorowe podświetlenie błędów i ostrzeżeń
- Suma łączna (Kobiety + Mężczyźni = Razem)
- Klikalne linki do plików Excel (otwieranie w systemie)

### Export
- **Zapis znalezionych osób do Excel** - tylko lista osób z 8 polami (imię, nazwisko, adresy, wiek, płeć, plik, ścieżka)
- **Zapis WSZYSTKIEGO do Excel** - kompletny eksport obejmujący do 10 arkuszy:
  1. **Znalezione osoby** - pełna lista z wszystkimi danymi
  2. **Podsumowanie** - podstawowe liczby (974 osoby, pliki, adresy, błędy, czas)
  3. **Statystyki wieku** - średnia, mediana, najmłodszy, najstarszy, rozstęp
  4. **Grupy wiekowe** - 6 grup z liczbami i procentami (0-17, 18-30, 31-50, 51-70, 71-90, 90+)
  5. **Urodziny w dekadach** - od najstarszych do najnowszych (1960s-2020s) z liczbami i procentami
  6. **Śluby w dekadach** - chronologicznie (1950s-2020s) z liczbami i procentami
  7. **Jubileusze** - nadchodzące jubileusze z pełnymi danymi
  8. **Śluby w zakresie lat** - wszystkie śluby w wybranym okresie
  9. **Nieznane imiona** - lista nierozpoznanych imion z lokalizacjami
  10. **Adresy** - statystyki adresowe
- **Wszystkie arkusze mają:**
  - Auto-dopasowanie szerokości kolumn do zawartości
  - Profesjonalne formatowanie (kolorowe nagłówki, obramowania)
  - Pełne dane widoczne bez przewijania
- **Sortowanie:** według wieku, adresu, nazwiska, alfabetycznie
- **Zachowanie starych adresów** we wszystkich eksportach

### Statystyki (okno z ciemnym motywem)
- **Kolorowe wyświetlanie** z podświetlaniem składni:
  - Niebieskie nagłówki sekcji
  - Zielone liczby i paski wypełnienia
  - Fioletowe procenty
  - Czerwone statystyki wiekowe (średnia, mediana, min, max)
- **Pełne dane widoczne w oknie:**
  - Urodziny w dekadach (od najstarszych) z paskami procentowymi
  - Śluby w dekadach (chronologicznie) z paskami procentowymi
  - Rozkład wieku: średnia, mediana, najmłodszy, najstarszy
  - 6 grup wiekowych z paskami i procentami
  - Statystyki adresów (unikalne, średnio osób na adres)
  - Jubileusze i śluby w zakresie
  - Problemy (błędy, ostrzeżenia, nieznane imiona)
  - Czas analizy i średni czas na plik
- **Przycisk eksportu w oknie statystyk:** "💾 Zapisz WSZYSTKO do Excel" - eksportuje pełne dane do Excela z wszystkimi arkuszami

### Śluby
- Osobne okno dialogowe "💍 Śluby w latach..."
- Wyszukiwanie ślubów małżonków i dziadków w zakresie lat
- Export wyników do Excel z auto-dopasowaniem kolumn
- Klikalne linki do kartotek

## 📝 Changelog

### v2.5 (Styczeń 2026)
- 💍 Usunięto wyświetlanie ślubów z głównej analizy
- 🔍 Dodano dialog wyszukiwania ślubów w zakresie lat
- 📊 Export ślubów do Excel z auto-szerokością kolumn
- 🖼️ Niestandardowe logo (logo.ico) zamiast ikony Python
- 🎨 Przeprojektowano interfejs użytkownika
- ✨ Dodano splash screen przy starcie
- 🔧 Ulepszono modularność kodu
- 💅 Dodano ikony i emoji w interfejsie
- 🐛 Naprawiono błąd zamykania okna wyników

### v3.1 (Styczeń 2026)
- Pierwsza wersja z GUI
- Podstawowe funkcje analizy

### v1.1 (Styczeń 2026)
- Pierwsza wersja stabilna
- Podstawowe funkcje analizy

## 🐛 Zgłaszanie błędów

W razie problemów:
1. Sprawdź plik `kartoteka_errors.log`
2. Upewnij się, że wszystkie zależności są zainstalowane
3. Sprawdź czy używasz Python 3.13+

## 📜 Licencja

Projekt prywatny - Parafia Przyborów

---

**© 2026 Parafia Przyborów | Wersja 2.5**
