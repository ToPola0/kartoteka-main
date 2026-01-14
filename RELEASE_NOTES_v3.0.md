# 🎉 Kartoteka Parafialna v3.0 - Obsługa Nieznanych Dat Urodzenia

## 📋 Nowe Funkcje

### 🎂 Obsługa dat 99/99/9999 jako "data nieznana"
- Osoby z datą urodzenia **99/99/9999** są teraz **prawidłowo liczone** w statystykach
- Program wyświetla **ostrzeżenie** o nieznanej dacie przy każdej takiej osobie
- Wiek jest **automatycznie obliczany** jako mediana wieku z całej parafii

### 📊 Dynamiczny wiek domyślny (mediana)
- Zamiast stałego wieku 50 lat, program używa **mediany wieku** z analizowanych danych
- Mediana jest obliczana **tylko na podstawie osób z prawidłowymi datami urodzenia**
- **System dwuetapowy** zapewnia spójność - wszystkie osoby z 99/99/9999 dostają ten sam wiek

### 📈 Ulepszona sekcja mediany w wynikach
- Sekcja mediany **przeniesiona na koniec raportu** (po wszystkich szczegółach)
- Pokazuje **liczbę osób z nieznanymi datami** urodzenia
- Wyświetla **ostateczną wartość mediany** użytą w obliczeniach

### 🔍 Ulepszona funkcja wyszukiwania
- **Klawisz Enter** uruchamia wyszukiwanie
- **Poprzednie wyniki automatycznie czyszczone** przy nowym wyszukiwaniu
- Wyniki pojawiają się w logicznym miejscu (po podsumowaniu, przed jubileuszami)

### 🪟 Poprawione zapisywanie pozycji okna
- Okno otwiera się **dokładnie tam gdzie zostało zamknięte**
- **Osobno zapisywany stan maksymalizacji** okna
- Pełna **synchronizacja z menedżerem okien Windows**

## 🐛 Poprawki Błędów

- ✅ Naprawiono błąd z **niekonsystentną medianą** dla osób z 99/99/9999
- ✅ Poprawiono **pozycjonowanie okna** przy ponownym uruchomieniu
- ✅ Usunięto **błędy indentacji** w kodzie wyszukiwania

## 🔧 Zmiany Techniczne

- Python **3.13.2**
- PyInstaller **6.12.0**
- Pillow **12.1.0**
- NumPy z **pełnym zestawem zależności** (collect_all)
- Implementacja **dwuetapowego systemu obliczania mediany**
- Ulepszone **zarządzanie geometrią okna** z zapisem stanu

## 📥 Instalacja

1. Pobierz plik **Kartoteka_v3.0_Release.zip**
2. Rozpakuj do dowolnego folderu
3. Uruchom **Kartoteka.exe** lub **Uruchom.bat**
4. Wybierz folder z plikami Excel parafii

## 💻 Wymagania

- Windows 10/11 (64-bit)
- **Brak konieczności instalacji Pythona** ani innych programów
- Wszystkie biblioteki dołączone w folderze `_internal`

## 🆘 Wsparcie

W razie problemów sprawdź plik **kartoteka_errors.log** w folderze programu.

---

**Rozmiar archiwum:** ~46 MB  
**Wersja:** 3.0  
**Data wydania:** 10 stycznia 2026  
**Logo:** Św. Jadwiga - Patronka Parafii Przyborów
