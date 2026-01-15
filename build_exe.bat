REM === AUTOMATYCZNE CZYSZCZENIE STARYCH BUILDÓW ===
echo Czyszczenie starych buildów i plików .py z podkatalogów...
if exist Kartoteka_Build (
    rmdir /S /Q Kartoteka_Build
)
if exist dist (
    rmdir /S /Q dist
)
if exist build (
    rmdir /S /Q build
)
if exist Kartoteka.spec del /Q /F Kartoteka.spec >nul 2>nul
echo   OK: Usunięto stare buildy, dist, build, spec
@echo off
echo ====================================================
echo   Kompilacja Kartoteka Parafialna v3.2 do EXE
echo ====================================================
echo.

echo Kompilowanie programu...
echo.

REM Kompiluj do EXE w trybie folderu (onedir) aby pliki były dostępne
REM Używamy logo.png zamiast nazwy z polskimi znakami
python -m PyInstaller --onedir --windowed --name="Kartoteka" --icon="logo.ico" --add-data="imiona.json;." --add-data="logo.png;." --add-data="logo.ico;." --hidden-import=statistics --hidden-import=export_statistics --hidden-import=analysis --hidden-import=data_processing --hidden-import=file_operations --hidden-import=gui_dialogs --hidden-import=gui_main --hidden-import=splash_screen --hidden-import=config --hidden-import=numpy --hidden-import=numpy.core._methods --hidden-import=numpy.lib.format --collect-all numpy main.py

echo.
echo ====================================================
echo   Tworzenie folderu Kartoteka_Build v3.2...
echo ====================================================

REM Usuń stary folder jeśli istnieje
if exist "Kartoteka_Build" (
    echo Usuwam stary folder Kartoteka_Build...
    rmdir /S /Q "Kartoteka_Build"
)

REM Utwórz nowy folder
mkdir "Kartoteka_Build"


REM Kopiuj plik EXE i cały folder _internal
echo Kopiowanie Kartoteka...
if exist "dist\Kartoteka\Kartoteka.exe" (
    copy "dist\Kartoteka\Kartoteka.exe" "Kartoteka_Build" >nul
    echo   OK: Kartoteka.exe skopiowany
    
    REM Kopiuj folder _internal
    if exist "dist\Kartoteka\_internal" (
        xcopy "dist\Kartoteka\_internal" "Kartoteka_Build\_internal" /E /I /Y >nul
        echo   OK: Folder _internal skopiowany
    )
) else (
    echo   BLAD: Nie znaleziono dist\Kartoteka\Kartoteka.exe
    echo   Kompilacja mogla sie nie udac!
    pause
    exit /b 1
)

REM Kopiuj logo (prosta nazwa bez polskich znaków)
copy "logo.png" "Kartoteka_Build\" >nul
if exist "Kartoteka_Build\logo.png" (
    echo   OK: logo.png
) else (
    echo   OSTRZEZENIE: logo.png nie skopiowane
)

REM Kopiuj logo przeźroczyste
copy "logo przeźroczyste.png" "Kartoteka_Build\" >nul
if exist "Kartoteka_Build\logo przeźroczyste.png" (
    echo   OK: logo przeźroczyste.png
) else (
    echo   OSTRZEZENIE: logo przeźroczyste.png nie skopiowane
)

REM Kopiuj ikonę
copy "logo.ico" "Kartoteka_Build\" >nul 2>nul
if exist "Kartoteka_Build\logo.ico" (
    echo   OK: logo.ico
)

REM Utwórz przykładowy plik settings.json jeśli nie istnieje
echo {"folder_path": "", "age_from": 0, "age_to": 150} > "Kartoteka_Build\settings.json"
echo   OK: settings.json utworzony (nadpisano)



REM Kopiuj skrypt uruchamiający
(
echo @echo off
echo cd /d "%%~dp0"
echo Kartoteka.exe
) > "Kartoteka_Build\Uruchom.bat"
echo   OK: Uruchom.bat utworzony

REM Utwórz plik README
echo ====================================== > "Kartoteka_Build\README.txt"
echo   KARTOTEKA PARAFIALNA - INSTRUKCJA OBSŁUGI >> "Kartoteka_Build\README.txt"
echo ====================================== >> "Kartoteka_Build\README.txt"
echo. >> "Kartoteka_Build\README.txt"
echo 1. URUCHOMIENIE PROGRAMU: >> "Kartoteka_Build\README.txt"
echo    - Kliknij Uruchom.bat lub bezposrednio Kartoteka.exe >> "Kartoteka_Build\README.txt"
echo. >> "Kartoteka_Build\README.txt"
echo 2. PIERWSZE KROKI: >> "Kartoteka_Build\README.txt"
echo    - Wybierz folder z kartotekami (pliki Excel) >> "Kartoteka_Build\README.txt"
echo    - Program zapamięta lokalizację i utworzy plik settings.json >> "Kartoteka_Build\README.txt"
echo    - Słownik imion (imiona.json) powinien znajdować się w folderze z kartotekami. >> "Kartoteka_Build\README.txt"
echo. >> "Kartoteka_Build\README.txt"
echo 3. MOŻLIWOŚCI PROGRAMU: >> "Kartoteka_Build\README.txt"
echo    - Analiza i przeglądanie danych osobowych z kartotek parafialnych (Excel) >> "Kartoteka_Build\README.txt"
echo    - Wyszukiwanie osób po imieniu, nazwisku, adresie >> "Kartoteka_Build\README.txt"
echo    - Filtrowanie według wieku, płci, jubileuszy >> "Kartoteka_Build\README.txt"
echo    - Eksport wyników do pliku Excel >> "Kartoteka_Build\README.txt"
echo    - Automatyczne wykrywanie i edycja nieznanych imion >> "Kartoteka_Build\README.txt"
echo    - Statystyki, wykrywanie błędów w danych, jubileusze >> "Kartoteka_Build\README.txt"
echo    - Obsługa ciemnego motywu >> "Kartoteka_Build\README.txt"
echo. >> "Kartoteka_Build\README.txt"
echo 4. EDYCJA PLIKÓW: >> "Kartoteka_Build\README.txt"
echo    - imiona.json: słownik imion i płci (możesz edytować w notatniku) >> "Kartoteka_Build\README.txt"
echo    - settings.json: ustawienia programu (lokalizacja kartotek, zakresy wieku itp.) >> "Kartoteka_Build\README.txt"
echo. >> "Kartoteka_Build\README.txt"
echo 5. UWAGI: >> "Kartoteka_Build\README.txt"
echo    - Folder "Excel" z wynikami zostanie utworzony automatycznie przy pierwszym eksporcie. >> "Kartoteka_Build\README.txt"
echo    - Program nie wymaga instalacji - wystarczy skopiować Kartoteka_Build na dowolny komputer z Windows. >> "Kartoteka_Build\README.txt"
echo    - W razie problemów sprawdź plik kartoteka_errors.log lub skontaktuj się z autorem. >> "Kartoteka_Build\README.txt"
echo. >> "Kartoteka_Build\README.txt"

echo.
echo ====================================================
echo   Kompilacja v3.2 zakonczona!
echo ====================================================
echo.
echo Folder: Kartoteka_Build\
echo   - Kartoteka.exe (program)
echo   - Uruchom.bat (szybkie uruchamianie)
echo   - imiona.json (slownik imion - mozesz edytowac)
echo   - settings.json (ustawienia)
echo   - logo.png (logo)
echo   - logo przeźroczyste.png (logo przeźroczyste)
echo   - logo.ico (ikona)
echo   - README.txt (instrukcja)
echo.
echo Mozesz teraz skopiowac caly folder Kartoteka_Build
echo na inny komputer lub do innej lokalizacji.
echo.
pause
