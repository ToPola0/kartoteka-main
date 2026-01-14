@echo off
echo ====================================================
echo   Kompilacja Kartoteka Parafialna do EXE
echo ====================================================
echo.

echo Kompilowanie programu...
echo.

REM Kompiluj do EXE w trybie folderu (onedir) aby pliki były dostępne
REM Używamy logo.png zamiast nazwy z polskimi znakami
python -m PyInstaller --onedir --windowed --name="Kartoteka" --icon="logo.ico" --add-data="imiona.json;." --add-data="logo.png;." --add-data="logo przeźroczyste.png;." --add-data="logo.ico;." --hidden-import=statistics --hidden-import=export_statistics --hidden-import=analysis --hidden-import=data_processing --hidden-import=file_operations --hidden-import=gui_dialogs --hidden-import=gui_main --hidden-import=splash_screen --hidden-import=config main.py

echo.
echo ====================================================
echo   Tworzenie folderu Kartoteka_Build...
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
    copy "dist\Kartoteka\Kartoteka.exe" "Kartoteka_Build\" >nul
    echo   OK: Kartoteka.exe skopiowany
    
    REM Kopiuj folder _internal
    if exist "dist\Kartoteka\_internal" (
        xcopy "dist\Kartoteka\_internal" "Kartoteka_Build\_internal\" /E /I /Y >nul
        echo   OK: Folder _internal skopiowany
    )
) else (
    echo   BLAD: Nie znaleziono dist\Kartoteka\Kartoteka.exe
    echo   Kompilacja mogla sie nie udac!
    pause
    exit /b 1
)

REM Kopiuj wszystkie potrzebne pliki
echo Kopiowanie plikow konfiguracyjnych...
copy "imiona.json" "Kartoteka_Build\" >nul
echo   OK: imiona.json

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
if not exist "Kartoteka_Build\settings.json" (
    echo {"folder_path": "", "age_from": 0, "age_to": 150} > "Kartoteka_Build\settings.json"
    echo   OK: settings.json utworzony
)

REM Kopiuj wszystkie moduły Python (dla pewności)
echo Kopiowanie modulow Python...
copy "statistics.py" "Kartoteka_Build\" >nul 2>nul
copy "export_statistics.py" "Kartoteka_Build\" >nul 2>nul
copy "analysis.py" "Kartoteka_Build\" >nul 2>nul
copy "data_processing.py" "Kartoteka_Build\" >nul 2>nul
copy "file_operations.py" "Kartoteka_Build\" >nul 2>nul
copy "gui_dialogs.py" "Kartoteka_Build\" >nul 2>nul
copy "gui_main.py" "Kartoteka_Build\" >nul 2>nul
copy "splash_screen.py" "Kartoteka_Build\" >nul 2>nul
copy "config.py" "Kartoteka_Build\" >nul 2>nul
echo   OK: Moduly Python skopiowane (opcjonalne, dla edycji)

REM Kopiuj skrypt uruchamiający
(
echo @echo off
echo cd /d "%%~dp0"
echo Kartoteka.exe
) > "Kartoteka_Build\Uruchom.bat"
echo   OK: Uruchom.bat utworzony

REM Utwórz plik README
echo ====================================== > "Kartoteka_Build\README.txt"
echo   KARTOTEKA PARAFIALNA - INSTRUKCJA >> "Kartoteka_Build\README.txt"
echo ====================================== >> "Kartoteka_Build\README.txt"
echo. >> "Kartoteka_Build\README.txt"
echo 1. Uruchom program klikajac: Uruchom.bat >> "Kartoteka_Build\README.txt"
echo    LUB bezposrednio: Kartoteka.exe >> "Kartoteka_Build\README.txt"
echo. >> "Kartoteka_Build\README.txt"
echo 2. Pliki do edycji: >> "Kartoteka_Build\README.txt"
echo    - imiona.json (slownik imion i plci) >> "Kartoteka_Build\README.txt"
echo    - settings.json (ustawienia - tworzone automatycznie) >> "Kartoteka_Build\README.txt"
echo. >> "Kartoteka_Build\README.txt"
echo 3. Folder "Excel" zostanie utworzony automatycznie >> "Kartoteka_Build\README.txt"
echo    przy pierwszym zapisie wynikow. >> "Kartoteka_Build\README.txt"
echo. >> "Kartoteka_Build\README.txt"

echo.
echo ====================================================
echo   Kompilacja zakonczona!
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
