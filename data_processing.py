import re
import unicodedata
import pandas as pd
from datetime import datetime

def remove_diacritics(text):
    """Usuwa polskie znaki diakrytyczne."""
    if not isinstance(text, str):
        return text
    return ''.join(c for c in unicodedata.normalize('NFKD', text) if not unicodedata.combining(c)).lower()

def format_person_name(name):
    """Formatuje imię/nazwisko z wielkimi literami."""
    if not name:
        return ""
    name = str(name).strip().lower()
    parts = re.split(r'([- ])', name)
    return "".join(p.capitalize() if p not in "- " else p for p in parts)

def extract_words(cell_value):
    """Ekstrakcja słów z komórki."""
    if pd.isna(cell_value):
        return []
    if isinstance(cell_value, str):
        tokens = re.findall(r'\b[\w-]+\b', cell_value.lower())
        tokens = [t.strip() for t in tokens if t and not t.isnumeric()]
        return tokens
    if isinstance(cell_value, (int, float)):
        return []
    if isinstance(cell_value, pd.Timestamp):
        return extract_words(str(cell_value.date()))
    return []

def validate_date_components(day, month, year):
    """Sprawdza czy komponenty daty są poprawne."""
    try:
        d, m, y = int(day), int(month), int(year)
        if not (1 <= m <= 12):
            # Specjalna obsługa daty umownej 99/99/9999
            if str(d) == '99' and str(m) == '99' and str(y) == '9999':
                return False, "Data umowna 99/99/9999 – przypisany wiek: mediana populacji. Osoba liczona w statystykach."
            return False, f"Nieprawidłowy miesiąc: {m}"
        if not (1 <= d <= 31):
            return False, f"Nieprawidłowy dzień: {d}"
        if y < 1800 or y > datetime.today().year + 1:
            return False, f"Nieprawidłowy rok: {y}"
        # Sprawdź czy data jest prawidłowa dla danego miesiąca/roku
        try:
            datetime(y, m, d)
            return True, None
        except ValueError as e:
            return False, f"Nieprawidłowa data {d}.{m}.{y}: {str(e)}"
    except (ValueError, TypeError):
        return False, "Błąd konwersji komponentów daty"

def normalize_date(value):
    """Normalizuje daty do formatu ISO."""
    if pd.isna(value):
        return None

    if isinstance(value, pd.Timestamp):
        return value.date().isoformat()

    if isinstance(value, str):
        text = value.lower()
        text = re.sub(r"\bś?lub\b", "", text).strip()
        m = re.search(r"(\d{1,2})[./-](\d{1,2})[./-](\d{4})", text)
        if m:
            d, mth, y = m.groups()
            is_valid, error_msg = validate_date_components(d, mth, y)
            if not is_valid:
                return None
            try:
                date_obj = datetime(int(y), int(mth), int(d))
                return date_obj.date().isoformat()
            except ValueError:
                return None
    return None

def extract_birth_date(cell):
    """Wyciąga datę urodzenia z komórki."""
    if isinstance(cell, (pd.Timestamp, datetime)):
        return cell.date()
    if isinstance(cell, str):
        m = re.search(r"(\d{1,2})[./-](\d{1,2})[./-](\d{4})", cell)
        if m:
            d, mth, y = m.groups()
            # Specjalna obsługa daty umownej
            if str(d) == '99' and str(mth) == '99' and str(y) == '9999':
                print("[INFO] Data 99/99/9999: przypisano wiek równy medianie populacji. Osoba liczona w statystykach.")
                return "MEDIANA_WIEKU"
            is_valid, error_msg = validate_date_components(d, mth, y)
            if not is_valid:
                return None
            try:
                return datetime(int(y), int(mth), int(d)).date()
            except ValueError:
                return None
    return None

def calculate_age(birth_date):
    """Oblicza wiek na podstawie daty urodzenia."""
    today = datetime.today().date()
    # Jeśli birth_date to liczba (wiek), zwróć ją bez zmian
    if isinstance(birth_date, (int, float)):
        return int(birth_date)
    # Jeśli birth_date to string "MEDIANA_WIEKU", zwróć None (nie powinno się zdarzyć, ale na wszelki wypadek)
    if isinstance(birth_date, str):
        return None
    # Standardowo licz wiek z daty
    return today.year - birth_date.year - (
        (today.month, today.day) < (birth_date.month, birth_date.day)
    )

def extract_number_from_text(s):
    """Wyciąga pierwszą liczbę z tekstu."""
    if not s:
        return None
    m = re.search(r'(\d+)', str(s))
    return int(m.group(1)) if m else None
