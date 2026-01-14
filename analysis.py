import re
import pandas as pd
from datetime import datetime
from data_processing import normalize_date

def get_upcoming_jubilees(marriage_date_str, surname, husband, wife, jub_type="MAŁŻONKOWIE", window_days=30, marriage_year_from=1900, marriage_year_to=2100):
    """Sprawdza nadchodzące jubileusze ślubu."""
    jubilees = []
    if not marriage_date_str:
        return jubilees

    try:
        marriage_date = datetime.fromisoformat(marriage_date_str).date()
    except Exception:
        return jubilees
    
    # Filtruj po roku ślubu
    if marriage_date.year < marriage_year_from or marriage_date.year > marriage_year_to:
        return jubilees

    today = datetime.today().date()

    try:
        anniversary_this_year = marriage_date.replace(year=today.year)
    except ValueError:
        try:
            anniversary_this_year = marriage_date.replace(year=today.year, day=28)
        except Exception:
            return jubilees

    years = today.year - marriage_date.year
    delta_days = (anniversary_this_year - today).days

    try:
        max_days = int(window_days)
    except Exception:
        max_days = 30

    if years in [10, 20, 25, 30, 40, 50, 60, 70] and 0 <= delta_days <= max_days:
        jubilees.append({
            "years": years,
            "date": anniversary_this_year.isoformat(),
            "surname": surname,
            "husband": husband,
            "wife": wife,
            "days": delta_days,
            "type": jub_type
        })

    return jubilees

def analyze_marriage_jubilees(sheet, file_path, surname="", window_days=30, marriage_year_from=1900, marriage_year_to=2100):
    """Analizuje jubileusze ślubu małżonków."""
    jubilees = []
    try:
        if sheet.shape[0] >= 10 and sheet.shape[1] >= 4:
            husband = sheet.iloc[8, 1] if not pd.isna(sheet.iloc[8, 1]) else None
            wife = sheet.iloc[9, 1] if not pd.isna(sheet.iloc[9, 1]) else None
            marriage_date_husband = sheet.iloc[8, 3] if not pd.isna(sheet.iloc[8, 3]) else None
            marriage_date_wife = sheet.iloc[9, 3] if not pd.isna(sheet.iloc[9, 3]) else None
            if husband and wife:
                marriage_date = marriage_date_husband or marriage_date_wife
                if isinstance(marriage_date, (pd.Timestamp, datetime)):
                    marriage_date = marriage_date.date().isoformat()
                elif isinstance(marriage_date, str):
                    marriage_date = marriage_date.split()[0]
                else:
                    marriage_date = None
                if marriage_date:
                    jubilees = get_upcoming_jubilees(marriage_date, surname, husband, wife, jub_type="MAŁŻONKOWIE", window_days=window_days, marriage_year_from=marriage_year_from, marriage_year_to=marriage_year_to)
    except Exception:
        pass
    return jubilees

def analyze_grandparents_jubilees(sheet, file_path, surname="", window_days=30, marriage_year_from=1900, marriage_year_to=2100):
    """Analizuje jubileusze ślubu dziadków."""
    jubilees = []
    try:
        start_col = 4
        end_col = min(sheet.shape[1] - 1, 17)
        regex_slub = re.compile(r'(?:ślub|slub)?\s*[:\-]?\s*(\d{1,2}[./-]\d{1,2}[./-]\d{4})', re.IGNORECASE)
        dziadek_idx = None
        babcia_idx = None
        for row_idx in range(len(sheet)):
            cell_D = str(sheet.iloc[row_idx, 3]).lower() if pd.notna(sheet.iloc[row_idx, 3]) else ""
            if "dziadek" in cell_D and "†" not in cell_D and "zm." not in cell_D:
                dziadek_idx = row_idx
            elif "babcia" in cell_D and "†" not in cell_D and "zm." not in cell_D:
                babcia_idx = row_idx
        if dziadek_idx is not None and babcia_idx is not None:
            marriage_date_gp = None
            for row_idx in range(max(0, dziadek_idx - 2), min(len(sheet), babcia_idx + 3)):
                for col in range(start_col, end_col + 1):
                    val = sheet.iloc[row_idx, col]
                    if pd.notna(val):
                        if isinstance(val, str):
                            m = regex_slub.search(val)
                            if m:
                                date_part = m.group(1)
                                marriage_date_gp = normalize_date(date_part)
                                if marriage_date_gp:
                                    break
                        else:
                            marriage_date_gp = normalize_date(val)
                            if marriage_date_gp:
                                break
                if marriage_date_gp:
                    break
            if marriage_date_gp:
                jubilees = get_upcoming_jubilees(marriage_date_gp, surname, "Dziadek", "Babcia", jub_type="DZIADKOWIE", window_days=window_days, marriage_year_from=marriage_year_from, marriage_year_to=marriage_year_to)
    except Exception:
        pass
    return jubilees

def extract_marriage_info(sheet):
    """Zwraca informacje o ślubie małżonków."""
    try:
        if sheet.shape[0] >= 10 and sheet.shape[1] >= 4:
            husband = sheet.iloc[8, 1] if not pd.isna(sheet.iloc[8, 1]) else None
            wife = sheet.iloc[9, 1] if not pd.isna(sheet.iloc[9, 1]) else None
            marriage_date = sheet.iloc[8, 3] if not pd.isna(sheet.iloc[8, 3]) else (sheet.iloc[9, 3] if not pd.isna(sheet.iloc[9, 3]) else None)
            if isinstance(marriage_date, (pd.Timestamp, datetime)):
                marriage_date = marriage_date.date().isoformat()
            elif isinstance(marriage_date, str):
                m = re.search(r"(\d{1,2}[./-]\d{1,2}[./-]\d{4})", marriage_date)
                marriage_date = normalize_date(m.group(1)) if m else None
            else:
                marriage_date = None
            return {"husband": husband, "wife": wife, "marriage_date": marriage_date}
    except Exception:
        pass
    return {"husband": None, "wife": None, "marriage_date": None}

def extract_grandparents_marriage_info(sheet):
    """Próbuje znaleźć datę ślubu dziadków."""
    try:
        start_col = 4
        end_col = min(sheet.shape[1] - 1, 17)
        regex_slub = re.compile(r'(?:ślub|slub)?\s*[:\-]?\s*(\d{1,2}[./-]\d{1,2}[./-]\d{4})', re.IGNORECASE)
        dziadek_idx = None
        babcia_idx = None
        for row_idx in range(len(sheet)):
            cell_D = str(sheet.iloc[row_idx, 3]).lower() if pd.notna(sheet.iloc[row_idx, 3]) else ""
            if "dziadek" in cell_D and "†" not in cell_D and "zm." not in cell_D:
                dziadek_idx = row_idx
            elif "babcia" in cell_D and "†" not in cell_D and "zm." not in cell_D:
                babcia_idx = row_idx
        if dziadek_idx is not None and babcia_idx is not None:
            for row_idx in range(max(0, dziadek_idx - 2), min(len(sheet), babcia_idx + 3)):
                for col in range(start_col, end_col + 1):
                    val = sheet.iloc[row_idx, col]
                    if pd.notna(val):
                        if isinstance(val, str):
                            m = regex_slub.search(val)
                            if m:
                                date_part = m.group(1)
                                marriage_date_gp = normalize_date(date_part)
                                if marriage_date_gp:
                                    return marriage_date_gp
                        else:
                            marriage_date_gp = normalize_date(val)
                            if marriage_date_gp:
                                return marriage_date_gp
    except Exception:
        pass
    return None
