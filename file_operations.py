import json
import os
from data_processing import remove_diacritics

def load_names(json_file):
    """Wczytuje plik JSON z imionami."""
    try:
        with open(json_file, "r", encoding="utf-8") as f:
            data = json.load(f)
            if not isinstance(data, dict) or not all(isinstance(k, str) and isinstance(v, str) for k, v in data.items()):
                raise ValueError("Nieprawidłowa struktura JSON.")
            return {
                remove_diacritics(k.strip().lower()): v.upper()
                for k, v in data.items()
            }
    except FileNotFoundError:
        return {}
    except (json.JSONDecodeError, ValueError):
        return {}

def save_names_to_json(names_dict, filepath):
    """Zapisuje słownik imion do pliku JSON."""
    try:
        if os.path.exists(filepath):
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)
        else:
            data = {}
        data.update(names_dict)
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False
