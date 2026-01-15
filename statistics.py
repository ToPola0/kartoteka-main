"""
Moduł do generowania statystyk z kartoteki parafialnej.
"""
from datetime import datetime
from collections import defaultdict


class Statistics:
    def get_family_age_ranges(self, found_people):
        """Zwraca zakres wieku (min/max) dla każdej kategorii rodzin (arkuszy)."""
        # Założenie: found_people zawiera 'file_path' i 'wiek' dla każdej osoby
        from collections import defaultdict
        family_ages = defaultdict(list)
        for person in found_people:
            fp = person.get('file_path')
            wiek = person.get('wiek')
            if fp and wiek is not None:
                try:
                    family_ages[fp].append(int(wiek))
                except Exception:
                    pass
        # Podział na kategorie
        ranges = {1: [], 2: [], 3: [], 4: []}
        for ages in family_ages.values():
            size = len(ages)
            if size == 1:
                ranges[1].append((min(ages), max(ages)))
            elif size == 2:
                ranges[2].append((min(ages), max(ages)))
            elif 3 <= size <= 4:
                ranges[3].append((min(ages), max(ages)))
            elif size >= 5:
                ranges[4].append((min(ages), max(ages)))
        # Zwróć min/max dla każdej kategorii
        def minmax(lst):
            if not lst:
                return None, None
            mins = [x[0] for x in lst]
            maxs = [x[1] for x in lst]
            return min(mins), max(maxs)
        return {
            'family_1': minmax(ranges[1]),
            'family_2': minmax(ranges[2]),
            'family_3_4': minmax(ranges[3]),
            'family_5plus': minmax(ranges[4])
        }

    def add_family_by_size(self, size):
        """Dodaje rodzinę (arkusz) do odpowiedniej kategorii wielkości."""
        if size == 1:
            self.family_count_1 += 1
        elif size == 2:
            self.family_count_2 += 1
        elif 3 <= size <= 4:
            self.family_count_3_4 += 1
        elif size >= 5:
            self.family_count_5plus += 1
    
    def __init__(self):
        self.names_counter = defaultdict(int)
        self.reset()
    
    def reset(self):
        """Resetuje wszystkie statystyki."""
        self.total_people = 0
        self.total_males = 0
        self.total_females = 0
        self.files_scanned = 0
        self.sheets_scanned = 0
        self.errors_count = 0
        self.warnings_count = 0
        self.unknown_names_count = 0
        self.jubilees_count = 0
        self.marriages_in_range_count = 0
        self.age_distribution = defaultdict(int)  # Przedziały wiekowe
        self.addresses = set()  # Unikalne adresy
        self.analysis_start_time = None
        self.analysis_end_time = None
        self.people_by_age_group = {
            "0-17": 0,
            "18-30": 0,
            "31-50": 0,
            "51-70": 0,
            "71-90": 0,
            "90+": 0
        }
        # Dekady urodzin i ślubów
        self.birth_decades = defaultdict(int)  # Dekady urodzin (np. 1940, 1950)
        self.marriage_decades = defaultdict(int)  # Dekady ślubów
        self.ages = []  # Lista wszystkich wieków do obliczeń statystycznych
        self.names_counter.clear()  # Czyszczenie licznika imion
        # Podział rodzin
        self.family_count_1 = 0
        self.family_count_2 = 0
        self.family_count_3_4 = 0
        self.family_count_5plus = 0
    
    def start_analysis(self):
        """Rozpoczyna pomiar czasu analizy."""
        self.reset()
        self.analysis_start_time = datetime.now()
    
    def end_analysis(self):
        """Kończy pomiar czasu analizy i zlicza rodziny po adresach."""
        self.analysis_end_time = datetime.now()
        # Nie nadpisuj liczników rodzin, są ustawiane przez update_family_stats
    
    def get_analysis_duration(self):
        """Zwraca czas trwania analizy."""
        if self.analysis_start_time and self.analysis_end_time:
            duration = self.analysis_end_time - self.analysis_start_time
            return duration.total_seconds()
        return 0
    
    def add_person(self, person):
        """Dodaje osobę do statystyk."""
        wiek = person.get("wiek")
        # Jeśli wiek to znacznik "MEDIANA_WIEKU", wylicz medianę z dotychczasowych osób
        if wiek == "MEDIANA_WIEKU":
            # Jeśli nie ma jeszcze żadnych osób, przyjmij 40 jako domyślną medianę
            if self.ages:
                sorted_ages = sorted(self.ages)
                count = len(sorted_ages)
                if count % 2 == 0:
                    wiek = (sorted_ages[count // 2 - 1] + sorted_ages[count // 2]) / 2
                else:
                    wiek = sorted_ages[count // 2]
                wiek = int(round(wiek))
            else:
                wiek = 40
            person["wiek"] = wiek
        self.total_people += 1
        # Płeć
        if person.get("plec") == "M":
            self.total_males += 1
        elif person.get("plec") == "K":
            self.total_females += 1
        # Wiek
        age = person.get("wiek")
        if age is not None:
            try:
                age_int = int(age)
            except Exception:
                age_int = None
            if age_int is not None:
                self.ages.append(age_int)  # Zapisz wiek do listy
                if age_int < 18:
                    self.people_by_age_group["0-17"] += 1
                elif age_int < 31:
                    self.people_by_age_group["18-30"] += 1
                elif age_int < 51:
                    self.people_by_age_group["31-50"] += 1
                elif age_int < 71:
                    self.people_by_age_group["51-70"] += 1
                elif age_int < 91:
                    self.people_by_age_group["71-90"] += 1
                else:
                    self.people_by_age_group["90+"] += 1
        # Adresy
        address = person.get("adres")
        if address:
            self.addresses.add(address)
        # Imię
        imie = person.get("imie")
        if imie:
            imie = str(imie).strip().capitalize()
            self.names_counter[imie] += 1
    
    def add_birth_year(self, year):
        """Dodaje rok urodzenia do statystyk dekadowych."""
        if year:
            decade = (year // 10) * 10
            self.birth_decades[decade] += 1
    
    def add_marriage_year(self, year):
        """Dodaje rok ślubu do statystyk dekadowych."""
        if year:
            decade = (year // 10) * 10
            self.marriage_decades[decade] += 1
    
    def get_age_stats(self):
        """Oblicza statystyki wiekowe: średnia, mediana, min, max."""
        if not self.ages:
            return {
                "average": 0,
                "median": 0,
                "min": 0,
                "max": 0
            }
        
        sorted_ages = sorted(self.ages)
        count = len(sorted_ages)
        
        # Średnia
        average = sum(sorted_ages) / count
        
        # Mediana
        if count % 2 == 0:
            median = (sorted_ages[count // 2 - 1] + sorted_ages[count // 2]) / 2
        else:
            median = sorted_ages[count // 2]
        
        return {
            "average": round(average, 1),
            "median": median,
            "min": sorted_ages[0],
            "max": sorted_ages[-1]
        }
    
    def add_file(self):
        """Zwiększa licznik przeskanowanych plików."""
        self.files_scanned += 1
    
    def add_sheet(self):
        """Zwiększa licznik przeskanowanych arkuszy."""
        self.sheets_scanned += 1
    
    def add_error(self):
        """Zwiększa licznik błędów."""
        self.errors_count += 1
    
    def add_warning(self):
        """Zwiększa licznik ostrzeżeń."""
        self.warnings_count += 1
    
    def add_unknown_name(self):
        """Zwiększa licznik nieznanych imion."""
        self.unknown_names_count += 1
    
    def add_jubilee(self):
        """Zwiększa licznik jubileuszy."""
        self.jubilees_count += 1
    
    def add_marriage_in_range(self):
        """Zwiększa licznik ślubów w zakresie."""
        self.marriages_in_range_count += 1
    
    def get_summary(self):
        """Zwraca słownik z podsumowaniem statystyk."""
        age_stats = self.get_age_stats()
        return {
            "total_people": self.total_people,
            "total_males": self.total_males,
            "total_females": self.total_females,
            "files_scanned": self.files_scanned,
            "sheets_scanned": self.sheets_scanned,
            "errors_count": self.errors_count,
            "warnings_count": self.warnings_count,
            "unknown_names_count": self.unknown_names_count,
            "jubilees_count": self.jubilees_count,
            "marriages_in_range_count": self.marriages_in_range_count,
            "unique_addresses": len(self.addresses),
            "age_groups": self.people_by_age_group.copy(),
            "birth_decades": dict(self.birth_decades),
            "marriage_decades": dict(self.marriage_decades),
            "age_average": age_stats["average"],
            "age_median": age_stats["median"],
            "age_min": age_stats["min"],
            "age_max": age_stats["max"],
                "analysis_duration": self.get_analysis_duration(),
                "family_count_1": self.family_count_1,
                "family_count_2": self.family_count_2,
                "family_count_3_4": self.family_count_3_4,
                "family_count_5plus": self.family_count_5plus
        }
    
    def format_statistics(self):
        """Formatuje statystyki do czytelnego tekstu z dekadami."""
        from datetime import date
        summary = self.get_summary()
        age_ranges = self.get_family_age_ranges(self.found_people)
        duration = summary["analysis_duration"]

        text = "\n"
        text += "=" * 78 + "\n"
        text += "                    STATYSTYKI ANALIZY                      \n"
        text += "=" * 78 + "\n\n"

        # LUDZIE
        text += "LUDZIE:\n"
        text += f"  Wszystkie osoby:   {summary['total_people']:>5}\n"
        text += f"  Kobiety:           {summary['total_females']:>5} ({self._percentage(summary['total_females'], summary['total_people']):>5.1f}%)\n"
        text += f"  Mezczyzni:         {summary['total_males']:>5} ({self._percentage(summary['total_males'], summary['total_people']):>5.1f}%)\n\n"

        # PLIKI
        text += "PLIKI I ARKUSZE:\n"
        text += f"  Przeskanowane pliki:         {summary['files_scanned']:>5}\n"
        text += f"  Przeskanowane arkusze:       {summary['sheets_scanned']:>5}\n"
        text += f"  Srednio arkuszy na plik:     {self._avg(summary['sheets_scanned'], summary['files_scanned']):>5}\n\n"

        # PODZIAŁ RODZIN
        text += "PODZIAŁ RODZIN:\n"
        min1, max1 = age_ranges['family_1']
        min2, max2 = age_ranges['family_2']
        min34, max34 = age_ranges['family_3_4']
        min5, max5 = age_ranges['family_5plus']
        text += f"  Rodziny 1-osobowe: {summary['family_count_1']:02d} (wiek: {min1 if min1 is not None else '-'}–{max1 if max1 is not None else '-'})\n"
        text += f"  Rodziny 2-osobowe: {summary['family_count_2']:02d} (wiek: {min2 if min2 is not None else '-'}–{max2 if max2 is not None else '-'})\n"
        text += f"  Rodziny 3–4-osobowe: {summary['family_count_3_4']:02d} (wiek: {min34 if min34 is not None else '-'}–{max34 if max34 is not None else '-'})\n"
        text += f"  Rodziny 5+-osobowe: {summary['family_count_5plus']:02d} (wiek: {min5 if min5 is not None else '-' }–{max5 if max5 is not None else '-'})\n\n"
        # URODZINY W DEKADACH
        if summary['birth_decades']:
            text += "URODZINY W DEKADACH (od najstarszych):\n"
            sorted_birth_decades = sorted(summary['birth_decades'].items())
            for decade, count in sorted_birth_decades:
                percentage = self._percentage(count, summary['total_people'])
                bar = self._create_bar(percentage, width=30)
                text += f"  {decade}s: {count:>4} osob  {bar} {percentage:>5.1f}%\n"
            text += "\n"

        # SLUBY W DEKADACH
        if summary['marriage_decades']:
            text += "SLUBY W DEKADACH (od najstarszych):\n"
            sorted_marriage_decades = sorted(summary['marriage_decades'].items())
            for decade, count in sorted_marriage_decades:
                total_marriages = sum(summary['marriage_decades'].values())
                percentage = self._percentage(count, total_marriages)
                bar = self._create_bar(percentage, width=30)
                text += f"  {decade}s: {count:>4} slubow {bar} {percentage:>5.1f}%\n"
            text += "\n"

        # ROZKLAD WIEKU
        text += "ROZKLAD WIEKU:\n"
        text += f"  Srednia wieku:              {summary['age_average']:>5.1f} lat\n"
        text += f"  Mediana wieku:              {summary['age_median']:>5.1f} lat\n"
        text += f"  Najmlodszy:                 {summary['age_min']:>5} lat\n"
        text += f"  Najstarszy:                 {summary['age_max']:>5} lat\n"
        text += "  Grupy wiekowe:\n"
        for age_group, count in summary['age_groups'].items():
            percentage = self._percentage(count, summary['total_people'])
            bar = self._create_bar(percentage, width=30)
            text += f"    {age_group:>6} lat: {count:>4} osob  {bar} {percentage:>5.1f}%\n"
        text += "\n"

        # TOP 20 IMION
        if self.names_counter:
            text += "TOP 20 IMION:\n"
            top_names = sorted(self.names_counter.items(), key=lambda x: x[1], reverse=True)[:20]
            max_count = top_names[0][1] if top_names else 1
            bar_width = 25
            for i, (name, count) in enumerate(top_names, 1):
                percent = self._percentage(count, summary['total_people'])
                # Skala słupka: długość zależna od największego imienia
                filled = int((count / max_count) * bar_width) if max_count else 0
                bar = "█" * filled + "░" * (bar_width - filled)
                text += f"  {i:2d}. {name:<15} {count:>4} [{bar}] {percent:>5.1f}%\n"
            text += "\n"
        # ADRESY
        text += "ADRESY:\n"
        text += f"  Unikalne adresy:             {summary['unique_addresses']:>5}\n"
        text += f"  Srednio osob na adres:       {self._avg(summary['total_people'], summary['unique_addresses']):>5}\n\n"

        # JUBILEUSZE
        text += "JUBILEUSZE I SLUBY:\n"
        text += f"  Nadchodzace jubileusze:      {summary['jubilees_count']:>5}\n"
        text += f"  Sluby w zakresie lat:        {summary['marriages_in_range_count']:>5}\n\n"

        # PROBLEMY
        text += "PROBLEMY:\n"
        text += f"  Bledy:                       {summary['errors_count']:>5}\n"
        text += f"  Ostrzezenia:                 {summary['warnings_count']:>5}\n"
        text += f"  Nieznane imiona:             {summary['unknown_names_count']:>5}\n\n"

        # CZAS
        text += "CZAS ANALIZY:\n"
        text += f"  Czas trwania:           {duration:>8.2f} sekund\n"
        if summary['files_scanned'] > 0:
            text += f"  Sredni czas na plik:    {duration/summary['files_scanned']:>8.2f} s\n"

        # Dni do końca roku
        today = date.today()
        end_of_year = date(today.year, 12, 31)
        days_left = (end_of_year - today).days
        text += f"\nDNI DO KOŃCA ROKU: {days_left:03d}\n"

        text += "\n" + "=" * 78 + "\n"

        return text
    
    def _percentage(self, part, total):
        """Oblicza procent z zaokrągleniem."""
        if total == 0:
            return 0
        return round((part / total) * 100, 1)
    
    def _avg(self, total, count):
        """Oblicza średnią z zaokrągleniem."""
        if count == 0:
            return "0.0"
        return f"{total/count:.1f}"
    
    def _create_bar(self, percentage, width=30):
        """Tworzy prosty pasek postępu z ramką."""
        filled = int((percentage / 100) * width)
        bar = "█" * filled + "░" * (width - filled)
        return f"[{bar}]"
