from collections import Counter

def count_families(people_list):
    """
    Zlicza rodziny według adresów na podstawie listy osób.
    Każda osoba to dict z kluczem 'adres'.
    Zwraca dict:
        {
            'family_count_1': ...,
            'family_count_2': ...,
            'family_count_3_4': ...,
            'family_count_5plus': ...
        }
    """
    addresses = [person.get('adres') for person in people_list if person.get('adres')]
    address_counter = Counter(addresses)
    result = {
        'family_count_1': 0,
        'family_count_2': 0,
        'family_count_3_4': 0,
        'family_count_5plus': 0
    }
    for count in address_counter.values():
        if count == 1:
            result['family_count_1'] += 1
        elif count == 2:
            result['family_count_2'] += 1
        elif 3 <= count <= 4:
            result['family_count_3_4'] += 1
        elif count >= 5:
            result['family_count_5plus'] += 1
    return result
