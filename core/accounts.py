"""
Lokálna cache IBAN → názov účtu.
Ukladá sa do accounts.json vedľa aplikácie.
"""
import json
import os

_PATH = os.path.join(os.path.dirname(__file__), "..", "accounts.json")


def get(iban: str) -> str:
    """Vráti uložený názov pre daný IBAN, alebo "" ak neexistuje."""
    try:
        with open(_PATH, encoding="utf-8") as f:
            return json.load(f).get(iban, "")
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return ""


def save(iban: str, name: str):
    """Uloží alebo aktualizuje názov pre daný IBAN. Atomický zápis — bezpečné pri páde."""
    try:
        with open(_PATH, encoding="utf-8") as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        data = {}

    data[iban] = name

    # Atomický zápis cez tmp súbor — ochrana pred korupciou pri páde
    tmp = _PATH + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, _PATH)
    except OSError:
        try:
            os.remove(tmp)
        except OSError:
            pass
