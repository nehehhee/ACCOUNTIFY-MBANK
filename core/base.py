from abc import ABC, abstractmethod
import pandas as pd

COLUMNS = [
    "banka",
    "nazov_uctu",
    "iban_uctu",
    "datum",
    "suma",
    "cd",
    "zostatok",
    "nazov_protiuctu",
    "iban_protiuctu",
    "typ",
    "poznamka",
]


class BankParser(ABC):
    name: str = ""

    @abstractmethod
    def can_parse(self, pdf) -> bool:
        """Vráti True ak tento parser vie spracovať dané PDF (otvorené cez pdfplumber)."""
        ...

    @abstractmethod
    def peek(self, pdf_path: str) -> dict:
        """
        Rýchle načítanie bez plnej konverzie.
        Vráti {"bank": str, "iban": str, "nazov": str}
        nazov môže byť "" ak sa nenašlo v PDF.
        """
        ...

    @abstractmethod
    def parse(self, pdf_path: str, nazov_uctu: str = "", progress_cb=None) -> tuple[str, pd.DataFrame]:
        """
        Spracuje PDF a vráti (iban_vlastneho_uctu, DataFrame).
        nazov_uctu — meno účtu zadané užívateľom (môže byť "").
        progress_cb(percent: int, text: str) je voliteľný callback pre UI.
        """
        ...
