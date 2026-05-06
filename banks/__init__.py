"""
Registre parserov. Pre pridanie novej banky:
1. Vytvor banks/nazov_banky.py s klasou ktorá dedí od BankParser
2. Importuj ju tu a pridaj do _PARSERS
"""
import pdfplumber

from .mbank import MBankParser

_PARSERS = [
    MBankParser(),
    # TatraBankaParser(),
    # CSOBParser(),
]


def get_parser(pdf_path: str):
    """Vráti správny parser pre dané PDF, alebo None ak banka nie je podporovaná."""
    with pdfplumber.open(pdf_path) as pdf:
        for parser in _PARSERS:
            if parser.can_parse(pdf):
                return parser
    return None


def supported_banks() -> list[str]:
    return [p.name for p in _PARSERS]
