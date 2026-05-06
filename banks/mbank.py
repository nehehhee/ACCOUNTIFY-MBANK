import re
import pdfplumber
import pandas as pd

from core.base import BankParser, COLUMNS
from core import accounts

TEXT_X_TOLERANCE = 2


class MBankParser(BankParser):
    name = "mBank"

    # ── Detekcia ──────────────────────────────────────────────────────────────

    def can_parse(self, pdf) -> bool:
        text = pdf.pages[0].extract_text(x_tolerance=TEXT_X_TOLERANCE) or ""
        return bool(re.search(r'mBank', text, re.IGNORECASE))

    # ── Peek (rýchle načítanie bez konverzie) ─────────────────────────────────

    def peek(self, pdf_path: str) -> dict:
        with pdfplumber.open(pdf_path) as pdf:
            iban = self._extract_account_iban(pdf)
            nazov = self._extract_account_name(pdf)

        # Ak sa meno nenašlo v PDF, skús cache
        if not nazov and iban:
            nazov = accounts.get(iban)

        return {"bank": self.name, "iban": iban, "nazov": nazov}

    # ── Parse (plná konverzia) ────────────────────────────────────────────────

    def parse(self, pdf_path: str, nazov_uctu: str = "", progress_cb=None) -> tuple[str, pd.DataFrame]:
        transactions = []
        ts = {"text_x_tolerance": TEXT_X_TOLERANCE}

        with pdfplumber.open(pdf_path) as pdf:
            iban = self._extract_account_iban(pdf)
            total = len(pdf.pages)

            for i, page in enumerate(pdf.pages):
                if progress_cb:
                    progress_cb(int((i + 1) / total * 90), f"Strana {i+1} / {total}")
                for table in (page.extract_tables(ts) or []):
                    for ri, row in enumerate(table):
                        if ri == 0 or not row or not row[0]:
                            continue
                        if not re.search(r'\d{2}\.\d{2}\.\d{4}', row[0]):
                            continue
                        datum = row[0].split("\n")[0].strip()
                        desc = row[1] if len(row) > 1 else ""
                        if len(row) == 5:
                            amt_s, bal_s = row[3] or "", row[4] or ""
                        elif len(row) == 4:
                            amt_s, bal_s = row[2] or "", row[3] or ""
                        else:
                            amt_s = bal_s = ""
                        suma, is_credit = self._parse_amount(amt_s)
                        if suma is None:
                            continue
                        bal_val, _ = self._parse_amount(bal_s)
                        typ, nazov_p, cislo, poz = self._parse_description(desc)
                        transactions.append({
                            "banka":           self.name,
                            "nazov_uctu":      nazov_uctu,
                            "iban_uctu":       iban,
                            "datum":           datum,
                            "suma":            suma,
                            "cd":              "C" if is_credit else "D",
                            "nazov_protiuctu": nazov_p,
                            "iban_protiuctu":  cislo,
                            "typ":             typ,
                            "poznamka":        poz,
                            "zostatok":        bal_val,
                        })

        df = pd.DataFrame(transactions, columns=COLUMNS)
        return iban, df

    # ── Pomocné metódy ────────────────────────────────────────────────────────

    def _extract_account_iban(self, pdf) -> str:
        text = pdf.pages[0].extract_text(x_tolerance=TEXT_X_TOLERANCE) or ""
        m = re.search(r'IBAN:\s*(SK\d{22})', text) or re.search(r'(SK\d{22})', text)
        return m.group(1) if m else ""

    def _extract_account_name(self, pdf) -> str:
        """Pokúsi sa nájsť meno majiteľa účtu na prvej strane PDF."""
        text = pdf.pages[0].extract_text(x_tolerance=TEXT_X_TOLERANCE) or ""
        lines = [l.strip() for l in text.splitlines() if l.strip()]

        # Vzor 1: explicitný label
        for label in (r'Majiteľ účtu[:\s]+(.+)', r'Držiteľ účtu[:\s]+(.+)', r'Klient[:\s]+(.+)'):
            m = re.search(label, text, re.IGNORECASE)
            if m:
                return m.group(1).strip()

        # Vzor 2: riadok tesne pred IBAN riadkom
        for i, line in enumerate(lines):
            if re.search(r'IBAN', line, re.IGNORECASE) and i > 0:
                candidate = lines[i - 1]
                # Meno vyzerá ako "Ján Novák" — aspoň dve slová, len písmená
                if re.match(r'^[A-ZÁČĎÉÍĽĹŇÓÔŔŠŤÚÝŽ][a-záčďéíľĺňóôŕšťúýž]+(?: [A-ZÁČĎÉÍĽĹŇÓÔŔŠŤÚÝŽ][a-záčďéíľĺňóôŕšťúýž]+)+$', candidate):
                    return candidate

        return ""

    def _parse_amount(self, s):
        if not s:
            return None, None
        s = re.sub(r'\s*EUR\s*$', '', s.strip()).strip()
        neg = s.startswith("-")
        s = s.lstrip("-+").replace(" ", "").replace(",", ".")
        try:
            return float(s), not neg
        except ValueError:
            return None, None

    def _extract_iban_from_text(self, text) -> str:
        for t in (text, text.replace(" ", "")):
            m = re.search(r'([A-Z]{2}\d{2}[\dA-Za-z]{10,30})', t)
            if m and len(m.group(1)) >= 15:
                return m.group(1)
        return ""

    def _parse_description(self, desc) -> tuple[str, str, str, str]:
        if not desc:
            return "", "", "", ""
        lines = [l.strip() for l in desc.split("\n") if l.strip()]
        if not lines:
            return "", "", "", ""

        typ = lines[0]
        nazov = cislo = poznamka = ""
        iban_idx = None

        for i in range(1, len(lines)):
            iban = self._extract_iban_from_text(lines[i])
            if iban:
                cislo, iban_idx = iban, i
                break

        if cislo:
            if iban_idx > 1:
                nazov = lines[1]
            if iban_idx < len(lines) - 1:
                poznamka = " ".join(lines[iban_idx + 1:])
        else:
            poznamka = " ".join(lines[1:])

        return typ, nazov, cislo, poznamka
