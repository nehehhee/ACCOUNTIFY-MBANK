"""
mBank PDF → Excel konvertor (100% lokálne, bez internetu)
Extrahuje bankové transakcie do štruktúrovanej tabuľky.

Stĺpce:
1. Dátum - prvý dátum z dvojice
2. Číslo účtu - IBAN z hlavičky PDF
3. Suma - absolútna hodnota (bez znamienka)
4. C/D - Credit/Debit
5. Účtovný zostatok - bez EUR
6. Číslo protiúčtu - IBAN z popisu transakcie
7. Názov protiúčtu - názov z popisu transakcie
8. Typ transakcie - typ z popisu transakcie
9. Poznámka - posledný riadok popisu (VS, KS, text)
"""

import pdfplumber
import pandas as pd
import re
import os
import sys

# Tolerancia pre spájanie znakov do slov.
# Hodnota 2 zachováva medzery medzi slovami (napr. "PLATBA KARTOU")
# Default (3) ich spája dokopy ("PLATBAKARTOU")
TEXT_X_TOLERANCE = 2


def extract_account_iban(pdf):
    """Extrahuje IBAN vlastného účtu z prvej strany PDF."""
    page = pdf.pages[0]
    text = page.extract_text(x_tolerance=TEXT_X_TOLERANCE) or ""

    iban_match = re.search(r'IBAN:\s*(SK\d{22})', text)
    if iban_match:
        return iban_match.group(1)

    account_match = re.search(r'(SK\d{22})', text)
    if account_match:
        return account_match.group(1)

    return ""


def parse_amount(amount_str):
    """
    Parsuje sumu z reťazca ako '-50,67 EUR', '1 675,25 EUR', '-10,00EUR'.
    Zvládne medzery v tisíckach (8 039,23) aj bez nich (8039,23).
    Vráti (absolútna_suma, je_kredit).
    """
    if not amount_str:
        return None, None

    amount_str = amount_str.strip()

    # Odstránenie "EUR"
    amount_str = re.sub(r'\s*EUR\s*$', '', amount_str).strip()

    # Zistenie znamienka
    is_negative = amount_str.startswith("-")

    # Odstránenie znamienka
    amount_str = amount_str.lstrip("-+").strip()

    # Odstránenie medzier (tisícové oddelovače: "8 039,23" → "8039,23")
    amount_str = amount_str.replace(" ", "")

    # Konverzia čiarky na bodku
    amount_str = amount_str.replace(",", ".")

    try:
        amount = float(amount_str)
    except ValueError:
        return None, None

    return amount, not is_negative


def parse_balance(balance_str):
    """Parsuje zostatok - odstráni EUR a medzery v tisíckach."""
    if not balance_str:
        return None

    balance_str = balance_str.strip()
    balance_str = re.sub(r'\s*EUR\s*$', '', balance_str).strip()

    # Odstránenie medzier (tisícové oddelovače)
    balance_str = balance_str.replace(" ", "")
    balance_str = balance_str.replace(",", ".")

    try:
        return float(balance_str)
    except ValueError:
        return None


def extract_iban_from_text(text):
    """
    Nájde IBAN v texte. Zvládne aj medzery v IBAN-e (napr. 'SK95 0200 0020 1400 0010 4512').
    Najprv skúsi priamy match, potom skúsi text bez medzier.
    """
    # 1. Priamy match (IBAN bez medzier)
    iban_match = re.search(r'([A-Z]{2}\d{2}[\dA-Za-z]{10,30})', text)
    if iban_match:
        candidate = iban_match.group(1)
        if len(candidate) >= 15:
            return candidate

    # 2. Skúsime text bez medzier (pre medzery v IBAN-e)
    text_no_spaces = text.replace(" ", "")
    iban_match = re.search(r'([A-Z]{2}\d{2}[\dA-Za-z]{10,30})', text_no_spaces)
    if iban_match:
        candidate = iban_match.group(1)
        if len(candidate) >= 15:
            return candidate

    return ""


def parse_description(desc_text):
    """
    Parsuje popis transakcie (multiline text oddelený \\n).

    Typická štruktúra (s IBAN-om):
    Riadok 1: Typ transakcie
    Riadok 2: Názov protiúčtu
    Riadok 3: IBAN protiúčtu
    Riadok 4+: Poznámka

    Ak IBAN nie je nájdený:
    - Názov protiúčtu ostane PRÁZDNY
    - Všetko okrem typu transakcie ide do Poznámky
    """
    if not desc_text:
        return "", "", "", ""

    lines = desc_text.split("\n")
    lines = [l.strip() for l in lines if l.strip()]

    typ_transakcie = ""
    nazov_protiuctu = ""
    cislo_protiuctu = ""
    poznamka = ""

    if len(lines) == 0:
        return typ_transakcie, nazov_protiuctu, cislo_protiuctu, poznamka

    # Riadok 1 = Typ transakcie (vždy)
    typ_transakcie = lines[0]

    if len(lines) == 1:
        return typ_transakcie, nazov_protiuctu, cislo_protiuctu, poznamka

    # Hľadáme IBAN v celom popise (riadky 2+)
    iban_line_idx = None
    for i in range(1, len(lines)):
        iban = extract_iban_from_text(lines[i])
        if iban:
            cislo_protiuctu = iban
            iban_line_idx = i
            break

    if cislo_protiuctu:
        # IBAN nájdený → Názov protiúčtu je riadok medzi typom a IBAN-om
        if iban_line_idx > 1:
            # Medzi typom a IBAN-om je názov (riadok 2)
            nazov_protiuctu = lines[1]
            # Špeciálny prípad: "0" ako artefakt (napr. ZSE)
            if nazov_protiuctu == "0" and iban_line_idx > 2:
                # "0" je artefakt, skutočný názov je za ním
                nazov_protiuctu = lines[1]  # necháme ako je
        # Poznámka = všetko za IBAN-om
        if iban_line_idx < len(lines) - 1:
            poznamka_lines = lines[iban_line_idx + 1:]
            poznamka = " ".join(poznamka_lines)
    else:
        # IBAN NEBOL nájdený → Názov protiúčtu ostane PRÁZDNY
        # Všetko okrem typu transakcie ide do Poznámky
        nazov_protiuctu = ""
        if len(lines) > 1:
            poznamka = " ".join(lines[1:])

    return typ_transakcie, nazov_protiuctu, cislo_protiuctu, poznamka


def process_mbank_pdf(pdf_path, output_path):
    """Hlavná funkcia na spracovanie mBank PDF."""

    print(f"\n📄 Spracovávam: {os.path.basename(pdf_path)}")
    print(f"   Cesta: {pdf_path}")

    transactions = []

    with pdfplumber.open(pdf_path) as pdf:
        print(f"   Počet strán: {len(pdf.pages)}")

        # 1. Extrahovanie IBAN vlastného účtu
        account_iban = extract_account_iban(pdf)
        print(f"   IBAN účtu: {account_iban}")

        # 2. Nastavenie tabuľky s nižšou x_tolerance pre správne medzery
        table_settings = {
            "text_x_tolerance": TEXT_X_TOLERANCE,
        }

        # 3. Prechádzame všetky strany
        for page_num, page in enumerate(pdf.pages):
            tables = page.extract_tables(table_settings)

            if not tables:
                continue

            for table in tables:
                for row_idx, row in enumerate(table):
                    # Preskočíme hlavičky
                    if row_idx == 0:
                        continue

                    if not row or not row[0]:
                        continue

                    dates_cell = row[0]

                    # Overenie dátumu
                    date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', dates_cell)
                    if not date_match:
                        continue

                    # Dátum - prvý dátum
                    dates = dates_cell.split("\n")
                    datum = dates[0].strip()

                    # Popis transakcie
                    desc_cell = row[1] if len(row) > 1 else ""

                    # Suma a zostatok - pozícia závisí od počtu stĺpcov
                    if len(row) == 5:
                        amount_str = row[3] or ""
                        balance_str = row[4] or ""
                    elif len(row) == 4:
                        amount_str = row[2] or ""
                        balance_str = row[3] or ""
                    else:
                        amount_str = ""
                        balance_str = ""

                    # Parsovanie sumy
                    suma, is_credit = parse_amount(amount_str)
                    if suma is None:
                        continue

                    cd = "C" if is_credit else "D"
                    zostatok = parse_balance(balance_str)

                    # Parsovanie popisu
                    typ_transakcie, nazov_protiuctu, cislo_protiuctu, poznamka = parse_description(desc_cell)

                    transactions.append({
                        "Dátum": datum,
                        "Číslo účtu": account_iban,
                        "Suma": suma,
                        "C/D": cd,
                        "Účtovný zostatok": zostatok,
                        "Číslo protiúčtu": cislo_protiuctu,
                        "Názov protiúčtu": nazov_protiuctu,
                        "Typ transakcie": typ_transakcie,
                        "Poznámka": poznamka,
                    })

    print(f"\n   ✅ Nájdených transakcií: {len(transactions)}")

    # DataFrame a Excel
    df = pd.DataFrame(transactions, columns=[
        "Dátum", "Číslo účtu", "Suma", "C/D", "Účtovný zostatok",
        "Číslo protiúčtu", "Názov protiúčtu", "Typ transakcie", "Poznámka",
    ])

    df.to_excel(output_path, sheet_name="Transakcie", index=False, engine="openpyxl")

    print(f"   💾 Uložené do: {output_path}")
    print(f"   📊 Jeden list 'Transakcie' s {len(df)} riadkami a 9 stĺpcami")

    # Ukážka
    print(f"\n   --- Ukážka prvých 5 transakcií ---")
    for i, row in df.head(5).iterrows():
        print(f"   {i+1}. {row['Dátum']} | {row['Suma']:.2f} {row['C/D']} | {row['Typ transakcie'][:35]} | {row['Názov protiúčtu'][:25]}")

    return df


def main():
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    pdf_path = os.path.join(desktop, "MBANK.pdf")

    if not os.path.exists(pdf_path):
        print(f"❌ Súbor nenájdený: {pdf_path}")
        sys.exit(1)

    output_path = os.path.join(desktop, "MBANK_transakcie_v3.xlsx")

    print("=" * 60)
    print("📋 mBank PDF → Excel konvertor (lokálne)")
    print("=" * 60)

    df = process_mbank_pdf(pdf_path, output_path)

    # Štatistiky
    credit_count = len(df[df["C/D"] == "C"])
    debit_count = len(df[df["C/D"] == "D"])
    credit_sum = df[df["C/D"] == "C"]["Suma"].sum()
    debit_sum = df[df["C/D"] == "D"]["Suma"].sum()

    print(f"\n{'='*60}")
    print(f"📊 ŠTATISTIKY")
    print(f"{'='*60}")
    print(f"   Kreditné (C): {credit_count} transakcií, spolu {credit_sum:,.2f} EUR")
    print(f"   Debetné (D):  {debit_count} transakcií, spolu {debit_sum:,.2f} EUR")
    print(f"   Celkom:       {len(df)} transakcií")
    print(f"\n✅ HOTOVO! Otvor súbor: {output_path}")


if __name__ == "__main__":
    main()
