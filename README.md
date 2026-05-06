# mBank PDF → Excel Konvertor

Lokálny nástroj na extrakciu bankových transakcií z mBank PDF výpisov do štruktúrovaného Excel súboru.

## Funkcie

- Extrakcia transakcií z mBank PDF výpisov
- 9 stĺpcov: Dátum, Číslo účtu, Suma, C/D, Účtovný zostatok, Číslo protiúčtu, Názov protiúčtu, Typ transakcie, Poznámka
- 100% lokálne spracovanie (bez internetu)
- Správne medzery medzi slovami
- Podpora IBAN-ov s medzerami

## Požiadavky

- Python 3.10+
- pdfplumber
- pandas
- openpyxl

## Inštalácia

```bash
pip install pdfplumber pandas openpyxl
```

## Použitie

1. Ulož PDF výpis z mBank na plochu
2. Spusti skript:

```bash
python pdf_to_excel.py
```

3. Výstupný Excel súbor sa uloží na plochu

## Výstup

Excel súbor s jedným listom "Transakcie" obsahujúcim stĺpce:

| Stĺpec | Popis |
|---------|-------|
| Dátum | Dátum zaúčtovania |
| Číslo účtu | IBAN vlastného účtu |
| Suma | Absolútna hodnota transakcie |
| C/D | C = kredit (príjem), D = debet (výdaj) |
| Účtovný zostatok | Zostatok po transakcii |
| Číslo protiúčtu | IBAN protistrany |
| Názov protiúčtu | Názov protistrany |
| Typ transakcie | Typ operácie |
| Poznámka | VS, KS, popis |
