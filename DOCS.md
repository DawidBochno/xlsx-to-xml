# Dokumentacja techniczna

## Struktura projektu

| Plik | Opis |
|---|---|
| `xlsx2xml.py` | Caly program: logika konwersji + GUI (Tkinter) + autotest |
| `requirements.txt` | Zaleznosci (`openpyxl`) |
| `start.bat` / `start.sh` | Uruchomienie z automatycznym `.venv` i instalacja paczek |
| `README.md` | Instalacja i instrukcja obslugi |

## Struktura XML

```xml
<workbook name="NAZWA_PLIKU.xlsx">
  <sheet name="NAZWA_ARKUSZA">
    <row index="NUMER_WIERSZA_W_EXCELU">
      <NAZWA_KOLUMNY>wartosc</NAZWA_KOLUMNY>
    </row>
  </sheet>
</workbook>
```

- `index` to rzeczywisty numer wiersza w arkuszu Excela (liczony od 1).
- Kazdy arkusz skoroszytu trafia do osobnego elementu `<sheet>`.
- Puste wiersze sa pomijane, puste komorki nie tworza tagow.

## Nazwy tagow

Naglowki kolumn sa zamieniane na poprawne nazwy XML:

| Naglowek w Excelu | Tag XML |
|---|---|
| `Imie i nazwisko` | `Imie_i_nazwisko` |
| `Cena [PLN]` | `Cena_PLN` |
| `2024` | `col_2024` |
| pusty | `col_3` (numer kolumny) |

Jesli opcja naglowkow jest **wylaczona**, wszystkie kolumny dostaja nazwy `col_1`, `col_2`, ...

## Konwersja wartosci

| Typ w Excelu | Wynik w XML |
|---|---|
| tekst / liczba | wartosc jako tekst |
| data / czas | format ISO 8601, np. `2024-05-01T00:00:00` |
| logiczny | `true` / `false` |
| pusta komorka | tag pomijany |
| formula | wynik formuly (`data_only=True`), nie tresc formuly |

> Uwaga: wynik formuly jest odczytywany z cache zapisanego przez Excela. Plik wygenerowany programowo (bez otwarcia w Excelu) moze nie miec zapisanych wynikow - wtedy komorka bedzie pusta.

## API (uzycie z poziomu kodu)

```python
from xlsx2xml import convert_file, find_inputs

for f in find_inputs("C:/dane"):        # plik albo folder
    convert_file(f, "C:/dane/xml", headers=True)
```

- `find_inputs(path)` - zwraca liste plikow `.xlsx`/`.xlsm` (pomija tymczasowe `~$`).
- `convert_file(xlsx_path, out_dir, headers=True)` - konwertuje jeden plik, zwraca sciezke wyniku.

## Uwagi implementacyjne

- Pliki czytane sa w trybie `read_only=True` - dziala z duzymi arkuszami przy malym zuzyciu pamieci.
- Konwersja dziala w osobnym watku, wiec GUI nie zamiera przy duzych plikach.
- XML zapisywany jest w UTF-8 z deklaracja i wcieciami (`ET.indent`).

## Test

```bash
python xlsx2xml.py --selftest
```

Tworzy tymczasowy plik XLSX, konwertuje go i sprawdza asercjami nazwy tagow, pomijanie pustych wierszy i wartosci komorek.
