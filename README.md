# XLSX to XML

Prosty program z interfejsem graficznym (GUI), ktory konwertuje pliki Excela (`.xlsx`, `.xlsm`) na pliki `.xml`.
Mozna wskazac pojedynczy plik albo caly folder, wybrac folder wynikowy i jednym kliknieciem przekonwertowac wszystko.

![Python](https://img.shields.io/badge/python-3.9%2B-blue) ![Platforma](https://img.shields.io/badge/platforma-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)

## Funkcje

- Wybor **pliku** lub **folderu** z plikami XLSX
- Wybor **folderu wyjsciowego** (tworzony automatycznie, jesli nie istnieje)
- Konwersja wsadowa - caly folder na raz
- Obsluga wielu arkuszy w jednym skoroszycie
- Opcja "pierwszy wiersz zawiera naglowki" - naglowki staja sie nazwami tagow XML
- Log postepu i bledow w oknie programu
- Wynik w UTF-8, z wcieciami (czytelny XML)

## Wymagania

- **Python 3.9 lub nowszy** ([python.org/downloads](https://www.python.org/downloads/)) - podczas instalacji na Windows zaznacz **"Add Python to PATH"**
- Biblioteka `openpyxl` (instaluje sie sama, patrz nizej)

## Instalacja i uruchomienie

### Windows - najprostszy sposob

1. Pobierz projekt: **Code -> Download ZIP** i rozpakuj (albo `git clone`).
2. Kliknij dwukrotnie **`start.bat`**.

Skrypt sam utworzy srodowisko wirtualne `.venv`, doinstaluje wymagane paczki i uruchomi program.
Przy kolejnych uruchomieniach startuje od razu.

### macOS / Linux

```bash
git clone https://github.com/DawidBochno/xlsx-to-xml.git
cd xlsx-to-xml
./start.sh
```

### Recznie (dowolny system)

```bash
pip install -r requirements.txt
python xlsx2xml.py
```

## Jak korzystac

1. Uruchom program.
2. Kliknij **Plik...** (jeden plik) lub **Folder...** (wszystkie pliki XLSX w folderze).
3. Sprawdz **folder wyjsciowy** - podpowiada sie automatycznie jako podfolder `xml`; mozesz go zmienic przyciskiem **Wybierz...**.
4. Zaznacz lub odznacz **"Pierwszy wiersz zawiera naglowki kolumn"**.
5. Kliknij **Konwertuj**. Postep i ewentualne bledy pojawia sie w oknie logu.

Dla kazdego pliku `nazwa.xlsx` powstaje `nazwa.xml` w folderze wyjsciowym.

## Format wyniku

```xml
<?xml version='1.0' encoding='utf-8'?>
<workbook name="dane.xlsx">
  <sheet name="Arkusz1">
    <row index="2">
      <Imie>Anna</Imie>
      <Wiek>30</Wiek>
    </row>
  </sheet>
</workbook>
```

Szczegoly w [DOCS.md](DOCS.md).

## Test

```bash
python xlsx2xml.py --selftest
```

## Rozwiazywanie problemow

| Problem | Rozwiazanie |
|---|---|
| `python` nie jest rozpoznawane | Zainstaluj Pythona i zaznacz "Add Python to PATH", potem otworz nowe okno konsoli |
| `No module named 'openpyxl'` | `pip install -r requirements.txt` albo uruchom przez `start.bat` / `start.sh` |
| `No module named 'tkinter'` (Linux) | `sudo apt install python3-tk` |
| Plik jest pomijany | Nazwy zaczynajace sie od `~$` to pliki tymczasowe Excela - zamknij plik w Excelu |

## Licencja

MIT
