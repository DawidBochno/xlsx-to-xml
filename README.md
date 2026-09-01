# XLSX to XML

Prosty program z interfejsem graficznym (GUI), ktory konwertuje pliki Excela (`.xlsx`, `.xlsm`, `.xls`) na pliki `.xml`.
Mozna wskazac pojedynczy plik albo caly folder, wybrac folder wynikowy i jednym kliknieciem przekonwertowac wszystko.

> **Program dziala w 100% lokalnie na Twoim komputerze.** Nie wysyla zadnych danych do internetu,
> nie korzysta z chmury ani zewnetrznych serwerow i nie wymaga zakladania konta.
> Po instalacji dziala **calkowicie offline** - Twoje pliki nigdy nie opuszczaja dysku.

![Python](https://img.shields.io/badge/python-3.9%2B-blue) ![Platforma](https://img.shields.io/badge/platforma-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey) ![Lokalnie](https://img.shields.io/badge/dziala-lokalnie%20%7C%20offline-brightgreen) ![Licencja](https://img.shields.io/badge/licencja-MIT-green)

## Funkcje

- **Dziala lokalnie i offline** - zero wysylki danych, zero telemetrii
- Wybor **pliku** lub **folderu** z plikami Excela
- Obsluga **nowego** (`.xlsx`, `.xlsm`) i **starego** (`.xls`) formatu Excela
- Wybor **folderu wyjsciowego** (tworzony automatycznie, jesli nie istnieje)
- Konwersja wsadowa - caly folder na raz
- Obsluga wielu arkuszy w jednym skoroszycie
- Opcja "pierwszy wiersz zawiera naglowki" - naglowki staja sie nazwami tagow XML
- Log postepu i bledow w oknie programu
- **Zapamietuje ostatnio uzyte sciezki** i ustawienia (plik `~/.xlsx2xml.json`)
- Uruchamia sie **bez okna konsoli** (`program_xlsx-to-xml.bat`)
- Wynik w UTF-8, z wcieciami (czytelny XML)

## Prywatnosc i bezpieczenstwo danych

Cala konwersja odbywa sie na Twoim komputerze:

- Program **nie nawiazuje polaczen sieciowych** - nie ma w nim ani jednego zapytania do internetu.
- Pliki wejsciowe sa tylko **odczytywane**, wyniki zapisywane do wskazanego przez Ciebie folderu.
- Jedyne dane zapisywane poza folderem wynikowym to ostatnio uzyte sciezki w pliku `~/.xlsx2xml.json` (na Twoim dysku).
- Internet potrzebny jest **wylacznie raz**, przy instalacji - do pobrania Pythona i paczek `openpyxl` / `xlrd` z PyPI.
  Pozniej program dziala bez sieci.

Dzieki temu mozna go bezpiecznie uzywac do plikow firmowych i danych wrazliwych.

## Wymagania

- **Python 3.9 lub nowszy** ([python.org/downloads](https://www.python.org/downloads/)) - podczas instalacji na Windows zaznacz **"Add Python to PATH"**
- Biblioteki `openpyxl` (dla `.xlsx`) i `xlrd` (dla `.xls`) - instaluja sie same, patrz nizej

## Instalacja i uruchomienie

### Windows - najprostszy sposob

1. Pobierz projekt: **Code -> Download ZIP** i rozpakuj (albo `git clone`).
2. Kliknij dwukrotnie **`install.bat`** - tworzy srodowisko `.venv` i instaluje wymagane paczki. Robisz to **tylko raz**.
3. Uruchamiaj program plikiem **`program_xlsx-to-xml.bat`** - startuje **bez okna konsoli**.

Mozesz zrobic skrot do `program_xlsx-to-xml.bat` na pulpicie.
Jesli uruchomisz go bez wczesniejszej instalacji, sam wywola `install.bat`.
Paczki doinstalowuja sie ponownie tylko wtedy, gdy zmieni sie `requirements.txt`.

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
3. Sprawdz **folder wyjsciowy** (przy kolejnym starcie obie sciezki podpowiadaja sie z ostatniego uzycia) - podpowiada sie automatycznie jako podfolder `xml`; mozesz go zmienic przyciskiem **Wybierz...**.
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
| `No module named 'openpyxl'` | Uruchom `install.bat` (Windows) / `start.sh` (Linux, macOS) albo `pip install -r requirements.txt` |
| `No module named 'tkinter'` (Linux) | `sudo apt install python3-tk` |
| `No module named 'xlrd'` | `pip install -r requirements.txt` - potrzebne tylko do plikow `.xls` |
| Plik jest pomijany | Nazwy zaczynajace sie od `~$` to pliki tymczasowe Excela - zamknij plik w Excelu |
| Chce zresetowac zapamietane sciezki | Usun plik `.xlsx2xml.json` z katalogu domowego uzytkownika |

## Licencja

MIT
