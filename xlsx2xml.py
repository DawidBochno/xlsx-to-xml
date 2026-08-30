"""XLSX -> XML converter with a small Tkinter GUI.

Uruchomienie:  python xlsx2xml.py
Autotest:      python xlsx2xml.py --selftest
"""
import datetime
import os
import re
import sys
import threading
import xml.etree.ElementTree as ET

try:
    import openpyxl
except ImportError:
    openpyxl = None


def _tag(name, idx):
    """Zamienia nazwe kolumny na poprawna nazwe tagu XML."""
    t = re.sub(r"\W+", "_", str(name).strip(), flags=re.UNICODE).strip("_")
    if not t:
        return "col_%d" % idx
    if not (t[0].isalpha() or t[0] == "_"):
        t = "col_" + t
    return t


def _text(value):
    if value is None:
        return ""
    if isinstance(value, (datetime.datetime, datetime.date, datetime.time)):
        return value.isoformat()
    if isinstance(value, bool):
        return "true" if value else "false"
    return str(value)


def convert_file(xlsx_path, out_dir, headers=True):
    """Konwertuje jeden plik XLSX na XML. Zwraca sciezke pliku wynikowego."""
    if openpyxl is None:
        raise RuntimeError("Brak biblioteki openpyxl - zainstaluj: pip install openpyxl")

    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    root = ET.Element("workbook", name=os.path.basename(xlsx_path))
    try:
        for ws in wb.worksheets:
            sheet_el = ET.SubElement(root, "sheet", name=ws.title)
            names = None
            for r, row in enumerate(ws.iter_rows(values_only=True), start=1):
                if all(c is None for c in row):
                    continue
                if headers and names is None:
                    names = [_tag(c, i) for i, c in enumerate(row, start=1)]
                    continue
                row_el = ET.SubElement(sheet_el, "row", index=str(r))
                for i, cell in enumerate(row, start=1):
                    if cell is None:
                        continue
                    tag = names[i - 1] if names and i <= len(names) else "col_%d" % i
                    ET.SubElement(row_el, tag).text = _text(cell)
    finally:
        wb.close()

    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, os.path.splitext(os.path.basename(xlsx_path))[0] + ".xml")
    ET.indent(root, space="  ")
    ET.ElementTree(root).write(out_path, encoding="utf-8", xml_declaration=True)
    return out_path


def find_inputs(path):
    """Plik -> [plik]. Folder -> wszystkie .xlsx/.xlsm w srodku (bez plikow tymczasowych ~$)."""
    if os.path.isfile(path):
        return [path]
    return sorted(
        os.path.join(path, f)
        for f in os.listdir(path)
        if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$")
    )


# ---------------------------------------------------------------- GUI

def main():
    import tkinter as tk
    from tkinter import filedialog, ttk

    root = tk.Tk()
    root.title("XLSX -> XML")
    root.geometry("640x420")
    root.minsize(560, 380)

    src = tk.StringVar()
    dst = tk.StringVar()
    headers = tk.BooleanVar(value=True)

    frm = ttk.Frame(root, padding=10)
    frm.pack(fill="both", expand=True)
    frm.columnconfigure(1, weight=1)

    ttk.Label(frm, text="Plik / folder wejsciowy:").grid(row=0, column=0, sticky="w")
    ttk.Entry(frm, textvariable=src).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 6))
    btns = ttk.Frame(frm)
    btns.grid(row=1, column=2, sticky="e", padx=(6, 0), pady=(0, 6))
    ttk.Button(btns, text="Plik...", width=9, command=lambda: _pick(src, dst, False)).pack(side="left")
    ttk.Button(btns, text="Folder...", width=9, command=lambda: _pick(src, dst, True)).pack(side="left", padx=(4, 0))

    ttk.Label(frm, text="Folder wyjsciowy:").grid(row=2, column=0, sticky="w")
    ttk.Entry(frm, textvariable=dst).grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 6))
    ttk.Button(frm, text="Wybierz...", width=20,
               command=lambda: dst.set(filedialog.askdirectory() or dst.get())
               ).grid(row=3, column=2, sticky="e", padx=(6, 0), pady=(0, 6))

    ttk.Checkbutton(frm, text="Pierwszy wiersz zawiera naglowki kolumn", variable=headers
                    ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(0, 8))

    log = tk.Text(frm, height=12, wrap="word", state="disabled")
    log.grid(row=6, column=0, columnspan=3, sticky="nsew")
    frm.rowconfigure(6, weight=1)
    sb = ttk.Scrollbar(frm, command=log.yview)
    sb.grid(row=6, column=3, sticky="ns")
    log.configure(yscrollcommand=sb.set)

    def say(msg):
        log.configure(state="normal")
        log.insert("end", msg + "\n")
        log.see("end")
        log.configure(state="disabled")

    def _pick(var, out_var, is_dir):
        p = filedialog.askdirectory() if is_dir else filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Wszystkie pliki", "*.*")])
        if p:
            var.set(os.path.normpath(p))
            if not out_var.get():
                base = p if is_dir else os.path.dirname(p)
                out_var.set(os.path.normpath(os.path.join(base, "xml")))

    def run():
        s, d = src.get().strip(), dst.get().strip()
        if not s or not os.path.exists(s):
            say("! Wskaz istniejacy plik lub folder wejsciowy.")
            return
        if not d:
            say("! Wskaz folder wyjsciowy.")
            return
        files = find_inputs(s)
        if not files:
            say("! Nie znaleziono plikow .xlsx/.xlsm.")
            return
        go.configure(state="disabled")
        say("Start: %d plik(ow) -> %s" % (len(files), d))

        def work():
            ok = 0
            for f in files:
                try:
                    out = convert_file(f, d, headers.get())
                    ok += 1
                    root.after(0, say, "  OK  %s -> %s" % (os.path.basename(f), os.path.basename(out)))
                except Exception as e:
                    root.after(0, say, "  BLAD %s: %s" % (os.path.basename(f), e))
            root.after(0, say, "Gotowe: %d/%d." % (ok, len(files)))
            root.after(0, lambda: go.configure(state="normal"))

        threading.Thread(target=work, daemon=True).start()

    go = ttk.Button(frm, text="Konwertuj", command=run)
    go.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(0, 8))

    if openpyxl is None:
        say("! Brak biblioteki openpyxl. Zainstaluj:  pip install openpyxl")

    root.mainloop()


# ---------------------------------------------------------------- selftest

def selftest():
    import tempfile
    assert _tag("Imie i nazwisko", 1) == "Imie_i_nazwisko"
    assert _tag("2024", 1) == "col_2024"
    assert _tag("   ", 3) == "col_3"
    assert _text(None) == "" and _text(True) == "true"

    tmp = tempfile.mkdtemp()
    xlsx = os.path.join(tmp, "test.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Imie", "Wiek"])
    ws.append(["Anna", 30])
    ws.append([None, None])
    ws.append(["Jan", 41])
    wb.save(xlsx)

    out = convert_file(xlsx, os.path.join(tmp, "out"))
    tree = ET.parse(out)
    rows = tree.findall(".//row")
    assert len(rows) == 2, rows
    assert rows[0].find("Imie").text == "Anna"
    assert rows[1].find("Wiek").text == "41"
    assert len(find_inputs(tmp)) == 1
    print("selftest OK ->", out)


if __name__ == "__main__":
    selftest() if "--selftest" in sys.argv else main()
