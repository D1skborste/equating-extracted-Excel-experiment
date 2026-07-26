# Tidjämförelse (EAM vs. Time manager)

Ett Python-verktyg för att automatiskt sammanställa och jämföra stämplade flex-tider från Personec med rapporterade projekttider från EAM.

---



## Systemkrav

- Python **3.10** eller nyare installerat på datorn.

---

## Snabbstart

### Alternativ 1: Direkt körning (Rekommenderat)

Startas via `./main/compare_times.py`, eller dra och släpp dina filer på py-ikonen direkt i Utforskaren. Skriptet installerar saknade paket automatiskt vid start.

### För en virtual python-miljlö (venv):
Using your preferred text editor, edit 'run_in_virtual_env.bat', to point to the python executable in your .venv folder.
```"c:\Users\PCMasterRace\Documents\GitHub\venv\Scripts\python.exe"```

Use that batch file to launch the script, or to drop your files into. This will also install dependencies in that virtual environment.

---

## How it works

It will process any excel or csv file, identify EAM project time by the key "AO-aktivitet" or "EAM" in the file (or sheet) name. Any other file(s) will be attributed using employee number or name. It can be separate files or a single Excel using multiple sheets.

When parsing the raw data it will create separate CSV files in ```.\temp_data``` (which is cleared at script start). Using the identiefied "EAM" file as a database, it will iterate the remaining files to match a file name, with existing person and dates in the database. If match is unsuccessful, user is asked to input employee number or name (file name and preview of the first 5 lines is shown). In the case of conflicting but valid name and number, number will take precedent. 

Output is saved to ```.\output_data\<employee_name>```, this isn't deleted but will be overwritten. Take care when employees have identical names.

---

## To do
- Handle output without overwriting when employees have identical names.
- Fix this readme, useability.
- Make more versatile, easier to modify for other systems and comparisons.

---
## Funktioner 
- There is some generic example data to test the functions. All identifiable information is fictional and do not represent a real person. Any correlation to such a person is purely coincidental.
- Can be used as Excel-to-CSV converter, for processing batches of files. Just abort the script when asked for employee info. The output to temp_data isn't deleted until the script is restarted.
---
- **Text below is AI bullshit**
- **Automatiskt filval (GUI / Drag & Drop):** Välj `.csv`, `.xlsx` eller `.xls`-filer via ett grafiskt gränssnitt eller genom att skicka med dem som argument.
- **Automatisk medarbetarmatchning:** Identifierar medarbetare baserat på 4-siffrigt anställningsnummer eller namn i filnamnet.
- **Interaktiv CLI-hantering:** Om en fil inte kan matchas automatiskt visas de första raderna ur filen och du kan ange namn/ID manuellt (eller hoppa över filen med `Enter`).
- **Data-sammanställning:** Beräknar total arbetad tid per dag, jämför avvikelser och exporterar resultatet till standardiserade `.csv`-filer under `output_data/`.
- **Automatisk beroendehantering:** Skriptet kontrollerar och installerar saknade Python-paket vid start.


Kör skriptet direkt via terminalen/kommandotolken:
```bash
python main.py
```

---
