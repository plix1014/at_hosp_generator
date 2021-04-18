# at_hosp_generator

creates Excel/Libre Office file from AGES CovidFallzahlen.csv

CSV Dateien werden einmal am Tag von der AGES Homepage runtergeladen.

### Prerequisites

* [Python 3 f. Windows](https://www.python.org/downloads/windows/) software
* [Python 3 f. MacOS](https://docs.python-guide.org/starting/install3/osx//) software
* Linux: apt-get install python3 || yum install python3

Installed with requirements.txt
* python 'pandas' module
* python 'requests' module
* python 'openpyxl' module


### Installation f. Windows

Die Dateien in ein Verzeichnis auspacken und das "generate.cmd" Skript starten.
Die Ergebnis Datei "AT_Hospitalisierung.xlsx" liegt dann im Unterverzeichnis "data"

e.g.: %systemdrive%%homepath%\Dokumente\ages
```
   generate.cmd
```

* Dateien für Windows
```
howto_windows.txt
requirements.txt
generate.cmd
prepare_icu_venv.cmd
at_hosp_csv2excel.py
```

* Dateien für Unix
```
howto_unix.txt
requirements.txt
at_hosp_csv2excel.py
```


## Author

* **plix1014** - [plix1014](https://github.com/plix1014)


## License

This project is licensed under the Attribution-NonCommercial-ShareAlike 4.0 International License - see the [LICENSE.md](LICENSE.md) file for details


## Acknowledgments

Mit diesem Skript wird eine Datei wie in diesem Video vorgeführt, erzeugt. Einmal täglich werden die aktuellen Daten von der AGES Homepage runtergeladen.
https://www.youtube.com/watch?v=A_8ZNvl2ZWQ

