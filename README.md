# CSV/TXT → XLSX/CSV Merger

Prosta aplikacja w Pythonie z interfejsem **Tkinter**, umożliwiająca scalanie danych z plików CSV lub TXT do jednego pliku wynikowego XLSX i CSV.

---

## Funkcjonalności

- Wybór jednego lub wielu plików CSV/TXT do scalenia  
- Wyświetlanie listy wybranych plików z numeracją i możliwością przewijania  
- Wybór pliku docelowego XLSX  
- Wybór konkretnych kolumn do kopiowania:
  - `Kod`  
  - `ProduktNazwa`  
  - `Cena`  
  - `VAT`  
- Obsługa brakujących kolumn i pustych pól  
- Sortowanie wynikowego pliku **od ostatniego wpisu** (najnowsze dane na górze)  
- Zachowanie tylko **unikalnych kodów** – duplikaty są usuwane, pozostaje najnowszy wpis  
- Zachowanie przecinków w liczbach  
- Obsługa polskich znaków w CSV (`cp1250`)  
- Zapis wyników zarówno do **XLSX**, jak i **CSV**  

---

## Wymagania

- Python 3.10+  
- Biblioteki:
  - `pandas`
  - `openpyxl` (do zapisu XLSX)
  - `tkinter` (standardowo w Pythonie)  

Instalacja zależności:

```bash
pip install pandas openpyxl
