# PUN GME – profilo orario giornaliero

Grafico HTML interattivo (ascissa: ore 1–24, ordinata: €/MWh).

## Funzioni
- Vista **Mese per mese** (curve giornaliere + media del mese) o **Aggregato** (media, mediana e banda P25–P75).
- Selezione dei mesi (Tutti, Nessuno, Inverno, Estate); colori per fase stagionale.
- **Carica serie PUN**: aggiunge file .xlsx/.csv GME direttamente nel browser.
- **Elimina mesi selezionati** / **Ripristina originali**.

## Rigenerare la pagina
```bash
pip install pandas openpyxl
python build_pun_dashboard.py CARTELLA_XLSX -o PUN_profilo_orario.html
```
I file di input hanno colonne `Data` (gg/mm/aaaa), `Ora`, `Periodo`, `€/MWh` (dati quartorari). Il prezzo orario è la media dei 4 quarti d'ora; la 25ª ora del cambio ora legale viene ignorata.

## Pubblicare su un'altra piattaforma
Il risultato è **un solo file .html**: caricalo su qualunque hosting statico (GitHub Pages, Netlify, SharePoint, sito aziendale) o incorporalo in un `<iframe>`.
Richiede accesso a internet per le librerie da CDN: Chart.js 4.4.1 e SheetJS 0.18.5 (cdnjs.cloudflare.com). Per uso offline scarica i due file e sostituisci i tag `<script src=...>` con percorsi locali.
Le serie caricate/eliminate dal browser sono salvate in `localStorage` (solo su quel dispositivo).
