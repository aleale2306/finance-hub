# Finance Hub 4.3 aggiornato — Mac

Pacchetto locale con gli aggiornamenti 5.2, 5.4, 5.5 e 5.7 applicati alla base di Finance Hub, mantenendo struttura, filtri, grafici, caricamenti e stile dell'interfaccia esistente.

## Prima installazione
1. Decomprimi lo ZIP in una cartella locale.
2. Fai doppio clic su `INSTALLA_E_AVVIA.command`.
3. Lo script crea un ambiente Python locale, installa le dipendenze e apre `http://localhost:8501`.
4. Se macOS blocca lo script: tasto destro sul file → **Apri**. In alternativa, dal Terminale:
   `xattr -dr com.apple.quarantine /percorso/Finance_Hub_4.3_Aggiornato`

Requisiti: macOS, Python 3.10 o successivo e connessione Internet solo durante la prima installazione delle dipendenze.

## Avvii successivi
- Avvio: `AVVIA_FINANCE_HUB.command`
- Riavvio: `RIAVVIA_FINANCE_HUB.command`
- Arresto: `ARRESTA_FINANCE_HUB.command`
- Log locale: `finance_hub.log`

Non copiare una vecchia cartella `.venv` da un'altra versione.

## Dati
- Configurazione inclusa: `Config_Finanze_Familiari_V2_Aggiornato.xlsx`.
- Carica manualmente gli export MoneyWiz CSV/XLSX dalla sidebar **Caricamenti**.
- Per i workbook storici viene selezionato il foglio transazionale più affidabile e viene usata prima la colonna `Importo in EUR`.
- Se `Importo in EUR` non è presente, vengono applicati i tassi del foglio `Saldi Iniziali`; restano disponibili i tassi storici per JPY, ISK ed EGP.
- I dati sono elaborati localmente e non sono inclusi nel pacchetto.

## Aggiornamenti inclusi
- **5.2 — Patrimonio storico:** serie mensile completa, tabella con variazioni ed export CSV.
- **5.4 — Storico vs Previsione:** confronto mensile per Entrate, Spese fisse inevitabili, Spese fisse evitabili, Spese necessarie, Spese variabili non necessarie e linee straordinarie Matrimonio, Casa nuova, Luna di miele (Giappone), Altri viaggi. Verde = favorevole; rosso = sfavorevole.
- **5.5 — Storico:** grafico mensile actual, dettaglio transazioni ed export CSV.
- **5.7 — Data quality:** elenco ed export delle transazioni senza categoria, escludendo i trasferimenti interni.

## Non inclusi in questa build
- Drill-down avanzato nella pagina Previsione.
- Automazione MoneyWiz.
- Budget avanzato nel file Config.
- Scenario analysis.
