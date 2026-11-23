# Riconcilia POS

Applicazione locale per confrontare POS e prima nota ed esportare Excel.

## Windows senza installazioni

Aprire Actions, selezionare l’ultima esecuzione riuscita di **Crea programma Windows** e scaricare **RiconciliaPOS-Windows**. Estrarre tutto lo ZIP e aprire **RiconciliaPOS.exe**. La cartella `_internal` deve restare accanto all’eseguibile. Non occorrono Python o Internet per usare il programma.

La compilazione prova l’eseguibile Windows con documenti sintetici e verifica anche l’export Excel.

## Documenti e riservatezza

Questo repository contiene solo codice e test sintetici. Nessun documento contabile, anagrafica o risultato reale è incluso. I file caricati nell’app sono elaborati sul computer locale.

## Sviluppo

Python 3.12, `pip install -r requirements.txt`, `python server.py --open`. Test: `python -m unittest -v test_core.py`.

PDF testuali nel formato supportato; scansioni e file XLS non supportati. Preautorizzate da verificare. L’Excel fotografa il confronto; le modifiche manuali non rieseguono le associazioni.
