
# AtecoExtractor

AtecoExtractor è un'applicazione GUI progettata per estrarre informazioni dettagliate sulle aziende corrispondenti a un determinato codice ATECO da fonti web pubbliche e integrarle in un file Excel e, opzionalmente, in un database MySQL.

## Caratteristiche principali

AtecoExtractor può raccogliere le seguenti informazioni:

- **Nome Azienda**: Nome dell'azienda.
- **Indirizzo**: Indirizzo fisico dell'azienda.
- **Codice CCRea**: Codice della Camera di Commercio.
- **Partita IVA**: Numero di Partita IVA.
- **Stato Attività**: Stato attuale dell'attività aziendale.
- **Fatturato**: Fatturato annuo dell'azienda.
- **Codice ATECO**: Codice ATECO dell'azienda.

## Installazione

Per installare le dipendenze necessarie, eseguire il seguente comando:

```bash
pip install -r requirements.txt
```
## Prerequisiti
Python 3.x

MySQL Server

Biblioteca requests per Python

## Utilizzo
##### Inserimento del Codice ATECO:
Avvia l'applicazione e inserisci il codice ATECO nel formato XX_X.
##### Specifica anche il nome del file Excel in cui salvare i dati estratti.
Estrazione e Salvataggio dei Dati:
##### Premi il pulsante "Estrai dati" per avviare l'estrazione delle informazioni.
Durante l'operazione, verrà visualizzata una barra di avanzamento che mostra lo stato dell'estrazione.
##### Salvataggio dei Risultati:
Dopo l'estrazione, i dati saranno salvati in un file Excel con il nome specificato.
##### Opzioni Avanzate:

Se desideri, puoi anche caricare i dati estratti direttamente in un database MySQL. 
Questa opzione è disponibile dopo il salvataggio dei dati in formato CSV.
##### Caricamento nel Database
Dopo il salvataggio in formato CSV, i dati possono essere caricati in un database MySQL configurato.
#####  Conclusione
Al termine del processo, l'utente riceverà una notifica di successo e avrà la possibilità di chiudere l'applicazione.
(https://github.com/Lorenzozero/AtecoExtractor/assets/77022961/dfd8f28a-650c-467f-9e46-1c999c202775)
(https://github.com/Lorenzozero/AtecoExtractor/assets/77022961/0a178087-802f-435a-864f-0a142d9f9ac6)



