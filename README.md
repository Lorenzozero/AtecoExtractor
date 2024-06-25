
# AtecoExtractor

AtecoExtractor è un'applicazione GUI progettata per l'estrazione e l'integrazione di informazioni dettagliate sulle aziende corrispondenti a specifici codici ATECO da fonti web pubbliche. L'applicazione consente di salvare i dati estratti in un file Excel e, opzionalmente, di caricarli in un database MySQL.

## Caratteristiche Principali

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

- Python 3.x
- MySQL Server
- Biblioteca requests per Python

## Utilizzo

### Inserimento del Codice ATECO

Avvia l'applicazione e inserisci il codice ATECO nel formato XX_X.

### Specifica del Nome del File Excel

Specifica il nome del file Excel in cui desideri salvare i dati estratti.

### Estrazione e Salvataggio dei Dati

Premi il pulsante "Estrai dati" per avviare l'estrazione delle informazioni. Durante l'operazione, una barra di avanzamento mostrerà lo stato dell'estrazione.

### Opzioni Avanzate

Dopo l'estrazione, i dati saranno salvati in un file Excel con il nome specificato. Puoi anche scegliere di caricare i dati estratti direttamente in un database MySQL, opzione disponibile dopo il salvataggio dei dati in formato CSV.

### Caricamento nel Database

I dati estratti in formato CSV possono essere caricati in un database MySQL configurato.

## Conclusione

Al termine del processo, riceverai una notifica di successo e avrai la possibilità di chiudere l'applicazione.

![Screenshot 1](https://github.com/Lorenzozero/AtecoExtractor/assets/77022961/dfd8f28a-650c-467f-9e46-1c999c202775)

![Screenshot 2](https://github.com/Lorenzozero/AtecoExtractor/assets/77022961/0a178087-802f-435a-864f-0a142d9f9ac6)

