# StampaScontrinoTorneo

Sistema di cassa e stampa comande per tornei/sagre.  
Una macro Excel gestisce le vendite e genera file di testo che vengono inviati
a una stampante termica ESC/POS tramite un eseguibile C++.

---

## Indice

1. [Requisiti](#requisiti)
2. [Struttura del progetto](#struttura-del-progetto)
3. [Configurazione stampante](#configurazione-stampante)
4. [Configurazione Excel (Macro VBA)](#configurazione-excel-macro-vba)
   - [Fogli necessari](#fogli-necessari)
   - [Costanti fogli e celle](#costanti-fogli-e-celle)
   - [Percorso EXE di stampa](#percorso-exe-di-stampa)
   - [Dimensioni font di stampa](#dimensioni-font-di-stampa)
   - [Pulsanti prodotto](#pulsanti-prodotto)
   - [Checkbox salse](#checkbox-salse)
5. [Tag di controllo nel file di testo](#tag-di-controllo-nel-file-di-testo)
6. [Flusso di funzionamento](#flusso-di-funzionamento)

---

## Requisiti

- **Windows** con Visual Studio (per compilare l'EXE) oppure l'EXE precompilato
- **Microsoft Excel** con macro abilitate (file `.xlsm`)
- **Stampante termica ESC/POS** raggiungibile via rete (TCP/IP, porta raw 9100)
- Cartella `%USERPROFILE%\Desktop\Comande\` esistente (viene usata per salvare i
  file `.txt` delle comande)

---

## Struttura del progetto

```
StampaScontrinoTorneo/
├── StampaScontrinoTorneo.cpp   # EXE: driver di stampa ESC/POS
├── macro.txt                   # Codice VBA da importare in Excel
└── README.md                   # Questo file
```

---

## Configurazione stampante

L'EXE si collega alla stampante via TCP/IP. L'indirizzo e la porta possono
essere configurati in due modi:

### Opzione 1 – Variabili d'ambiente (consigliato)

Impostare le variabili d'ambiente `PRINTER_IP` e `PRINTER_PORT`.
Se non sono impostate, vengono usati i valori predefiniti.

**Temporanee (sessione corrente di PowerShell):**

```powershell
$env:PRINTER_IP   = "192.168.1.100"
$env:PRINTER_PORT = "9100"
```

**Permanenti (per tutto il sistema):**

```powershell
[System.Environment]::SetEnvironmentVariable("PRINTER_IP",   "192.168.1.100", "User")
[System.Environment]::SetEnvironmentVariable("PRINTER_PORT", "9100",           "User")
```

Oppure: **Impostazioni di sistema** → **Variabili d'ambiente** → Nuova variabile utente.

### Opzione 2 – Valori predefiniti nel codice sorgente

Se le variabili d'ambiente non sono impostate, l'EXE usa questi valori
(modificabili nel file `.cpp`):

```cpp
static const char* DEFAULT_PRINTER_IP   = "192.168.80.121";
static const int   DEFAULT_PRINTER_PORT = 9100;
```

All'avvio l'EXE stampa in console l'indirizzo usato, ad esempio:

```
Printer: 192.168.80.121:9100
```

---

## Configurazione Excel (Macro VBA)

Il codice VBA si trova in `macro.txt`. Va copiato in un modulo standard del
file Excel (`.xlsm`): **Alt+F11** → Inserisci → Modulo → Incolla il contenuto.

### Fogli necessari

Il workbook deve contenere i seguenti fogli:

| Foglio           | Descrizione                                              |
|------------------|----------------------------------------------------------|
| `Config`         | Contatore progressivo numero comanda (cella `B1`)        |
| `Prodotti`       | Tabella `tblProdotti` con nome, prezzo e categoria       |
| `Cassa`          | Interfaccia di vendita con pulsanti e vendita corrente   |
| `Vendite`        | Storico di tutte le vendite registrate                   |
| `Comanda`        | Template usato per costruire il testo di ogni comanda    |
| `ComandaNumero`  | Stampa extra con solo il numero comanda                  |

### Costanti fogli e celle

Tutte le costanti sono raggruppate in cima al modulo VBA. I principali gruppi:

| Costante                             | Valore predefinito | Descrizione                                    |
|--------------------------------------|--------------------|------------------------------------------------|
| `NOME_FOGLIO_CONFIG`                 | `"Config"`         | Nome del foglio configurazione                 |
| `CELLA_ULTIMO_NUMERO_COMANDA`        | `"B1"`             | Cella con l'ultimo numero comanda usato        |
| `NOME_FOGLIO_PRODOTTI`               | `"Prodotti"`       | Nome del foglio prodotti                       |
| `NOME_TABELLA_PRODOTTI`              | `"tblProdotti"`    | Nome della Tabella Excel nel foglio Prodotti   |
| `COL_PRODOTTI_CATEGORIA`             | `3`                | Colonna della categoria (C)                    |
| `NOME_FOGLIO_CASSA`                  | `"Cassa"`          | Nome del foglio cassa                          |
| `COL_PRODOTTO_CASSA`                 | `"E"`              | Colonna nome prodotto nella vendita corrente   |
| `COL_QTA_CASSA`                      | `"F"`              | Colonna quantità                               |
| `COL_PREZZO_UNITARIO_CASSA`          | `"G"`              | Colonna prezzo unitario                        |
| `COL_SUBTOTALE_CASSA`                | `"H"`              | Colonna subtotale                              |
| `CELLA_TOTALE_VENDITA_CORRENTE_CASSA`| `"H21"`            | Cella con la formula SOMMA del totale          |
| `CELLA_IMPORTO_PAGATO_CASSA`         | `"H22"`            | Cella importo pagato dal cliente               |

> **Per modificare il layout del foglio Cassa**, basta aggiornare le costanti
> corrispondenti senza toccare il resto del codice.

### Percorso EXE di stampa

Nella costante `PERCORSO_EXE_STAMPA` è specificato il percorso completo
dell'eseguibile compilato:

```vba
Const PERCORSO_EXE_STAMPA As String = _
	"C:\Users\WiZ\Desktop\StampaScontrino\StampaScontrinoTorneo\x64\Debug\StampaScontrinoTorneo.exe"
```

**Aggiornare questo percorso** se si sposta l'EXE o si compila in modalità Release.

### Dimensioni font di stampa

Le dimensioni di ogni sezione dello scontrino sono configurabili tramite
costanti nel formato `"W,H"`:

| Costante           | Valore | Sezione dello scontrino                     |
|--------------------|--------|---------------------------------------------|
| `FONT_INTESTAZIONE`| `"1,1"`| Numero comanda nell'intestazione (2×2)      |
| `FONT_DETTAGLI`    | `"0,1"`| Titolo comanda, data, separatori (1×2)      |
| `FONT_ORDINE`      | `"1,1"`| Testo "Ordine N. ###" (2×2)                |
| `FONT_ITEMS`       | `"0,0"`| Righe prodotti (1×1, normale)               |
| `FONT_TOTALE`      | `"0,1"`| Riga totale (1×2)                           |

**W** = moltiplicatore larghezza, **H** = moltiplicatore altezza.  
I valori vanno da `0` (1×) a `7` (8×).

**Esempio:** per rendere i prodotti più grandi, cambiare:

```vba
Const FONT_ITEMS = "0,1"   ' da 1x1 a 1x2 (doppia altezza)
```

### Pulsanti prodotto

Ogni prodotto ha una Sub dedicata collegata a un pulsante sul foglio Cassa.
Per **aggiungere un nuovo prodotto**:

1. Aggiungere il prodotto nella tabella `tblProdotti` (nome, prezzo, categoria)
2. Aggiungere una nuova Sub nel modulo VBA:
   ```vba
   Sub Pulsante_NuovoProdotto_Click(): Call AddSpecificItemToSale("Nome Prodotto"): End Sub
   ```
3. Inserire un pulsante nel foglio Cassa e assegnargli la macro

Il nome passato a `AddSpecificItemToSale` deve corrispondere **esattamente**
al nome nella tabella `tblProdotti`.

### Checkbox salse

Il pulsante "Patate Fritte" controlla automaticamente le checkbox salse.
Le celle di collegamento sono configurabili:

| Costante              | Cella  | Salsa            |
|-----------------------|--------|------------------|
| `CELLA_CHECK_KETCHUP` | `AB1`  | Ketchup (extra)  |
| `CELLA_CHECK_MAIONESE`| `AB2`  | Maionese (extra) |
| `CELLA_CHECK_SENAPE`  | `AB3`  | Senape (extra)   |

---

## Tag di controllo nel file di testo

La macro genera file `.txt` con tag speciali che l'EXE interpreta come
comandi ESC/POS. Ogni tag deve essere su una riga a sé:

### Font

```
@@FONT:W,H@@
```

Imposta la dimensione carattere. `W` = larghezza (0-7), `H` = altezza (0-7).

Sono disponibili anche i preset con nome:

| Tag           | Dimensione         |
|---------------|--------------------|
| `@@LARGE@@`   | 2× larghezza, 2× altezza |
| `@@MEDIUM@@`  | 1× larghezza, 2× altezza |
| `@@SMALL@@`   | 1× larghezza, 1× altezza |

### Allineamento

| Tag            | Effetto                              |
|----------------|--------------------------------------|
| `@@CENTER@@`   | Centra il testo successivo           |
| `@@LEFT@@`     | Allinea a sinistra (predefinito)     |

### Taglio carta

```
====
```

Una riga che inizia con `====` fa avanzare la carta e taglia.

---

## Flusso di funzionamento

```
 Pulsante prodotto
		│
		▼
 AddSpecificItemToSale()     ← aggiunge/incrementa riga nel foglio Cassa
		│
		▼
 FinalizzaVendita()          ← pulsante "Finalizza"
		│
		├─ Classifica gli articoli (cibo, stinco, focaccia, bevande, staff birra)
		├─ Genera comande separate per ogni tipologia (BuildComandaText)
		├─ Registra tutto nel foglio Vendite
		├─ Pulisce la cassa
		│
		▼
 ScriviFileComandaTesto()    ← salva il file .txt in Desktop\Comande\
		│
		▼
 StampaScontrinoTorneo.exe   ← legge il .txt, invia i comandi ESC/POS alla stampante
		│
		▼
	Stampante termica         ← stampa le comande con tagli separati
```

### Categorie comande

La macro genera comande separate (con taglio carta fra ognuna) in base alla
tipologia:

| Comanda              | Filtro                                                  |
|----------------------|---------------------------------------------------------|
| COMANDA CUCINA       | Categoria "cibo", esclusi stinco/focaccia/staff birra   |
| COMANDA STINCO       | Nome contiene "stinco"                                  |
| COMANDA FOCACCE      | Nome contiene "focaccia"                                |
| COMANDA BEVANDE      | Categoria "bere", escluso "staff birra"                 |

**Staff Birra** viene registrato nelle vendite ma **non** genera comanda.

### Annullamento

Il pulsante `AnnullaVenditaCorrente` svuota la vendita corrente senza
registrarla nel foglio Vendite.
