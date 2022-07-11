# SCADE
Un sistema di gestione scadenze incredibilmente rudimentale, ma sperabilmente efficace.



## Obiettivo

L'obiettivo di SCADE è mettere a disposizione uno scadenziario condiviso per piccole organizzazioni o team, che si affianchi, senza sostituirle, applicazioni più strutturate.


## Prerequisiti per lo sviluppo

Uno dei prerequisiti è l'utilizzo di strumenti semplici e molto elementari, che sfruttino risorse native di Windows o comunque preinstallate e di utilizzo comune, senza necessità di installazione.

Altro prerequisito, necessario per essere accettato con favore dalla propria organizzazione interna, IT o responsabili della sicurezza, è la comprensibilità del codice, commentato anche in modo prolisso, per favorire una piena consapevolezza di come SCADE funziona.


## Comportamento

Il comportamento di SCADE è piuttosto rudimentale: una volta opportunamente configurato, legge da un file Excel le scadenze di attività proprie o - nel caso di responsabili - dei propri collaboratori, e mostra degli alert, variabilmente attivabili, rispetto alle attività scadute o in scadenza.

## Prerequisiti per l'utilizzo

* Per poter utilizzare SCADE occorre avere la possibilità (e quindi le autorizzazioni)  di lanciare script Powershell. Non occorre lanciare gli script come  amministratori, nè mai SCADE chiederà di eseguire alcunché come amministratore.

  
* Occorre inoltre aver installato Office ed in particolare Excel in una versione  relativamente recente: è necessario per poter accedere e leggere i file Excel  sfruttando le librerie di Office.

* E' importante avere i permessi per la configurazione di task automatici in Pianificazione Attività di Windows (Task Scheduler).

* Occorre infine avere accesso ad un Excel, basato sul modello incluso in SCADE, da manutenere all'interno del team in una cartella condivisa, o sul proprio PC per un utilizzo
solo personale. La struttura ed i nomi di colonna non vanno modificate. Le nuove righe vanno aggiunte in calce.

NOTA: non c'è uno stretto controllo formale sui dati: è essenziale quindi che la compilazione dell'excel sia puntuale ed attenta da parte di tutto il team.


## Dettaglio funzionale

La parte iniziale dello script contiene le variabili di configurazione.
Senz'altro non una delle pratiche migliori (meglio un file dedicato) ma è altrettanto vero che in questo modo si segue il principio di semplicità 
indicato più sopra, includendo in un unico file, eventualmente modificabile all'occorrenza, tutte le funzionalità.
Le variabili sono descritte. Fra queste, il nome del file Excel.

Seguono le funzioni, Show-Notification() per mostrare le notifiche di Windows, e getColumnID(), per il recupero dell'indice di colonna dell'Excel ( quest'ultima
commentata e non utilizzata nella versione corrente di SCADE)


L'area contraddistinta dal figlet MAIN è il punto dove lo script cicla su ciascuna riga della tabella Excel, e, dato l'utente windows corrente, determina e ricorda le scadenze per l'utente, ed ev.te per gli utenti di cui si è responsabili, e le attività scadute non chiuse.


## Dettaglio operativo

E' necessario che un responsabile o un componente del team con una minima competenza in powershell effettui i seguenti passi di configurazione:

1. copia dell'Excel di template al percorso desiderato (percorso locale o di rete) 
2. configurazione sul proprio pc di SCADE.ps1, indicando il percorso nella    variabile $path, ed eventualmente attivando o disattivando i diversi flag, commentati e descritti.
   
3. pulire l'Excel dai dati demo, e inserire una scadenza relativa al proprio    utente (all'interno di un dominio dovrà essere nella forma: NOMEDOMINIO\NOMEUTENTE)
3. verifica del funzionamento, lanciando SCADE.ps1 da powershell
4. una volta soddisfatti del funzionamento, configurare i pc dei colleghi/collaboratori  (o spiegando loro come fare).
5. Configurare per l'esecuzione automatica, ad esempio al mattino e al pomeriggio, di SCADE.ps1 all'interno della Pianificazione Attività (Task Scheduler) di Windows.

## Utilizzo e descrizione dell'Excel

E' importantissimo:
* effettuare frequenti backup dell'Excel
* Aggiungere/modificare le proprie attività evitando di sovrapporsi (ancorché Excel lo possa consentire)
* Aggiornare le attività completate, indicandolo.
* Non cambiare il formato delle celle.
* Sfruttare se possibile tutti i campi disponibili: Ipotesi Recupero e Confidenza Successo consentono di effettuare una valutazione sommaria delle attività più appetibili, e - a parità di scadenza e con poco tempo a disposizione - decidere per quali attività vale la pena proseguire e quali archiviare perché poco redditizie o con scarsa probabilità di successo.

## TODO:

1. Occorre implementare l'inserimento automatico dei Task.
2. Pensare a un semplice script di configurazione.



