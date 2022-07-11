
########################################################################################################
#                                                                                                      #
#        Gabriele Motta - SCADE - v.0.7 beta                                                           #
#        https://github.com/trincio/SCADE                                                              #
#                                                                                                      #
#                                                                                                      #
#        Obiettivo                                                                                     #
#                                                                                                      #
#        L'obiettivo di SCADE è mettere a disposizione uno scadenziario condiviso                      #
#        per piccole organizzazioni o team, che si affianchi, senza sostituirle,                       #
#        applicazioni più strutturate.                                                                 #
#                                                                                                      #
#                                                                                                      #
#        Prerequisiti per lo sviluppo                                                                  #
#                                                                                                      #
#        Uno dei prerequisiti è l'utilizzo di strumenti semplici e molto elementari, che               #
#        sfruttino risorse native di Windows o comunque preinstallate                                  #
#        e di utilizzo comune, senza necessità di installazione.                                       #
#                                                                                                      #
#        Altro prerequisito, necessario per essere accettato con favore dalla                          #
#        propria organizzazione interna, IT o responsabili della sicurezza,                            #
#        è la comprensibilità del codice, commentato anche in modo prolisso,                           #
#        per favorire una piena consapevolezza di come SCADE funziona.                                 #
#                                                                                                      #
#                                                                                                      #
#        Comportamento                                                                                 #
#                                                                                                      #
#        Il comportamento di SCADE è piuttosto rudimentale: una volta opportunamente                   #
#        configurato, legge da un file Excel le scadenze di attività proprie o - nel                   #
#        caso di responsabili - dei propri collaboratori, e mostra degli alert,                        #
#        variabilmente attivabili, rispetto alle attività scadute o in scadenza.                       #
#                                                                                                      #
#        Prerequisiti per l'utilizzo                                                                   #
#                                                                                                      #
#        Per poter utilizzare SCADE occorre avere la possibilità (e quindi le autorizzazioni)          #
#         di lanciare script Powershell. Non occorre lanciare gli script come                          #
#         amministratori, nè mai SCADE chiederà di eseguire alcunché come amministratore.              #
#                                                                                                      #
#                                                                                                      #
#        Occorre inoltre aver installato Office ed in particolare Excel in una versione                #
#         relativamente recente: è necessario per poter accedere e leggere i file Excel                #
#         sfruttando le librerie di Office.                                                            #
#                                                                                                      #
#        E' importante avere i permessi per la configurazione di task automatici in Pianificazione     #
#        Attività di Windows (Task Scheduler).                                                         #
#                                                                                                      #
#        Occorre infine avere accesso ad un Excel, basato sul modello incluso in SCADE, da manutenere  #
#        all'interno del team in una cartella condivisa, o sul proprio PC per un utilizzo              #
#        solo personale. La struttura ed i nomi di colonna non vanno modificate.                       #
#        Le nuove righe vanno aggiunte in calce.                                                       #
#        Non c'è uno stretto controllo formale sui dati: è essenziale quindi che la compilazione       #
#        dell'excel sia puntuale ed attenta da parte di tutto il team.                                 #
#                                                                                                      #
#                                                                                                      #
#        Dettaglio funzionale                                                                          #
#                                                                                                      #
#        La parte iniziale dello script contiene le variabili di configurazione.                       #
#        Senz'altro non una delle pratiche migliori (meglio un file dedicato) ma                       #
#        è altrettanto vero che in questo modo si segue il principio di semplicità                     #
#        indicato più sopra, includendo in un unico file, eventualmente modificabile                   #
#        all'occorrenza, tutte le funzionalità.                                                        #
#        Le variabili sono descritte. Fra queste, il nome del file Excel.                              #
#                                                                                                      #
#        Seguono le funzioni, Show-Notification() per mostrare le notifiche di Windows,                #
#        e getColumnID(), per il recupero dell'indice di colonna dell'Excel ( quest'ultima             #
#        commentata e non utilizzata nella versione corrente di SCADE)                                 #
#                                                                                                      #
#                                                                                                      #
#        L'area contraddistinta dal figlet MAIN è il punto dove lo script cicla                        #
#        su ciascuna riga della tabella Excel, e, dato l'utente windows corrente,                      #
#        determina e ricorda le scadenze per l'utente, ed ev.te per gli utenti di cui si è             #
#        responsabili, e le attività scadute non chiuse.                                               #
#                                                                                                      #
#                                                                                                      #
#        Dettaglio operativo                                                                           #
#                                                                                                      #
#        E' necessario che un responsabile o un componente del team con una minima                     #
#        competenza in powershell effettui i seguenti passi di configurazione:                         #
#                                                                                                      #
#        1. copia dell'Excel di template al percorso desiderato (percorso locale o di rete)            #
#        2. configurazione sul proprio pc di SCADE.ps1, indicando il percorso nella                    #
#           variabile $path, ed eventualmente attivando o disattivando i diversi flag,                 #
#           commentati e descritti.                                                                    #
#                                                                                                      #
#        3. pulire l'Excel dai dati demo, e inserire una scadenza relativa al proprio                  #
#           utente (all'interno di un dominio dovrà essere nella forma: NOMEDOMINIO\NOMEUTENTE)        #
#        3. verifica del funzionamento, lanciando SCADE.ps1 da powershell                              #
#        4. una volta soddisfatti del funzionamento, configurare i pc dei colleghi/collaboratori       #
#           (o spiegando loro come fare).                                                              #
#        5. Configurare per l'esecuzione automatica, ad esempio al mattino e al pomeriggio,            #
#           di SCADE.ps1 all'interno della Pianificazione Attività (Task Scheduler) di Windows.        #
#                                                                                                      #
########################################################################################################


 


cls


#NOTA BENE: IL TERMINE NOTIFICA VIENE UTILIZZATO IN QUESTA SEDE, NEI COMMENTI, AD INDICARE LE NOTIFICHE WINDOWS

$oldci = [System.Threading.Thread]::CurrentThread.CurrentCulture
$newci = [System.Globalization.CultureInfo]"it-IT"
$worksheetName = "Scadenziariotest"


#Giorni per soglie scadenza. Default a 15 e 30 giorni.

$daysbeforealertlev1 = 15

$daysbeforealertlev2 = 30

#Flag per notifica Windows, diversificata per tipo. Se impostati a false ignorano la notifica Windows. Default tutti true.

$_FLAG_NOTIFY_EXPIRED = $true       #notifica scadenza

$_FLAG_NOTIFY_1ST_ALERT = $true     #notifica primo termine

$_FLAG_NOTIFY_2ND_ALERT = $true     #notifica secondo termine



$_FLAG_NOTIFY_EXPIRED = $false   
                                
$_FLAG_NOTIFY_1ST_ALERT = $false 
                                
$_FLAG_NOTIFY_2ND_ALERT = $false 



#Flag per notifica via popup.
$_FLAG_POPUP = $true



#carattere per wrap/ritorno a capo
$wrap ="`r`n"


#offset indicante l'inizio dei dati.
$offsetData= 1 

$path=".\SCADENZIARIOtest.xlsx"

#se non sono presenti le librerie, usare psexcel https://www.c-sharpcorner.com/article/read-excel-file-using-psexcel-in-powershell2/
#altro suggerimento https://github.com/maravento/winzenity




function Show-Notification {
    [cmdletbinding()]
    Param (
        [string]
        $ToastTitle,
        [string]
        [parameter(ValueFromPipeline)]
        $ToastText
    )

    [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
    $Template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)
    #$Template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastImageAndText02)

    $RawXml = [xml] $Template.GetXml()
    ($RawXml.toast.visual.binding.text|where {$_.id -eq "1"}).AppendChild($RawXml.CreateTextNode($ToastTitle)) > $null
    ($RawXml.toast.visual.binding.text|where {$_.id -eq "2"}).AppendChild($RawXml.CreateTextNode($ToastText)) > $null

    $SerializedXml = New-Object Windows.Data.Xml.Dom.XmlDocument
    $SerializedXml.LoadXml($RawXml.OuterXml)

    $Toast = [Windows.UI.Notifications.ToastNotification]::new($SerializedXml)
    $Toast.Tag = "PowerShell"
    $Toast.Group = "PowerShell"
    $Toast.ExpirationTime = [DateTimeOffset]::Now.AddMinutes(1)

    $Notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("PowerShell")
    $Notifier.Show($Toast);
}


	
# function getColumnID {
# 	
#     param (
# 	[Parameter()]
# 	
#         $worksheet,
#         $nameToSearch
#     )	
# 	
# 	
# Write-Host getColumnID su  $worksheet e  $nameToSearch
# $columnNamesRow = 1           # or whichever row your names are in
# #$nameToSearch = "Fab Hours"   # or whatever name you want to search for
# $columnToUse = -1
# $columnsCount = ($WorkSheet.UsedRange.Columns).count # $worksheet.Cells(1, $worksheet.Columns.Count).End(xlToLeft).Column
# Write-Host individuate $columnsCount colonne attive
# 
# 
# For ( $col = 1; $col -le $columnsCount; $col++){
# 	
#  
# 	
# 	$n = $worksheet.Cells($columnNamesRow, $col).Name
#    $ColumnName = $worksheet.columns($col).Name
#    Write-Host valore $n and $ColumnName
#    If ($worksheet.Cells($columnNamesRow, $col).Name -eq  $nameToSearch) {  $columnToUse = $col;}
# 
# 
# 	
# }
# 
# Write-Host -----------------------------------------
# Write-Host
# 
# return $columnToUse;	
# 	
# 	
# }
 
 
#       __  __    _    ___ _   _
#      |  \/  |  / \  |_ _| \ | |
#      | |\/| | / _ \  | ||  \| |
#      | |  | |/ ___ \ | || |\  |
#      |_|  |_/_/   \_|___|_| \_|
#      

#Istanzia nuovo oggetto Application Excel
$objExcel = New-Object -ComObject "Excel.Application"


#Nasconde l'oggetto, ne configura la parte internazionale, apre il file al percorso $path e apre il foglio $worksheetName
$objExcel.visible = $false
[system.threading.Thread]::CurrentThread.CurrentCulture = $newci
$objWorkbook =$objExcel.Workbooks.Open($path)  
$worksheet = $objWorkbook.sheets.item($worksheetName)

 
$listObject = $worksheet.ListObjects | where-object { $_.DisplayName -eq "Scadenziario" }    

 
$rowCount=$listObject.listRows.Count

Write-Host Numero righe individuate:  -$rowCount-


#GETCOLUMNID() va resa funzionante per poter accedere ai valori tramite nome di colonna e non tramite indice.
#chiama la funzione getColumnID, per individuare all'interno del foglio, la posizione della colonna 
#$description = getColumnID $worksheet "Descrizione Pratica" ;
#$description = getColumnID $worksheet "Scadenza" ;


#Recupero nome utente corrente per verifica
$currentuser=[System.Security.Principal.WindowsIdentity]::GetCurrent().Name



Write-Output   ""
Write-Output   "Vengono prese in considerazione unicamente le righe aventi come responsabile dell'attività l'utente: $currentuser"
Write-Output   ""



### Messaggi

    $EXPIRED_TASKS_TEXT = ""
    $FIRSTALERT_TEXT = ""
    $SECONDALERT_TEXT = ""


### Messaggi per il responsabile

    $MANAGER_EXPIRED_TASKS_TEXT = ""
    $MANAGER_FIRSTALERT_TEXT = ""
    $MANAGER_SECONDALERT_TEXT = ""



##Ciclo sulle date per verificare le scadute
 
##CAVEAT 1. Versione piuttosto rudimentale: accedo tramite indice della colonna, non tramite nome.
##CAVEAT 2. Utilizzo di Cells, non usa ListObject. 
for ($row = 1+$offsetData; $row -le $rowCount+$offsetData; $row++) {


    #per comodità recupero il valore in giorni (value2)
    $activitydate_number = $WorkSheet.Cells.Item($row, 1).value2
    $activitydate = [DateTime]::FromOADate($activitydate_number)
    $txtactivitydate =$activitydate.ToString("dd-MM-yyyy")

    $id = $WorkSheet.Cells.Item($row, 2).Text

    $customer = $WorkSheet.Cells.Item($row, 3).Text

    $description = $WorkSheet.Cells.Item($row, 4).Text

    $amount = $WorkSheet.Cells.Item($row, 5).value2

    $success = $WorkSheet.Cells.Item($row, 6).value2
 
    $status = $WorkSheet.Cells.Item($row, 7).Text

    $note = $WorkSheet.Cells.Item($row, 8).Text

    $owner = $WorkSheet.Cells.Item($row, 9).Text

    $coordinator = $WorkSheet.Cells.Item($row, 10).Text

    $lastmodified = $WorkSheet.Cells.Item($row, 11).Text

    $lastmodifier = $WorkSheet.Cells.Item($row, 12).Text



	#calcolo del valore recuperabile sulla base del valore totale e della confidenza
    $EXPECTED_VALUE=$amount*$success
    
 
 
    #valori data per confronto con i valori provenienti dall'Excel
 
	$Today = (Get-Date)  

    $datealert1 = (Get-Date).AddDays($daysbeforealertlev1)

    $datealert2 = (Get-Date).AddDays($daysbeforealertlev2)

    

 

    #Per velocizzare vengono salvati i testi in variabili da utilizzare nella generazione della notifica.
    #TODO: Valutare l'utilizzo di array


   
    #SE L'UTENTE RESPONSABILE DELL'ATTIVITA' ALL'INTERNO DELLA RIGA EXCEL CORRISPONDE ALL'UTENTE WINDOWS CORRENTE 
    if  ($currentuser -eq  $owner){

    
		#SEGUONO LE VERIFICHE SULLE DATE PER INDIVIDUARE ATTIVITA' SCADUTE O IN SCADENZA
        
        if($today -ge $activitydate){
            
            $currtext = "$row ) La tua attività '$id' sul soggetto '$customer' è scaduta in data $txtactivitydate. Hai perso € $EXPECTED_VALUE che sarebbero stati ragionevolmente ottenibili."
            $EXPIRED_TASKS_TEXT = $EXPIRED_TASKS_TEXT + $currtext + $wrap

            #se il flag _FLAG_NOTIFY_EXPIRED è impostato a true aggiunge una notifica
            if ($_FLAG_NOTIFY_EXPIRED){ Show-Notification -ToastTitle "PRATICA SCADUTA" -ToastText  $currtext}
    
            }    

        elseif($datealert1 -gt $activitydate ){
            $currtext = "$row ) La tua attività '$id' sul soggetto '$customer' su cui sono ragionevolmente ottenibili € $EXPECTED_VALUE sta per scadere in data $txtactivitydate"
            $FIRSTALERT_TEXT = $FIRSTALERT_TEXT + $currtext + $wrap

            #se il flag _FLAG_NOTIFY_1ST_ALERT è impostato a true aggiunge una notifica
            if ($_FLAG_NOTIFY_1ST_ALERT){ Show-Notification -ToastTitle "PRATICA IN SCADENZA" -ToastText  $currtext}
  
          }      

        elseif($datealert2 -gt $activitydate ){
          
          $currtext = "$row ) La tua attività '$id' sul soggetto '$customer' su cui sono ragionevolmente ottenibili € $EXPECTED_VALUE sta per scadere in data $txtactivitydate"
            $SECONDALERT_TEXT = $SECONDALERT_TEXT + $currtext + $wrap

             #se il flag _FLAG_NOTIFY_2ND_ALERT è impostato a true aggiunge una notifica
            if ($_FLAG_NOTIFY_2ND_ALERT) { Show-Notification -ToastTitle "PRATICA IN SCADENZA" -ToastText  $currtext}

          }      
  
 
    }
	


 #SE L'UTENTE COORDINATORE ALL'INTERNO DELLA RIGA EXCEL CORRISPONDE ALL'UTENTE WINDOWS CORRENTE
    if  ($currentuser -eq  $coordinator){


		#SEGUONO LE VERIFICHE SULLE DATE PER INDIVIDUARE ATTIVITA' SCADUTE O IN SCADENZA
    
        if($today -ge $activitydate){
            
            $currtext = "$row ) L'attività '$id' in carico a $owner sul soggetto '$customer' è scaduta in data $txtactivitydate. Stima: € $EXPECTED_VALUE (valore:  € $amount con stima successo: $success )."
            $MANAGER_EXPIRED_TASKS_TEXT = $EXPIRED_TASKS_TEXT + $currtext + $wrap

            #se il flag _FLAG_NOTIFY_EXPIRED è impostato a true aggiunge una notifica
            if ($_FLAG_NOTIFY_EXPIRED){ Show-Notification -ToastTitle "PRATICA SCADUTA PER UN TUO COLLABORATORE " -ToastText  $currtext}
    
            }    

        elseif($datealert1 -gt $activitydate ){
            $currtext = "$row ) L'attività '$id' in carico a $owner sul soggetto '$customer' su cui sono ragionevolmente ottenibili € $EXPECTED_VALUE sta per scadere in data $txtactivitydate"
            $MANAGER_FIRSTALERT_TEXT = $FIRSTALERT_TEXT + $currtext + $wrap

            #se il flag _FLAG_NOTIFY_1ST_ALERT è impostato a true aggiunge una notifica
            if ($_FLAG_NOTIFY_1ST_ALERT){ Show-Notification -ToastTitle "PRATICA IN SCADENZA PER UN TUO COLLABORATORE" -ToastText  $currtext}
  
          }      

        elseif($datealert2 -gt $activitydate ){
          
          $currtext = "$row ) L'attività '$id' in carico a $owner sul soggetto '$customer' su cui sono ragionevolmente ottenibili € $EXPECTED_VALUE sta per scadere in data $txtactivitydate"
            $MANAGER_SECONDALERT_TEXT = $SECONDALERT_TEXT + $currtext + $wrap

             #se il flag _FLAG_NOTIFY_2ND_ALERT è impostato a true aggiunge una notifica
            if ($_FLAG_NOTIFY_2ND_ALERT) { Show-Notification -ToastTitle "PRATICA IN SCADENZA PER UN TUO COLLABORATORE" -ToastText  $currtext}

          }      
  
 
    }


 
}


#CHIUSURA DELL'EXCEL: E' FONDAMENTALE CHE L'EXCEL NON RIMANGA APERTO INDEFINITAMENTE

$objWorkbook.Close() 


#DEFINIZIONE DEI MESSAGGI E PRESENTAZIONE ALL'UTENTE SECONDO LE MODALITA' SELEZIONATE OCN LE VARIABILI-FLAG NELL'HEADER DELLO SCRIPT

$AllNotification  ="Dati relativi al file Excel al percorso: $path" +  "$wrap"
$AllNotification  =$AllNotification   +" ---- ACCERTATI SEMPRE CHE IL FILE SIA AGGIORNATO E NE VENGA EFFETTUATO REGOLARMENTE IL BACKUP ----"
$AllNotification  = $AllNotification  + "$wrap" +  "$wrap"

$PersonalHeader =                   "                      +--------------------------+         " +   "$wrap"
$PersonalHeader = $PersonalHeader + "                      |   ATTIVITA' PERSONALI    |         " +   "$wrap"
$PersonalHeader = $PersonalHeader + "                      +--------------------------+         " +   "$wrap" +   "$wrap"



$personalNote =                 "---- ATTIVITA' SCADUTE:" + "$wrap" + "$EXPIRED_TASKS_TEXT" + "$wrap" +   "$wrap"
$personalNote = $personalNote + "---- ATTIVITA' IN SCADENZA ENTRO I PROSSIMI $daysbeforealertlev1 GIORNI:" + "$wrap" + "$FIRSTALERT_TEXT"    + "$wrap" +   "$wrap"
$personalNote = $personalNote + "---- ATTIVITA' IN SCADENZA ENTRO I PROSSIMI $daysbeforealertlev2 GIORNI:" + "$wrap" + "$SECONDALERT_TEXT"   + "$wrap" +   "$wrap"


$coworkersHeader =                    "               +--------------------------------------+   " +   "$wrap"
$coworkersHeader = $coworkersHeader + "               |   EV.LI ATTIVITA' DI COLLABORATORI   |   " +   "$wrap"
$coworkersHeader = $coworkersHeader + "               +--------------------------------------+   " +   "$wrap" +   "$wrap"

$coworkersNote =                  "---- ATTIVITA' SCADUTE:" + "$wrap" + "$MANAGER_EXPIRED_TASKS_TEXT" + "$wrap" +   "$wrap"
$coworkersNote = $coworkersNote + "---- ATTIVITA' IN SCADENZA ENTRO I PROSSIMI $daysbeforealertlev1 GIORNI:" + "$wrap" + "$MANAGER_FIRSTALERT_TEXT"    + "$wrap" +   "$wrap"
$coworkersNote = $coworkersNote + "---- ATTIVITA' IN SCADENZA ENTRO I PROSSIMI $daysbeforealertlev2 GIORNI:" + "$wrap" + "$MANAGER_SECONDALERT_TEXT"   + "$wrap" +   "$wrap"


Write-Output $AllNotification + $PersonalHeader + $personalNote + $coworkersHeader + $coworkersNote




if($_FLAG_POPUP) { [System.Windows.MessageBox]::Show($AllNotification + $personalNote  + $coworkersNote)}

$rightnow = (Get-Date). ToString("yyyyMMdd_hhmmss")
$filename = "$rightnow"+"_SCADENZE.txt"


echo $AllNotification + $PersonalHeader + $personalNote + $coworkersHeader + $coworkersNote > $filename

start notepad $filename

exit




 