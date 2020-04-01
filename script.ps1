param (
[String]$server,
[String]$datenbank,
[String]$benutzername,
[String]$passwort,
[switch]$autoImport,
[switch]$hidden
)
##### Parameter entgegennehmen ^ #####

$path = $PSScriptRoot
function Write-Files(){
    ########## Dateien erstellen Funtkion##########

    # Variablen definieren
    $outpath = "$path\output\"
    $adresse_file = "$outpath\adresse.csv"
    $betreuende_file = "$outpath\betreuende.csv"
    $elternteil_file = "$outpath\elternteil.csv"
    $elternteil_zu_kind_file = "$outpath\elternteil_zu_kind.csv"
    $kind_file = "$outpath\kind.csv"
    $kita_file = "$outpath\kita.csv"
    $schuleinheit_file = "$outpath\schuleinheit.csv"

    # CSV Pfade festlegen
    $csv1 = "$path\src\data1.csv"
    $csv2 = "$path\src\data2.csv"
    
    Write-Log "Dateien werden importiert..."

    # Dateien importieren
    $data1 = Import-Csv -Path $csv1 -Delimiter ','
    $data2 = Import-Csv -Path $csv2 -Delimiter ','

    # Aufteilen in verschiedene Variablen
    $adresse = $data1 | Select-Object "strasse","haus_nr*","plz","ort"
    $betreuende = $data1 | Select-Object @{N="vorname"; E={$_.betr_vorname}},@{N="nachname"; E={$_.betr_nachname}},@{N="geburtsdatum"; E={$_.betr_geburtsdatum}},@{N="tel_nr"; E={$_.betr_tel_nr}},"kita_name" -Unique | Where-Object {$_.vorname -ne "NULL" -and $_.nachname -ne "NULL" -and $_.geburtsdatum -ne "NULL" -and $_.tel_nr -ne "NULL"}
    $elternteil = $data2 | Select-Object @{N="vorname"; E={$_.elt_vorname}},@{N="nachname"; E={$_.elt_nachname}},@{N="geschlecht"; E={$_.elt_geschlecht}},@{N="strasse"; E={$_.elt_strasse}},@{N="haus_nr"; E={$_.elt_haus_nr}},@{N="haus_nr_zusatz"; E={$_.elt_haus_nr_zusatz}},@{N="plz"; E={$_.elt_plz}},@{N="ort"; E={$_.elt_ort}} -Unique
    $kind = $data2 | Select-Object @{N="vorname"; E={$_.kind_vorname}},@{N="geburtsdatum"; E={$_.kind_geburtsdatum}},@{N="geschlecht"; E={$_.kind_geschlecht}},"kita_name" -Unique | Where-Object {$_.vorname -ne "NULL"}
    $elternteil_zu_kind = $data2 | Select-Object @{N="vorname"; E={$_.kind_vorname}},@{N="geburtsdatum"; E={$_.kind_geburtsdatum}},@{N="geschlecht"; E={$_.kind_geschlecht}},"elt*" | Where-Object {$_.vorname -ne "NULL"}
    $kita = $data1 | Select-Object "kita_name","strasse","haus_nr*","plz","ort","Schuleinheit_Name" -Unique | Where-Object {$_.kita_name -ne "NULL"}
    $schuleinheit = $data1 | Select-Object "schuleinheit_name" -Unique | Where-Object {$_.schuleinheit_name -ne "NULL"}


    ################################
    # Neue Spalten und Foreign Keys
    ################################
    
    # Adresse ######################
    Write-Log "Tabelle Adresse wird erstellt..."
    $i = 1
    foreach ($line in $adresse){ #Für jede Zeile in der Tabelle "Adresse"
        $line | Add-Member -NotePropertyName "adresse_id" -NotepropertyValue $i #ID erstellen und hinzufügen (neue Spalte)
        $i++
    }
    # Schuleinheit #################
    Write-Log "Tabelle Schuleinheit wird erstellt..."
    $i = 1
    foreach ($line in $schuleinheit){
        $line | Add-Member -NotePropertyName "schuleinheit_id" -NotepropertyValue $i
        $i++
    }
    # Elternteil ###################
    Write-Log "Tabelle Elternteil wird erstellt..."
    $i = 1
    foreach ($line in $elternteil){
        $line | Add-Member -NotePropertyName "elternteil_id" -NotepropertyValue $i
        $i++
    }
    foreach ($line in $elternteil){
        #FK_Adresse_ID herauslesen
        $fk_adresse_id = $adresse | 
            Where-Object {$_.strasse -eq $line.strasse -and $_.plz -eq $line.plz -and $_.ort -eq $line.ort -and $_.haus_nr -eq $line.haus_nr -and $_.Haus_Nr_Zusatz -eq $line.haus_nr_zusatz} |
            Select-Object -ExpandProperty "adresse_id"
        
        $line | Add-Member -NotePropertyName "fk_adresse_id" -NotepropertyValue $fk_adresse_id
    }


    # Kita #########################
    Write-Log "Tabelle Kita wird erstellt..."
    $i = 1
    foreach ($line in $kita){
        $line | Add-Member -NotePropertyName "kita_id" -NotepropertyValue $i
        $i++
    }
    foreach ($line in $kita){
        #FK_Adresse_ID herauslesen
        $fk_adresse_id = $adresse | 
            Where-Object {$_.strasse -eq $line.strasse -and $_.plz -eq $line.plz -and $_.ort -eq $line.ort -and $_.haus_nr -eq $line.haus_nr -and $_.Haus_Nr_Zusatz -eq $line.haus_nr_zusatz} |
            Select-Object -ExpandProperty "adresse_id"
        
        #FK_Schuleinheit_ID
        $fk_schuleinheit_id = $schuleinheit | 
            Where-Object {$_.schuleinheit_name -eq $line.schuleinheit_name} | 
            Select-Object -ExpandProperty "schuleinheit_id"

        #Beide IDs hinzufügen
        $line | Add-Member -NotePropertyName "fk_adresse_id" -NotepropertyValue $fk_adresse_id
        $line | Add-Member -NotePropertyName "fk_schuleinheit_id" -NotepropertyValue $fk_schuleinheit_id
    }

    # Kind #########################
    Write-Log "Tabelle Kind wird erstellt..."
    $i = 1
    foreach ($line in $kind){
        $line | Add-Member -NotePropertyName "kind_id" -NotepropertyValue $i
        $i++
    }
    foreach ($line in $kind){

        #fk_kita_id herauslesen
        $fk_kita_id = $kita | 
            Where-Object {$_.kita_name -eq $line.kita_name} |
            Select-Object -ExpandProperty "kita_id"
        
        $line | Add-Member -NotePropertyName "fk_kita_id" -NotepropertyValue $fk_kita_id
    }

    # Elternteil_zu_Kind ###########
    Write-Log "Tabelle Elternteil_zu_Kind wird erstellt..."
    $i = 1
    foreach ($line in $elternteil_zu_kind){
        $line | Add-Member -NotePropertyName "elternteil_zu_kind_id" -NotepropertyValue $i
        $i++
    }
    foreach ($line in $elternteil_zu_kind){
        #fk_kind_id herauslesen
        $fk_kind_id = $kind | 
        Where-Object {$_.vorname -eq $line.vorname -and $_.geburtsdatum -eq $line.geburtsdatum -and $_.geschlecht -eq $line.geschlecht -and $_.haus_nr -eq $line.haus_nr -and $_.Haus_Nr_Zusatz -eq $line.haus_nr_zusatz} |
        Select-Object -ExpandProperty "kind_id"

        #fk_adresse_id auslesen
        $fk_adresse_id = $adresse | 
        Where-Object {$_.strasse -eq $line.elt_strasse -and $_.plz -eq $line.elt_plz -and $_.ort -eq $line.elt_ort -and $_.haus_nr -eq $line.elt_haus_nr -and $_.Haus_Nr_Zusatz -eq $line.elt_haus_nr_zusatz} |
        Select-Object -ExpandProperty "adresse_id"

        #fk_elternteil_id auslesen
        $fk_elternteil_id = $elternteil | 
        Where-Object {$_.vorname -eq $line.elt_vorname -and $_.nachname -eq $line.elt_nachname -and $_.geschlecht -eq $line.elt_geschlecht -and $_.fk_adresse_id -eq $fk_adresse_id} |
        Select-Object -ExpandProperty "elternteil_id"

        $line | Add-Member -NotePropertyName "fk_kind_id" -NotepropertyValue $fk_kind_id
        $line | Add-Member -NotePropertyName "fk_elternteil_id" -NotepropertyValue $fk_elternteil_id
        
    }
    
    # Betreuende ###################
    Write-Log "Tabelle Betreuende wird erstellt..."
    $i = 1
    foreach ($line in $betreuende){
        $line | Add-Member -NotePropertyName "betreuende_id" -NotepropertyValue $i
        $i++
    }
    foreach ($line in $betreuende){
        #fk_kita_id herauslesen
        $fk_kita_id = $kita | 
            Where-Object {$_.kita_name -eq $line.kita_name} |
            Select-Object -ExpandProperty "kita_id"
        
        $line | Add-Member -NotePropertyName "fk_kita_id" -NotepropertyValue $fk_kita_id
    }


    ##################################
    # Variablen in Dateien exportieren
    ##################################
    Write-Log "Dateien werden exportiert"
    if (!(Test-Path $outpath)){
        mkdir -path $outpath
    }

    # Adresse
    $adresse | Select-Object "adresse_id","strasse","haus_nr","haus_nr_zusatz","plz","ort" | 
    Export-Csv -Path $adresse_file -NoTypeInformation -Delimiter ',' -Encoding UTF8
    #ConvertTo-Csv -NoTypeInformation -Delimiter ',' | # Convert to CSV string data without the type metadata
    #Select-Object -Skip 1 | # Trim header row, leaving only data columns
    #Out-File -FilePath $adresse_file -Encoding utf8

    # Schuleinheit
    $schuleinheit | Select-Object "schuleinheit_id","schuleinheit_name" | 
    Export-Csv -Path $schuleinheit_file -NoTypeInformation -Delimiter ',' -Encoding UTF8
    #ConvertTo-Csv -NoTypeInformation -Delimiter ',' | # Convert to CSV string data without the type metadata
    #Select-Object -Skip 1 | # Trim header row, leaving only data columns
    #Out-File -FilePath $schuleinheit_file -Encoding utf8

    # Kita
    $kita | Select-Object "kita_id","kita_name","fk_adresse_id","fk_schuleinheit_id" | 
    Export-Csv -Path $kita_file -NoTypeInformation -Delimiter ',' -Encoding UTF8
    #ConvertTo-Csv -NoTypeInformation -Delimiter ',' | # Convert to CSV string data without the type metadata
    #Select-Object -Skip 1 | # Trim header row, leaving only data columns
    #Out-File -FilePath $kita_file -Encoding utf8

    # Betreuende
    $betreuende | Select-Object "betreuende_id","vorname","nachname","geburtsdatum","tel_nr","fk_kita_id" | 
    Export-Csv -Path $betreuende_file -NoTypeInformation -Delimiter ',' -Encoding UTF8
    #ConvertTo-Csv -NoTypeInformation -Delimiter ',' | # Convert to CSV string data without the type metadata
    #Select-Object -Skip 1 | # Trim header row, leaving only data columns
    #Out-File -FilePath $betreuende_file -Encoding utf8

    # Elternteil
    $elternteil | Select-Object "elternteil_id","vorname","nachname","geschlecht","fk_adresse_id" | 
    Export-Csv -Path $elternteil_file -NoTypeInformation -Delimiter ',' -Encoding UTF8
    #ConvertTo-Csv -NoTypeInformation -Delimiter ',' | # Convert to CSV string data without the type metadata
    #Select-Object -Skip 1 | # Trim header row, leaving only data columns
    #Out-File -FilePath $elternteil_file -Encoding utf8

    # Kind
    $kind | Select-Object "kind_id","vorname","geburtsdatum","geschlecht","fk_kita_id" | 
    Export-Csv -Path $kind_file -NoTypeInformation -Delimiter ',' -Encoding UTF8
    #ConvertTo-Csv -NoTypeInformation -Delimiter ',' | # Convert to CSV string data without the type metadata
    #Select-Object -Skip 1 | # Trim header row, leaving only data columns
    #Out-File -FilePath $kind_file 

    # Elternteil zu Kind
    $elternteil_zu_kind | Select-Object "elternteil_zu_kind_id","fK_elternteil_id","fk_kind_id" | 
    Export-Csv -Path $elternteil_zu_kind_file -NoTypeInformation -Delimiter ',' -Encoding UTF8
    #ConvertTo-Csv -NoTypeInformation -Delimiter ',' | # Convert to CSV string data without the type metadata
    #Select-Object -Skip 1 | # Trim header row, leaving only data columns
    #Out-File -FilePath $elternteil_zu_kind_file -Encoding utf8  
}
function Import-ToDatabase(){
    ########## Import in die Datenbank Funtkion##########

    $dataError = $false
    if ($server -eq ""){ 
        #Daten werden aus dem Form ausgelesen
        $server = $serverNameBox.Text
        $database = $databaseNameBox.Text
        $username = $usernameBox.Text
        $password = $passwordBox.Text
        if ($server -eq "" -or $database -eq "" -or $username -eq ""){
            #Fehlermeldung wenn angaben fehlen
            $MessageBody = "Mindestens ein Feld wurde nicht ausgefuellt"
            [System.Windows.Forms.MessageBox]::Show($MessageBody,"Fehler!",0,[System.Windows.Forms.MessageBoxIcon]::Error)
            $dataError = $true
        }
    }
    else { 
        #Daten wurden als Parameter übergeben
        $server = $server
        $database = $datenbank
        $username = $benutzername
        $password = $passwort
    }

    $port = 3306

    #https://dev.mysql.com/downloads/connector/net/
    if ($dataError -eq $false){
        if (![System.Reflection.Assembly]::LoadWithPartialName("MySql.Data")){
        
            $url = "https://dev.mysql.com/get/Downloads/Connector-Net/mysql-connector-net-8.0.19.msi"
            $ButtonType = [System.Windows.MessageBoxButton]::YesNo
            $MessageIcon = [System.Windows.MessageBoxImage]::Error
            $MessageBody = "Der MySQL Data Connector ist nicht installiert.`rDie Verbindung zur Datenbank kann nur mit dem installierten Connector hergestellt werden.`rSoll diese Seite im Browser geoeffnet werden?"
            $MessageTitle = "Mysql.Data konnte nicht geladen werden"
            
            $choice = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
    
            if ($choice -eq "Yes"){
                Start-Process $url
            }
        }
        else{
            try {
                Write-Log "Datenbank verbindung hergestellt...`r"
                # Direkter Link zur Datei:
                #$mySQLDataDLL = "$path/lib/MySQL.Data.dll"
                #[void][system.reflection.Assembly]::LoadFrom($mySQLDataDLL)
                [void][System.Reflection.Assembly]::LoadWithPartialName("MySql.Data") #Mysql.Data laden (Mysql Connector, welcher installiert sein muss)
                $ConnectionString = "Server=$server;Port=$port;Database=$database;Uid=$username;Pwd=$password;SslMode=Preferred"  #Connection String erstellen
                $connection = New-Object MySql.Data.MySqlClient.MySqlConnection
                $connection.ConnectionString = $ConnectionString
                $connection.Open() #Verbindung mit Datenbank herstellen
                $tables = @("kind","elternteil","betreuende","adresse","elternteil_zu_kind","kita","schuleinheit") #Alle Tabellen
                foreach ($table in $tables){
                    Write-log "$table wird importiert...`r"
    
                    #Tabelle leeren --> Delete Query
                    $query = "Delete from $table"
                    $command = New-Object MySql.Data.MySqlClient.MySqlCommand($query, $connection)
                    $dataAdapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($command)
                    $dataSet = New-Object System.Data.DataSet
                    $recordCount = $dataAdapter.Fill($dataSet, "data") #Anzahl 
                    $dataSet.Tables["data"] | Format-Table
    
                    #CSV importieren    
                    $includesfieldnames = $true #Spaltenüberschriften in der Datei?
                    $newtablename = $table #Tabellen Name
                    $csvfile = "$path\output\$table.csv" #Pfad zur Datei
                    $csvdata = Get-Content -Path $csvfile #CSV-Datei einlesen
                    $fieldnames = $csvdata[0].Split(',') #Spaltenüberschriften
                    $csvdata = $csvdata #| Select-Object -Skip 1 #Erste Zeile überspringen
                    
                    if ($includesfieldnames -eq $true) {$start=1} else {$start=0} #Spaltenüberschriften in der Datei? Wenn ja, erste Zeile weglassen
                    for ($i=$start;$i -le ($csvdata.Count-$start); $i++) {
                        $insertsql = 'INSERT INTO `'+$newtablename+'` (' # Sql-Befehl erstellen
                        $fieldcount = $fieldnames.Count # Anzahl Spalten
                        $commatrack = 1
                        foreach ($field in $fieldnames) { # Jede Spalte im Insert Into befehl einfügen
                            $insertsql = $insertsql+$field
                            $commatrack++ 
                            if ($commatrack -le $fieldcount) {
                                $insertsql = $insertsql+',' # Wenn es nicht die letzte Spalte ist, ein Komma hinzufügen
                            }
                        }
                        $insertsql = $insertsql -replace '"','`' # " mit ` ersetzen
                        $insertsql = $insertsql+') VALUES (' # Values
                        $commatrack = 1
                        foreach ($itemrow in $csvdata[$i]) {
                            $item = $itemrow.Split(',') #Zeile aufteilen ben Commas
                            foreach ($data in $item) {
                                $insertsql = $insertsql+''+$data+'' # Jede Spalte im Values Teil einfügen
                                $commatrack++
                                if ($commatrack -le $fieldcount) {
                                    $insertsql = $insertsql+',' # Wenn es nicht die letzte Spalte ist, ein Komma hinzufügen
                                }
                            }
                        }
                        $notnull = $false
                        if (($insertsql -Split "VALUES")[1].Trim() -ne "("){
                            $notnull = $true
                        }
                        $insertsql = "$insertsql)"
                        if ($notnull -eq $true){
                            $Query = $insertsql
                            $command = New-Object MySql.Data.MySqlClient.MySqlCommand($query, $connection)
                            $dataAdapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($command)
                            $dataSet = New-Object System.Data.DataSet
                            $recordCount = $dataAdapter.Fill($dataSet, "data")
                        }
                    }
                        
                    #Testen ob es importiert wurde --> Durch Select Query
                    $Query = "Select * FROM $table"
                    $command = New-Object MySql.Data.MySqlClient.MySqlCommand($query, $connection)
                    $dataAdapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($command)
                    $dataSet = New-Object System.Data.DataSet
                    $recordCount = $dataAdapter.Fill($dataSet, "data")
                    $dataSet.Tables["data"] | Format-Table
    
                    Write-Log "Geschriebene Datensaetze $recordCount" #Anzahl geschriebene Datensätze ausgeben
                }
                
            }
            catch {
                $errormessage = $Error[0] #Error Message speichern
                Write-Log "Query konnte nicht ausgeführt werden: $errormessage" #Errormessage ausgeben
                #Error Message als Popup anzeigen
                $MessageBody = "Es gab folgenden Fehler: $errormessage"
                [System.Windows.Forms.MessageBox]::Show($MessageBody,"Fehler!",0,[System.Windows.Forms.MessageBoxIcon]::Error)
                
            }
            Finally {
                #Connection schliessen, egal ob es einen Fehler gab oder nicht
                Write-log "Verbindung geschlossen`r"
                $connection.Close() #Verbindung mit Datenbank schliessen
            }
    
        }
    }
}

###### Funktion für das Schreiben des Logs im GUI/Terminal ######
function Write-Log([string]$text){
    if ($hidden.IsPresent){ #Wenn das Script ohne Gui gestartet wurde
        Write-Host "$text" #In Console schreiben
    }
    else {
        $logbox.AppendText("$text `r`n") #In Textbox im Form schreiben
    }
}
if ($hidden.IsPresent){ #Wenn das Script ohne Gui gestartet wurde
    Write-Files #CSV-Dateien erstellen
    if ($autoImport.IsPresent){ #Wenn der Parameter -autoImport verwendet wird
        Import-ToDatabase #Import in die Datenbank durchführen
    }
}
else {
    Add-Type -AssemblyName PresentationCore,PresentationFramework
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $form                            = New-Object system.Windows.Forms.Form
    $form.ClientSize                 = '535,290'
    $form.text                       = "Datenbank aktualisierungs Tool"
    $form.TopMost                    = $false
    $form.MaximizeBox                = $false
    $form.FormBorderStyle            = "Fixed3D"

    $serverNameBox                   = New-Object system.Windows.Forms.TextBox
    $serverNameBox.multiline         = $false
    $serverNameBox.width             = 192
    $serverNameBox.height            = 20
    $serverNameBox.location          = New-Object System.Drawing.Point(297,38)
    $serverNameBox.Font              = 'Microsoft Sans Serif,10'

    $serverNameLabel                 = New-Object System.Windows.Forms.Label
    $serverNameLabel.width           = 192
    $serverNameLabel.height          = 20
    $serverNameLabel.location        = New-Object System.Drawing.Point(297,23)
    $serverNameLabel.Text            = "Server"

    $databaseNameBox                 = New-Object system.Windows.Forms.TextBox
    $databaseNameBox.multiline       = $false
    $databaseNameBox.width           = 192
    $databaseNameBox.height          = 20
    $databaseNameBox.location        = New-Object System.Drawing.Point(297,81)
    $databaseNameBox.Font            = 'Microsoft Sans Serif,10'

    $databaseNameLabel               = New-Object System.Windows.Forms.Label
    $databaseNameLabel.width         = 192
    $databaseNameLabel.height        = 20
    $databaseNameLabel.location      = New-Object System.Drawing.Point(297,66)
    $databaseNameLabel.Text          = "Datenbank"

    $usernameBox                     = New-Object system.Windows.Forms.TextBox
    $usernameBox.multiline           = $false
    $usernameBox.width               = 192
    $usernameBox.height              = 20
    $usernameBox.location            = New-Object System.Drawing.Point(297,126)
    $usernameBox.Font                = 'Microsoft Sans Serif,10'

    $usernameLabel                   = New-Object System.Windows.Forms.Label
    $usernameLabel.width             = 192
    $usernameLabel.height            = 20
    $usernameLabel.location          = New-Object System.Drawing.Point(297,111)
    $usernameLabel.Text              = "Benutzername"

    $passwordBox                     = New-Object system.Windows.Forms.TextBox
    $passwordBox.multiline           = $false
    $passwordBox.width               = 192
    $passwordBox.height              = 20
    $passwordBox.location            = New-Object System.Drawing.Point(297,170)
    $passwordBox.Font                = 'Microsoft Sans Serif,10'
    $passwordBox.PasswordChar        = '*';

    $passwordLabel                   = New-Object System.Windows.Forms.Label
    $passwordLabel.width             = 192
    $passwordLabel.height            = 20
    $passwordLabel.location          = New-Object System.Drawing.Point(297,155)
    $passwordLabel.Text              = "Passwort"

    $startUploadButton               = New-Object system.Windows.Forms.Button
    $startUploadButton.text          = "In Datenbank laden"
    $startUploadButton.width         = 157
    $startUploadButton.height        = 42
    $startUploadButton.location      = New-Object System.Drawing.Point(297,206)
    $startUploadButton.Font          = 'Microsoft Sans Serif,12'

    $logbox                          = New-Object system.Windows.Forms.TextBox
    $logbox.multiline                = $true
    $logbox.width                    = 251
    $logbox.height                   = 152
    $logbox.ReadOnly                 = $true
    $logbox.location                 = New-Object System.Drawing.Point(24,38)
    $logbox.Font                     = 'Microsoft Sans Serif,10'
    $logbox.ScrollBars               = "Vertical"

    $startCreateButton               = New-Object system.Windows.Forms.Button
    $startCreateButton.text          = "Dateien erstellen"
    $startCreateButton.width         = 150
    $startCreateButton.height        = 43
    $startCreateButton.location      = New-Object System.Drawing.Point(24,207)
    $startCreateButton.Font          = 'Microsoft Sans Serif,12'

    $form.controls.AddRange(@($serverNameBox,$serverNameLabel,$databaseNameBox,$databaseNameLabel,$usernameBox,$usernameLabel,$passwordBox,$passwordLabel,$startUploadButton,$logbox,$startCreateButton))
    $startCreateButton.Add_Click({
        $logbox.Text = "" #Logbox leeren
        Write-Files
    })
    $startUploadButton.Add_Click({  
        $logbox.Text = "" #Logbox leeren
        Import-ToDatabase
    })
    $form.showDialog() #Form anzeigen, sobald alles geladen ist. (Script wird an diesem Punkt angehalten und macht dann hier weiter, sobald das Script geschlossen wurde)
}
