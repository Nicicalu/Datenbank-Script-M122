# Datenbank-Script V1.0

In diesem Readme wird beschrieben, wie das Script funktioniert, was die Voraussetzungen sind und wie man es automatisiert.
[Github Repository](https://github.com/Nicicalu/Datenbank-Script-M122)


## Voraussetzungen

 - Die Dateien `data1.csv` und `data2.csv` müssen sich im Ordner `src` befinden, welcher eine Stufe unter dem Script im selben Verzeichnis ist z.B. `C:\pfad\zu\script\src\`
 - Die Dateien werden in einem Ordner mit dem Namen "output" ausgegeben. Dieser wird automatisch erstellt, wenn er nicht schon vorhanden ist.
 - **Direkter Import in die Datenbank (Nur Mysql/MariaDB):** Dafür muss auf dem PC der MYSQL-Connector installiert sein. Dieser wird benötigt, dass das Script eine Verbindung mit der Datenbank herstellen kann. Wenn er nicht installiert ist, werden Sie beim Start des Imports darauf hingewiesen.
 Link: https://dev.mysql.com/downloads/connector/net/
## Automatisieren
Das Script automatisch laufen zu lassen, ist ganz einfach. Muss niemand etwas von Hand machen, sondern das Script erstellt die Dateien und importiert Sie automatisch in die Datenbank. 
Der Befehl um das Script ohne GUI zu starten ist folgender:

    C:\pfad\zu\script\script.ps1 -hidden
Wenn das Script die Dateien automatisch in die Datenbank importieren soll, muss noch der Parameter `-autoImport` und die Daten für die MYSQL Datenbank mitgegeben werden.

    C:\pfad\zu\script\script.ps1 -hidden -autoImport -server "127.0.0.1" -datenbank "datenbankname" -benutzername "username" -passwort "password"
Um das Script mit diesen Parametern an gewissen Tagen oder Events zu wiederholen können Sie folgende Schritte durchführen:

### Aufgabenplanung öffnen
Suchen Sie in der Windows-Suche nach `Aufgabenplanung` und öffnen Sie die Aufgabenplanung. 
Drücken Sie mit der rechten Maustaste auf die Aufgabenplanungsbibliothek und wählen Sie `Aufgabe erstellen`.
### Trigger konfigurieren
Geben Sie der Aufgabe im Reiter `Allgemein` einen Namen und wechseln in den Reiter `Trigger`.
Hier werden Trigger festgelegt. Hier kann man im Dropdown auswählen, dass diese Aufgabe an gewünschten Tagen, beim Start es Computers oder bei einem anderen Event ausgeführt wird. Wählen Sie hier `Nach einem Zeitplan`, wenn das Script jeden Tag oder an spezifischen Tagen ausgeführt werden soll.  
Unten können Sie jetzt auswählen, wann es ausgeführt werden soll. Wählen Sie hier täglich, wöchentlich oder monatlich, wenn das Script mehrere Male ausgeführt werden soll. Auf der rechten Seite können Sie die gewünschte Zeit auswählen, und die Wiederholung konfigurieren.
### Aktionen konfigurieren
Wechseln Sie dann in den Tab `Aktionen`. Hier drücken Sie auf `Neu` um eine neue Aktion zu erstellen. Im Dropdown-Menü wählen Sie `Programm starten`. Im Feld `Programm/Skript` geben Sie `PowerShell` ein.  Bei Argumente geben Sie jetzt die Parameter ein folgendes ein: 

    -command " & C:\Pfad\zu\Script\script.ps1"
Ersetzen Sie den Pfad mit dem richtigen Pfad in ihrer Umgebung.
Fügen Sie am Ende dann noch die Parameter hinzu welche Sie benötigen. Der Parameter `-hidden` **muss** verwendet werden, da das GUI nicht angezeigt werden würde.
Am Ende könnte das dann z.B. so aussehen (Beispiel mit automatischem Import in die DB):

    -command " & C:\Pfad\zu\Script\script.ps1" -hidden -autoImport -server "172.0.0.1" -datenbank "schule" -benutzername "username" -passwort "passwort"

### Aufgabe abschliessen
In den Reitern `Bedingungen` und `Einstellungen` können Sie noch weitere Einstellungen dieser Aufgabe vornehmen. Dies ist aber normalerweise nicht notwendig.
## Tests
Das Script wurde von mir folgendermassen getestet:
- [x] Die erstellten CSV-Dateien werden erstellt
- [x] Der Import in eine Datenbank mit MYSQL Workbench und PHPMYADMIN (Herr Schneller war während des Tests anwesend. Dieser wurde für gut befunden.)
- [x] Der automatische Import in eine MYSQL oder MARIADB Datenbank
- [x] Das Script kann im Hintergrund ausgeführt werden (Parameter `-hidden`)
- [x] Errror-Handling funktioniert

## Weiterführende Links
 - [https://www.deskmodder.de/wiki/index.php/Aufgabenplanung_Aufgabe_erstellen_unter_Windows_10](https://www.deskmodder.de/wiki/index.php/Aufgabenplanung_Aufgabe_erstellen_unter_Windows_10)
 - Github Repository: [https://github.com/Nicicalu/Datenbank-Script-M122](https://github.com/Nicicalu/Datenbank-Script-M122)

***
*von Nicolas Caluori © 2020*
