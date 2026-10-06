# SmartMove für Outlook

Im Kontextmenü einer E-Mail erscheint das Untermenü **SmartMove**. Es kann eine
lokale Ablagestatistik manuell erstellen, nach JSON exportieren und daraus
dynamische Zielordner vorschlagen. Ein Klick auf einen Vorschlag verschiebt die
ausgewählte E-Mail direkt. Außerdem lässt sich der Absendername mit einem Klick
in der LinkedIn-Personensuche öffnen.

## Voraussetzung

- Windows mit **klassischem Outlook** (nicht „Neues Outlook“)
- .NET Framework 4.8
- .NET SDK 8 oder Visual Studio 2022 mit .NET-Desktopentwicklung

Ein echter Eintrag im Rechtsklickmenü wird von Outlook-Web-Add-ins und dem neuen
Outlook nicht angeboten. Deshalb ist diese Version ein lokales COM-Add-in für das
klassische Outlook unter Windows.

## Bauen und installieren

1. Klassisches Outlook vollständig schließen.
2. Eine normale PowerShell öffnen; Administratorrechte sind nicht nötig.
3. In diesen Projektordner wechseln und ausführen:

   ```powershell
   .\install.ps1
   ```

4. Klassisches Outlook starten.
5. In der Nachrichtenliste eine E-Mail rechtsklicken. Dort steht **SmartMove**
   mit dynamischen **Move To …**-Vorschlägen sowie den Befehlen
   **Absender auf LinkedIn suchen**, **Statistik initial erstellen** und
   **Statistik als JSON exportieren**.

Der erste Build lädt einmalig die .NET-4.8-Referenzassemblies über NuGet. Die
Installation wird ausschließlich für den aktuellen Windows-Benutzer registriert.
Bei einer Aktualisierung gilt derselbe Ablauf: Outlook vollständig schließen und
`install.ps1` erneut ausführen.

Falls PowerShell das Skript ausdrücklich als nicht vertrauenswürdig blockiert,
heben Sie nur die Markierung der drei lokalen Skripte auf und versuchen es erneut:

```powershell
Unblock-File .\build.ps1, .\install.ps1, .\uninstall.ps1
.\install.ps1
```

## Deinstallieren

Outlook schließen und dann ausführen:

```powershell
.\uninstall.ps1
```

## Fehlerdiagnose

Falls Outlook den Eintrag nicht zeigt, unter
`Datei > Optionen > Add-Ins > Verwalten: COM-Add-Ins` prüfen, ob **SmartMove**
aktiv ist. Der Eintrag erscheint beim Rechtsklick auf eine E-Mail-Zeile in der
Nachrichtenliste, nicht beim Rechtsklick auf den Text im Lesebereich. Das Add-in
funktioniert nur im klassischen Outlook.

Das Diagnoseprotokoll liegt unter
`%LOCALAPPDATA%\SmartMove\SmartMove.log`. Es enthält keine E-Mail-Daten.

Falls Outlook das Add-in nach einem Fehler deaktiviert hat, unter
`Datei > Langsame und deaktivierte COM-Add-Ins` beziehungsweise unter den
COM-Add-In-Einstellungen prüfen, ob SmartMove blockiert ist. Das
Installationsskript entfernt ausschließlich alte SmartMove-Einträge aus
Outlooks `CrashingAddinList` und `DisabledItems`.

## Eigene E-Mail-Adressen konfigurieren

SmartMove erkennt gesendete beziehungsweise selbst weitergeleitete Nachrichten
an einer frei konfigurierbaren Liste eigener E-Mail-Adressen. Die Datei wird
beim ersten Start von Outlook automatisch angelegt und liegt unter:

```text
%LOCALAPPDATA%\SmartMove\settings.json
```

`%LOCALAPPDATA%` steht dabei für das lokale Anwendungsdatenverzeichnis des
aktuellen Windows-Benutzers. Zum Bearbeiten kann die Datei beispielsweise mit
folgendem PowerShell-Befehl in Editor geöffnet werden:

```powershell
notepad "$env:LOCALAPPDATA\SmartMove\settings.json"
```

Unter `ownAddresses` werden alle eigenen Adressen und Aliase eingetragen. Es
können beliebig viele Adressen konfiguriert werden:

```json
{
  "ownAddresses": [
    "vorname.nachname@example.org",
    "alias@example.org"
  ]
}
```

Groß- und Kleinschreibung spielt keine Rolle; Leerraum am Anfang und Ende wird
entfernt und doppelte Einträge werden ignoriert. Die Konfiguration gilt nur für
den jeweiligen Windows-Benutzer und wird nicht im Projektverzeichnis abgelegt.

Nach jeder Änderung an `settings.json` muss **Statistik initial erstellen**
erneut ausgeführt werden. Dadurch wird die bestehende Statistik vollständig mit
der neuen Adressliste aufgebaut. Eine Nachricht, deren Absender in
`ownAddresses` steht, wird anhand ihrer „An“-Empfänger ausgewertet. Bei allen
anderen Nachrichten verwendet SmartMove weiterhin die Absenderadresse. Ist die
Liste leer, findet keine besondere Behandlung eigener Nachrichten statt.

## Statistik

Der Scan wird ausschließlich manuell über
`SmartMove > Statistik initial erstellen` gestartet. Er durchsucht alle
auswertbaren E-Mail-Ablageordner in den eingebundenen Outlook-Speichern.
Posteingang, Gesendet, Postausgang, Entwürfe, Papierkorb, Junk, Suchordner und
technische Synchronisationsordner werden ausgelassen. Unterordner des
Posteingangs werden als mögliche Ablageordner berücksichtigt.

SmartMove liest pro Nachricht Nachrichtenklasse, Absenderadresse und Betreff.
Steht der Absender in der konfigurierten Liste `ownAddresses`, liest es
zusätzlich die „An“-Empfänger und verwendet deren E-Mail-Adressen statt der
Absenderadresse. Bei mehreren „An“-Empfängern zählt jede unterschiedliche
Adresse einmal pro Nachricht. Nachrichtentext und Anhänge werden nicht gelesen.
Der vollständige Betreff wird nicht gespeichert. Gespeichert werden:

- je Absender- beziehungsweise Empfänger-/Ordnerpaar die Anzahl der E-Mails,
- je signifikantem Betreffwort/Ordnerpaar die Anzahl der E-Mails,
- Outlook-Ordnerkennung und lesbarer Ordnerpfad.

Betreffwörter werden kleingeschrieben, müssen aus Buchstaben bestehen und
länger als drei Zeichen sein. Häufige deutsche Wörter werden anhand der
etablierten Snowball-Stopwortliste entfernt. Zusätzlich filtert SmartMove
typische E-Mail-Floskeln wie `hallo`, `danke`, `bitte`, `grüße` und
`weitergeleitet`. Ein Wort wird pro E-Mail höchstens einmal gezählt.

Die SQLite-Datenbank liegt unter:

```text
%LOCALAPPDATA%\SmartMove\smartmove.db
```

Ein erneuter initialer Scan ersetzt die vorhandene Statistik. Der Scan zeigt
seinen Fortschritt an, kann abgebrochen werden und schreibt in Blöcken. Zwischen
kurzen Arbeitsphasen pausiert SmartMove bewusst, damit Outlook bedienbar bleibt
und das Add-in nicht unnötig als langsam eingestuft wird. Pro Arbeitsphase werden
höchstens 100 Tabellenzeilen beziehungsweise ungefähr 25 Millisekunden Arbeit
ausgeführt; anschließend pausiert der Scanner mindestens 150 Millisekunden.

Der JSON-Export wird über das SmartMove-Untermenü gestartet. Der Speicherort ist
frei wählbar. Adressen und Betreffwörter werden nach ihrer gesamten
Nachrichtenanzahl und deren Ordner jeweils absteigend nach Häufigkeit ausgegeben.
Die Adressen stehen aus Kompatibilitätsgründen weiterhin unter dem JSON-Schlüssel
`senders`.

Beispiel für das Exportformat:

```json
{
  "version": 2,
  "lastScanUtc": "2026-07-13T10:30:00Z",
  "senders": {
    "kunde@example.de": {
      "\\Postfach\\Projekte\\Kunde": 12
    }
  },
  "subjectWords": {
    "rechnung": {
      "\\Postfach\\Buchhaltung\\Rechnungen": 37
    }
  }
}
```

## Ordnervorschläge und Verschieben

Oben im SmartMove-Untermenü erscheinen, soweit genügend historische Daten
vorhanden sind:

- die zwei häufigsten Ziele für den Absender beziehungsweise, bei einer eigenen
  konfigurierten Absenderadresse, für die „An“-Empfänger,
- die zwei am besten bewerteten Ziele für die Signalwörter im Betreff.

Bei der Betreffbewertung zählt neben der Häufigkeit auch, wie eindeutig ein Wort
auf einen Ordner verweist. Zeigen mehrere verschiedene Signalwörter auf dasselbe
Ziel, erhält dieses Ziel pro zusätzlichem gemeinsamen Wort einen Bonus von
50 Prozent auf die kombinierte Evidenz. Doppelte Ziele zwischen Absender- und
Betreffvorschlägen werden nach Möglichkeit vermieden.

Die Signalwortwertung verwendet sinngemäß:

```text
Evidenz je Wort = P(Ordner | Wort) × ln(1 + Trefferzahl)
Gesamtwertung   = Summe der Evidenzen × (1 + 0,5 × (Signalwörter - 1))
```

Ein Klick auf **Move To …** verschiebt die ausgewählte E-Mail unmittelbar und
ohne zusätzliche Bestätigung. Bei einer Mehrfachauswahl werden alle ausgewählten
E-Mails verschoben; die Vorschläge basieren dabei auf der ersten Nachricht. Zur
kompakten Anzeige enthält der Menütext nur die letzten beiden Ebenen des
Ordnerpfads. Intern bleibt der Zielordner über StoreID und EntryID eindeutig.

## LinkedIn-Suche

**SmartMove > Absender auf LinkedIn suchen** öffnet die LinkedIn-Personensuche
für den Anzeigenamen des Absenders im Standardbrowser. SmartMove speichert keine
LinkedIn-Zugangsdaten und greift nicht automatisiert auf LinkedIn-Profile zu.
Eine vorhandene LinkedIn-Anmeldung des Browsers wird verwendet.

## Aktuelle Grenzen und mögliche Ausbaustufen

Aktuell lernt SmartMove nur durch einen erneuten manuellen Komplettscan. Mögliche
nächste Schritte sind:

1. bestätigte SmartMove-Verschiebungen sofort in der Statistik nachführen,
2. manuell abweichende Zielentscheidungen als Korrektur lernen,
3. Absender- und Signalwortwertung zu einer gemeinsamen Gesamtwertung verbinden,
4. Scan-Tempo und Stopwortliste konfigurierbar machen,
5. optional eine Rückfrage oder Rückgängig-Funktion vor dem Verschieben anbieten.

## Datenschutz

Alle Statistikdaten bleiben lokal unter `%LOCALAPPDATA%\SmartMove` und werden
nicht an externe Dienste übertragen. Gespeichert werden Absenderadressen oder
bei Nachrichten von konfigurierten eigenen Adressen „An“-Empfängeradressen,
normalisierte Signalwörter, Ordnerkennungen, Ordnerpfade und aggregierte
Trefferzahlen. Auch `settings.json` bleibt lokal. Nachrichtentexte, Anhänge und
vollständige Betreffzeilen werden nicht gespeichert. Nur wenn **Absender auf
LinkedIn suchen** angeklickt wird, wird der Anzeigename des Absenders als
Suchbegriff an LinkedIn übergeben.

## Lizenz

Der SmartMove-eigene Quellcode steht unter der
[MIT-Lizenz](LICENSE). Hinweise und Lizenzbedingungen für verwendete
Drittbestandteile, insbesondere die deutsche Snowball-Stopwortliste, stehen in
[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md). Bei binären Veröffentlichungen
muss diese Datei zusammen mit der Software ausgeliefert werden.
