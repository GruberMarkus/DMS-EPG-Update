Recording Service EPG Update Script
http://www.dvbviewer.tv/forum/topic/41624-epg-update-script/

[Deutsch]
Seit der Version 1.23.0.0 bietet das Recording Service einen recht schnellen
Weg um das EPG aller Sender zu aktualisieren.

Mit einem internen Timer ("EPGStart") ist es möglich, dieses EPG-Update zu
starten - dabei muss allerdings eine fixe Laufzeit angegeben werden.

Das Recording Service EPG Update Script startet das das EPG-Update und wartet
solange, bis das Update abgeschlossen ist - unabhängig davon, wie lange es dauert.

Nach dem Abschluss des Updates wird ein AutoTimer-Task ausgeführt. Der Rechner,
auf dem das Recording Service läuft, kann wahlweise in den Standby versetzt
werden (wenn gewisse Voraussetzungen erfüllt sind).

Details zu den Konfigurationsmöglichkeiten finden sich in der Datei "sample.ini".

Das Script wird von der Eingabeaufforderung aus gestartet und erwartet die Angabe
einer ini-Datei: "cscript.exe RS-EPG-Update.vbs /ini:sample.ini".

Standardmässig werden alle Meldungen des Script in der Datei
"RS-EPG-Update.log" protokolliert.

Wer mehr Konfigurationsmöglichkeiten benötigt, sollte auf das DVBViewer EPG
Update Script von http://www.dvbviewer.tv/forum/topic/41624-epg-update-script/
zurückgreifen.



[English]
Since version 1.23.0.0, the Recording Service offers a quite fast way to update
the EPG of all channels.

With an internal timer ("EPGStart"), this EPG update can be started - but you
have to define a fixed runtime.

The Recording Service EPG Update Script starts this EPG update and waits until
the update is finished - no matter how long it takes.

The AutoTimer task is executed after the update. The computer hosting the
Recording Service can be put into standby mode (wenn certain prerequisites are met).

Details regarding the configuration options can be found in the file "sample.ini".

The script is started via the command prompt and expects the path to an ini file
as parameter: "cscript.exe RS-EPG-Update.vbs /ini:sample.ini".

Per default, all script messages are logged in the file "RS-EPG-Update.log".

If you need more configuration options, you should use the DVBViewer EPG Update
script from http://www.dvbviewer.tv/forum/topic/41624-epg-update-script/.