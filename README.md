# Tagesrapport App

Web-App auf Basis der Excel-Vorlage `Tagesrapport App.xlsx`.

## Funktionen

- Schrittweise Erfassung für Fahrhabe, Gebäude und Trocknung
- Verzweigung: Es werden nur die ausgewählten Arbeitsbereiche angezeigt
- Eigene Mitarbeiter-Stammliste für Polygon-Mitarbeitende
- Temporärmitarbeitende als Freitext mit ergänzbarem Temporärbüro
- Arbeitsstunden, Zuschläge, Fahrzeuge, Geräte, Produkte, Verbrauchsmaterial und Entsorgung
- Bemerkungsseite und Abschlussübersicht
- Projektleiter-Ordner mit Korrektur- und Freigabeprozess zum Frontoffice
- Tagesübersicht der geleisteten Stunden je Mitarbeiter
- Rapporte speichern, erneut öffnen, korrigieren und löschen
- Export als ausgefüllte Excel-Datei im Aufbau der Originalvorlage
- Responsive Darstellung für PC, Tablet und Smartphone
- Optionaler Zugriffsschutz über `APP_PASSWORD`
- Oberfläche auf Deutsch, Französisch und Italienisch
- Persönlicher Login für jeden Projektleiter
- Automatische E-Mail an `vorname.nachname@polygongroup.com`
- PDF-Archiv: `Projektleiter / aktuelles Datum / Rapport.pdf`

## Lokal starten

```bash
python -m pip install -r requirements.txt
python app.py
```

Danach `http://localhost:8080` öffnen.

## Auf Railway veröffentlichen

1. Den Ordner als GitHub-Repository hochladen.
2. In Railway **New Project → Deploy from GitHub repo** auswählen.
3. Das Repository auswählen. Railway erkennt `railway.json` und startet die App automatisch.
4. Unter **Settings → Networking → Generate Domain** eine Webadresse erzeugen.
5. Für dauerhafte Rapporte unter **Volumes** ein Volume mit Mount Path `/data` hinzufügen.
6. Unter **Variables** `DATA_DIR=/data` setzen.
7. Optional `APP_PASSWORD` als Zugriffsschlüssel setzen.

### Erforderliche Railway-Variablen

```text
DATA_DIR=/data
APP_PASSWORD=<starker gemeinsamer App-Schlüssel>
BASE_URL=https://<deine-railway-domain>
SMTP_HOST=smtp.office365.com
SMTP_PORT=587
SMTP_USERNAME=<M365-Absenderkonto>
SMTP_PASSWORD=<Passwort oder App-Kennwort>
SMTP_FROM=<M365-Absenderadresse>
```

Das verwendete Microsoft-365-Konto muss SMTP AUTH verwenden dürfen. Falls SMTP AUTH in der Organisation deaktiviert ist, muss der Versand später auf Microsoft Graph oder einen internen SMTP-Relay umgestellt werden.

## Projektleiter-Logins einrichten

1. App mit dem gemeinsamen `APP_PASSWORD` öffnen.
2. **Projektleiter-Zugänge** wählen.
3. Für jeden Projektleiter ein persönliches Passwort mit mindestens zehn Zeichen setzen.
4. Der Link in der Benachrichtigungs-E-Mail öffnet direkt den Projektleiter-Login und danach den zugehörigen Rapport.

## Workflow

1. Ein Rapport wird zunächst als Entwurf gespeichert.
2. **An Projektleiter übermitteln** erzeugt ein vollständiges PDF im Tages-Unterordner des Projektleiters und sendet die Benachrichtigungs-E-Mail.
3. Der Projektleiter öffnet und korrigiert den Rapport; insbesondere kann er bei Temporärmitarbeitenden das Temporärbüro ergänzen.
4. **An Frontoffice** verschiebt den kontrollierten Rapport in den gemeinsamen Frontoffice-Ordner.
5. Das Frontoffice kann den Rapport als Excel herunterladen und als erledigt markieren.

Ohne Railway-Volume funktioniert die App ebenfalls, gespeicherte Rapporte können bei einem neuen Deployment jedoch verloren gehen.
