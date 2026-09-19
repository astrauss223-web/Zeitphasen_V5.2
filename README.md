# Zeitphasenmodell V5.2.1

**Bestehendes Portfolio importieren, Zeitphasen planen & Operative Manager steuern**

---

## 🎯 Highlights & Neuerungen in Version 5.2.1

- **🏷️ Anlageschwerpunkt in der Fondsauswahlliste (VEM-Liste 11.09.2026):**
  - In der Fondsauswahlliste (beim Öffnen eines Topfes) wird hinter jedem Fonds **in derselben Zeile** neben dem Managementansatz (z. B. `EB Aktien/Anleihen`) und der Ertragsverwendung (`ausschüttend` / `thesaurierend`) der offizielle **Anlageschwerpunkt** aus der aktuellen VEM-Liste vom 11.09.2026 angezeigt (z. B. `Anleihen Welt`, `Aktien USA`, `Vermögensverwalter - ausgewogen`, `Anleihen Euro Kurz`, `Aktien Technologie` etc.).
  - Die Kennzeichnung erfolgt über einen eleganten, gut lesbaren Badge mit dezentem Rahmen, der sich harmonisch in das MLP-Design einfügt.
  - Auch in den Sofort-Suchergebnissen der Fonds-Suche wird der Anlageschwerpunkt direkt mit aufgeführt.

- **👁️ Sichtbarkeits-Umschalter für „Auswahl operativer Manager“:**
  - Neues Icon (Auge) oben links im Panel „Auswahl operativer Manager“ (auf gleicher Höhe wie die Aktions-Icons für Import, Export, Reset und PDF).
  - **Funktion:** Blendet die detaillierten Fondskarten in der rechten Seitenleiste per Klick komplett aus und zeigt einen unaufdringlichen Statushinweis an.
  - **Werteerhalt:** Alle Schichten und Zylinder unter den Zeitphasen behalten ihre Fonds-Badges (z. B. `1 Fonds`), Einmalbeträge und Sparraten. Die Anlageplanung (Gesamtinvestition, Verplant, Verbleibend) sowie die Fondsmarkierungen (`✓ Ausgewählt`) in den Modals bleiben zu 100 % aktiv.
  - Ein erneuter Klick auf das Auge blendet die Auswahl sofort wieder ein.

- **⚡ Nahtloser & pop-up-freier JSON- / Bibliotheks-Import:**
  - Kein störendes oder blockierendes Browser-Alert-Popup mehr beim Laden von `.json`-Setups.
  - Sofortige, flüssige Übernahme aller Daten (Einmalanlagen, monatliche Sparraten, Töpfe und Fonds).
  - Volle Kompatibilität mit flachen und geschachtelten `globals`-JSON-Strukturen.
  - Zuverlässige Lade- und Speicherroutine (`loadPortfolioSetup`) sowohl über Datei-Upload als auch über die integrierte Depot-Bibliothek.

- **📄 Anlagevorschlag als PDF exportieren:**
  - Exportfunktion über den Button „Anlagevorschlag PDF“ in der rechten Seitenleiste.
  - Generiert ein formatiertes PDF-Dokument mit Datum, Gesamtanlagevolumen und tabellarischer Übersicht aller Fonds, WKNs, Ansätze, Einmalbeiträge und Sparraten.

- **⏱️ Zeitphasenmodell-Skala & Schichten (V5.2 Layout):**
  - **Zeitphase 1:** `< 1 Jahr` · `Geldmarkt`
  - **Zeitphase 2:** `> 2 Jahre` · `Zielrendite`, `WB Anleihen`, `EB Anleihen` *(3 Zeilen)*
  - **Zeitphase 3:** `> 4 Jahre` · `Zielrendite`, `WB Anleihen`, `EB Anleihen` *(3 Zeilen)*
  - **Zeitphase 4:** `> 6 Jahre` · `WB Aktien/Anleihen`
  - **Zeitphase 5:** `> 8 Jahre` · `WB Aktien`
  - **Zeitphase 6:** `> 10 Jahre` · `EB Aktien`
  - **Losgelöste Schichten:** `Tagesgeld` und `Echte / unechte Anlageklassen mit unternehmerischen Risiken` (für Managementansätze außerhalb der Phasen 1–6).

- **ℹ️ Kursives `i` Info-Icon bei Managementansatz-Badges:**
  - Bei Fonds mit näherer Unterbezeichnung in der VEM-Liste (z. B. `DWS04F` mit *„Hochzins mittlere Restlaufzeiten“* oder `848105` mit *„Staats- und Unternehmensanleihen mittlere Restlaufzeiten“*) befindet sich am Badge ein blaues kursives **`i`** Info-Icon mit Hover-Tooltip.

- **🔍 Suchfunktion mit automatischer Ansteuerung & Flash-Highlight:**
  - Schnelles Auffinden von Fonds nach Name, WKN oder ISIN.
  - Bei Klick auf das Suchergebnis öffnet sich automatisch die entsprechende Schicht und der Fonds wird für 2,5 Sekunden golden hervorgehoben.

- **📥 Depot-Import Engine (CSV / Excel / PDF / JSON):**
  - Verarbeitet Depotauszüge im CSV-, Excel- (`.xlsx`, `.xls`), PDF- und JSON-Format.
  - Erkennt mehrzeilige Tabellenlayouts und deutsche Währungsformate automatisch.

---

## 🧭 Bedienung der Hauptfunktionen

### 1. Operative Manager ein- und ausblenden
1. Klicken Sie in der rechten Spalte bei **„Auswahl operativer Manager“** oben links auf das **Auge-Icon** (`👁`).
2. Die detaillierte Fondsliste wird ausgeblendet; das Icon wechselt zu einem durchgestrichenen Auge (`👁‍🗨`) mit aktivem Akzent.
3. Die Beträge und Zylinder unter den Zeitphasen sowie die Werte in der Anlageplanung bleiben unverändert sichtbar und aktiv.
4. Ein erneuter Klick auf das Icon blendet die Liste wieder ein.

### 2. Depot oder Setup laden (.json / .csv / .xlsx / .pdf)
* **Über die Toolbar:** Oben auf den grünen Button **„Depot laden (CSV / Excel)“** klicken.
* **Über die rechte Leiste:** Auf das Upload-Icon (`[ ⬆ ]`) klicken.
* **Über die Bibliothek:** Auf **„Bibliothek“** klicken, um gespeicherte Setups per Klick auf **„▶ Laden“** abzurufen oder eigene `.json`-Dateien zu importieren.
* Das Setup wird unmittelbar ohne blockierende Dialogfenster angewendet.

### 3. Setup in der Bibliothek speichern oder als JSON exportieren
* **In Bibliothek speichern:** Auf das Disketten-Icon (`[ 💾 ]`) in der rechten Leiste klicken und einen Namen vergeben.
* **Als JSON herunterladen:** Auf das Download-Icon (`[ ⬇ ]`) klicken, um die Konfiguration als `.json`-Datei lokal zu sichern.

### 4. Anlagevorschlag als PDF herunterladen
* Klicken Sie unten im rechten Panel auf den Button **„Anlagevorschlag PDF“**.
* Das PDF wird automatisch mit allen im Portfolio befindlichen Positionen und Beträgen generiert und heruntergeladen.

---

## 💻 Systemvoraussetzungen & Browser

Reine Client-seitige Webanwendung – läuft direkt im Browser ohne Server-Installation oder Cloud-Zwang.

| System | Empfohlene Browser |
|--------|---------------------|
| 🍎 macOS | Safari, Chrome, Firefox, Edge |
| 🪟 Windows | Chrome, Edge, Firefox |

---

## 🗂️ Dateistruktur Version 5.2

| Datei | Beschreibung |
|-------|-------------|
| `index.html` | Hauptseite des Zeitphasenmodells V5.2 mit Steuerungspanel und Zylindern |
| `styles_v5.2.css` | Styling für 6 Zeitphasen, Zylinder, Badges, Icons und Animationen |
| `app_v5.2.js` | Gesamte Anwendungslogik (Berechnungen, Sichtbarkeits-Toggle, PDF-Export, Bibliothek, Import-Engine) |
| `fund_data.js` | Fondsdatenbank (120+ Fonds mit WKN, ISIN, Managementansätzen, Unterkategorien und Ländergewichtungen) |
| `portfolio_import.json` | Beispiel-Setup zur schnellen Demonstration |
| `README.md` | Diese Dokumentation |

---

*Zeitphasenmodell V5.2 · Stand: September 2026*
