# ALPS/PASS Tools for Microsoft Visio

Ein VSTO-Add-in für Microsoft Visio zum Modellieren, Importieren, Prüfen,
Anordnen und Konvertieren von PASS- und ALPS-Prozessmodellen.

Das Add-in ergänzt Visio um den Ribbon-Tab **ALPS/PASS ADDIN**. Es verbindet
die vorhandenen ALPS/PASS-Visio-Schablonen mit einem OWL/RDF-Importer,
automatischem Graph-Layout, semantischem SID-/SBD-Snapping, Modellnavigation,
einer lokalen Prüfung von Elementnamen, ALPS-SID-Verifikation und einem
BPMN-2.0-Export.

> **Plattformhinweis:** Das Add-in basiert auf .NET Framework 4.8 und VSTO.
> Ausführung, vollständiger Build und manuelle Abnahme benötigen Windows,
> Microsoft Visio und die Office/VSTO-Entwicklungswerkzeuge. Das Projekt ist
> nicht für Visio im Web oder macOS bestimmt.

## Inhalt

- [Funktionsumfang](#funktionsumfang)
- [Voraussetzungen](#voraussetzungen)
- [Installation](#installation)
- [Erste Schritte](#erste-schritte)
- [Funktionen im Detail](#funktionen-im-detail)
- [Datenschutz und externe Dienste](#datenschutz-und-externe-dienste)
- [Entwicklung](#entwicklung)
- [Tests und Qualitätssicherung](#tests-und-qualitätssicherung)
- [Projektstruktur](#projektstruktur)
- [Bekannte Grenzen](#bekannte-grenzen)
- [Fehlerbehebung](#fehlerbehebung)
- [Weitere Dokumentation](#weitere-dokumentation)
- [Lizenz und Drittkomponenten](#lizenz-und-drittkomponenten)

## Funktionsumfang

| Ribbon-Gruppe | Befehl | Ergebnis |
| --- | --- | --- |
| Standard Functions | **Open ALPS/PASS Stencils** | Öffnet die benötigten SID-/SBD-Schablonen für die interaktive Modellierung. |
| ALPS Layer Editing | **Show layer Explorer** | Zeigt Modelle, Layer und verknüpfte SID-/SBD-Seiten in einem navigierbaren Explorer. |
| OWL PASS Tools | **Import OWL** | Importiert PASS-/ALPS-Modelle aus `.owl` oder `.rdf` nach Visio. |
| OWL PASS Tools | **Auto-Arrange → Top-down** | Ordnet das aktive SID- oder SBD-Diagramm von oben nach unten an. |
| OWL PASS Tools | **Auto-Arrange → Left-right** | Ordnet das aktive SID- oder SBD-Diagramm von links nach rechts an. |
| OWL PASS Tools | **Verify ALPS Models** | Vergleicht eine abstrakte ALPS-Spezifikation mit einem implementierenden Modell. |
| Model Conversion | **Convert PASS to BPMN** | Konvertiert ein PASS-/ALPS-OWL/RDF-Modell in eine BPMN-2.0-Datei mit Diagrammkoordinaten. |
| NLP PASS Checking | **Check Model Naming** | Bewertet unterstützte Elementnamen mit einem lokalen, eingebetteten Klassifikator. |
| NLP PASS Checking | **Retrain** | Trainiert den lokalen Klassifikator erneut aus 680 mitgelieferten Beispielen. |
| NLP PASS Checking | **Provider Settings** | Konfiguriert optionale Vorschlagsanbieter und deren Modelle. |
| Automatisch | **SID-/SBD-SnapHandler** | Koppelt Extension-Shapes geometrisch und semantisch an Elemente einer erweiterten Hintergrundseite. |

Die Kernfunktionen zum Importieren, Prüfen und lokalen Klassifizieren arbeiten
offline. Nur optionale NLP-Namensvorschläge greifen auf einen vom Benutzer
aktivierten externen Anbieter zu.

## Voraussetzungen

### Für die Nutzung

- Windows 10 oder Windows 11
- Microsoft Visio Desktop
- .NET Framework 4.8
- Microsoft Visual Studio Tools for Office Runtime (VSTO Runtime)
- die zum Modellieren verwendeten ALPS/PASS-SID- und SBD-Schablonen im
  Visio-Ordner **My Shapes**

Die Bitness von Visio und der Office-Installation muss zusammenpassen. Für das
Add-in ist in der Solution standardmäßig `Any CPU` konfiguriert.

### Für die Entwicklung

- Visual Studio 2022
- Workload **Office/SharePoint development** bzw. die VSTO-Komponenten
- .NET Framework 4.8 Developer Pack
- NuGet/MSBuild aus Visual Studio
- Microsoft Visio zum Debuggen der COM-/VSTO-Integration
- optional .NET SDK 8.0.414 oder neuer für Coverage mit Coverlet

## Installation

### Veröffentlichtes Installationspaket

1. Visio vollständig schließen.
2. Das vollständige veröffentlichte Paket lokal entpacken. `setup.exe` und der
   Ordner `Application Files` müssen zusammenbleiben.
3. `setup.exe` starten und die VSTO-/Add-in-Installation abschließen.
4. Visio öffnen und prüfen, ob der Tab **ALPS/PASS ADDIN** sichtbar ist.
5. Über **Open ALPS/PASS Stencils** prüfen, ob die Schablonen gefunden werden.

Historische Pakete können mit einem Testzertifikat signiert sein. Importiere ein
solches Zertifikat nur, wenn Herkunft und Fingerabdruck geprüft wurden, nur für
den aktuellen Benutzer und nur für die erforderliche Installationsdauer. Die
ausführliche historische Anleitung erklärt auch den Windows-Fehler
„deployment and application do not have matching security zones“:
[Installations- und Zertifikatsanleitung](docs/latex/main.pdf).

> Das Repository enthält Quellcode, aber kein allgemein vertrauenswürdig
> signiertes Release-Paket. Für produktive Verteilung sollte ein eigenes,
> kontrolliertes Codesigning-Zertifikat verwendet werden.

### Direkt aus Visual Studio

1. `ALPS_Visio_Tools.sln` in Visual Studio 2022 öffnen.
2. NuGet-Pakete wiederherstellen.
3. `Debug | Any CPU` auswählen.
4. Das Projekt `ALPS_Visio_AddIn-rewrite` als Startprojekt setzen.
5. Mit **F5** starten. Visual Studio registriert das VSTO-Add-in für den
   Entwicklungsbenutzer und startet bzw. verbindet Visio.

## Erste Schritte

1. In Visio eine neue Zeichnung öffnen.
2. Im Tab **ALPS/PASS ADDIN** auf **Open ALPS/PASS Stencils** klicken.
3. Ein Modell mit den SID-/SBD-Schablonen erstellen oder über **Import OWL**
   eine Beispieldatei laden:
   - `docs/[Test]_Vacation_Request_2D.owl` enthält Layoutkoordinaten;
   - `docs/[Test]_Vacation_Request.owl` demonstriert das Fallback-Layout.
4. Mit **Show layer Explorer** zwischen Modellebenen und verknüpften Seiten
   navigieren.
5. Für eine Layer-Erweiterung im Explorer eine Hintergrundseite auswählen und
   ein `ActorExtension`- bzw. `StateExtension`-Shape nahe an seine Referenz
   bewegen. Den angebotenen Snap bestätigen.
6. Bei Bedarf **Auto-Arrange** ausführen und anschließend die fachliche
   Modellstruktur prüfen.
7. Vor der Weitergabe des Modells **Check Model Naming**, **Verify ALPS
   Models** oder **Convert PASS to BPMN** verwenden.

## Funktionen im Detail

### Schablonen und Modellierung

**Open ALPS/PASS Stencils** sucht die benötigten Visio-Schablonen und öffnet sie
für die interaktive Verwendung. Beim OWL-Import werden Schablonen getrennt
verwaltet, damit Makros nur dort aktiviert werden, wo sie für den Import nötig
sind. Abhängig von Windows-Vertrauensstellung und Schablonenherkunft kann Visio
eine Makro-Sicherheitsabfrage anzeigen.

Die bestehenden Modellierungsfunktionen der Schablonen bleiben erhalten,
darunter SID-/SBD-Elemente, Seitenverknüpfungen, ShapeSheet-Metadaten,
Connectoren und das dokumentbezogene Snapping.

### Layer Explorer

**Show layer Explorer** öffnet den WPF-basierten Modell-Explorer. Er zeigt die
Modellhierarchie, SID-Layer und verknüpfte SBD-Seiten. Der Explorer aktualisiert
sich beim Wechsel von Dokumenten und Seiten und unterstützt Navigation sowie
die vorhandenen Bearbeitungsdialoge.

Die zugehörigen Controller sind dokumentbezogen. Beim Aktualisieren werden alte
COM-Events gelöst, damit Aktionen nicht mehrfach ausgelöst werden.

### SnapHandler: semantisches SID-/SBD-Snapping

Der SnapHandler ist eine automatisch aktive Modellierungsfunktion und besitzt
keinen eigenen Ribbon-Befehl. Er unterstützt ALPS-Layer-Erweiterungen, indem er
ein Extension-Shape auf der Vordergrundseite mit dem fachlich entsprechenden
Element der eingeblendeten Hintergrundseite koppelt. Das ist mehr als das
native Visio-Snapping: Neben Position und Größe werden die für `extends` und
die Navigation benötigten ShapeSheet-Daten gepflegt.

#### Voraussetzungen und Aktivierung

Snapping wird nur auf vollständig initialisierten ALPS/PASS-Seiten aktiviert:

- Eine SID-Seite wird über ihre PageSheet-Zellen für Modell-URI, Seitentyp,
  Layer, Modellversion und Priorität erkannt.
- Eine SBD-Seite wird über die PageSheet-Referenz auf ihr verknüpftes Subjekt
  erkannt.
- Die Vordergrundseite muss über ihre `extends`-Eigenschaft eine andere SID-
  bzw. SBD-Seite erweitern. Der Controller setzt diese Seite zugleich als
  Visio-`BackPage`.
- Das bewegte Shape benötigt die richtige Kategorie aus der ALPS/PASS-Schablone;
  gewöhnliche Visio-Shapes werden ignoriert.

Die Seitenbeziehung kann über den Layer Explorer, die vorhandenen
Eigenschaftsdialoge, die Schablonenmakros oder gültige ShapeSheet-Daten
entstehen. Noch nicht vollständig aufgebaute Seiten werden nach `PageAdded`
weiter beobachtet und registriert, sobald ihre erforderlichen Zellen vorhanden
sind.

#### Gemeinsamer Ablauf

1. Ein Page-Controller beobachtet ShapeSheet-Änderungen der zugehörigen Seite.
   Bewegungen werden über `PinX` und `PinY`, explizite `extends`-Änderungen über
   die jeweilige Property- bzw. Hyperlink-Zelle erkannt.
2. Der Handler betrachtet nur passende Referenz-Shapes auf der aktuell
   erweiterten Hintergrundseite.
3. Liegen die Mittelpunkte in X- **und** Y-Richtung jeweils höchstens 20 mm
   auseinander, öffnet sich der lokalisierte Dialog **Snap-Einstellungen**.
   Eine bereits praktisch identische X-Position wird dabei unterdrückt, damit
   das exakt ausgerichtete Shape nicht sofort erneut angeboten wird.
4. **Ja** speichert die Zuordnung und richtet das Extension-Shape exakt auf der
   Referenz aus. Seine Breite und Höhe werden jeweils auf die Referenzgröße plus
   5 mm gesetzt; dadurch bleibt das überlagernde Extension-Shape sichtbar.
5. Bewegt sich die Referenz später auf einer als Hintergrund verwendeten Seite,
   wird das gekoppelte Extension-Shape erneut ausgerichtet und skaliert.
6. **Nein** entfernt eine bestehende Zuordnung und bereinigt die
   typabhängigen `extends`-Daten.

Wenn mehrere passende Referenzen gleichzeitig im 20-mm-Bereich liegen, kann
der Bestätigungsdialog nacheinander für jede mögliche Zuordnung erscheinen.
Die fachlich richtige Referenz ist dann anhand der angezeigten Shape-Namen zu
wählen.

#### Unterschiede zwischen SID und SBD

| Verhalten | SID-Snapping | SBD-Snapping |
| --- | --- | --- |
| Vordergrund-Shape | Kategorie `ActorExtension` | Kategorie `StateExtension` |
| Referenz auf der Hintergrundseite | Kategorie `StandardActor` | Kategorie `alpsSBDstate` |
| Explizite Auflösung | Shape-Name aus `extendedSubject` | `modelComponentID` aus `Prop.extends.Value` |
| Persistierte Referenz | Hyperlink `extendedSubject` als `<Layer>/<ShapeNameU>` und `Prop.extends.Value` als `<ModelURI>#<ShapeNameU>` | `Prop.extends.Value` mit der `modelComponentID` des Referenzzustands |
| Entfernen aus dem Fangbereich | sofortiges Unsnap | Dialog, ob die Kopplung beibehalten werden soll |
| Zusätzliche Semantik | synchronisiert nach Möglichkeit die verknüpften SBD-Seiten beider Subjekte | betrifft die Zustandsreferenz innerhalb der bereits festgelegten SBD-Seitenbeziehung |

##### SID: Subjekt- und Verhaltensvererbung

Ein `ActorExtension` kann nur auf einen `StandardActor` der erweiterten
SID-Hintergrundseite snappen. Nach der Bestätigung schreibt der Handler:

- den Navigations-Hyperlink `extendedSubject` mit Hintergrund-Layer und
  Shape-`NameU`;
- die vollständige `extends`-Referenz aus Modell-URI und Shape-`NameU`;
- die Vordergrund-/Hintergrundbeziehung der beiden verknüpften SBD-Seiten,
  sofern beide Subjekte einen gültigen `linkedSBD`-Hyperlink besitzen.

Wird eine bestehende SID-Kopplung gelöst, werden Hyperlink und `extends`-Wert
geleert. Die SBD-Seite des Vordergrundsubjekts erweitert dann nicht länger das
Verhalten der alten Referenz; auch die bisher als erweitert markierte
Hintergrund-SBD wird freigegeben. Beim Wechsel auf eine andere Referenz wird
eine vorherige SBD-Beziehung entsprechend entfernt.

`MacroExtension` ist ein Sonderfall: Geometrie, Subjekt-Hyperlink und
Shape-`extends` werden gesetzt, die SBD-Seiten werden jedoch bewusst nicht als
Verhaltensvererbung miteinander verbunden.

##### SBD: Zustandsvererbung

Ein `StateExtension` kann auf jedes Hintergrund-Shape mit der Kategorie
`alpsSBDstate` snappen. Die semantische Referenz verwendet nicht den sichtbaren
Shape-Namen, sondern die stabile `modelComponentID` des Zustands. Wird eine ID
manuell in `Prop.extends.Value` eingetragen, versucht der Handler dieselbe
Zuordnung wiederherzustellen; bei einer unbekannten ID erscheint eine
Fehlermeldung.

Wird ein bereits gekoppeltes `StateExtension` aus dem 20-mm-Bereich gezogen,
fragt **Snap-Einstellungen**, ob es gekoppelt bleiben soll:

- **Ja** behält die `extends`-Referenz und richtet Größe und Position bei Bedarf
  wieder an der Referenz aus.
- **Nein** entfernt die Zuordnung und leert `Prop.extends.Value`.

#### Hintergrunddarstellung und Seitenbeziehungen

Beim Erweitern einer SID- oder SBD-Seite verwaltet der Page-Controller neben
dem SnapHandler die Visio-Hintergrundseite. Ein Separator-Shape aus der
SID-Schablone visualisiert die Trennung zwischen Vorder- und Hintergrund und
wird auf die Größe der erweiterten Seite gebracht. Die Eigenschaftsdialoge
bieten dafür drei Darstellungen:

- **No separation:** vollständig transparent;
- **Standard separation:** stark transparent;
- **Full separation:** undurchsichtig.

Ändert sich die Größe der Hintergrundseite, wird auch der Separator neu
dimensioniert. Beim Entfernen der Seitenbeziehung werden Hintergrund und
Separator entfernt. Die Shape-Zuordnungen werden über die oben beschriebenen
SID-/SBD-spezifischen Unsnap-Pfade und ihre ShapeSheet-Referenzen bereinigt.

#### Lebenszyklus und Wiederherstellung

`ThisAddIn` besitzt genau einen `ModelController` für das zuletzt aktive
Zeichnungsdokument. Er registriert SID-/SBD-Seiten, erstellt pro Seite den
passenden Controller und rekonstruiert vorhandene Beziehungen aus den
ShapeSheet-Zellen. Beim Dokumentwechsel, Import oder Explorer-Refresh werden
alte Page-Controller zuerst freigegeben und ihre COM-Events abgemeldet. Dadurch
bleiben Snapping-Aktionen dokumentbezogen und alte Dokumente reagieren nach
einem Refresh nicht weiter.

Die gespeicherten Hyperlinks und `extends`-Werte sind daher entscheidend: Die
In-Memory-Zuordnung selbst ist nur für die aktuelle Controller-Laufzeit gültig
und wird beim Neuaufbau aus dem Visio-Dokument wiederhergestellt.

### OWL-/RDF-Import

**Import OWL** akzeptiert `.owl` und `.rdf`. Der Importer:

- lädt PASS-/ALPS-Modelle mit den eingebetteten Standard-PASS- und
  ALPS-Ontologien;
- exportiert alle lesbaren Prozessmodelle in Visio;
- erzeugt SID- und verknüpfte SBD-Seiten mit eindeutigen Namen;
- übernimmt IDs, Labels, Kommentare, Hyperlinks und vorhandene
  Visualisierungspunkte;
- verbindet Nachrichtenaustausche und Zustandsübergänge mit ihren semantischen
  Endpunkten;
- platziert SID-Nachrichten in den zugehörigen Message Boxes;
- nutzt bei fehlenden Koordinaten ein deterministisches Fallback-Layout;
- aktualisiert anschließend den Model Explorer.

Die Ontologien werden bevorzugt aus dem ausgelieferten `Resources`-Ordner
gelesen. Fehlen sie dort, materialisiert das Add-in die eingebetteten Kopien in
`%LOCALAPPDATA%\ALPS-Visio-Add-In\Ontologies`.

### Auto-Arrange

Der Split-Button **Auto-Arrange** arbeitet auf der aktiven Zeichnungsseite:

- **Top-down** ordnet Ränge von oben nach unten an und richtet die Druckseite
  im Hochformat aus.
- **Left-right** ordnet Ränge von links nach rechts an, berücksichtigt lange
  Nachrichtenlabels und verwendet Querformat.

Auf SBD-Seiten werden Zustände gerankt, Alternativen geordnet, Rückkanten in
äußere Korridore gelegt und Ports deterministisch verteilt. Auf SID-Seiten
nehmen nur Subjekte am Graph-Layout teil; Message Boxes und ihre Listenelemente
werden anschließend an den passenden Kommunikationskanälen positioniert.
Connectoren werden nach der Anordnung an die ursprünglichen semantischen
Endpunkte zurückgebunden. Wiederholte Läufe sollen dasselbe Ergebnis erzeugen.

### Lokale Prüfung von PASS-Elementnamen

**Check Model Naming** sammelt unterstützte Shapes aus allen Seiten und zeigt
Seite, Shape-ID, Elementtyp, Label, Klassifikation und Konfidenz. Unterstützt
werden derzeit:

- `FullySpecifiedSubject`
- `InterfaceSubject`
- `MultiSubject`
- `MessageSpecification`
- `DoState`
- `SendState`
- `ReceiveState`
- `DoTransition`

Der Klassifikator ist eine deterministische, im Prozess laufende Textklassifikation
ohne ML-Runtime. Er trainiert bei der ersten Verwendung aus der eingebetteten
TSV-Datei und funktioniert vollständig ohne API-Schlüssel oder Netzwerk.
**Retrain** verwirft das geladene Modell und trainiert es erneut aus den 680
mitgelieferten Beispielen.

#### Optionale Vorschlagsanbieter

In **Provider Settings** stehen OpenAI, Anthropic und UniGPT als eingebaute
Profile zur Verfügung. Zusätzlich können OpenAI-kompatible und
Anthropic-kompatible Endpunkte angelegt werden. Pro Profil lassen sich Basis-URL,
API-Schlüssel, verfügbares Modell und Aktivstatus verwalten.

Modelle werden authentifiziert über den jeweiligen `/models`-Endpunkt geladen.
Vorschläge werden nur für Einträge angefragt, die der lokale Klassifikator zur
Prüfung markiert. Ein Anbieterfehler wird am betroffenen Ergebnis angezeigt und
bricht die übrige lokale Prüfung nicht ab.

### ALPS-Verifikation

**Verify ALPS Models** fordert zuerst eine abstrakte Spezifikation und danach
das implementierende OWL/RDF-Modell an. Die UI-unabhängige Prüfung arbeitet
offline und kontrolliert den im integrierten Thesis-Prototyp unterstützten
SID-Umfang:

- `implements`-Abdeckung für Subjekte, Nachrichtenaustausche und
  Kommunikationsakte;
- Erhalt von `FullySpecifiedSubject`;
- Kommunikationsbeschränkungen in beiden Richtungen;
- mehrdeutige Mehrfachimplementierungen;
- nicht auflösbare Korrespondenten, Sender und Empfänger.

Das Ergebnisfenster trennt Fehler, Warnungen und Scope-Hinweise. Ein
tab-separierter Bericht inklusive beider Dateipfade kann in die Zwischenablage
kopiert werden. Deterministische Regression-Fixtures liegen in
`docs/[Test]_ALPS_Verification_*.owl`; der ursprüngliche Thesis-Korpus befindet
sich in `docs/verification-thesis/`.

Diese Funktion ist eine gezielte, nachvollziehbare SID-Prüfung und kein
vollständiger formaler ALPS-Konformitätsbeweis. Weitere Grenzen stehen unter
[Bekannte Grenzen](#bekannte-grenzen).

### PASS nach BPMN 2.0

**Convert PASS to BPMN** liest das erste lesbare PASS-Prozessmodell aus einer
`.owl`- oder `.rdf`-Datei und schreibt eine `.bpmn`-Datei. Der integrierte
Converter erzeugt unter anderem:

- Kollaboration, Teilnehmer/Pools und Prozesse;
- Tasks, Events, Gateways, Sequenz- und Nachrichtenflüsse;
- BPMN-DI-Shapes und orthogonal geroutete Kanten;
- kompakte Pool- und Prozesslayouts;
- XML-konforme, eindeutige IDs und auflösbare Referenzen.

Vor der Erfolgsmeldung validiert das Add-in das serialisierte Ergebnis. Es
verwirft unter anderem ungültige oder doppelte IDs, unbekannte Referenzen,
nicht-finite bzw. extreme Koordinaten, Kanten mit weniger als zwei Wegpunkten,
diagonale Kanten und bestimmte inkonsistente Gateway-Bedingungen.

## Datenschutz und externe Dienste

- OWL-Import, Auto-Arrange, Layer Explorer, ALPS-Verifikation und der lokale
  Namensklassifikator arbeiten ohne externe Übertragung.
- Ohne aktiv konfigurierten Vorschlagsanbieter führt **Check Model Naming**
  keinen API-Aufruf aus.
- Bei aktivierten Vorschlägen werden nur Shape-Typ und aktuelles Label der zur
  Prüfung markierten Elemente an den gewählten Anbieter gesendet. Dokumentpfad,
  Seitenname und Shape-ID gehören nicht zum Prompt.
- Providerprofile, Modellauswahl und API-Schlüssel werden für den aktuellen
  Windows-Benutzer gemeinsam mit DPAPI verschlüsselt unter
  `%LOCALAPPDATA%\ALPS Visio Add-In\nlp-providers.dat` gespeichert.
- Ein vorhandener älterer, DPAPI-geschützter UniGPT-Schlüssel wird aus
  `nlp-api-key.dat` migriert.
- API-Schlüssel oder personenbezogene Modelle dürfen nicht eingecheckt werden.

Die Datenschutz- und Nutzungsbedingungen des jeweils konfigurierten externen
Anbieters gelten zusätzlich.

## Entwicklung

### Wiederherstellen und Bauen

In einer **Developer Command Prompt for VS 2022**:

```powershell
nuget restore ALPS_Visio_Tools.sln
msbuild ALPS_Visio_Tools.sln /p:Configuration=Debug
msbuild ALPS_Visio_Tools.sln /p:Configuration=Release
```

Das Produktionsprojekt verwendet .NET Framework 4.8, C# 10, VSTO und die
Visio-/Office-Interop-Assemblies. NuGet-Abhängigkeiten werden klassisch in
`packages.config` verwaltet; generierte Ordner wie `packages/`, `bin/` und
`obj/` gehören nicht ins Repository.

### Debuggen

Das Add-in aus Visual Studio mit **F5** starten. Tests ohne COM-Zugriff können
im Test Explorer ausgeführt werden; Ribbon, ShapeSheet, Makros, Snapping und
visuelle Layoutqualität müssen zusätzlich in einer echten Visio-Instanz geprüft
werden.

## Tests und Qualitätssicherung

Die Solution enthält das separate Projekt
`ALPS_Visio_AddIn-rewrite.Tests`. Es verwendet MSTest und deckt die ohne aktive
Visio-Instanz testbaren Kernbereiche ab:

- Snapping-Geometrie;
- deterministische SBD-Ränge, Reihenfolge und Connector-Ports;
- Ontologie-Fallback und Cache-Aktualisierung;
- OWL-Modellauflösung und ALPS-Verifikationsregeln;
- lokales NLP-Training und Provider-Konfiguration;
- OpenAI- und Anthropic-kompatible Request-/Response-Verträge ohne echte
  Netzwerkaufrufe;
- BPMN-Validierung sowie die Konvertierung des Regression-Modells.

### Automatisierte Tests

```powershell
nuget restore ALPS_Visio_Tools.sln
msbuild ALPS_Visio_Tools.sln /p:Configuration=Debug
dotnet test ALPS_Visio_AddIn-rewrite.Tests\ALPS_Visio_AddIn-rewrite.Tests.csproj `
  --configuration Debug --no-build --no-restore
```

Alternativ können alle Tests direkt im Visual Studio Test Explorer gestartet
werden. Für Coverage:

```powershell
dotnet test ALPS_Visio_AddIn-rewrite.Tests\ALPS_Visio_AddIn-rewrite.Tests.csproj `
  --configuration Debug --no-build --no-restore `
  --collect:"XPlat Code Coverage" `
  --settings ALPS_Visio_AddIn-rewrite.Tests\coverlet.runsettings
```

Coverage misst die verwaltete, automatisierbare Kernlogik. COM-/VSTO-Code wird
bewusst über die manuelle Abnahme abgesichert und darf nicht durch eine hohe
Gesamtprozentzahl vorgetäuscht werden.

Für den SnapHandler prüfen die automatisierten Tests die COM-freie
Abstandsgeometrie und ihre 20-mm-Grenze. Die vollständige Ereigniskette aus
`CellChanged`, Visio-Hintergrundseite, Bestätigungsdialog, ShapeSheet-Schreibzugriff
und Weitergabe an verknüpfte SBD-Seiten benötigt eine echte Visio-Instanz und
gehört deshalb zusätzlich in die manuelle Abnahme.

### Manuelle Visio-Abnahme

Nach Änderungen an Ribbon, Import, Stencils, Connectoren, Auto-Arrange,
Snapping oder UI ist zusätzlich die vollständige
[manuelle Abnahme](MANUAL_TESTING.md) auf Windows/Visio auszuführen. Sie deckt
beide Vacation-Request-Modelle, SID-/SBD-Verknüpfungen, alle Layout-Richtungen,
NLP-Provider, Verifikation und BPMN-Sichtprüfung ab.

## Projektstruktur

```text
ALPS_Visio_Tools.sln
├── ALPS_Visio_AddIn-rewrite/          VSTO-Produktionsprojekt
│   ├── ThisAddIn.cs                   Lebenszyklus und Dokument-Controller
│   ├── ALPSRibbon.cs                  sichtbare Ribbon-Befehle
│   ├── Importing/                     OWL-Import und Ontologie-Auflösung
│   ├── OWLShapes/                     Visio-Adapter und Exportlogik
│   │   └── Layout/                    COM-freie Fallback-Layoutlogik
│   ├── VisioInfrastructure/           ShapeSheet, Stencils, Routing, Layout
│   ├── Snapping/                      dokumentbezogenes SID-/SBD-Snapping
│   ├── NlpChecking/                   lokaler Klassifikator und Provider
│   ├── Verification/                  offline ALPS-SID-Verifikation
│   ├── BpmnConversion/                integrierter BPMN-2.0-Converter
│   ├── Compatibility/                 Legacy-Konstanten und Fassaden
│   ├── _old/UI/                       weiterhin aktive Legacy-WPF-Oberfläche
│   └── Resources/                     Ontologien, Bilder, Texte, Trainingsdaten
├── ALPS_Visio_AddIn-rewrite.Tests/    automatisierte Unit-/Regressionstests
├── docs/                              Beispiele und Hintergrunddokumentation
├── MANUAL_TESTING.md                  Windows-/Visio-Abnahmecheckliste
├── FUNCTIONAL_SCOPE.md                Abgrenzung zum früheren Stand
└── THIRD_PARTY_NOTICES.md              Herkunft integrierter Komponenten
```

`VisioHelper.cs` bleibt eine Kompatibilitätsfassade. Wiederverwendbare
COM-Operationen gehören nach `VisioInfrastructure/`; neue Tests in das separate
`*Tests`-Projekt.

## Bekannte Grenzen

- Build und Ausführung des VSTO-Add-ins sind Windows-/Visio-gebunden.
- Die ALPS-Verifikation deckt den implementierten SID-Teil des
  Thesis-Prototyps ab. SBD-Ablauflogik, Präzedenz-/Trigger-Transitionen,
  abstrakte Kommunikationskanäle, Finalized-Message-Semantik,
  Subjekt-Multiplizitäten sowie Start-/Endeigenschaften sind nicht vollständig
  formal verifiziert.
- Der BPMN-Befehl konvertiert das erste lesbare Prozessmodell der ausgewählten
  Datei, nicht mehrere Modelle in eine gemeinsame BPMN-Datei.
- Die Qualität der optionalen Namensvorschläge hängt vom gewählten externen
  Anbieter und Modell ab; der lokale Klassifikator bleibt die einzige
  Offline-Komponente dieses Vorschlagspfads.
- Der SnapHandler unterstützt ausschließlich die vorgesehenen Kategorien und
  ShapeSheet-Zellen der ALPS/PASS-Schablonen. Shapes ohne diese Kategorien und
  unvollständig initialisierte Seiten werden nicht semantisch gekoppelt.
- Das automatische Nachführen einer verschobenen Hintergrundreferenz ist für
  die fachlich vorgesehene eindeutige Extension-Zuordnung ausgelegt. Mehrere
  Extension-Shapes auf derselben Referenz müssen in der manuellen Abnahme
  besonders geprüft werden.
- Visuelle Qualität, Makro-Vertrauen und COM-Event-Lebenszyklen lassen sich
  nicht vollständig in einem Headless-Test nachbilden.

Weitere Refactoring-Grenzen stehen in [REFACTORING.md](REFACTORING.md), die
exakte Funktionsabgrenzung in [FUNCTIONAL_SCOPE.md](FUNCTIONAL_SCOPE.md).

## Fehlerbehebung

### Der Ribbon-Tab erscheint nicht

- Visio schließen und die Installation bzw. den F5-Start erneut ausführen.
- Unter **File → Options → Add-ins** prüfen, ob das Add-in deaktiviert wurde.
- VSTO Runtime, .NET Framework 4.8 und passende Office-Bitness prüfen.
- In Visual Studio die richtige Debug-Konfiguration und das VSTO-Projekt als
  Startprojekt wählen.

### Schablonen werden nicht gefunden oder Makros reagieren nicht

- SID-/SBD-Schablonen in **My Shapes** installieren.
- Windows-Dateiblockierung und Visio Trust Center prüfen.
- Schablonen über **Open ALPS/PASS Stencils** interaktiv neu öffnen.
- Makros nur aus geprüften Quellen aktivieren.

### OWL-Import schlägt fehl

- Dateiendung und RDF/OWL-Syntax prüfen.
- Mit `docs/[Test]_Vacation_Request_2D.owl` gegenprüfen.
- Sicherstellen, dass `%LOCALAPPDATA%\ALPS-Visio-Add-In\Ontologies` für den
  Benutzer beschreibbar ist.
- Ausnahme, Eingabedatei und betroffene Seite für einen Fehlerbericht sichern.

### Extension-Shapes snappen nicht

- Prüfen, ob die Vordergrundseite tatsächlich über `extends` mit der richtigen
  SID- bzw. SBD-Hintergrundseite verbunden ist und diese in Visio sichtbar ist.
- Ausschließlich die Extension-Shapes aus den ALPS/PASS-Schablonen verwenden:
  `ActorExtension` im SID oder `StateExtension` im SBD.
- Das Shape so bewegen, dass sein Mittelpunkt in beiden Achsen höchstens 20 mm
  von der Referenz entfernt ist. Für einen neuen Snap nicht bereits exakt auf
  derselben X-Position starten.
- Im SID muss die Referenz ein `StandardActor` sein; im SBD muss sie die
  Kategorie `alpsSBDstate` und eine `modelComponentID` besitzen.
- Bei SID-Verhaltensvererbung zusätzlich die `linkedSBD`-Hyperlinks beider
  Subjekte prüfen. `MacroExtension` erzeugt absichtlich keine SBD-Vererbung.
- Den Layer Explorer neu öffnen oder aktualisieren, damit die dokumentbezogenen
  Controller aus den ShapeSheet-Daten neu aufgebaut werden.
- Falls ein Dialog oder eine Bewegung mehrfach verarbeitet wird, Dokumentname,
  Seite, Shape-Namen und die genaue Aktion notieren und den Event-Lebenszyklus
  nach [MANUAL_TESTING.md](MANUAL_TESTING.md) prüfen.

### Provider oder Modelle können nicht geladen werden

- Basis-URL, Protokolltyp, API-Schlüssel und Modell prüfen.
- Proxy-/Firewall-Regeln sowie TLS 1.2 berücksichtigen.
- Die Namensprüfung ohne aktiven Anbieter erneut ausführen; der lokale
  Klassifikator muss weiterhin funktionieren.

### ClickOnce meldet unterschiedliche Sicherheitszonen

Das Paket vollständig in einen lokalen Ordner entpacken und die
Windows-Dateiblockierung nur für verifizierte Installationsdateien entfernen.
Details stehen in der
[Installations- und Zertifikatsanleitung](docs/latex/main.pdf).

## Weitere Dokumentation

- [Manuelle Abnahmetests](MANUAL_TESTING.md)
- [Funktionsumfang und Abgrenzung](FUNCTIONAL_SCOPE.md)
- [Refactoring-Notizen](REFACTORING.md)
- [ShapeSheet-Daten](docs/Data%20in%20ShapeSheet.md)
- [ALPS-Verifikations-Fixtures](docs/verification-thesis/README.md)
- [Drittkomponenten und Lizenzen](THIRD_PARTY_NOTICES.md)

## Lizenz und Drittkomponenten

Der Repository-Code steht unter der in [LICENSE](LICENSE) enthaltenen Lizenz.
Der integrierte PASS-zu-BPMN-Converter und die übernommene
ALPS-Verifikationslogik haben eigene Herkunfts- und Lizenzhinweise. Vor einer
Weitergabe sind [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) und die
Upstream-Lizenz unter
`ALPS_Visio_AddIn-rewrite/BpmnConversion/Upstream/LICENSE` zu beachten.
