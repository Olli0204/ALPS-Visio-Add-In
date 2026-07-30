# Hallo
... und herzlich Willkommen im Visio ALPS AddIn Repository!

## Aufbau
Damit dein Start *etwas* leichter als meiner ist, erkläre ich dir,
wie das AddIn aufgebaut ist und funktioniert.

### Entrypoint
![ThisAddIn.svg](../docs/ThisAddIn.svg)
Beim Start wird `ThisAddIn#ThisAddIn_Startup` ausgeführt. Das initialisiert vor allem den Model Explorer und die Snap Handler.
Die produktiven Komponenten liegen in `_old/UI/` und `Snapping/`; weitere
verhaltensnahe Änderungen benötigen Windows-/Visio-Charakterisierungstests. Der
WPF-Pfad `_old/UI/` bleibt wegen der Pfadabhängigkeit des alten
WinFX-Markup-Compilers bestehen.

Außerdem werden die Buttons erzeugt, siehe dazu `ALPSRibbon`.

### Importer
![OWLImporter.svg](../docs/OWLImporter.svg)
Der Importer wird aufgerufen, wenn der `Import OWL`-Button gedrückt und eine Datei ausgewählt wurde. Der API-Parser lädt
alle darin definierten Modelle; `OwlImportService` exportiert anschließend jedes unterstützte Modell nach Visio.

### Modell
![PASSProcessModel.svg](../docs/PASSProcessModel.svg)
Bestensfalls könnte man die Codestruktur durch mehrfaches Vererben vereinfachen, aber wir sind leider wegen Visio gezwungen,
in dieser veralteten C#-Version zu bleiben. Es sei ebenfalls erwähnt, dass die Diagramme nicht den üblichen Konventionen
entsprechen, sondern lediglich einen Ansatz zur Orientierung im Code liefern sollen.

Jedes Modell enthält Layer. Sind es mehr als `1`, so ist das ein Feature von ALPS (gegenüber PASS).

Layer enthalten Subjekte und Nachrichten; letztere sind dabei in der API als Nachrichten-Liste in den Verbinder und die
Nachrichten aufgeteilt. Über die Umsetzung in Visio reden wir lieber nicht...

Subjekte können vielerlei verschiedener Art sein. Das einzige, welches fast vollständig implementiert ist, ist das
`FullySpecifiedSubject`.

Subjekte exportieren sich selber über eine Hilfsklasse, `SubjectExport`. Diese erzeugt dann wiederum das SubjectBehavior,
welches Zustände und Transitionen enthält (welche wiederum ihre eigenen Hilfsklassen zum Export haben).

### Konstanten und Hilfsklassen
`Constants` enthält die Konstanten der refaktorierten Visio-Implementierung. Die umfangreichere Übergangsversion
`ALPSConstants` und `ALPSGlobalFunctions` liegt in `Compatibility/`; eine offene Aufgabe ist, noch benötigte Werte
schrittweise in die neue Struktur zu übertragen.

`VisioHelper` ist eine schmale, quellkompatible Fassade. Die Implementierungen für ShapeSheet, Stencils, Seiten,
Positionierung, Routing und Auto-Arrange liegen in `VisioInfrastructure/`. `ShapeFinder` bleibt als quellkompatible
Fassade erhalten und delegiert die Dateisystemsuche sowie die Auswahl der neuesten installierten Stencil-Version an
`VisioInfrastructure/StencilFileLocator`.

Das deterministische Ersatzlayout bleibt über `OWLShapes/VisioLayout` erreichbar.
Graph-Rangberechnung, Connector-Portplanung, Bounds-Auflösung und schwach
referenzierter Laufzeitzustand liegen getrennt unter `OWLShapes/Layout/`.

`ModelController` besitzt die Snapping-Controller eines Dokuments. Beim Refresh,
Dokumentwechsel und Add-in-Shutdown werden deren Visio-Events explizit abgemeldet;
es gibt keinen globalen SID-Controller-Cache mehr.

Der Model Explorer verbleibt aus Build-Kompatibilitätsgründen in `_old/UI/`.
Baumaufbau und Prioritätsnormalisierung sind jedoch aus dem WPF-Code-behind in
`LayerExplorerTreeBuilder` ausgelagert.

## nächste Schritte
Jetzt da du dich hoffentlich in angemessenerer Zeit einarbeiten konntest, kommen die nächsten Aufgaben auf dich zu.

### Erledigte Stabilitätsarbeiten
- Texte werden zentral als Visio-Formel-Literale geschrieben und Anführungszeichen korrekt escaped.
- SID- und SBD-Seiten erhalten bei Namenskollisionen einen eindeutigen Suffix.
- Modelle ohne 2D-Koordinaten erhalten ein deterministisches, einfaches Rasterlayout.
- Die Ontologien werden aus `Resources/` geladen; der Import öffnet SID makroaktiv und SBD ohne zweiten Sicherheitsdialog.
- Dokumentation ist teilweise unvollständig oder fehlt komplett. Ein einheitliches Schema wäre von Vorteil - ich habe
bisher die JavaDoc Konventionen übernommen. Inline-Kommentare sollten reduziert werden und nur für die aktive Entwicklung
(z.B. Notiz von Aufgaben) benutzt werden. Nur in Ausnahmefällen dürfen einzelne Zeilen mit einem Kommentar erklärt werden;
im Allgemeinen ist sprechender Code besser.
- `Constants` sollte weiter vervollständigt und mit `Compatibility/ALPSConstants.cs` konsolidiert werden.
  ShapeSheet-Zugriffe und kollisionssichere Seitennamen sind bereits zentralisiert.

## offene Aufgaben
- Die API stimmt nicht immer mit der Ontologie überein. Daher können manche Features aktuell nicht implementiert werden.
Eine Dokumentation der API wäre sehr von Vorteil.
	- Subjekte: `hasSubjectExecutionMapping` wird nur für `FullySpecifiedSubject` implementiert.
	- Bei manchen Eigenschaften habe ich einen Kommentar `// alps.net.api` dazugeschrieben, diese habe ich zwar nicht
gefunden, konnte aber auch nicht verifizieren, dass sie nicht in der API existieren.
- Alle Eigenschaften aus der Ontologie sollten implementiert werden.
- `Snapping/` und der Model Explorer in `_old/UI/` funktionieren produktiv.
  Weitere Verhaltensänderungen benötigen Windows-/Visio-Charakterisierungstests;
  der XAML-Pfad muss wegen des alten WinFX-Compilers erhalten bleiben.

## Empfehlungen und persönliche Hinweise
Ich habe im Laufe meiner Entwicklung mehrere *Mini-Dokumentationen* geschrieben, diese habe ich alle mit in den `docs`
Ordner im Wurzelordner gelegt. (`combined-onts.notes`, `Data in ShapeSheet.md`) Ebenfalls beigefügt ist
eine Syntax-Highlight Erweiterung für VSCodium (wahrscheinlich auch VSCode) für die `.notes` Datei.

Protégé ist manchmal etwas komisch, dennoch hilft der Reasoner sehr gut dabei, die Ontologie zu verstehen.

Dein Computer ist nicht langsam, das ist Visio.

Die nächsten fachlichen Ausbaupunkte sind die noch markierten Ontologieeigenschaften in `SubjectExport`, `StateExport`
und `TransitionExport`. Strukturell folgen Charakterisierungstests für `ModelController`, Model Explorer und Snapping.

Die Regressionseingaben `docs/[Test]_Vacation_Request_2D.owl` und `docs/[Test]_Vacation_Request.owl` decken Modelle mit
und ohne Koordinaten ab. Beide gehören zum manuellen Abnahmelauf in `MANUAL_TESTING.md`.

---

Falls du mehr Fragen hast: Irgendwer (@MatthesElstermann) hat bestimmt meinen Kontakt.

Ansonsten wünsche ich frohes Entwickeln! :D
