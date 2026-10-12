PowerPoint Stay On Slide Points Add-In v1.0
===================================================

Überblick:
Das PowerPointStayOnSlidePointsAddIn löst ein zentrales Problem in Microsoft PowerPoint:
Im normalen Folien-Bearbeitungsmodus führt das Scrollen mit dem Mausrad oder die Betätigung
von BildAuf / BildAb am Folienrand standardmäßig zum sofortigen, unbeabsichtigten
Wechseln auf die nächste oder vorherige Folie.

Mit aktiviertem Stay on Slide Modus:
- Bleibt die Ansicht fest auf der aktuellen Folie fixiert.
- Das Scrollen stoppt am oberen und unteren Rand, anstatt die Folie zu wechseln.
- Ein Folienwechsel erfolgt nur noch bewusst durch Klicken oder Navigieren in der
  linken Folien-Miniaturleiste (Thumbnails Sidebar).
- Automatischer Montagebereich (Virtual Pasteboard):
  Das Add-In erweitert den scrollbaren Bereich ober- und unterhalb der Folie
  automatisch um konfigurierbare Puffer (Standard: 1500 pt). Dadurch können Inhalte
  problemlos weit über den Folienrand hinaus aufgebaut werden, ohne dass manuell
  Dummy-Objekte platziert oder verschoben werden müssen.
- Temporäre Montage-Anker werden vor dem Speichern, Drucken oder Schließen automatisch
  rückstandslos entfernt, sodass die Präsentationsdatei sauber bleibt.

Menüband (Ribbon):
- Tab: Add-Ins
- Gruppe: Stay on Slide v1.0
- Toggle-Button: Stay on Slide (Großer Ein/Aus-Schalter)
- Button: Settings (Einstellungsdialog für Puffer und Tastatursperre)
- Button: About (Produktinfo, Versionshinweise und Pfade)

Systemanforderungen:
- Microsoft Windows 10 / 11
- Microsoft PowerPoint 2016 / 2019 / 2021 / Microsoft 365 (Desktop-Version)
- .NET Framework 4.8
- Visual Studio 2010 Tools for Office Runtime (VSTO Runtime)
