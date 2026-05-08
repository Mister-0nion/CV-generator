# 📄 CV Generator — Mister-0nion's-Layout

Professioneller Lebenslauf-Generator mit automatischem Layout-Management und 5 Farbpaletten.

## 🚀 Deployment auf Streamlit Cloud (kostenlos, öffentliche URL)

### Schritt 1: GitHub-Repository anlegen

1. Gehe zu [github.com](https://github.com) → **New repository**
2. Name: `cv-generator` (oder beliebig)
3. Sichtbarkeit: **Public** ✓ (Streamlit Cloud braucht das für kostenloses Hosting)
4. **Create repository** klicken

### Schritt 2: Diese Dateien hochladen

Lade alle folgenden Dateien in dein Repository hoch (drag & drop auf GitHub):

```
📁 dein-repo/
├── streamlit_app.py      ← Haupt-App
├── generate_cv.js        ← CV-Generator (Node.js)
├── requirements.txt      ← Python-Abhängigkeiten
├── packages.txt          ← System-Pakete (Node.js)
├── package.json          ← npm-Konfiguration
└── .streamlit/
    └── config.toml       ← Streamlit-Konfiguration
```

> **Tipp:** Den Ordner `.streamlit/` mit `config.toml` auf GitHub anlegen:  
> Klicke "Add file" → "Create new file" → tippe `.streamlit/config.toml` als Dateiname.

### Schritt 3: Mit Streamlit Cloud verbinden

1. Gehe zu [share.streamlit.io](https://share.streamlit.io)
2. Mit GitHub-Account einloggen (kostenlos)
3. **"New app"** klicken
4. Repository auswählen → Branch: `main`
5. Main file path: `streamlit_app.py`
6. **"Deploy!"** klicken

Nach ca. 2–3 Minuten bekommst du eine URL wie:  
`https://dein-name-cv-generator-streamlit-app-xyz.streamlit.app`

---

## 💻 Lokal starten

```bash
# Einmalig installieren:
pip install streamlit pillow
npm install

# Starten:
streamlit run streamlit_app.py
# → http://localhost:8501
```

---

## 🎨 Features

- **5 Farbpaletten:** Forest Green, Navy Blue, Bordeaux, Slate, Cognac
- **Eigene Farben:** Sidebar, Akzent, Überschriften individuell anpassbar
- **Alle 5 auf einmal** generieren und vergleichen
- **Profilbild** mit automatischem Zuschnitt (Gesichts-Heuristik)
- **KPI-Badges** pro Job-Eintrag
- **Layout-Engine:** Schriftgröße und Abstände passen sich automatisch an die Inhaltsmenge an
- **Download** als .docx (Word-kompatibel)

---

## 📋 Eingabe-Struktur

Das Formular ist in 4 Tabs gegliedert:

| Tab | Inhalt |
|-----|--------|
| 👤 Persönliche Daten | Name, Kontakt, Kurzprofil, Foto |
| 💼 Berufserfahrung | Stellen mit KPI-Badges und Bullet-Beschreibungen |
| ⭐ Skills & Tools | Kategorisierte Skills mit Punkte-Rating (1–5) |
| 🎓 Ausbildung & Sprachen | Bildungsweg und Sprachkenntnisse |

---

## 🏗 Technischer Stack

- **Frontend:** Streamlit (Python)
- **DOCX-Generierung:** Node.js + docx-js
- **Bildverarbeitung:** Pillow (Python)
- **Layout-Engine:** JavaScript (generate_cv.js)
- **Hosting:** Streamlit Cloud (kostenlos)
