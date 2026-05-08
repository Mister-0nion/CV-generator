"""
streamlit_app.py — CV Generator (Zetsche-Layout)
Starten lokal:  streamlit run streamlit_app.py
Hosting:        streamlit.io (kostenlos)
"""

import streamlit as st
import json, uuid, subprocess, tempfile, os
from pathlib import Path
from PIL import Image
import io

# ─── Einmalig: node_modules installieren (Streamlit Cloud) ───────────────────
@st.cache_resource
def setup_node_modules():
    """Installiert docx via npm – läuft einmalig beim App-Start."""
    app_dir = Path(__file__).parent
    node_modules = app_dir / "node_modules" / "docx"
    if not node_modules.exists():
        result = subprocess.run(
            ["npm", "install", "--prefix", str(app_dir)],
            capture_output=True, text=True, timeout=120
        )
        if result.returncode != 0:
            return f"npm Fehler: {result.stderr[:300]}"
        return "npm install OK"
    return "node_modules bereits vorhanden"

_npm_status = setup_node_modules()

# ─── Seitenkonfiguration ─────────────────────────────────────────────────────
st.set_page_config(
    page_title="CV Generator",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── Paletten ────────────────────────────────────────────────────────────────
PALETTES = {
    "green":    {"label": "🌿 Forest Green",  "hex": "#0A5C46"},
    "navy":     {"label": "🌊 Navy Blue",      "hex": "#1E3A5F"},
    "bordeaux": {"label": "🍷 Bordeaux",       "hex": "#6B1A2A"},
    "slate":    {"label": "🪨 Slate",          "hex": "#2D3748"},
    "cognac":   {"label": "🍂 Cognac",         "hex": "#7B3F1E"},
}

PALETTE_COLORS = {
    "green":    {"sidebarBg":"0A5C46","sidebarAccent":"2EB88A","mainAccent":"0A5C46","mainHeading":"164D3A","tagBg":"D6F0E8"},
    "navy":     {"sidebarBg":"1E3A5F","sidebarAccent":"4A90D9","mainAccent":"1E3A5F","mainHeading":"0D2035","tagBg":"D0E4F7"},
    "bordeaux": {"sidebarBg":"6B1A2A","sidebarAccent":"C0435A","mainAccent":"6B1A2A","mainHeading":"4A0E1A","tagBg":"F5D5DB"},
    "slate":    {"sidebarBg":"2D3748","sidebarAccent":"68D391","mainAccent":"2D3748","mainHeading":"1A202C","tagBg":"E2E8F0"},
    "cognac":   {"sidebarBg":"7B3F1E","sidebarAccent":"E07B39","mainAccent":"7B3F1E","mainHeading":"5A2D0C","tagBg":"F5E6D3"},
}

# ─── CSS ─────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* Allgemein */
[data-testid="stSidebar"] { background: #f8f7f4; }
h1 { font-size: 1.6rem !important; }
h2 { font-size: 1.1rem !important; border-bottom: 2px solid #0A5C46; padding-bottom: 4px; }
h3 { font-size: 0.95rem !important; color: #374151; }

/* Farbpaletten-Cards */
.palette-row { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 1rem; }
.palette-chip {
    display: inline-flex; align-items: center; gap: 8px;
    padding: 6px 14px; border-radius: 20px;
    border: 2px solid #e4e4e7; background: white;
    font-size: 0.82rem; font-weight: 500; cursor: pointer;
}

/* Trennlinie */
hr { border: none; border-top: 1px solid #e4e4e7; margin: 1.2rem 0; }

/* Vorschau-Box */
.preview-box {
    border-radius: 8px; overflow: hidden;
    border: 1px solid #e4e4e7; margin-top: 0.5rem;
}
.preview-sidebar {
    padding: 10px 14px;
}
.preview-main {
    background: white; padding: 10px 14px;
}

/* Expander-Styling */
.stExpander { border: 1px solid #e4e4e7 !important; border-radius: 8px !important; }
</style>
""", unsafe_allow_html=True)

# ─── Session State initialisieren ────────────────────────────────────────────
def init_state():
    defaults = {
        "experience": [{"title":"","company":"","start":"","end":"heute","location":"","kpis":"","description":""}],
        "skills":     [{"category":"Marketing","items":[{"name":"","level":4}]}],
        "languages":  [{"name":"","level":""}],
        "education":  [{"degree":"","institution":"","year":"","details":""}],
        "palette":    "green",
        "custom_mode": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# ─── Sidebar: Design ─────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🎨 Farbpalette")

    # Palette auswählen
    selected = st.radio(
        "Stil wählen",
        options=list(PALETTES.keys()),
        format_func=lambda k: PALETTES[k]["label"],
        index=list(PALETTES.keys()).index(st.session_state.palette),
        key="palette_radio",
        label_visibility="collapsed",
    )
    st.session_state.palette = selected
    st.session_state.custom_mode = False

    # Live-Vorschau
    pc = PALETTE_COLORS[selected]
    st.markdown(f"""
    <div class="preview-box">
      <div class="preview-sidebar" style="background:#{pc['sidebarBg']}">
        <div style="color:#{pc['sidebarAccent']};font-size:0.6rem;font-weight:700;text-transform:uppercase;letter-spacing:0.1em">Sidebar</div>
        <div style="color:white;font-size:1rem;font-weight:700;margin-top:2px">Max Mustermann</div>
        <div style="color:#{pc['sidebarAccent']};font-size:0.7rem">Senior Manager</div>
      </div>
      <div class="preview-main">
        <div style="color:#{pc['mainHeading']};font-size:0.7rem;font-weight:700;text-transform:uppercase;border-bottom:2px solid #{pc['mainAccent']};padding-bottom:2px;display:inline-block">Berufserfahrung</div>
        <div style="margin-top:5px;font-size:0.7rem;color:#1F2937;font-weight:600">Senior Designer</div>
        <div style="color:#{pc['mainAccent']};font-size:0.68rem">Firma GmbH · 2022 – heute</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    # Eigene Farben
    with st.expander("⚙️ Eigene Farben anpassen"):
        st.session_state.custom_mode = True
        c_sidebar  = st.color_picker("Sidebar-Farbe",  f"#{pc['sidebarBg']}")
        c_accent   = st.color_picker("Sidebar-Akzent", f"#{pc['sidebarAccent']}")
        c_main     = st.color_picker("Haupt-Akzent",   f"#{pc['mainAccent']}")
        c_heading  = st.color_picker("Überschriften",  f"#{pc['mainHeading']}")
        c_tag      = st.color_picker("Tag-Hintergrund",f"#{pc['tagBg']}")

    st.markdown("---")
    st.markdown("""
    <div style="font-size:0.75rem;color:#6B7280;line-height:1.6">
    💡 <b>Tipp:</b> Wähle eine Palette oder passe einzelne Farben manuell an.<br><br>
    📄 Der Lebenslauf wird als <b>.docx</b> (Word) heruntergeladen und kann in Word, LibreOffice oder Google Docs weiter bearbeitet werden.
    </div>
    """, unsafe_allow_html=True)

# ─── Hauptbereich ────────────────────────────────────────────────────────────
st.markdown("# 📄 CV Generator")
st.caption("Professioneller Lebenslauf im Zetsche-Layout · automatisch layout-optimiert")
st.markdown("---")

# ── Tab-Struktur ──────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "👤 Persönliche Daten",
    "💼 Berufserfahrung",
    "⭐ Skills & Tools",
    "🎓 Ausbildung & Sprachen",
    "✦ Generieren",
])

# ═════════════════════════════════════════════════════════════════════════════
# TAB 1: Persönliche Daten
# ═════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown("## 👤 Persönliche Daten")

    col1, col2 = st.columns(2)
    with col1:
        name     = st.text_input("Vollständiger Name *", placeholder="Max Mustermann")
        email    = st.text_input("E-Mail *", placeholder="max@beispiel.de")
        location = st.text_input("Ort / Adresse", placeholder="Berlin, Deutschland")
    with col2:
        title    = st.text_input("Berufsbezeichnung", placeholder="Senior Marketing Manager")
        phone    = st.text_input("Telefon", placeholder="+49 170 1234567")
        linkedin = st.text_input("LinkedIn / Portfolio", placeholder="linkedin.com/in/max")

    st.markdown("---")
    summary = st.text_area(
        "Kurzprofil / Über mich",
        placeholder="Erfahrener Spezialist mit 6+ Jahren in B2B SaaS-Marketing...",
        height=120,
    )

    st.markdown("---")
    st.markdown("### 📷 Profilbild (optional)")
    col_photo, col_hint = st.columns([1, 2])
    with col_photo:
        photo_file = st.file_uploader(
            "Bild hochladen",
            type=["jpg","jpeg","png","webp"],
            label_visibility="collapsed",
        )
        if photo_file:
            img = Image.open(photo_file)
            w, h = img.size
            size = min(w, h)
            left = (w - size) // 2
            top  = max(0, int((h - size) * 0.35))
            top  = min(top, h - size)
            img_crop = img.crop((left, top, left+size, top+size))
            img_crop = img_crop.resize((200, 200), Image.LANCZOS)
            st.image(img_crop, width=130, caption="Vorschau (quadratisch zugeschnitten)")
    with col_hint:
        st.markdown("""
        <div style="font-size:0.82rem;color:#6B7280;margin-top:0.5rem;line-height:1.7">
        • JPG oder PNG, max. 15 MB<br>
        • Wird automatisch quadratisch zugeschnitten<br>
        • Gesichts-Heuristik: oberer Bildbereich bevorzugt<br>
        • Erscheint oben in der Sidebar
        </div>
        """, unsafe_allow_html=True)

# ═════════════════════════════════════════════════════════════════════════════
# TAB 2: Berufserfahrung
# ═════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown("## 💼 Berufserfahrung")
    st.caption("Neueste Stelle zuerst. Zeilenumbrüche in der Beschreibung werden als Bullet-Punkte dargestellt.")

    exp_list = st.session_state.experience

    for i, job in enumerate(exp_list):
        with st.expander(f"**Stelle {i+1}:** {job['title'] or '(leer)'} @ {job['company'] or ''}", expanded=(i==0)):
            c1, c2 = st.columns(2)
            with c1:
                exp_list[i]["title"]   = st.text_input("Jobtitel *", value=job["title"], key=f"exp_title_{i}", placeholder="Senior Marketing Manager")
                exp_list[i]["start"]   = st.text_input("Start", value=job["start"], key=f"exp_start_{i}", placeholder="Jan 2022")
                exp_list[i]["location"]= st.text_input("Ort", value=job["location"], key=f"exp_loc_{i}", placeholder="Berlin")
            with c2:
                exp_list[i]["company"] = st.text_input("Unternehmen *", value=job["company"], key=f"exp_company_{i}", placeholder="Acme GmbH")
                exp_list[i]["end"]     = st.text_input("Ende", value=job["end"], key=f"exp_end_{i}", placeholder="heute")

            exp_list[i]["kpis"] = st.text_input(
                "KPI-Highlights (optional, kommasepariert)",
                value=job["kpis"], key=f"exp_kpis_{i}",
                placeholder="8+ Märkte bespielt, 20+ Stakeholder koordiniert, 100% Globale Reichweite",
                help="Erscheinen als farbige Zahlen-Badges über der Beschreibung"
            )
            exp_list[i]["description"] = st.text_area(
                "Aufgaben & Erfolge (jede Zeile = ein Bullet-Punkt)",
                value=job["description"], key=f"exp_desc_{i}",
                height=130,
                placeholder="Konzeption & Rollout einer Customer Journey...\nAufbau eines Retention-Systems...\nROI-Überwachung via PowerBI...",
            )

            if st.button(f"🗑 Stelle {i+1} entfernen", key=f"del_exp_{i}"):
                st.session_state.experience.pop(i)
                st.rerun()

    if st.button("➕ Stelle hinzufügen", use_container_width=True):
        st.session_state.experience.append({"title":"","company":"","start":"","end":"heute","location":"","kpis":"","description":""})
        st.rerun()

# ═════════════════════════════════════════════════════════════════════════════
# TAB 3: Skills & Tools
# ═════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown("## ⭐ Kernkompetenzen")
    st.caption("Erscheinen in der Sidebar mit Punkte-Rating (1–5 Sterne).")

    skill_list = st.session_state.skills

    for ci, cat in enumerate(skill_list):
        with st.expander(f"**Kategorie {ci+1}:** {cat['category'] or '(leer)'}", expanded=(ci==0)):
            skill_list[ci]["category"] = st.text_input(
                "Kategorie-Name", value=cat["category"],
                key=f"scat_{ci}", placeholder="Marketing"
            )
            st.markdown("**Skills mit Niveau:**")
            items = cat["items"]
            for ii, item in enumerate(items):
                c1, c2, c3 = st.columns([3, 1, 0.4])
                with c1:
                    items[ii]["name"] = st.text_input(
                        "Skill", value=item["name"],
                        key=f"skill_name_{ci}_{ii}", placeholder="B2B SaaS Marketing",
                        label_visibility="collapsed"
                    )
                with c2:
                    items[ii]["level"] = st.select_slider(
                        "Level", options=[1,2,3,4,5], value=item.get("level",4),
                        key=f"skill_lvl_{ci}_{ii}",
                        label_visibility="collapsed",
                        format_func=lambda x: "●"*x + "○"*(5-x)
                    )
                with c3:
                    if st.button("✕", key=f"del_skill_{ci}_{ii}"):
                        skill_list[ci]["items"].pop(ii)
                        st.rerun()

            if st.button(f"+ Skill hinzufügen", key=f"add_skill_{ci}"):
                skill_list[ci]["items"].append({"name":"","level":4})
                st.rerun()

            if st.button(f"🗑 Kategorie entfernen", key=f"del_cat_{ci}"):
                st.session_state.skills.pop(ci)
                st.rerun()

    if st.button("➕ Kategorie hinzufügen", use_container_width=True):
        st.session_state.skills.append({"category":"","items":[{"name":"","level":4}]})
        st.rerun()

    st.markdown("---")
    st.markdown("## 🔧 Tools")
    st.caption("Erscheinen als fett gesetzte Badges in der Sidebar.")
    tools_raw = st.text_input(
        "Tools (kommasepariert)",
        placeholder="PowerBI, Looker, Salesforce, HubSpot, Jira",
        label_visibility="collapsed"
    )

# ═════════════════════════════════════════════════════════════════════════════
# TAB 4: Ausbildung & Sprachen
# ═════════════════════════════════════════════════════════════════════════════
with tab4:
    col_edu, col_lang = st.columns(2)

    with col_edu:
        st.markdown("## 🎓 Ausbildung")
        edu_list = st.session_state.education
        for i, edu in enumerate(edu_list):
            with st.expander(f"**Abschluss {i+1}:** {edu['degree'] or '(leer)'}", expanded=(i==0)):
                edu_list[i]["degree"]      = st.text_input("Abschluss *", value=edu["degree"], key=f"edu_deg_{i}", placeholder="B.A. Marketing")
                edu_list[i]["institution"] = st.text_input("Institution *", value=edu["institution"], key=f"edu_inst_{i}", placeholder="HS Mittweida")
                edu_list[i]["year"]        = st.text_input("Jahr / Zeitraum", value=edu["year"], key=f"edu_yr_{i}", placeholder="2011–2014")
                edu_list[i]["details"]     = st.text_input("Details (optional)", value=edu["details"], key=f"edu_det_{i}", placeholder="Note: 1,8 · Schwerpunkt KI")
                if st.button(f"🗑 Entfernen", key=f"del_edu_{i}"):
                    st.session_state.education.pop(i)
                    st.rerun()

        if st.button("➕ Abschluss hinzufügen", use_container_width=True):
            st.session_state.education.append({"degree":"","institution":"","year":"","details":""})
            st.rerun()

    with col_lang:
        st.markdown("## 🌍 Sprachen")
        lang_list = st.session_state.languages
        for i, lang in enumerate(lang_list):
            c1, c2, c3 = st.columns([2, 2, 0.5])
            with c1:
                lang_list[i]["name"]  = st.text_input("Sprache", value=lang["name"], key=f"lang_name_{i}", placeholder="Englisch", label_visibility="collapsed" if i>0 else "visible")
            with c2:
                lang_list[i]["level"] = st.text_input("Niveau", value=lang["level"], key=f"lang_lvl_{i}", placeholder="Fließend / C1", label_visibility="collapsed" if i>0 else "visible")
            with c3:
                st.markdown("<div style='margin-top:28px'>" if i==0 else "", unsafe_allow_html=True)
                if st.button("✕", key=f"del_lang_{i}"):
                    st.session_state.languages.pop(i)
                    st.rerun()

        if st.button("➕ Sprache hinzufügen", use_container_width=True):
            st.session_state.languages.append({"name":"","level":""})
            st.rerun()

# ═════════════════════════════════════════════════════════════════════════════
# TAB 5: Generieren
# ═════════════════════════════════════════════════════════════════════════════
with tab5:
    st.markdown("## ✦ Lebenslauf generieren")

    # Zusammenfassung anzeigen
    col_sum1, col_sum2, col_sum3 = st.columns(3)
    with col_sum1:
        st.metric("Stellen", len([e for e in st.session_state.experience if e.get("title")]))
    with col_sum2:
        total_skills = sum(len([s for s in cat["items"] if s.get("name")]) for cat in st.session_state.skills)
        st.metric("Skills", total_skills)
    with col_sum3:
        st.metric("Palette", PALETTES[st.session_state.palette]["label"])

    st.markdown("---")

    # Palette-Auswahl für den Generier-Button
    gen_col1, gen_col2 = st.columns([2,1])
    with gen_col1:
        gen_mode = st.radio(
            "Was generieren?",
            ["Gewählte Palette", "Alle 5 Farbvarianten"],
            horizontal=True,
        )
    with gen_col2:
        st.markdown(f"<div style='padding:0.5rem;background:{PALETTES[st.session_state.palette]['hex']};border-radius:6px;color:white;text-align:center;font-weight:600'>{PALETTES[st.session_state.palette]['label']}</div>", unsafe_allow_html=True)

    st.markdown("")
    generate_clicked = st.button(
        "✦ Lebenslauf jetzt erstellen",
        type="primary",
        use_container_width=True,
    )

    if generate_clicked:
        if not name:
            st.error("Bitte mindestens einen Namen eingeben (Tab 1).")
        else:
            with st.spinner("Lebenslauf wird erstellt..."):

                # ── Foto vorbereiten ──────────────────────────────────────────
                photo_path = None
                if photo_file:
                    img = Image.open(photo_file)
                    if img.mode in ('RGBA','P'): img = img.convert('RGB')
                    w, h = img.size
                    size = min(w, h)
                    left = (w - size) // 2
                    top  = max(0, int((h - size) * 0.35))
                    top  = min(top, h - size)
                    img_crop = img.crop((left, top, left+size, top+size))
                    img_crop = img_crop.resize((300, 300), Image.LANCZOS)
                    tmp_photo = tempfile.NamedTemporaryFile(suffix=".jpg", delete=False)
                    img_crop.save(tmp_photo.name, "JPEG", quality=92)
                    photo_path = tmp_photo.name

                # ── KPIs parsen ───────────────────────────────────────────────
                def parse_kpis(raw):
                    if not raw: return []
                    result = []
                    for part in raw.split(","):
                        part = part.strip()
                        if not part: continue
                        tokens = part.split(" ", 1)
                        result.append({"value": tokens[0], "label": tokens[1] if len(tokens)>1 else ""})
                    return result

                # ── CV-Daten zusammenstellen ──────────────────────────────────
                cv_data = {
                    "name": name, "title": title, "email": email,
                    "phone": phone, "location": location, "linkedin": linkedin,
                    "summary": summary,
                    "photo": photo_path,
                    "experience": [
                        {**e, "kpis": parse_kpis(e.get("kpis",""))}
                        for e in st.session_state.experience if e.get("title")
                    ],
                    "skills":    [c for c in st.session_state.skills    if any(s.get("name") for s in c["items"])],
                    "tools":     [t.strip() for t in tools_raw.split(",") if t.strip()] if tools_raw else [],
                    "languages": [l for l in st.session_state.languages if l.get("name")],
                    "education": [e for e in st.session_state.education if e.get("degree")],
                }

                # ── Paletten bestimmen ────────────────────────────────────────
                if gen_mode == "Alle 5 Farbvarianten":
                    palettes_to_gen = list(PALETTES.keys())
                else:
                    palettes_to_gen = [st.session_state.palette]

                # Custom-Farben einmischen wenn nötig
                custom_overrides = {}
                if st.session_state.custom_mode:
                    custom_overrides = {
                        "sidebarBg":     c_sidebar.replace("#",""),
                        "sidebarAccent": c_accent.replace("#",""),
                        "mainAccent":    c_main.replace("#",""),
                        "mainHeading":   c_heading.replace("#",""),
                        "tagBg":         c_tag.replace("#",""),
                    }

                # ── Für jede Palette generieren ────────────────────────────────
                generated = []
                js_path   = Path(__file__).parent / "generate_cv.js"
                progress  = st.progress(0)

                for idx, pal_key in enumerate(palettes_to_gen):
                    progress.progress((idx) / len(palettes_to_gen), text=f"Erstelle {PALETTES[pal_key]['label']}...")

                    # JSON-Datei
                    tmp_json = tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False, encoding="utf-8")
                    json.dump(cv_data, tmp_json, ensure_ascii=False)
                    tmp_json.close()

                    # Palette-Argument
                    if custom_overrides:
                        merged = {**PALETTE_COLORS[pal_key], **custom_overrides}
                        tmp_pal = tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False)
                        json.dump(merged, tmp_pal)
                        tmp_pal.close()
                        pal_arg = tmp_pal.name
                    else:
                        pal_arg = pal_key
                        tmp_pal = None

                    # Output-Pfad
                    out_file = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
                    out_file.close()

                    # Node aufrufen
                    result = subprocess.run(
                        ["node", str(js_path), tmp_json.name, out_file.name, pal_arg],
                        capture_output=True, text=True, timeout=60
                    )

                    # Cleanup
                    Path(tmp_json.name).unlink(missing_ok=True)
                    if tmp_pal: Path(tmp_pal.name).unlink(missing_ok=True)

                    if result.returncode == 0:
                        with open(out_file.name, "rb") as f:
                            docx_bytes = f.read()
                        generated.append({
                            "palette": pal_key,
                            "label":   PALETTES[pal_key]["label"],
                            "bytes":   docx_bytes,
                            "filename": f"lebenslauf_{name.split()[0].lower() if name else 'cv'}_{pal_key}.docx",
                        })
                    else:
                        st.warning(f"Fehler bei Palette {pal_key}: {result.stderr[:300]}")

                    Path(out_file.name).unlink(missing_ok=True)

                if photo_path:
                    Path(photo_path).unlink(missing_ok=True)

                progress.progress(1.0, text="✓ Fertig!")

            # ── Download-Buttons ──────────────────────────────────────────────
            if generated:
                st.success(f"✓ {len(generated)} Lebenslauf{'/' if len(generated)>1 else ''} erstellt!")

                if len(generated) == 1:
                    g = generated[0]
                    st.download_button(
                        label=f"⬇ {g['label']} herunterladen",
                        data=g["bytes"],
                        file_name=g["filename"],
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary",
                        use_container_width=True,
                    )
                else:
                    st.markdown("**Alle Varianten herunterladen:**")
                    cols = st.columns(len(generated))
                    for col, g in zip(cols, generated):
                        with col:
                            st.download_button(
                                label=f"⬇ {g['label']}",
                                data=g["bytes"],
                                file_name=g["filename"],
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True,
                            )

                with st.expander("📋 Layout-Log anzeigen"):
                    st.caption("Content-Score und Layout-Entscheidungen des letzten Durchlaufs")
                    st.code(result.stdout)
