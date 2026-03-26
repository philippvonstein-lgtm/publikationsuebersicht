#!/usr/bin/env python3
"""
Streamlit Web-App — Publikationsübersicht
Klinik III für Innere Medizin, Uniklinik Köln
"""

import io
import sys
import time
from datetime import date, datetime

import streamlit as st

# Fix Encoding auf Windows
if sys.stdout.encoding != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

# Importiere alle Kernfunktionen aus dem bestehenden Skript
from publikationsuebersicht import (
    scrape_staff,
    search_pubmed,
    search_scholar,
    deduplicate_articles,
    create_word_document,
    normalize_name,
    _split_name,
    _umlaut_variants,
    PUBMED_DELAY,
    Entrez,
)

# ---------------------------------------------------------------------------
# Seitenconfig
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Publikationsübersicht — Kardiologie UK Köln",
    page_icon="🫀",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS
st.markdown("""
<style>
    .main-title { text-align: center; color: #1a5276; }
    .stProgress .st-bo { background-color: #1a5276; }
    div[data-testid="stMetricValue"] { font-size: 2rem; }
    .pub-title { font-weight: bold; font-size: 14px; }
    .pub-authors { font-size: 12px; color: #333; }
    .pub-meta { font-size: 11px; color: #666; font-style: italic; }
    .clinic-author { font-weight: bold; color: #1a5276; }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Zusätzliche Mitarbeiter (nicht auf der Homepage)
# ---------------------------------------------------------------------------
EXTRA_STAFF_NAMES = [
    "Per Arkenberg",
    "Harshal Nemade",
    "Kristel Martinez Lagunas",
    "Yein Park",
    "Elvina Santhamma Philip",
    "Suchitra Narayan",
    "Chantal Wientjens",
    "Holger Winkels",
    "Martin Mollenhauer",
    "Valerie Lohner",
    "Thomas Riffelmacher",
    "Arian Sultan",
    "Simon Braumann",
    "Christoph Adler",
]

# Session state für eigene Autoren
if "custom_authors" not in st.session_state:
    st.session_state.custom_authors = []

# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.title("⚙️ Einstellungen")
    st.caption("🫀 Kardiologie UK Köln")

    st.subheader("📅 Zeitraum")
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input(
            "Von",
            value=date(2025, 1, 1),
            min_value=date(2000, 1, 1),
            max_value=date.today(),
            format="DD.MM.YYYY",
        )
    with col2:
        end_date = st.date_input(
            "Bis",
            value=date.today(),
            min_value=date(2000, 1, 1),
            max_value=date.today(),
            format="DD.MM.YYYY",
        )

    st.subheader("🔍 Suchoptionen")
    use_scholar = st.checkbox("Google Scholar einbeziehen", value=False,
                               help="Deutlich langsamer wegen Rate-Limiting")

    # --- Autoren-Auswahl ---
    st.subheader("👥 Autoren")

    use_med3 = st.checkbox("Med III Homepage (alle)", value=True,
                            help="Alle Mitarbeiter von der Klinik-Website")

    all_extra = EXTRA_STAFF_NAMES + st.session_state.custom_authors
    selected_extra = st.multiselect(
        "Zusätzliche Mitarbeiter",
        options=all_extra,
        default=all_extra,
        help="Wähle welche zusätzlichen Autoren durchsucht werden sollen",
    )

    # Eigene Autoren hinzufügen
    with st.expander("➕ Neuen Autor hinzufügen"):
        new_author = st.text_input("Name (Vorname Nachname)",
                                    placeholder="z.B. Maria Müller",
                                    key="new_author_input")
        if st.button("Hinzufügen", key="add_author_btn"):
            if new_author.strip() and new_author.strip() not in st.session_state.custom_authors:
                st.session_state.custom_authors.append(new_author.strip())
                st.rerun()

    st.divider()

    st.subheader("🎯 Autorenfilter")
    position_filter_sidebar = st.selectbox("Autorenposition", [
        "Alle Publikationen",
        "Nur Erstautorschaften",
        "Nur Letztautorschaften",
        "Nur Erst- & Letztautorschaften",
        "Nur Ko-Autorenschaften",
    ], help="Filter gilt für Anzeige UND Word-Export")

    st.subheader("📧 PubMed E-Mail")
    email = st.text_input("E-Mail für PubMed API",
                          value="publikationen@uk-koeln.de",
                          help="PubMed verlangt eine E-Mail-Adresse")

    st.divider()

    start_search = st.button("🔍 Suche starten", type="primary", use_container_width=True)

    st.divider()
    st.caption("Klinik III für Innere Medizin\nHerzzentrum der Uniklinik Köln")


# ---------------------------------------------------------------------------
# Hauptbereich
# ---------------------------------------------------------------------------
st.markdown("<h1 class='main-title'>🫀 Publikationsübersicht</h1>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align:center; color:#666;'>Klinik III für Innere Medizin — Kardiologie</h3>",
            unsafe_allow_html=True)
st.divider()


def build_clinic_author_checker(staff):
    """Baut die _is_clinic_author Hilfsfunktion mit staff-Kontext."""
    from difflib import SequenceMatcher

    staff_by_last = {}
    for s in staff:
        parts = s["clean_name"].split()
        if len(parts) >= 2:
            _, _, s_last = _split_name(s["clean_name"])
            for variant in _umlaut_variants(s_last.lower()):
                if variant not in staff_by_last:
                    staff_by_last[variant] = []
                staff_by_last[variant].append(s["clean_name"])

    def is_clinic_author(fau_name: str) -> bool:
        if "," in fau_name:
            parts = fau_name.split(",", 1)
            fau_last = parts[0].strip().lower()
            fau_first = parts[1].strip().split()[0].lower() if parts[1].strip() else ""
        else:
            parts = fau_name.strip().split()
            fau_first = parts[0].lower() if parts else ""
            fau_last = parts[-1].lower() if len(parts) >= 2 else ""

        fau_last_variants = list(_umlaut_variants(fau_last))
        fau_last_parts = fau_last.split()
        if len(fau_last_parts) > 1:
            fau_last_variants.extend(_umlaut_variants(fau_last_parts[-1]))

        for variant in fau_last_variants:
            if variant in staff_by_last:
                for clean_name in staff_by_last[variant]:
                    cn_first = clean_name.split()[0].lower()
                    cn_first_variants = [v.lower() for v in _umlaut_variants(cn_first)]
                    for cfv in cn_first_variants:
                        if len(cfv) >= 4 and len(fau_first) >= 4:
                            ratio = SequenceMatcher(None, cfv, fau_first).ratio()
                            if ratio > 0.7:
                                return True
                        elif len(fau_first) <= 2 and fau_first:
                            if cfv and cfv[0] == fau_first[0]:
                                return True
                        elif cfv == fau_first:
                            return True
        return False

    return is_clinic_author


def format_authors_html(authors_list, is_clinic_fn):
    """Formatiert Autorenliste als HTML mit fetten Klinik-Autoren."""
    parts = []
    for author in authors_list:
        if is_clinic_fn(author):
            parts.append(f"<b style='color:#1a5276;'>{author}</b>")
        else:
            parts.append(author)
    return "; ".join(parts)


# ---------------------------------------------------------------------------
# Session State initialisieren
# ---------------------------------------------------------------------------
if "articles" not in st.session_state:
    st.session_state.articles = None
if "staff" not in st.session_state:
    st.session_state.staff = None
if "search_done" not in st.session_state:
    st.session_state.search_done = False


# ---------------------------------------------------------------------------
# Suche durchführen
# ---------------------------------------------------------------------------
if start_search:
    Entrez.email = email

    start_date_str = start_date.strftime("%Y/%m/%d")
    end_date_str = end_date.strftime("%Y/%m/%d")

    # Schritt 1: Mitarbeiter zusammenstellen
    with st.status("🏥 Mitarbeiterliste wird erstellt...", expanded=True) as status:
        staff = []

        # a) Med III Homepage (falls aktiviert)
        if use_med3:
            from publikationsuebersicht import scrape_staff as _scrape_raw
            # Scrape, aber OHNE die EXTRA_STAFF aus dem Skript
            # (die steuern wir jetzt über die Sidebar)
            import publikationsuebersicht as _mod
            orig_extra = getattr(_mod, '_ORIG_EXTRA', None)
            raw_staff = _scrape_raw()
            # Entferne die im Skript hartcodierten Extras
            # (wir fügen sie separat über die Sidebar hinzu)
            extra_clean_names = {n.lower() for n in EXTRA_STAFF_NAMES}
            staff = [s for s in raw_staff
                     if s["clean_name"].lower() not in extra_clean_names]
            st.write(f"✅ **{len(staff)}** Mitarbeiter von der Homepage")

        # b) Ausgewählte Zusatz-Mitarbeiter
        for name in selected_extra:
            staff.append({
                "full_name": name, "clean_name": name,
                "title": "", "position": "", "category": "Zusätzlich",
            })
        if selected_extra:
            st.write(f"✅ **{len(selected_extra)}** zusätzliche Mitarbeiter")

        # c) Eigene Autoren
        for name in st.session_state.custom_authors:
            staff.append({
                "full_name": name, "clean_name": name,
                "title": "", "position": "", "category": "Eigener Autor",
            })
        if st.session_state.custom_authors:
            st.write(f"✅ **{len(st.session_state.custom_authors)}** eigene Autoren")

        # Duplikate entfernen
        seen = set()
        unique_staff = []
        for s in staff:
            key = s["clean_name"].lower()
            if key not in seen:
                seen.add(key)
                unique_staff.append(s)
        staff = unique_staff

        status.update(label=f"✅ {len(staff)} Mitarbeiter gesamt", state="complete")

    st.session_state.staff = staff

    # Schritt 2: PubMed
    all_articles = []
    progress_bar = st.progress(0, text="PubMed-Suche...")
    status_text = st.empty()

    for i, person in enumerate(staff):
        progress = (i + 1) / len(staff)
        progress_bar.progress(progress,
                              text=f"🔍 PubMed: {person['clean_name']} ({i+1}/{len(staff)})")
        articles = search_pubmed(person, start_date_str, end_date_str)
        all_articles.extend(articles)
        status_text.text(f"  → {len(articles)} Treffer für {person['clean_name']}")

    progress_bar.progress(1.0, text=f"✅ PubMed: {len(all_articles)} Artikel gefunden")

    # Schritt 3: Google Scholar (optional)
    if use_scholar:
        scholar_progress = st.progress(0, text="Google Scholar-Suche...")
        start_year = start_date.year
        end_year = end_date.year
        scholar_count = 0

        for i, person in enumerate(staff):
            progress = (i + 1) / len(staff)
            scholar_progress.progress(
                progress,
                text=f"📚 Scholar: {person['clean_name']} ({i+1}/{len(staff)})"
            )
            articles = search_scholar(person, start_year, end_year)
            all_articles.extend(articles)
            scholar_count += len(articles)

        scholar_progress.progress(1.0,
                                  text=f"✅ Scholar: {scholar_count} Artikel gefunden")

    # Schritt 4: Deduplizierung
    with st.spinner("🔄 Duplikate bereinigen..."):
        unique_articles = deduplicate_articles(all_articles)

    st.session_state.articles = unique_articles
    st.session_state.search_done = True
    st.session_state.start_date = start_date
    st.session_state.end_date = end_date

    st.success(f"🎉 **{len(unique_articles)} eindeutige Publikationen** gefunden "
               f"(aus {len(all_articles)} Rohergebnissen)")
    st.balloons()


# ---------------------------------------------------------------------------
# Ergebnisse anzeigen
# ---------------------------------------------------------------------------
def apply_position_filter(articles, pos_filter):
    """Wendet den Sidebar-Autorenfilter global an."""
    if pos_filter == "Alle Publikationen":
        return articles
    elif pos_filter == "Nur Erstautorschaften":
        return [a for a in articles if a.get("author_position") == "Erstautor"]
    elif pos_filter == "Nur Letztautorschaften":
        return [a for a in articles if a.get("author_position") == "Letztautor"]
    elif pos_filter == "Nur Erst- & Letztautorschaften":
        return [a for a in articles if a.get("author_position") in ("Erstautor", "Letztautor")]
    elif pos_filter == "Nur Ko-Autorenschaften":
        return [a for a in articles
                if a.get("author_position", "").startswith("Ko-Autor")
                or a.get("author_position") in ("Vorletzter Autor", "Investigator")]
    return articles


if st.session_state.search_done and st.session_state.articles is not None:
    articles_all = st.session_state.articles
    articles = apply_position_filter(articles_all, position_filter_sidebar)
    staff = st.session_state.staff
    is_clinic = build_clinic_author_checker(staff)

    st.divider()

    # Metriken
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📄 Publikationen", len(articles))
    with col2:
        st.metric("👥 Mitarbeiter", len(staff))
    with col3:
        # Zähle Erstautorschaften
        first_author = sum(1 for a in articles if a.get("author_position") == "Erstautor")
        st.metric("✍️ Erstautorschaften", first_author)
    with col4:
        last_author = sum(1 for a in articles if a.get("author_position") == "Letztautor")
        st.metric("🏆 Letztautorschaften", last_author)

    st.divider()

    # Textfilter und Mitarbeiter-Filter (vor den Tabs, damit überall verfügbar)
    col_filter1, col_filter2 = st.columns([3, 1])
    with col_filter1:
        search_term = st.text_input("🔍 Suche in Titeln/Autoren",
                                    placeholder="z.B. Baldus, tricuspid, mitral...")
    with col_filter2:
        mitarbeiter_filter = st.selectbox(
            "Mitarbeiter",
            ["Alle"] + sorted(set(a.get("assigned_to", "") for a in articles if a.get("assigned_to"))),
        )

    # Artikel filtern
    filtered = articles
    if search_term:
        term = search_term.lower()
        filtered = [a for a in filtered
                    if term in a.get("title", "").lower()
                    or any(term in str(au).lower() for au in a.get("full_authors", a.get("authors", [])))]
    if mitarbeiter_filter != "Alle":
        filtered = [a for a in filtered if a.get("assigned_to") == mitarbeiter_filter]

    st.write(f"**{len(filtered)} Publikationen** angezeigt")

    # Tabs für verschiedene Ansichten
    tab1, tab2, tab3 = st.tabs(["📋 Publikationsliste", "📊 Statistiken", "📥 Export"])

    with tab1:

        # Artikel anzeigen
        for i, art in enumerate(filtered, 1):
            with st.container():
                # Titel
                pmid = art.get("pmid", "")
                title = art.get("title", "Kein Titel")
                if pmid:
                    st.markdown(f"**{i}. [{title}](https://pubmed.ncbi.nlm.nih.gov/{pmid}/)**")
                else:
                    st.markdown(f"**{i}. {title}**")

                # Autoren
                display_authors = art.get("full_authors", art.get("authors", []))
                if isinstance(display_authors, list):
                    authors_html = format_authors_html(display_authors, is_clinic)
                    st.markdown(authors_html, unsafe_allow_html=True)

                # Meta
                journal = art.get("journal", "")
                date_raw = art.get("date_raw", "")
                doi = art.get("doi", "")
                meta_parts = []
                if journal:
                    meta_parts.append(f"*{journal}*")
                if date_raw:
                    meta_parts.append(f"({date_raw})")
                if doi:
                    meta_parts.append(f"DOI: [{doi}](https://doi.org/{doi})")
                if pmid:
                    meta_parts.append(f"PMID: {pmid}")

                st.caption(" | ".join(meta_parts))
                st.divider()

    with tab2:
        import pandas as pd

        st.subheader("📊 Publikationen pro Mitarbeiter")

        # Zähle Publikationen pro Mitarbeiter
        pub_counts = {}
        for art in articles:
            name = art["assigned_to"]
            if name not in pub_counts:
                pub_counts[name] = {"Gesamt": 0, "Erstautor": 0, "Letztautor": 0, "Ko-Autor": 0}
            pub_counts[name]["Gesamt"] += 1
            pos = art.get("author_position", "")
            if pos == "Erstautor":
                pub_counts[name]["Erstautor"] += 1
            elif pos == "Letztautor":
                pub_counts[name]["Letztautor"] += 1
            else:
                pub_counts[name]["Ko-Autor"] += 1

            # Zähle auch andere Klinik-Autoren
            for other in art.get("other_clinic_authors", []):
                oname = normalize_name(other["name"])
                if oname not in pub_counts:
                    pub_counts[oname] = {"Gesamt": 0, "Erstautor": 0, "Letztautor": 0, "Ko-Autor": 0}
                pub_counts[oname]["Gesamt"] += 1
                pub_counts[oname]["Ko-Autor"] += 1

        if pub_counts:
            df = pd.DataFrame.from_dict(pub_counts, orient="index")
            df = df.sort_values("Gesamt", ascending=False)
            df.index.name = "Mitarbeiter"

            # Balkendiagramm
            st.bar_chart(df["Gesamt"], height=400, color="#1a5276")

            # Tabelle
            st.dataframe(df, use_container_width=True, height=600)

        # Journal-Verteilung
        st.subheader("📰 Top Journals")
        journal_counts = {}
        for art in articles:
            j = art.get("journal", "Unbekannt")
            journal_counts[j] = journal_counts.get(j, 0) + 1

        if journal_counts:
            df_j = pd.DataFrame(
                sorted(journal_counts.items(), key=lambda x: x[1], reverse=True)[:20],
                columns=["Journal", "Anzahl"]
            )
            st.bar_chart(df_j.set_index("Journal"), height=400, color="#2e86c1")

    with tab3:
        st.subheader("📥 Als Word-Dokument exportieren")
        filter_info = f"(**{position_filter_sidebar}**)" if position_filter_sidebar != "Alle Publikationen" else ""
        st.write(f"Exportiert **{len(filtered)} Publikationen** {filter_info} als formatiertes Word-Dokument.")

        if st.button("📄 Word-Dokument generieren", type="primary"):
            with st.spinner("Word-Dokument wird erstellt..."):
                # In-Memory erstellen
                output_buffer = io.BytesIO()
                output_path = f"Publikationsuebersicht_{datetime.now().strftime('%Y-%m-%d')}.docx"

                s_date = st.session_state.start_date.strftime("%Y-%m-%d")
                e_date = st.session_state.end_date.strftime("%Y-%m-%d")

                create_word_document(filtered, staff, s_date, e_date, output_path)

                # Datei zum Download bereitstellen
                with open(output_path, "rb") as f:
                    docx_bytes = f.read()

            st.download_button(
                label="⬇️ Download Word-Dokument",
                data=docx_bytes,
                file_name=output_path,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
            )
            st.success(f"✅ Dokument bereit: {output_path}")

else:
    # Startseite wenn noch keine Suche durchgeführt
    st.info("👈 Wähle einen **Zeitraum** in der Sidebar und klicke **Suche starten**")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        ### 🔍 Automatische Suche
        Durchsucht **PubMed** und optional **Google Scholar**
        für alle Mitarbeiter der Klinik.
        """)
    with col2:
        st.markdown("""
        ### 📊 Statistiken
        Übersicht der Publikationen pro Mitarbeiter,
        Erstautorschaften und Top-Journals.
        """)
    with col3:
        st.markdown("""
        ### 📥 Word-Export
        Exportiert die Ergebnisse als formatiertes
        Word-Dokument mit fetten Klinik-Autoren.
        """)
