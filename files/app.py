# -*- coding: utf-8 -*-
"""
Greenwashing-Check nach EmpCo-Richtlinie / Bioland-Checkliste
==============================================================

Dieses Programm prüft eingefügte Texte, hochgeladene Dateien und Webseiten
auf mögliche kritische Umweltaussagen (Greenwashing) im Sinne der
EU-Richtlinie EmpCo (2024/825) und der Bioland-Checkliste vom 14.07.2026.

WICHTIG: Dies ist KEINE Rechtsberatung. Das Tool liefert nur automatisierte
Anhaltspunkte auf Basis von Stichwort- und Kontextsuche. Jeder markierte
Treffer muss von einem Menschen geprüft werden.
"""

import io
import json
import re
from datetime import datetime

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup
from docx import Document
from pypdf import PdfReader

# ---------------------------------------------------------------------------
# 1) BEGRIFFSLISTEN
# ---------------------------------------------------------------------------
# Kategorie 1: Allgemeine Umweltaussagen -> nur zulässig mit Begründung ODER
# Bio-/Bioland-Bezug (siehe Bioland-Checkliste, Flowchart Seite 2).
DEFAULT_KAT1 = [
    {"begriff": "umweltfreundlich", "pattern": r"umweltfreundlich\w*"},
    {"begriff": "naturnah", "pattern": r"naturnah\w*"},
    {"begriff": "klimaschonend", "pattern": r"klimaschonend\w*"},
    {"begriff": "klimafreundlich", "pattern": r"klimafreundlich\w*"},
    {"begriff": "bienenfreundlich", "pattern": r"bienenfreundlich\w*"},
    {"begriff": "grundwasserschutz / -schonend", "pattern": r"grundwasserschützend\w*|grundwasserschonend\w*|grundwasserschutz\w*"},
    {"begriff": "bodenschonend", "pattern": r"bodenschonend\w*"},
    {"begriff": "ökologisch / öko-", "pattern": r"ökologisch\w*|öko-\w*|\böko\b"},
    {"begriff": "grün (als Umweltclaim)", "pattern": r"\bgrün\w*\b"},
    {"begriff": "nachhaltig", "pattern": r"nachhaltig\w*"},
    {"begriff": "naturbelassen", "pattern": r"naturbelassen\w*"},
    {"begriff": "biologisch abbaubar / kompostierbar", "pattern": r"biologisch abbaubar|bio-abbaubar\w*|kompostierbar\w*"},
]

# Kategorie 2: Klimaneutralitäts-/Kompensations-Claims -> laut EmpCo
# grundsätzlich kritisch, unabhängig vom Kontext.
DEFAULT_KAT2 = [
    {"begriff": "klimaneutral", "pattern": r"klimaneutral\w*"},
    {"begriff": "CO2-/CO₂-neutral", "pattern": r"co2-neutral\w*|co₂-neutral\w*|co2\s*neutral\w*|kohlendioxidneutral\w*"},
    {"begriff": "klimapositiv", "pattern": r"klimapositiv\w*"},
    {"begriff": "treibhausgasneutral / THG-neutral", "pattern": r"treibhausgasneutral\w*|thg-neutral\w*"},
]

# Signalwörter für den Kontext-Check (Kategorie 1 und 3)
ERKLAERUNG_PATTERN = r"\bweil\b|\bdenn\b|,\s*da\b|\baufgrund\b|\bdank\b"
BIO_PATTERN = r"\bbio(?!land)\w*\b|öko-verordnung|eu-bio|kontrolliert\s+biologisch"
BIOLAND_PATTERN = r"\bbioland\b"

# ---------------------------------------------------------------------------
# 2) HILFSFUNKTIONEN: TEXT-ANALYSE
# ---------------------------------------------------------------------------

def get_all_terms():
    """Führt Standardbegriffe (Kat. 1 + 2) und eigene Begriffe zusammen."""
    terms = []
    for t in DEFAULT_KAT1:
        terms.append({**t, "kategorie": 1})
    for t in DEFAULT_KAT2:
        terms.append({**t, "kategorie": 2})
    terms.extend(st.session_state.get("custom_terms", []))
    return terms


def get_context_window(text, start, end, max_chars=200):
    """Extrahiert den Satz/Textabschnitt rund um einen Treffer (für die
    Anzeige UND für den Kontext-Check nach Begründung/Bio-Bezug)."""
    left_bound = max(0, start - max_chars)
    right_bound = min(len(text), end + max_chars)
    left_search = text[left_bound:start]
    right_search = text[end:right_bound]

    left_matches = list(re.finditer(r"[.!?\n]", left_search))
    sent_start = left_bound + (left_matches[-1].end() if left_matches else 0)

    right_match = re.search(r"[.!?\n]", right_search)
    sent_end = end + (right_match.end() if right_match else len(right_search))

    return text[sent_start:sent_end].strip()


def evaluate_kat1(window):
    """Ampel-Logik für Kategorie 1, angelehnt an das Bioland-Flowchart."""
    if re.search(ERKLAERUNG_PATTERN, window, re.IGNORECASE):
        return "🟢 Grün", (
            "Mögliche Begründung in der Nähe entdeckt (z. B. „weil / denn / "
            "aufgrund …“). Bitte prüfen, ob die Begründung konkret und "
            "nachprüfbar ist."
        )
    if re.search(BIO_PATTERN, window, re.IGNORECASE):
        return "🟢 Grün", (
            "Bio-Bezug in der Nähe gefunden. Die Aussage stützt sich "
            "vermutlich auf eine anerkannte Umweltleistung (Bio-Standard)."
        )
    if re.search(BIOLAND_PATTERN, window, re.IGNORECASE):
        return "🟡 Gelb", (
            "Nur Bioland-Bezug ohne zusätzlichen Bio-Verweis gefunden. Laut "
            "Checkliste rechtlich noch nicht abschließend gesichert – "
            "zusätzlichen Verweis auf den Bio-Standard empfehlen."
        )
    return "🔴 Rot", (
        "Keine Begründung und kein Bio-/Bioland-Bezug in der Nähe gefunden. "
        "Aussage ist wahrscheinlich unzulässig – Formulierung prüfen/anpassen."
    )


def evaluate_kat2():
    """Kategorie 2 gilt laut EmpCo grundsätzlich als kritisch."""
    return "🔴 Rot", (
        "Klimaneutralitäts-/Kompensations-Claim. Laut EmpCo grundsätzlich "
        "kritisch, wenn er (auch teilweise) auf Kompensation statt echter "
        "Emissionsreduktion beruht – unabhängig vom Kontext."
    )


def check_text(text, quelle):
    """Prüft einen Text auf alle Kategorien und gibt eine Liste von
    Fundstellen (Dictionaries) zurück."""
    findings = []
    if not text or not text.strip():
        return findings

    for term in get_all_terms():
        try:
            pattern = re.compile(term["pattern"], re.IGNORECASE | re.UNICODE)
        except re.error:
            continue  # ungültiges Muster überspringen, statt die App abstürzen zu lassen
        for match in pattern.finditer(text):
            window = get_context_window(text, match.start(), match.end())
            if term["kategorie"] == 2:
                ampel, hinweis = evaluate_kat2()
            else:
                ampel, hinweis = evaluate_kat1(window)
            findings.append({
                "Quelle": quelle,
                "Kategorie": term["kategorie"],
                "Gefundener Begriff": match.group(),
                "Textstelle": window,
                "Ampel": ampel,
                "Hinweis": hinweis,
            })

    # Kategorie 3: eigenständiger Bioland-Check (auch ohne Kat.-1-Treffer)
    for match in re.finditer(BIOLAND_PATTERN, text, re.IGNORECASE):
        window = get_context_window(text, match.start(), match.end())
        if not re.search(BIO_PATTERN, window, re.IGNORECASE):
            findings.append({
                "Quelle": quelle,
                "Kategorie": 3,
                "Gefundener Begriff": match.group(),
                "Textstelle": window,
                "Ampel": "🟡 Gelb",
                "Hinweis": (
                    "Bioland-Bezug ohne zusätzlichen Bio-Verweis in der Nähe. "
                    "Laut Checkliste rechtlich noch nicht abschließend "
                    "gesichert – zusätzlichen Bio-Standard-Verweis empfehlen."
                ),
            })

    return findings


# ---------------------------------------------------------------------------
# 3) HILFSFUNKTIONEN: DATEIEN EINLESEN
# ---------------------------------------------------------------------------

def extract_text_from_txt(file):
    raw = file.read()
    for enc in ("utf-8", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="ignore")


def extract_text_from_docx(file):
    doc = Document(file)
    return "\n".join(p.text for p in doc.paragraphs)


def extract_text_from_csv(file):
    df = pd.read_csv(file)
    text_cols = df.select_dtypes(include="object").columns
    if len(text_cols) == 0:
        return ""
    return "\n".join(
        df[text_cols].astype(str).apply(lambda row: " ".join(row), axis=1)
    )


def extract_text_from_pdf(file):
    reader = PdfReader(file)
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def extract_text_from_file(uploaded_file):
    name = uploaded_file.name.lower()
    uploaded_file.seek(0)
    if name.endswith(".txt"):
        return extract_text_from_txt(uploaded_file)
    elif name.endswith(".docx"):
        return extract_text_from_docx(uploaded_file)
    elif name.endswith(".csv"):
        return extract_text_from_csv(uploaded_file)
    elif name.endswith(".pdf"):
        return extract_text_from_pdf(uploaded_file)
    return ""


def fetch_url_text(url):
    if not url.startswith("http"):
        url = "https://" + url
    try:
        headers = {"User-Agent": "Mozilla/5.0 (Greenwashing-Check-Tool)"}
        resp = requests.get(url, headers=headers, timeout=10)
        resp.raise_for_status()
    except Exception as e:
        return None, str(e)

    soup = BeautifulSoup(resp.text, "html.parser")
    for tag in soup(["script", "style", "noscript", "header", "footer", "nav"]):
        tag.decompose()
    text = soup.get_text(separator=" ")
    text = re.sub(r"\s+", " ", text).strip()
    return text, None


# ---------------------------------------------------------------------------
# 4) STREAMLIT-OBERFLÄCHE
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Greenwashing-Check (EmpCo)", page_icon="🌱", layout="wide")

if "custom_terms" not in st.session_state:
    st.session_state.custom_terms = []

st.title("🌱 Greenwashing-Check nach EmpCo-Richtlinie")
st.caption(
    "Prüft Texte auf mögliche kritische Umweltaussagen gemäß der EmpCo-Richtlinie "
    "(EU 2024/825, gültig ab 27.09.2026) und der Bioland-Checkliste vom 14.07.2026."
)
st.warning(
    "⚠️ **Keine Rechtsberatung:** Dieses Tool liefert nur automatisierte Anhaltspunkte "
    "auf Basis von Stichwort- und Kontextsuche. Es ersetzt keine rechtliche Prüfung. "
    "Jeder markierte Treffer sollte von einem Menschen geprüft werden – im Zweifel "
    "durch Rechtsberatung.",
    icon="⚠️",
)

# --- Sidebar: eigene Begriffe verwalten ------------------------------------
with st.sidebar:
    st.header("⚙️ Eigene Begriffe verwalten")

    with st.form("add_term_form", clear_on_submit=True):
        neuer_begriff = st.text_input("Neuer Begriff / Formulierung")
        kategorie_wahl = st.selectbox(
            "Kategorie",
            [1, 2],
            format_func=lambda k: (
                "Kategorie 1 – allgemeine Umweltaussage (braucht Begründung/Bio-Bezug)"
                if k == 1
                else "Kategorie 2 – Klimaneutralität/Kompensation (immer kritisch)"
            ),
        )
        submitted = st.form_submit_button("➕ Hinzufügen")
        if submitted and neuer_begriff.strip():
            pattern = re.escape(neuer_begriff.strip().lower()) + r"\w*"
            st.session_state.custom_terms.append({
                "begriff": neuer_begriff.strip(),
                "pattern": pattern,
                "kategorie": kategorie_wahl,
            })
            st.success(f"„{neuer_begriff}“ hinzugefügt.")

    if st.session_state.custom_terms:
        st.write("**Aktuelle eigene Begriffe:**")
        for i, term in enumerate(st.session_state.custom_terms):
            c1, c2 = st.columns([4, 1])
            c1.write(f"{term['begriff']} (Kat. {term['kategorie']})")
            if c2.button("🗑️", key=f"del_{i}"):
                st.session_state.custom_terms.pop(i)
                st.rerun()

    st.divider()
    st.write("**Begriffsliste speichern/laden**")
    st.caption(
        "Eigene Begriffe gehen beim Schließen der App verloren, außer du "
        "speicherst sie hier als Datei und lädst sie beim nächsten Mal wieder hoch."
    )
    if st.session_state.custom_terms:
        json_bytes = json.dumps(
            st.session_state.custom_terms, ensure_ascii=False, indent=2
        ).encode("utf-8")
        st.download_button(
            "💾 Eigene Begriffe exportieren (JSON)",
            data=json_bytes,
            file_name="eigene_begriffe.json",
            mime="application/json",
        )

    imported = st.file_uploader("📂 Eigene Begriffe laden (JSON)", type="json", key="import_terms")
    if imported is not None:
        file_id = getattr(imported, "file_id", imported.name)
        if st.session_state.get("last_imported_id") != file_id:
            try:
                data = json.loads(imported.read().decode("utf-8"))
                st.session_state.custom_terms.extend(data)
                st.session_state["last_imported_id"] = file_id
                st.success(f"{len(data)} Begriffe geladen.")
            except Exception as e:
                st.error(f"Datei konnte nicht gelesen werden: {e}")

    st.divider()
    with st.expander("📋 Standard-Begriffsliste anzeigen"):
        st.write("**Kategorie 1 – allgemeine Umweltaussagen:**")
        st.table(pd.DataFrame(DEFAULT_KAT1)[["begriff"]].rename(columns={"begriff": "Begriff"}))
        st.write("**Kategorie 2 – Klimaneutralität/Kompensation:**")
        st.table(pd.DataFrame(DEFAULT_KAT2)[["begriff"]].rename(columns={"begriff": "Begriff"}))

# --- Haupt-Eingabebereich ---------------------------------------------------
tab1, tab2, tab3 = st.tabs(["📝 Text einfügen", "📁 Dateien hochladen", "🌐 Webseite(n) prüfen"])

with tab1:
    pasted_text = st.text_area("Text hier einfügen", height=250, key="pasted_text")

with tab2:
    uploaded_files = st.file_uploader(
        "Dateien hochladen (.txt, .docx, .csv, .pdf) – mehrere Dateien möglich",
        type=["txt", "docx", "csv", "pdf"],
        accept_multiple_files=True,
    )
    st.caption(
        "Bei .csv-Dateien werden alle Textspalten automatisch zusammengefasst."
    )

with tab3:
    urls_raw = st.text_area(
        "Eine Web-Adresse pro Zeile",
        height=150,
        placeholder="https://www.beispiel.de/nachhaltigkeit\nhttps://www.beispiel.de/produkte",
    )

st.divider()

if st.button("🔍 Jetzt prüfen", type="primary"):
    sources = {}

    if pasted_text and pasted_text.strip():
        sources["Eingefügter Text"] = pasted_text

    for f in uploaded_files or []:
        try:
            sources[f.name] = extract_text_from_file(f)
        except Exception as e:
            st.warning(f"Datei „{f.name}“ konnte nicht gelesen werden: {e}")

    for url in [u.strip() for u in (urls_raw or "").splitlines() if u.strip()]:
        with st.spinner(f"Rufe {url} ab …"):
            text, error = fetch_url_text(url)
        if error:
            st.warning(f"„{url}“ konnte nicht abgerufen werden: {error}")
        else:
            sources[url] = text

    if not sources:
        st.warning("Bitte zuerst Text einfügen, Dateien hochladen oder eine Web-Adresse angeben.")
    else:
        all_findings = []
        with st.spinner("Texte werden geprüft …"):
            for quelle, text in sources.items():
                all_findings.extend(check_text(text, quelle))
        st.session_state["results"] = all_findings
        st.session_state["geprueft_am"] = datetime.now().strftime("%d.%m.%Y %H:%M")
        st.session_state["anzahl_quellen"] = len(sources)

# --- Ergebnisanzeige ---------------------------------------------------------
if "results" in st.session_state:
    findings = st.session_state["results"]
    st.subheader(
        f"📊 Ergebnisse – {st.session_state.get('anzahl_quellen', 0)} Quelle(n) geprüft "
        f"(Stand: {st.session_state.get('geprueft_am', '')})"
    )

    if not findings:
        st.success("Keine kritischen Formulierungen gefunden.")
    else:
        df = pd.DataFrame(findings)

        col1, col2, col3 = st.columns(3)
        col1.metric("🔴 Rot (wahrscheinlich unzulässig)", int((df["Ampel"] == "🔴 Rot").sum()))
        col2.metric("🟡 Gelb (zu prüfen)", int((df["Ampel"] == "🟡 Gelb").sum()))
        col3.metric("🟢 Grün (vermutlich unauffällig)", int((df["Ampel"] == "🟢 Grün").sum()))

        nur_kritisch = st.checkbox("Nur Rot/Gelb anzeigen (Grün ausblenden)", value=True)
        anzeige_df = df[df["Ampel"] != "🟢 Grün"] if nur_kritisch else df

        st.dataframe(anzeige_df, use_container_width=True, hide_index=True)

        csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "📥 Alle Ergebnisse als CSV herunterladen",
            data=csv_bytes,
            file_name="greenwashing_check_ergebnisse.csv",
            mime="text/csv",
        )
