import os
import pandas as pd
import requests
from bs4 import BeautifulSoup
import fitz  # PyMuPDF
import re
import streamlit as st
from SCRAPER_DEF import miao as scraper_main
import io
from pdfminer.high_level import extract_text_to_fp
from pdfminer.layout import LAParams
import spacy
import base64

# Carica il modello di lingua inglese per NLP
nlp = spacy.load("en_core_web_sm")

# Configurazione della pagina Streamlit (DEVE essere il primo comando Streamlit)
st.set_page_config(page_title="Analizzatore di Ingredienti", page_icon="🧪", layout="wide")

# CSS personalizzato per migliorare lo stile
st.markdown("""
    <style>
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    .stButton>button {
        width: 100%;
        border-radius: 20px;
        border: 2px solid #1E88E5;
        background-color: white;
        color: #1E88E5;
        transition: all 0.3s;
    }
    .stButton>button:hover {
        background-color: #1E88E5;
        color: white;
    }
    .reportview-container .main .block-container {
        max-width: 1200px;
    }
    h1 {
        color: #1E88E5;
        font-size: 3rem;
        text-align: center;
        margin-bottom: 2rem;
        position: relative;
        z-index: 1;
    }
    .background-image {
        position: absolute;
        top: 30%;
        left: 50%;
        transform: translate(-50%, -30%);
        opacity: 0.1;
        z-index: 0;
        width: 50%;
    }
    h2 {
        color: #43A047;
    }
    h3 {
        color: #FBC02D;
    }
    .sidebar .sidebar-content {
        background-color: #f0f0f0;
    }
    .symbol-button {
        margin: 2px;
        padding: 5px 10px;
        border-radius: 5px;
        background-color: #1E88E5;
        color: white;
        cursor: pointer;
    }
    .symbol-button:hover {
        background-color: #1565C0;
    }
    </style>
    """, unsafe_allow_html=True)


# Funzione per leggere, generare e salvare il PDF
def read_generate_and_save_pdf(selected_item, df, ingredient_column):
    pdf_dir = "pdf_momentari"
    pdf_path = os.path.join(pdf_dir, f"{selected_item}.pdf")
    
    # Controlla se il PDF esiste già
    if os.path.exists(pdf_path):
        return pdf_path, fitz.open(pdf_path)
    
    # Se il PDF non esiste, lo scarica
    link = df.loc[df[ingredient_column] == selected_item, 'Link'].values[0]
    try:
        # Scarica il contenuto HTML della pagina
        html_content = requests.get(link).text
        page_soup = BeautifulSoup(html_content, 'html.parser')
        pdf_link_element = page_soup.find(id="ContentContainer_ContentBottom_ingredientReferences").find('a')
        if not pdf_link_element:
            raise Exception("PDF non trovato nella pagina.")

        # Costruisci il link completo del PDF
        pdf_link = pdf_link_element['href']
        pdf_link = pdf_link if pdf_link.startswith('https://') else 'https://cir-reports.cir-safety.org/' + pdf_link[3:]
        response = requests.get(pdf_link)

        if not response.ok:
            raise Exception(f"Errore durante il download del PDF: {response.status_code}")

        # Salva il PDF localmente
        os.makedirs(pdf_dir, exist_ok=True)
        with open(pdf_path, 'wb') as file:
            file.write(response.content)

        return pdf_path, fitz.open(pdf_path)
    except Exception as e:
        st.error(f"Errore durante l'apertura del link: {e}")
        return None, None

# Funzione per estrarre il testo dal PDF
def extract_text_from_pdf(pdf_path):
    output_string = io.StringIO()
    with open(pdf_path, 'rb') as fin:
        extract_text_to_fp(fin, output_string, laparams=LAParams(), output_type='text', codec='utf-8')
    return output_string.getvalue()

# Funzione per estrarre valori numerici dal testo
def extract_numeric_value(text):
    match = re.search(r"([<>]?)\s*(\d+(?:\.\d+)?)\s*(mg/kg|mg/m³|g/kg)", text, re.IGNORECASE)
    if match:
        comparator, value, unit = match.groups()
        value = float(value)
        if comparator == '>':
            value += 0.1
        elif comparator == '<':
            value -= 0.1
        return value, unit
    return None, None

# Funzione per estrarre le frasi di contesto
def extract_context_sentences(text, pattern):
    doc = nlp(text)
    sentences = list(doc.sents)
    context_sentences = []
    
    for i, sent in enumerate(sentences):
        if re.search(pattern, sent.text, re.IGNORECASE):
            start_index = max(0, i - 2)
            end_index = min(len(sentences), i + 2)
            context_sentences.append(" ".join([s.text for s in sentences[start_index:end_index]]))
            
    return context_sentences

# Funzione per analizzare il PDF
def analyze_pdf(pdf_path):
    noael_list = []
    ld50_list = []

    full_text = extract_text_from_pdf(pdf_path)
    pdf_document = fitz.open(pdf_path)

    # Definizione dei pattern per NOAEL e LD50
    noael_patterns = [r"NOAEL", r"no[- ]observed[- ]adverse[- ]effect[- ]level", r"NOEL", r"no[- ]effect[- ]level"]
    ld50_patterns = [r"LD50", r"lethal[- ]dose", r"median[- ]lethal[- ]dose"]

    # Analisi di ogni pagina del PDF
    for i, page in enumerate(pdf_document):
        page_text = page.get_text("text")
        for pattern in noael_patterns:
            if re.search(pattern, page_text, re.IGNORECASE):
                value, unit = extract_numeric_value(page_text)
                if value is not None:
                    context = extract_context_sentences(page_text, pattern)
                    for sentence in context:
                        noael_list.append((i, sentence, value, unit))
        for pattern in ld50_patterns:
            if re.search(pattern, page_text, re.IGNORECASE):
                value, unit = extract_numeric_value(page_text)
                if value is not None:
                    context = extract_context_sentences(page_text, pattern)
                    for sentence in context:
                        ld50_list.append((i, sentence, value, unit))

    # Ordina e prendi solo il primo risultato per NOAEL e LD50
    noael_list = sorted(noael_list, key=lambda x: x[2])[:1]
    ld50_list = sorted(ld50_list, key=lambda x: x[2])[:1]

    return noael_list, ld50_list

# Funzione per visualizzare una pagina del PDF
def display_pdf_page(pdf_path, page_number):
    doc = fitz.open(pdf_path)
    page = doc.load_page(page_number - 1)  # i numeri di pagina iniziano da 0
    zoom = 4.0  # aumenta il valore per migliorare la risoluzione
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    img_bytes = pix.tobytes("png")
    encoded = base64.b64encode(img_bytes).decode()
    return f'<img src="data:image/png;base64,{encoded}" style="width:100%">'

def toggle_noael():
    st.session_state.show_noael = not st.session_state.show_noael

def toggle_ld50():
    st.session_state.show_ld50 = not st.session_state.show_ld50

# Funzione per visualizzare l'analisi
def display_analysis(pdf_path, noael_list, ld50_list):
    if 'show_noael' not in st.session_state:
        st.session_state.show_noael = False
    if 'show_ld50' not in st.session_state:
        st.session_state.show_ld50 = False

    if noael_list or ld50_list:
        st.success("Analisi completata con successo!")

        col1, col2 = st.columns(2)

        with col1:
            st.subheader("NOAEL (Livello di Effetto Avverso Non Osservato)")
            if noael_list:
                page_num, sentence, value, unit = noael_list[0]
                st.info(f"**Pagina numero: {page_num + 1}**")
                st.write(sentence)
                
                button_text = "Nascondi pagina NOAEL" if st.session_state.show_noael else "Mostra pagina NOAEL"
                st.button(button_text, key='noael_button', on_click=toggle_noael)
                
                if st.session_state.show_noael:
                    st.markdown(display_pdf_page(pdf_path, page_num + 1), unsafe_allow_html=True)

        with col2:
            st.subheader("LD50 (Dose Letale, 50%)")
            if ld50_list:
                page_num, sentence, value, unit = ld50_list[0]
                st.info(f"**Pagina numero: {page_num + 1}**")
                st.write(sentence)
                
                button_text = "Nascondi pagina LD50" if st.session_state.show_ld50 else "Mostra pagina LD50"
                st.button(button_text, key='ld50_button', on_click=toggle_ld50)
                
                if st.session_state.show_ld50:
                    st.markdown(display_pdf_page(pdf_path, page_num + 1), unsafe_allow_html=True)
            else:
                st.warning("LD50 non trovato.")
    
    else:
        st.warning("Nessun valore NOAEL o LD50 trovato nel documento.")

# Funzione principale dell'applicazione
def main_app():
    st.title("🧪 Analizzatore di Ingredienti")

    # Sidebar
    with st.sidebar:
        st.header("Opzioni")
        if st.button('🔄 Aggiorna Database', key='update_db_sidebar'):
            with st.spinner("Aggiornamento del database in corso..."):
                scraper_main()  # da SCRAPER_DEF
            st.success("Database aggiornato con successo!")
            st.info("Dati ottenuti dal sito web CIR https://www.cir-safety.org/")
        
        st.header("Filtra per lettera o numero")
        symbols = "!ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        
        if 'selected_letter' not in st.session_state:
            st.session_state.selected_letter = None
        
        cols = st.columns(7)  # 7 colonne per un layout più compatto
        for i, symbol in enumerate(symbols):
            if cols[i % 7].button(symbol, key=f"symbol_{symbol}", help=f"Filtra per {symbol}"):
                st.session_state.selected_letter = symbol
        
        if st.button("Cancella", key="clear_button"):
            st.session_state.selected_letter = None

    # Contenuto principale
    st.write("Seleziona un ingrediente per visualizzare i valori NOAEL e LD50")

    # Leggi il file Excel con gli ingredienti
    input_excel = 'cir_reports.xlsx'
    if os.path.exists(input_excel):
        df = pd.read_excel(input_excel)
        ingredient_column = 'Ingredient_Name' if 'Ingredient_Name' in df.columns else 'INCI_Name'
        if ingredient_column not in df.columns:
            st.error(f"La colonna '{ingredient_column}' non è presente nel file Excel.")
            return
    else:
        st.error(f"Il file '{input_excel}' non è stato trovato.")
        return

    # Filtra il DataFrame in base alla lettera selezionata
    if st.session_state.selected_letter:
        selected_letter = st.session_state.selected_letter
        if selected_letter == "!":
            filtered_df = df[df[ingredient_column].str.match(r'^\d', na=False)]
        else:
            filtered_df = df[df[ingredient_column].str.startswith(selected_letter, na=False)]
    else:
        filtered_df = df

    # Ricerca assistita con autocompletamento
    selected_item = st.selectbox('🔍 Cerca ingrediente', filtered_df[ingredient_column].tolist(), key='ingredient_search')

    # Pulsante di ricerca
    if st.button('🔬 Analizza Ingrediente', key='analyze_button'):
        with st.spinner(f"Analisi di {selected_item} in corso..."):
            pdf_path, pdf_file = read_generate_and_save_pdf(selected_item, df, ingredient_column)
            if pdf_file:
                st.session_state.noael_list, st.session_state.ld50_list = analyze_pdf(pdf_path)
                st.session_state.pdf_path = pdf_path
                st.session_state.analysis_done = True
            else:
                st.error("Impossibile analizzare il PDF. Controlla la connessione e riprova.")

    # Visualizza l'analisi e il pulsante di download del PDF
    if 'analysis_done' in st.session_state and st.session_state.analysis_done:
        display_analysis(st.session_state.pdf_path, st.session_state.noael_list, st.session_state.ld50_list)
        
        # Aggiungi il pulsante di download del PDF
        with open(st.session_state.pdf_path, "rb") as pdf_file:
            PDFbyte = pdf_file.read()
        
        st.download_button(
            label="📥 Scarica PDF della Ricerca Scientifica",
            data=PDFbyte,
            file_name=f"{selected_item}_ricerca.pdf",
            mime="application/pdf"
        )

    # Footer
    st.markdown("---")
    st.markdown("Sviluppato da Tedesco e Fenzi")

# Punto di ingresso dell'applicazione
if __name__ == "__main__":
    main_app()
