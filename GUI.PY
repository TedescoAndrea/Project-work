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
import time

# Load the English language model for NLP
nlp = spacy.load("en_core_web_sm")

# Streamlit page configuration (MUST be the first Streamlit command)
st.set_page_config(page_title="Ingredient Analyzer", page_icon="🧪", layout="wide")

# Custom CSS for improved styling
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


def read_generate_and_save_pdf(selected_item, df, ingredient_column):
    pdf_dir = "pdf_momentari"
    pdf_path = os.path.join(pdf_dir, f"{selected_item}.pdf")
    
    if os.path.exists(pdf_path):
        return pdf_path, fitz.open(pdf_path)
    
    link = df.loc[df[ingredient_column] == selected_item, 'Link'].values[0]
    try:
        html_content = requests.get(link).text
        page_soup = BeautifulSoup(html_content, 'html.parser')
        pdf_link_element = page_soup.find(id="ContentContainer_ContentBottom_ingredientReferences").find('a')
        if not pdf_link_element:
            raise Exception("PDF non trovato nella pagina.")

        pdf_link = pdf_link_element['href']
        pdf_link = pdf_link if pdf_link.startswith('https://') else 'https://cir-reports.cir-safety.org/' + pdf_link[3:]
        response = requests.get(pdf_link)

        if not response.ok:
            raise Exception(f"Errore durante il download del PDF: {response.status_code}")

        os.makedirs(pdf_dir, exist_ok=True)
        with open(pdf_path, 'wb') as file:
            file.write(response.content)

        return pdf_path, fitz.open(pdf_path)
    except Exception as e:
        st.error(f"Errore durante l'apertura del link: {e}")
        return None, None

def extract_text_from_pdf(pdf_path):
    output_string = io.StringIO()
    with open(pdf_path, 'rb') as fin:
        extract_text_to_fp(fin, output_string, laparams=LAParams(), output_type='text', codec='utf-8')
    return output_string.getvalue()

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

def analyze_pdf(pdf_path):
    noael_list = []
    ld50_list = []

    full_text = extract_text_from_pdf(pdf_path)
    pdf_document = fitz.open(pdf_path)

    noael_patterns = [r"NOAEL", r"no[- ]observed[- ]adverse[- ]effect[- ]level", r"NOEL", r"no[- ]effect[- ]level"]
    ld50_patterns = [r"LD50", r"lethal[- ]dose", r"median[- ]lethal[- ]dose"]

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

    noael_list = sorted(noael_list, key=lambda x: x[2])[:1]
    ld50_list = sorted(ld50_list, key=lambda x: x[2])[:1]

    return noael_list, ld50_list

def display_pdf_page(pdf_path, page_number):
    doc = fitz.open(pdf_path)
    page = doc.load_page(page_number - 1)  # page numbers start from 0
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

def display_analysis(pdf_path, noael_list, ld50_list):
    if 'show_noael' not in st.session_state:
        st.session_state.show_noael = False
    if 'show_ld50' not in st.session_state:
        st.session_state.show_ld50 = False

    if noael_list or ld50_list:
        st.success("Analisi completata con successo!")

        col1, col2 = st.columns(2)

        with col1:
            st.subheader("NOAEL (No Observed Adverse Effect Level)")
            if noael_list:
                page_num, sentence, value, unit = noael_list[0]
                st.info(f"**Pagina numero: {page_num + 1}**")
                st.write(sentence)
                
                button_text = "Nascondi pagina NOAEL" if st.session_state.show_noael else "Mostra pagina NOAEL"
                st.button(button_text, key='noael_button', on_click=toggle_noael)
                
                if st.session_state.show_noael:
                    st.markdown(display_pdf_page(pdf_path, page_num + 1), unsafe_allow_html=True)

        with col2:
            st.subheader("LD50 (Lethal Dose, 50%)")
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

def main_app():
    st.title("🧪 Ingredient Analyzer")

    # Sidebar
    with st.sidebar:
        st.header("Options")
        if st.button('🔄 Update Database', key='update_db_sidebar'):
            with st.spinner("Updating database..."):
                scraper_main()  # from SCRAPER_DEF
            st.success("Database updated successfully!")
            st.info("Data obtained from CIR website https://www.cir-safety.org/")
        
        st.header("Filter by letter or number")
        symbols = "!ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        
        if 'selected_letter' not in st.session_state:
            st.session_state.selected_letter = None
        
        cols = st.columns(7)  # 7 columns for a more compact layout
        for i, symbol in enumerate(symbols):
            if cols[i % 7].button(symbol, key=f"symbol_{symbol}", help=f"Filter by {symbol}"):
                st.session_state.selected_letter = symbol
        
        if st.button("Clear", key="clear_button"):
            st.session_state.selected_letter = None

    # Main content
    st.write("Select an ingredient to view NOAEL and LD50 values")

    # Read the Excel file with ingredients
    input_excel = 'cir_reports.xlsx'
    if os.path.exists(input_excel):
        df = pd.read_excel(input_excel)
        ingredient_column = 'Ingredient_Name' if 'Ingredient_Name' in df.columns else 'INCI_Name'
        if ingredient_column not in df.columns:
            st.error(f"The column '{ingredient_column}' is not present in the Excel file.")
            return
    else:
        st.error(f"The file '{input_excel}' was not found.")
        return

    # Filter the DataFrame based on the selected letter
    if st.session_state.selected_letter:
        selected_letter = st.session_state.selected_letter
        if selected_letter == "!":
            filtered_df = df[df[ingredient_column].str.match(r'^\d', na=False)]
        else:
            filtered_df = df[df[ingredient_column].str.startswith(selected_letter, na=False)]
    else:
        filtered_df = df

    # Assisted search with autocomplete
    selected_item = st.selectbox('🔍 Search ingredient', filtered_df[ingredient_column].tolist(), key='ingredient_search')

    # Search button
    if st.button('🔬 Analyze Ingredient', key='analyze_button'):
        with st.spinner(f"Analyzing {selected_item}..."):
            pdf_path, pdf_file = read_generate_and_save_pdf(selected_item, df, ingredient_column)
            if pdf_file:
                st.session_state.noael_list, st.session_state.ld50_list = analyze_pdf(pdf_path)
                st.session_state.pdf_path = pdf_path
                st.session_state.analysis_done = True
            else:
                st.error("Unable to analyze the PDF. Check your connection and try again.")

    # Show the analysis
    if 'analysis_done' in st.session_state and st.session_state.analysis_done:
        display_analysis(st.session_state.pdf_path, st.session_state.noael_list, st.session_state.ld50_list)

    # Footer
    st.markdown("---")
    st.markdown("Developed by Tedesco e Fenzi")

if __name__ == "__main__":
    main_app()
