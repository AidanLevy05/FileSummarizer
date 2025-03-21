import pdfplumber
from pptx import Presentation
from nltk.tokenize import sent_tokenize
import spacy
import os

current_dir = os.path.dirname(os.path.abspath(__file__))
model_path = os.path.join(current_dir, "en_core_web_sm")
nlp = spacy.load(model_path)

def extract_text_from_pdf(pdf_path):
    text_per_slide = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text_per_slide.append(page_text)
            else:
                print(f"Failed to extract text from page {pdf.pages.index(page)}")
    return text_per_slide

def extract_text_from_pptx(pptx_path):
    text_per_slide = []
    presentation = Presentation(pptx_path)
    for slide in presentation.slides:
        slide_text = []
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                slide_text.append(shape.text)
        text_per_slide.append("\n".join(slide_text))
    return text_per_slide

def get_sentence_count(word_count):
    if word_count <= 500:
        return 3
    elif 501 <= word_count <= 1000:
        return 7
    else:
        return 15

def summarize_text_spacy(text):
    doc = nlp(text)
    
    sentences = list(doc.sents)
    important_sentences = []
    
    for sent in sentences:
        if any(ent.label_ in ["PERSON", "ORG", "DATE"] for ent in sent.ents):
            important_sentences.append(sent.text)
    
    if len(important_sentences) < 3:
        important_sentences = [sent.text for sent in sentences[:5]]
    
    return important_sentences

def process_file(file_path):
    if file_path.endswith(".pdf"):
        slides_text = extract_text_from_pdf(file_path)
    elif file_path.endswith(".pptx"):
        slides_text = extract_text_from_pptx(file_path)
    else:
        raise ValueError("Unsupported file type. Please provide a .pdf or .pptx file.")


    summaries = []
    for slide in slides_text:
        summary = summarize_text_spacy(slide)
        summaries.append(summary)
    
    return summaries