import os
import sys
import io
import re
import threading
from dataclasses import dataclass
from typing import Dict, List, Tuple, Optional, Any
from datetime import datetime
from docx import Document

import fitz
from PIL import Image
import pytesseract
import customtkinter as ctk
from transformers import pipeline, AutoTokenizer
from collections import defaultdict


prompt_gpt = (  
            "Tu es un agent spécifiquement entraîné pour résumer avec rigueur. L'utilisateur va te fournir un procès-verbal à résumer.\n"

            "Règles impératives à respecter :\n"

            "Tu ne dois JAMAIS modifier, réécrire ou reformuler les informations factuelles : noms, prénoms, lieux, montants, etc. doivent rester identiques au texte source.\n"

            "Quand il est demandé de recopier un texte mot à mot, tu dois le copier exactement comme il apparaît. Aucun ajout, suppression, correction ou reformulation n'est autorisé.\n"

            "Les dates doivent TOUJOURS être converties au format JJ/MM/AAAA. Exemple : '20 JUIN 2025' devient '20/06/2025'.\n"

            "Si une information est absente, ignore simplement la rubrique concernée dans ton rendu.\n"

            "Structure attendue du rendu :\n"
            "Fichier : [insère ici le nom du fichier PDF traité]\n"
            "Date : [insère ici la date du procès-verbal trouvée en en-tête au format JJ/MM/AAAA]\n"
            "Décisions :\n"
            "- Approbation des comptes : [indique ici si les comptes sont approuvés ou non, et copie la phrase exacte du procès-verbal le prouvant]\n"
            "- Affectation des résultats / dividendes : [si cette section est mentionnée, résume la décision en une phrase extraite ou synthétisée à partir du texte]\n"

            "Résumé : [utilise cette section SEULEMENT si les deux sections précédentes sont absentes. Rédige un paragraphe résumant les points clés du procès-verbal, sans reformulation des faits.]\n"

            "Conventions : [copie ici mot à mot tout passage du procès-verbal concernant les conventions, si présent]\n"
        )


if getattr(sys, 'frozen', False):
    # Running as a PyInstaller bundled executable
    bundle_dir = sys._MEIPASS
else:
    # Running as a normal Python script (e.g., during development)
    bundle_dir = os.path.dirname(os.path.abspath(__file__))

# Configure Tesseract path if necessary
tesseract_path = os.path.join(bundle_dir, 'assets', 'tesseract', 'tesseract.exe')
if os.path.exists(tesseract_path):
    pytesseract.pytesseract.tesseract_cmd = tesseract_path

# -------------- NER --------------
_ner_pipeline: Optional[Any] = None
_tokenizer: Optional[Any] = None
MODEL_NAME = "cmarkea/distilcamembert-base-ner"  # French model

def _initialize_ner_pipeline():
    global _ner_pipeline, _tokenizer
    if _ner_pipeline is None or _tokenizer is None:
        # Check if the model/tokenizer are already available, otherwise download
        try:
            _tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME)
            _ner_pipeline = pipeline("ner", model=MODEL_NAME, tokenizer=_tokenizer, aggregation_strategy="simple")
        except Exception as e:
            # Handle potential connection/download issues gracefully
            print(f"Error initializing NER pipeline: {e}")
            raise

# -------------- Data --------------
@dataclass
class MappingEntry:
    tag: str
    type: str
    original_value: str

# -------------- Mapping Manager (New Class) --------------
class MappingManager:
    """Manages global mapping and entity counters for consistent tagging across multiple texts/pages."""
    def __init__(self):
        # Maps (type, original_value) to tag string (e.g., PERSON_1)
        self.seen: Dict[Tuple[str, str], str] = {}
        # Maps tag string (e.g., PERSON_1) to MappingEntry
        self.global_mapping: Dict[str, MappingEntry] = {}
        # Counter for each entity type (e.g., {'PERSON': 5})
        self.counters: Dict[str, int] = defaultdict(int)

    def get_or_create_tag(self, typ: str, original: str) -> str:
        key = (typ, original)
        # 1. Check if we've seen this exact entity (type + value) before
        if key in self.seen:
            tag = self.seen[key]
        else:
            # 2. Create a new tag
            self.counters[typ] += 1
            tag = f"{typ}_{self.counters[typ]}"
            # 3. Store the new mapping globally
            self.global_mapping[tag] = MappingEntry(tag=tag, type=typ, original_value=original)
            self.seen[key] = tag
        return tag


# -------------- Extraction --------------
def extract_text_from_pdf(pdf_path: str, use_ocr: bool = True, ocr_lang: str = "fra") -> List[Tuple[int, str]]:
    """Extract text page by page, with OCR fallback, returns a list of (page_num, text)."""
    page_texts: List[Tuple[int, str]] = []
    with fitz.open(pdf_path) as doc:
        for i, page in enumerate(doc):
            text = page.get_text("text").strip()
            # If text is empty or OCR is forced, try OCR
            if (not text or use_ocr):
                # Using a higher DPI (e.g., 300/72) for better OCR quality
                pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72), alpha=False)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                ocr_text = pytesseract.image_to_string(img, lang=ocr_lang).strip()
                if ocr_text:
                    text = ocr_text
            if text:
                # Store text along with its page number (1-based)
                page_texts.append((i + 1, text))
    return page_texts

# -------------- Anonymization --------------
SENSITIVE_LABELS = {"PER": "PERSON", "LOC": "LOCATION", "ORG": "ORG"}
REGEX_PATTERNS = [
    ("EMAIL", r"\b[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}\b"),
    ("PHONE", r"\b(?:\+?\d{1,3}[\s.\-]?)?(?:\(?\d{2}\)?[\s.\-]?){4,5}\d{2}\b"),
    # Date pattern adapted to be more robust
    ("DATE", r"\b(?:\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4}|\d{4}[/.\-]\d{1,2}[/.\-]\d{1,2}|\d{1,2}\s?(?:janv\.?|févr\.?|mars|avr\.?|mai|juin|juil\.?|août|sept\.?|oct\.?|nov\.?|déc\.?)\s?\d{2,4})\b"),
]

def find_regex_entities(text: str) -> List[Tuple[int, int, str, str]]:
    spans = []
    for typ, pattern in REGEX_PATTERNS:
        # Use re.IGNORECASE for matching
        for m in re.finditer(pattern, text, flags=re.IGNORECASE):
            spans.append((m.start(), m.end(), m.group(0), typ))
    return spans

def anonymize_text(text: str, manager: MappingManager) -> Tuple[str, Dict[str, MappingEntry]]:
    """
    Anonymizes a single block of text (e.g., a page).
    Uses the provided MappingManager to ensure consistent tagging.
    Returns the anonymized text and the mapping *created/updated* during this process.
    """
    _initialize_ner_pipeline()
    
    # 1. Collect NER entities
    ner_spans = []
    # The 'ner' pipeline returns a list of dictionaries.
    for ent in _ner_pipeline(text):
        grp = ent.get("entity_group")
        if grp in SENSITIVE_LABELS:
            # We use SENSITIVE_LABELS[grp] (e.g., "PERSON") as the standard type
            ner_spans.append((ent["start"], ent["end"], ent["word"], SENSITIVE_LABELS[grp]))
    
    # 2. Collect Regex entities
    reg_spans = find_regex_entities(text)
    
    # 3. Combine and sort all spans. Sorting reverse is crucial for safe string replacement.
    all_spans = ner_spans + reg_spans
    all_spans.sort(key=lambda x: x[0], reverse=True)

    # mapping_local will only contain new/used tags for this text block
    mapping_local: Dict[str, MappingEntry] = {}
    text_list = list(text)

    for start, end, original, typ in all_spans:
        # Use the MappingManager to get a consistent tag
        tag = manager.get_or_create_tag(typ, original)
        
        # Add the used tag to the local mapping (optional, but useful for debugging/display)
        if tag not in mapping_local:
            mapping_local[tag] = manager.global_mapping[tag]
        
        # Perform replacement in the text (done in reverse order for safety)
        text_list[start:end] = list(tag)
        
    return "".join(text_list), mapping_local

def deanonymize_text(anonymized_text: str, mapping: Dict[str, MappingEntry]) -> str:
    """Reverses anonymization using a provided mapping."""
    result = anonymized_text
    # Sorting by length reverse is crucial to replace longer tags first (e.g., PERSON_10 before PERSON_1)
    for tag in sorted(mapping.keys(), key=len, reverse=True):
        # Use re.escape to handle any special characters in the tag name
        result = re.sub(re.escape(tag), mapping[tag].original_value, result)
    return result


# -------------- GUI --------------
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("green")

class AnonApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Anonymiseur de PDF (FR)")
        self.geometry("1100x750")

        # Replaced single pdf_path with a list
        self.pdf_paths: List[str] = []
        
        # Combined text from all pages/files
        self.text = "" 
        self.anonymized = ""
        
        # Global mapping manager instance
        self.mapping_manager = MappingManager()
        
        # Storing the original text, page by page, file by file for reconstruction
        # List of (filename, page_num, text)
        self.extracted_pages: List[Tuple[str, int, str]] = []

        self._build_gui()

    # ====================== GUI BUILD ======================
    def _build_gui(self):
        self.tabs = ctk.CTkTabview(self)
        self.tabs.pack(fill="both", expand=True, padx=15, pady=15)

        self.tab_anon = self.tabs.add("Anonymiser")
        self.tab_deanon = self.tabs.add("Déanonymiser")

        self._build_anon_tab()
        self._build_deanon_tab()

        # Status label
        self.status_label = ctk.CTkLabel(self, text="Prêt", text_color="green", anchor="w")
        self.status_label.pack(fill="x", padx=10, pady=(0, 8))

    # ====================== ANONYMISER TAB ======================
    def _build_anon_tab(self):
        frame = self.tab_anon

        # Top bar
        top = ctk.CTkFrame(frame)
        top.pack(fill="x", pady=(10, 10))

        self.btn_browse = ctk.CTkButton(top, text="📂 Importer PDF(s)", command=self.browse_pdf)
        self.btn_browse.pack(side="left", padx=5, pady=5)

        self.btn_extract = ctk.CTkButton(top, text="🪄 Extraire texte", command=self.extract_text_threaded, state="disabled")
        self.btn_extract.pack(side="left", padx=5)

        self.btn_anonymize = ctk.CTkButton(top, text="🔒 Anonymiser", command=self.anonymize_threaded, state="disabled")
        self.btn_anonymize.pack(side="left", padx=5)

        self.btn_copy = ctk.CTkButton(top, text="📋 Copier sortie", command=self.copy_output, state="disabled")
        self.btn_copy.pack(side="left", padx=5)

        # Text previews
        bottom = ctk.CTkFrame(frame)
        bottom.pack(fill="both", expand=True, padx=10, pady=10)

        # Original Textbox - now displays combined text with file/page headers
        self.text_original = ctk.CTkTextbox(bottom, wrap="word", font=("Inter", 13))
        self.text_original.pack(side="left", fill="both", expand=True, padx=(0, 5))

        self.text_anonymized = ctk.CTkTextbox(bottom, wrap="word", font=("Inter", 13))
        self.text_anonymized.pack(side="right", fill="both", expand=True, padx=(5, 0))

    # ====================== DEANONYMISER TAB ======================
    def _build_deanon_tab(self):
        frame = self.tab_deanon
        
        # Frame for input/output textboxes
        middle = ctk.CTkFrame(frame)
        middle.pack(fill="both", expand=True, padx=10, pady=10)

        self.text_deanon_input = ctk.CTkTextbox(middle, wrap="word", font=("Inter", 13))
        self.text_deanon_input.pack(side="left", fill="both", expand=True, padx=(0, 5))

        self.text_deanon_output = ctk.CTkTextbox(middle, wrap="word", font=("Inter", 13))
        self.text_deanon_output.pack(side="right", fill="both", expand=True, padx=(5, 0))

        # Bottom bar for buttons
        bottom = ctk.CTkFrame(frame)
        bottom.pack(fill="x", pady=(0, 10))
        
        self.btn_deanon = ctk.CTkButton(bottom, text="🔓 Déanonymiser et Afficher", command=self.deanonymize_and_display)
        self.btn_deanon.pack(side="left", padx=10, pady=5)
        
        self.btn_save = ctk.CTkButton(bottom, text="💾 Sauvegarder sortie", command=self.save_deanon_output)
        self.btn_save.pack(side="right", padx=10, pady=5)

    # ====================== STATUS ======================
    def set_status(self, msg: str, color="green"):
        self.status_label.configure(text=msg, text_color=color)
        self.update_idletasks()

    # ====================== FILE HANDLERS ======================
    def browse_pdf(self):
        from tkinter import filedialog
        # Allow multiple file selection
        paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if not paths:
            return
        
        self.pdf_paths = list(paths)
        
        # Reset state for new files
        self.text = ""
        self.anonymized = ""
        self.extracted_pages = []
        self.mapping_manager = MappingManager() # Reset mapping manager
        self.text_original.delete("1.0", "end")
        self.text_anonymized.delete("1.0", "end")
        
        filenames = ", ".join([os.path.basename(p) for p in self.pdf_paths])
        self.set_status(f"{len(self.pdf_paths)} PDF(s) chargé(s): {filenames[:80]}...", color="green")
        
        self.btn_extract.configure(state="normal", fg_color="green")
        self.btn_anonymize.configure(state="disabled", fg_color="gray")
        self.btn_copy.configure(state="disabled", fg_color="gray")

    def extract_text_threaded(self):
        if not self.pdf_paths:
            self.set_status("Aucun PDF sélectionné", color="red")
            return
        threading.Thread(target=self._extract_text, daemon=True).start()

    def _extract_text(self):
        self.extracted_pages = []
        full_text_original = []
        total_pages = 0
        
        try:
            self.set_status(f"Extraction du texte de {len(self.pdf_paths)} fichier(s)...", color="orange")
            
            for pdf_path in self.pdf_paths:
                filename = os.path.basename(pdf_path)
                
                # Get text page by page
                page_texts = extract_text_from_pdf(pdf_path)
                
                for page_num, text in page_texts:
                    total_pages += 1
                    header = f"\n\n--- FICHIER: {filename} / PAGE {page_num} ---\n"
                    
                    full_text_original.append(header)
                    full_text_original.append(text)
                    
                    self.extracted_pages.append((filename, page_num, text))

            self.text = "\n\n".join(full_text_original)
            
            self.text_original.delete("1.0", "end")
            self.text_original.insert("1.0", self.text)
            
            self.btn_anonymize.configure(state="normal", fg_color="green")
            self.set_status(f"Texte extrait avec succès ({total_pages} pages).", color="green")
        except Exception as e:
            self.set_status(f"Erreur d'extraction: {e}", color="red")

    # ====================== ANONYMIZATION ======================
    def anonymize_threaded(self):
        if not self.extracted_pages:
            self.set_status("Aucun texte extrait à anonymiser", color="red")
            return
        threading.Thread(target=self._anonymize, daemon=True).start()

    def _anonymize(self):
        # Reset the global mapping manager and output texts
        self.mapping_manager = MappingManager()
        self.anonymized = ""
        self.text_anonymized.delete("1.0", "end")
        
        full_text_anonymized = []
        total_entities = 0
        
        try:
            self.set_status(f"Anonymisation de {len(self.extracted_pages)} pages en cours…", color="orange")
            
            for filename, page_num, original_text in self.extracted_pages:
                
                # Use the page text for anonymization
                anon_text, mapping_local = anonymize_text(original_text, self.mapping_manager)
                total_entities += len(mapping_local)

                header = f"\n\n--- FICHIER: {filename}"
                full_text_anonymized.append(header)
                full_text_anonymized.append(anon_text)
                full_text_anonymized.append(f"--- FIN ---\n")
                
                # Update status for user feedback on progress
                self.set_status(f"Anonymisation... (Fichier: {filename}, Page: {page_num}, Entités: {total_entities})", color="orange")

            self.anonymized = "\n\n".join(full_text_anonymized)
            self.text_anonymized.insert("1.0", self.anonymized)
            
            self.btn_copy.configure(state="normal", fg_color="green")
            self.set_status(f"Anonymisation terminée ({len(self.pdf_paths)} fichiers, {total_entities} entités uniques).", color="green")
            
        except Exception as e:
            self.set_status(f"Erreur d’anonymisation: {e}", color="red")
            print(f"Anonymization Error: {e}")

    def copy_output(self):
        self.clipboard_clear()
        output_to_copy = prompt_gpt + "\n\n" + self.text_anonymized.get("1.0", "end").strip()
        self.clipboard_append(output_to_copy)
        self.set_status("Texte copié dans le presse-papiers.", color="green")

    # ====================== DEANONYMIZATION ======================
    def deanonymize_and_display(self):
        """Processes the input text using the global mapping and displays the result."""
        text_in = self.text_deanon_input.get("1.0", "end").strip()
        
        # Use the global mapping from the MappingManager
        global_mapping = self.mapping_manager.global_mapping 
        
        if not text_in:
            self.set_status("Aucun texte à déanonymiser", color="red")
            return
        if not global_mapping:
            self.set_status("Aucune cartographie d'entités disponible (anonymisez d'abord)", color="red")
            return

        try:
            self.set_status("Déanonymisation en cours...", color="orange")
            # Use the complete, global mapping for deanonymization
            text_out = deanonymize_text(text_in, global_mapping)
            self.text_deanon_output.delete("1.0", "end")
            self.text_deanon_output.insert("1.0", text_out)
            self.set_status("Texte déanonymisé avec succès.", color="green")
        except Exception as e:
            self.set_status(f"Erreur de déanonymisation: {e}", color="red")

    def manage_all_summaries_in_document(self,document_path: str, new_summaries: List[Dict]):
        """
        Manages all summaries in a Word document: extracts existing, combines with new,
        sorts them, clears the document, and rewrites all summaries in date order.

        Args:
            document_path (str): The path to the Word document.
            new_summaries (List[Dict]): A list of dictionaries, where each dict represents
                                    a new summary:
                                    {'date': datetime_obj, 'content': 'summary text'}
        """
        def extract_existing_summaries(document: Document) -> List[Dict]:
            """
            Extract existing summaries from the document by finding content between Level 1 headings
            and looking for dates within that content.
            """
            summaries = []
            current_summary = None
            
            for paragraph in document.paragraphs:
                # Check if this is a Level 1 heading
                if paragraph.style.name == 'Heading 1':
                    # Save the previous summary if it exists and has a valid date
                    if current_summary and current_summary.get('date'):
                        summaries.append(current_summary)
                    
                    # Start a new summary
                    current_summary = {
                        'heading_text': paragraph.text,
                        'content_paragraphs': [],
                        'date': None
                    }
                    
                    # Try to extract date from the heading
                    date_from_heading = extract_date_from_text(paragraph.text)
                    if date_from_heading:
                        current_summary['date'] = date_from_heading
                
                elif current_summary is not None:
                    # This paragraph belongs to the current summary
                    current_summary['content_paragraphs'].append(paragraph.text)
                    
                    # If we haven't found a date yet, try to extract it from this paragraph
                    if not current_summary['date']:
                        date_from_content = extract_date_from_text(paragraph.text)
                        if date_from_content:
                            current_summary['date'] = date_from_content
            
            # Don't forget the last summary
            if current_summary and current_summary.get('date'):
                summaries.append(current_summary)
            
            return summaries

        def extract_date_from_text(text: str) -> Optional[datetime]:
            """
            Extract a date from text using various date patterns.
            Returns the first valid date found, or None if no date is found.
            """
            # Common date patterns
            date_patterns = [
                r'\b(\d{4}-\d{2}-\d{2})\b',           # YYYY-MM-DD
                r'\b(\d{2}/\d{2}/\d{4})\b',           # MM/DD/YYYY
                r'\b(\d{2}-\d{2}-\d{4})\b',           # MM-DD-YYYY
                r'\b(\d{1,2}/\d{1,2}/\d{4})\b',       # M/D/YYYY or MM/D/YYYY etc.
                r'\b(\d{4}/\d{2}/\d{2})\b',           # YYYY/MM/DD
            ]
            
            date_formats = [
                '%Y-%m-%d',
                '%m/%d/%Y',
                '%m-%d-%Y',
                '%m/%d/%Y',
                '%Y/%m/%d'
            ]
            
            for i, pattern in enumerate(date_patterns):
                matches = re.findall(pattern, text)
                for match in matches:
                    try:
                        # Handle the flexible format for pattern index 3
                        if i == 3:
                            # Try different formats for flexible date patterns
                            for fmt in ['%m/%d/%Y', '%m/%d/%Y']:
                                try:
                                    return datetime.strptime(match, fmt)
                                except ValueError:
                                    continue
                        else:
                            return datetime.strptime(match, date_formats[i])
                    except ValueError:
                        continue
            
            return None

        def clear_document_content(document: Document):
            """
            Safely clear all content from the document while preserving the document structure.
            """
            # Remove all paragraphs
            for paragraph in document.paragraphs[:]:
                p = paragraph._element
                p.getparent().remove(p)
            
            # Remove all tables
            for table in document.tables[:]:
                t = table._element
                t.getparent().remove(t)


        try:
            document = Document(document_path)
        except FileNotFoundError:
            print(f"Document not found: {document_path}. Creating new document.")
            document = Document()
        except Exception as e:
            print(f"Error loading document {document_path}: {e}")
            return

        all_summaries_to_write = []

        # 1. Extract existing summaries
        existing_summaries = extract_existing_summaries(document)
        all_summaries_to_write.extend(existing_summaries)

        # 2. Add new summaries to the list
        for new_sum_data in new_summaries:
            if 'date' not in new_sum_data or 'summary' not in new_sum_data:
                print(f"Warning: Invalid new summary data: {new_sum_data}")
                continue
                
            new_summary_date = new_sum_data['date']
            new_summary_filename = new_sum_data['filename']
            new_summary_heading = f"Fichier: {new_summary_filename}"
            
            all_summaries_to_write.append({
                'date': new_summary_date,
                'heading_text': new_summary_heading,
                'content_paragraphs': [new_sum_data['summary']]
            })

        # 3. Sort all summaries by date in descending order (most recent first)
        all_summaries_to_write.sort(key=lambda x: x['date'], reverse=True)

        # 4. Clear the document content safely
        clear_document_content(document)

        # 5. Rewrite all sorted summaries back into the document
        document.add_heading('Résumés des procès verbaux', 0)
        for summary_item in all_summaries_to_write:
            document.add_heading(summary_item['heading_text'], level=1)
            for content_p_text in summary_item['content_paragraphs']:
                if content_p_text.strip():  # Only add non-empty paragraphs
                    document.add_paragraph(content_p_text)
                    document.add_paragraph("\n")

        # 6. Save the modified document
        try:
            document.save(document_path)
            print(f"Document '{document_path}' modifié avec {len(all_summaries_to_write)} nouveaux résumés.")
        except Exception as e:
            print(f"Erreur lors de la sauvegarde du document {document_path}: {e}")

    def save_deanon_output(self):
        from tkinter import filedialog, messagebox
        from datetime import datetime
        import os
        import re

        # Get the full text from the déanonymisé output box
        text_out = self.text_deanon_output.get("1.0", "end").strip()

        if not text_out:
            self.set_status("Rien à sauvegarder (déanonymisez d'abord)", color="red")
            return

        # 1. Ask where to save the Word document
        path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")],
            title="Sauvegarder le résultat déanonymisé"
        )

        if not path:
            self.set_status("Sauvegarde annulée.", color="red")
            return

        # --- Helper: Extract date from the LLM's structured output ---
        # This function must be able to parse the date formatted as DD/MM/AAAA from the LLM output.
        def extract_date_from_summary_block(summary_block: str) -> Optional[datetime]:
            """
            Extracts the date (DD/MM/AAAA) from the 'Date :' line in a summary block.
            """
            # Regex to find 'Date : XX/XX/XXXX'
            date_regex = r"Date\s*:\s*(\d{2}/\d{2}/\d{4})"
            match = re.search(date_regex, summary_block, re.IGNORECASE)

            if match:
                date_str = match.group(1)
                try:
                    # Date format is explicitly required as JJ/MM/AAAA (i.e., %d/%m/%Y)
                    return datetime.strptime(date_str, "%d/%m/%Y")
                except ValueError:
                    return None
            return None

        # --- Helper: Parse the full output text into individual summaries ---
        def parse_llm_output(full_text: str) -> List[Dict]:
            """
            Splits the LLM output (which includes the prompt) into individual summaries 
            based on the 'Fichier : [filename]' header defined in the prompt's output structure.
            """
            summaries_list: List[Dict] = []
            
            # 1. Separate the Prompt from the Content
            # The full output starts with the prompt, which ends before the first "Fichier :".
            content_start_index = full_text.find("Fichier : ")
            if content_start_index == -1:
                # If no 'Fichier :' header is found, treat the whole thing as one block 
                # (this handles the case where the LLM might have failed to format correctly)
                return [{"date": datetime.now(), "summary": full_text, "filename": "UNKNOWN_FILE"}]
            
            content_text = full_text[content_start_index:].strip()

            # 2. Split content by the 'Fichier :' header
            # Prepend a unique marker to the start of the content to catch the first file.
            split_content = ("\nFichier : " + content_text).split("\nFichier : ")
            
            # The first element will be empty or contain non-summary junk; skip it.
            for block in split_content[1:]:
                if not block.strip():
                    continue

                # 3. Extract Filename
                # The filename is the first line/word after the header.
                filename_match = re.match(r"([^\n]+)", block)
                filename = filename_match.group(1).strip() if filename_match else "UNKNOWN_FILE"
                
                # 4. Extract Date and Summary Content
                summary_content = f"Fichier : {filename}\n{block.strip()}"
                
                summary_date = extract_date_from_summary_block(summary_content)
                
                # If the date is missing, use current date and print a warning
                if not summary_date:
                    print(f"Warning: Could not extract date for file: {filename}. Using current date.")
                    summary_date = datetime.now()

                summaries_list.append({
                    "date": summary_date,
                    "summary": summary_content,
                    "filename": filename
                })

            return summaries_list

        # --- Execution ---
        
        # Parse the full déanonymisé text to get a list of individual summary dictionaries
        new_summaries_data = parse_llm_output(text_out)

        if not new_summaries_data:
            self.set_status("Erreur de parsing : aucun résumé individuel trouvé dans la sortie.", color="red")
            return

        # Use the existing manager function to handle Word document logic
        try:
            # The manage_all_summaries_in_document function is designed to handle this list
            self.manage_all_summaries_in_document(path, new_summaries_data)
            
            messagebox.showinfo("Succès", f"Sauvegardé sous {path}")
            self.set_status(f"Fichier Word sauvegardé : {os.path.basename(path)} avec {len(new_summaries_data)} résumés.", color="green")
            
        except Exception as e:
            messagebox.showerror("Erreur de Sauvegarde", f"Impossible de sauvegarder le fichier : {e}")
            self.set_status("Erreur de Sauvegarde", color="red")



# ====================== MAIN ======================
if __name__ == "__main__":
    app = AnonApp()
    app.mainloop()