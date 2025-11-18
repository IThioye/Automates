import os
import threading
import customtkinter as ctk
import pytesseract
from pdf2image import convert_from_path
from tkinter import filedialog, scrolledtext
import tkinter as tk
from llama_cpp import Llama
from docx import Document # Import for Word document generation
import re # For regex to extract date from LLM output
from datetime import datetime # For date parsing and sorting
from typing import List, Dict, Optional
import sys
import fitz
from PIL import Image

# Set appearance mode and default color theme
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("dark-blue")


MODELS= [
            "qwen2-1_5b-instruct-q4_0.gguf",
            "gemma-2-2b-it-Q4_K_M.gguf" 
        ]
if getattr(sys, 'frozen', False):
    # Running as a PyInstaller bundled executable
    bundle_dir = sys._MEIPASS
else:
    # Running as a normal Python script (e.g., during development)
    bundle_dir = os.path.dirname(os.path.abspath(__file__)) # os.path.dirname(os.path.dirname(__file__)) 

MODEL_PATH = os.path.join(bundle_dir, 'assets', "models", MODELS[0]) 
tesseract_path = os.path.join(bundle_dir, 'assets', 'tesseract', 'tesseract.exe')
poppler_dir = os.path.join(bundle_dir, 'assets', 'poppler','Library', 'bin')
pytesseract.pytesseract.tesseract_cmd = tesseract_path



prompt1 = (
            "Tu es un agent spécifiquement entraîné pour résumer avec rigueur. "
            "L'utilisateur va te fournir un procès-verbal à résumer.\n"
            "Règles impératives :\n"
            "- Tu ne dois JAMAIS modifier, réécrire ou reformuler les informations factuelles : noms, prénoms, lieux, montants, etc. doivent rester identiques au texte source.\n"
            "- Quand il est demandé de recopier un texte mot à mot, tu dois la copier exactement comme elle apparaît. Aucun ajout, suppression, correction ou reformulation.\n"
            "- Quand il s'agit des dates, mets les TOUJOURS sous le format JJ/MM/AAAA. (ex: si tu vois 20 JUIN 2025 tu devras écrire 20/06/2025)\n"
            "- Si une information demandée est absente du texte, ignore simplement la rubrique correspondante.\n\n"
            "Ton rendu se fera sous ce format :\n"
            "Date : copier ici la date du procès-verbal trouvée en en-tête (format JJ/MM/AAAA et uniquement dans ce format)\n"
            "Décisions :\n"
            "- Approbation des comptes : indiquer si les comptes sont approuvés ou non et citer la phrase exacte prouvant cela\n"
            "- Affectation des résultats / dividendes : si mentionné, reprendre en une phrase les décisions prises (sinon, ignorer)\n"
            "- Résumé : Si il n' y a pas mentions d'approbation de comptes ou d'affectation de résultats/dividendes, faire un résumé" 
            " en un paragraphe du document avec les points essentiels et ne pas ajouter les points Approbation des comptes"
            " et Affectation des résultats / dividendes\n" 
            "- Conventions : recopier le texte mot à mot parlant des conventions,si présent\n"
        )

prompt2 = (  
            "Tu es un agent spécifiquement entraîné pour résumer avec rigueur. L'utilisateur va te fournir un procès-verbal à résumer.\n"

            "Règles impératives à respecter :\n"

            "Tu ne dois JAMAIS modifier, réécrire ou reformuler les informations factuelles : noms, prénoms, lieux, montants, etc. doivent rester identiques au texte source.\n"

            "Quand il est demandé de recopier un texte mot à mot, tu dois le copier exactement comme il apparaît. Aucun ajout, suppression, correction ou reformulation n'est autorisé.\n"

            "Les dates doivent TOUJOURS être converties au format JJ/MM/AAAA. Exemple : '20 JUIN 2025' devient '20/06/2025'.\n"

            "Si une information est absente, ignore simplement la rubrique concernée dans ton rendu.\n"

            "Structure attendue du rendu :\n"
            "Date : [insère ici la date du procès-verbal trouvée en en-tête au format JJ/MM/AAAA]\n"

            "Décisions :\n"
            "- Approbation des comptes : [indique ici si les comptes sont approuvés ou non, et copie la phrase exacte du procès-verbal le prouvant]\n"
            "- Affectation des résultats / dividendes : [si cette section est mentionnée, résume la décision en une phrase extraite ou synthétisée à partir du texte]\n"

            "Résumé : [utilise cette section SEULEMENT si les deux sections précédentes sont absentes. Rédige un paragraphe résumant les points clés du procès-verbal, sans reformulation des faits.]\n"

            "Conventions : [copie ici mot à mot tout passage du procès-verbal concernant les conventions, si présent]\n"
        )

templates = {
                "PV Attribution des comptes": prompt2
            }

class PDFToLLMApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.model_path = MODEL_PATH
        self.llm = None  # Initialize the LLM object

        # Configure main window
        self.title("IA de Résumé de PDF")
        self.geometry("1000x810")
        self.minsize(1000, 810)
        
        # Initialize variables
        self.pdf_paths = [] 
        self.extracted_texts = {} 
        self.summaries = []
        self.is_generating = False
        self.stream_active = False

        self.api_url = "http://127.0.0.1:1234/v1/chat/completions"
        self.llama_cpp_url = "http://localhost:8080/v1/chat/completions"
        self.model_entry = "gemma-2-2b-it"
        self.temperature = 0.2
        self.stream_mode = True

        # --- COMBINE UI STEPS ---
        # 1. Create the overall structure (Main UI elements, but hidden/disabled)
        self.create_ui()
        
        # 2. Overlay the loading UI
        self.create_loading_overlay() 

        # 3. Start loading the model
        self.load_model()

    def create_loading_overlay(self):
        """Create an overlay frame to show while the model loads."""
        # This frame covers the main content area
        self.loading_frame = ctk.CTkFrame(self, fg_color=self.cget("fg_color"))
        self.loading_frame.place(relx=0.5, rely=0.5, anchor="center", relwidth=1.0, relheight=1.0)
        
        self.loading_label = ctk.CTkLabel(
            self.loading_frame, 
            text="Veuillez patienter, le modèle se charge...", 
            font=("Arial", 24, "bold"), 
            text_color="gray"
        )
        self.loading_label.place(relx=0.5, rely=0.4, anchor="center")
        
        self.progress_bar = ctk.CTkProgressBar(self.loading_frame, mode="indeterminate", width=300)
        self.progress_bar.place(relx=0.5, rely=0.5, anchor="center")
        self.progress_bar.start()

    def load_model(self):
        def load():
            """Load the LLaMA model in a background thread"""
            try:
                # Simulate loading time (Remove this in actual implementation)
                import time
                time.sleep(5) 
                
                self.llm = Llama(
                    model_path=self.model_path,
                    use_mmap=True,
                    n_ctx=2048,
                    n_threads=os.cpu_count()
                )
                print(f"✅ Model loaded successfully: {self.model_path}")
                self.after(100, self.on_model_loaded) # Use self.after to run in the main thread
            except Exception as e:
                print(f"❌ Model loading failed: {e}")
                self.after(100, lambda: self.loading_label.configure(text=f"❌ Erreur: Impossible de charger le modèle. Vérifiez le chemin. \n Détail: {e}", text_color="red", wraplength=400))
                self.after(100, self.progress_bar.stop)
                
        threading.Thread(target=load, daemon=True).start()

    def on_model_loaded(self):
        """Called when the model is successfully loaded"""
        
        # 1. Stop progress and destroy the overlay
        self.progress_bar.stop()
        self.loading_frame.destroy() 

        # 2. Enable/Configure necessary UI elements if needed
        # self.generate_button.configure(state="normal") 
        
        self.status_bar.configure(text="Prêt (Modèle Chargé)", fg_color="green")
        print("UI is now fully active.")

    # --- ORIGINAL create_ui METHOD (with necessary placeholder/dummy methods) ---
    def create_ui(self):
        # Create a container frame for all main content
        self.main_container = ctk.CTkFrame(self)
        self.main_container.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Create main frame with padding
        main_frame = ctk.CTkFrame(self.main_container)
        main_frame.pack(fill="both", expand=True) # Removed main_container padx/pady to use main_frame instead
        
        # Top frame for controls
        top_frame = ctk.CTkFrame(main_frame)
        top_frame.pack(fill="x", padx=10, pady=10)
        
        # PDF selection section
        self.save_summaries_button = ctk.CTkButton(top_frame, text="Sauvegarder Résumés", command=self.save_summaries, state="disabled")
        self.save_summaries_button.pack(side="right", padx=10, pady=10)

        self.generate_all_button = ctk.CTkButton(top_frame, text="Générer tout", command=self.generate_all_summaries, state="disabled")
        self.generate_all_button.pack(side="right", padx=10, pady=10)
        
        self.browse_button = ctk.CTkButton(top_frame, text="Sélectionner PDFs", command=self.browse_pdfs)
        self.browse_button.pack(side="right", padx=10, pady=10)
    
        self.pdf_label = ctk.CTkLabel(top_frame, text="Aucun PDF sélectionné", width=400, anchor="w")
        self.pdf_label.pack(side="right", padx=(10, 20), pady=10)

        # Instructions frame
        instructions_frame = ctk.CTkFrame(main_frame)
        instructions_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        ctk.CTkLabel(instructions_frame, text="Instructions pour l'IA:", anchor="w").pack(side="top", fill="x", padx=10, pady=(10, 5))
        
        self.instruction_text = ctk.CTkTextbox(instructions_frame, height=120) 
        self.instruction_text.pack(fill="x", padx=10, pady=(0, 10))
        self.instruction_text.insert("1.0", templates[list(templates.keys())[0]])
        
        # Add some default instruction templates
        templates_frame = ctk.CTkFrame(instructions_frame)
        templates_frame.pack(fill="x", padx=10, pady=(0, 10))

        self.delete_instructions_button = ctk.CTkButton(templates_frame, text="Supprimer le texte", command=lambda: self.instruction_text.delete("1.0", "end"))
        self.delete_instructions_button.pack(side="left", padx=10, pady=5)

        apply_button = ctk.CTkButton(templates_frame, text="Appliquer la template", command=lambda: print("Apply template"))
        apply_button.pack(side="right", padx=10, pady=5)
        
        template_options = ["PV Attribution des comptes"]
        
        self.template_var = ctk.StringVar(value=template_options[0])
        template_dropdown = ctk.CTkOptionMenu(
            templates_frame, 
            values=template_options,
            variable=self.template_var,
            command=lambda x: print(f"Template selected: {x}")
        )
        template_dropdown.pack(side="right", padx=10, pady=5)

        ctk.CTkLabel(templates_frame, text="Templates:").pack(side="right", padx=(5, 10), pady=5)

        # Text areas section - now only for viewing current PDF and LLM output
        text_frame = ctk.CTkFrame(main_frame)
        text_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Configure grid for text areas
        text_frame.grid_columnconfigure(0, weight=1)
        text_frame.grid_columnconfigure(1, weight=1)
        text_frame.grid_rowconfigure(1, weight=1)
        
        # PDF text area
        custom_font = ("Helvetica", 10)
        ctk.CTkLabel(text_frame, text="Texte extrait des PDFs:").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 0))
        self.pdf_text = scrolledtext.ScrolledText(text_frame, wrap=tk.WORD, width=40, height=20, bg="#f0f0f0",font=custom_font)
        self.pdf_text.grid(row=1, column=0, sticky="nsew", padx=(10, 5), pady=10)
        
        # LLM output area
        ctk.CTkLabel(text_frame, text="Réponse de l'IA:").grid(row=0, column=1, sticky="w", padx=10, pady=(10, 0))
        self.llm_output = scrolledtext.ScrolledText(text_frame, wrap=tk.WORD, width=40, height=20, bg="#f0f0f0",font=custom_font)
        self.llm_output.grid(row=1, column=1, sticky="nsew", padx=(5, 10), pady=10)

        self.timer_running = False
        self.elapsed_seconds = 0
        self.timer_label = ctk.CTkLabel(self.main_container, text="⏱️ Temps écoulé: 00:00", anchor="e")
        self.timer_label.pack(side="bottom", fill="x", padx=10)
        
        # Status bar
        bold_font = ctk.CTkFont(family="Helvetica",weight="bold")
        # Initially set status to indicate model loading
        self.status_bar = ctk.CTkLabel(self, text="Chargement du modèle...", anchor="w", height=25,text_color="white",fg_color="orange",font=bold_font)
        self.status_bar.pack(side="bottom", fill="x", padx=10, pady=(0, 5))

    def update_timer(self):
        """Update the timer label every second when the timer is running."""
        if self.timer_running:
            self.elapsed_seconds += 1
            mins, secs = divmod(self.elapsed_seconds, 60)
            time_str = f"⏱️ Temps écoulé: {mins:02d}:{secs:02d}"
            self.timer_label.configure(text=time_str,text_color="red")
            # Store the reference to the scheduled callback
            self.timer_id = self.after(1000, self.update_timer)

    def start_timer(self):
        """Start the timer, reset elapsed seconds to 0, and start calling update_timer every second."""
        self.elapsed_seconds = 0
        self.timer_running = True
        self.timer_id = None  # Initialize timer_id
        self.update_timer()

    def stop_timer(self):
        """Stop the timer by setting timer_running to False and canceling the scheduled callback."""
        self.timer_running = False
        self.timer_label.configure(text_color="green")
        # Cancel the scheduled callback if it exists
        if self.timer_id is not None:
            self.after_cancel(self.timer_id)
            self.timer_id = None

    def delete_text(self):
        """Delete text from text areas"""
        self.instruction_text.delete("1.0", tk.END)

    def apply_template(self, choice):
        """Store the selected template"""
        self.template_var.set(choice)
    
    def apply_selected_template(self):
        """Apply the selected template to the instruction box"""
        
        template = self.template_var.get()        
        self.instruction_text.delete("1.0", tk.END)
        self.instruction_text.insert("1.0", templates.get(template, "Tu es un assistant IA. Réponds à cette demande de la manière la plus utile possible."))
        
    def browse_pdfs(self): # Changed to browse_pdfs
        """Open file dialog to select multiple PDF files"""
        file_paths = filedialog.askopenfilenames( # Changed to askopenfilenames
            title="Sélectionner les fichiers PDF",
            filetypes=[("Fichier PDF", "*.pdf"), ("Tous les fichiers", "*.*")]
        )
        
        if file_paths:
            self.pdf_paths = list(file_paths) # Store as a list
            filenames = [os.path.basename(p) for p in self.pdf_paths]
            self.pdf_label.configure(text=f"Selectionné: {len(filenames)} PDF(s)")
            self.save_summaries_button.configure(state="disabled") # Disable save until generated
            self.extracted_texts = {} # Clear previously extracted texts
            self.summaries = [] # Clear previous summaries
            self.pdf_text.delete(1.0, tk.END) # Clear display text
            self.llm_output.delete(1.0, tk.END) # Clear display text

            # Immediately extract text from all selected PDFs in a background thread
            threading.Thread(target=self._extract_all_texts_thread, daemon=True).start()
            self.status_bar.configure(text=f"Prêt à extraire {len(self.pdf_paths)} PDF(s).",text_color="white",fg_color="green")
    
    def _extract_all_texts_thread(self):
        """Extract text from all selected PDF files in a background thread."""
        self.status_bar.configure(text=f"Extraction du texte de {len(self.pdf_paths)} PDF(s)...")
        self.update_idletasks()

        self.extracted_texts = {} # Reset for new batch

        for i, pdf_path in enumerate(self.pdf_paths):
            filename = os.path.basename(pdf_path)
            self.status_bar.configure(text=f"Extraction du texte: {filename} ({i+1}/{len(self.pdf_paths)})...",text_color="white",fg_color="orange")
            self.update_idletasks()
            
            try:
                doc = fitz.open(pdf_path)
                text_parts = []
                for page_num in range(len(doc)):
                    pix = doc.load_page(page_num).get_pixmap()
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    text = pytesseract.image_to_string(img, lang="fra")
                    text_parts.append(text)
                
                self.extracted_texts[pdf_path] = "\n".join(text_parts)
                self.pdf_text.insert(tk.END,  "\n\n" + "\n".join(text_parts))

            except Exception as e:
                self.status_bar.configure(text=f"Erreur durant l'extraction de {filename}: {str(e)}",text_color="white",fg_color="red")
                self.extracted_texts[pdf_path] = f"[Extraction Error]: {str(e)}" # Store error for that PDF
        
        self.status_bar.configure(text=f"Extraction du texte de tous les PDF terminée ! Prêt à générer.",text_color="white",fg_color="green")
        self.generate_all_button.configure(state="normal", fg_color="#28a745", hover_color="#218838") # Ensure button is enabled after extraction

    def generate_all_summaries(self):
        """Initiate generation of summaries for all extracted texts."""
        if not self.extracted_texts:
            self.status_bar.configure(text="Aucun texte extrait à traiter.")
            return

        if self.is_generating:
            # If already generating, allow cancellation of the current batch
            self.stream_active = False # Signal to stop current streaming (if any)
            self.is_generating = False
            self.generate_all_button.configure(text="Générer tout", fg_color="#28a745", hover_color="#218838")
            self.status_bar.configure(text="Génération de la tâche annulée.",text_color="white",fg_color="red")
            self.stop_timer()
            return
        
        self.is_generating = True
        self.generate_all_button.configure(text="Annuler génération", fg_color="#dc3545", hover_color="#c82333")
        self.save_summaries_button.configure(state="disabled") # Disable save during generation
        self.llm_output.delete(1.0, tk.END) # Clear previous LLM output display
        self.summaries = [] # Clear previous summaries list

        self.status_bar.configure(text="Démarrage de la génération des résumés...",text_color="white",fg_color="orange")
        self.start_timer()

        # Start a new thread for batch processing
        threading.Thread(target=self._batch_generate_summaries_thread, daemon=True).start()

    def _batch_generate_summaries_thread(self):
        """Thread function to generate summaries for all PDFs."""
        system_prompt = self.instruction_text.get("1.0", tk.END).strip()
        if not system_prompt:
            self.status_bar.configure(text="Veuillez donner des instructions à l'IA.")
            self.is_generating = False
            self.generate_all_button.configure(text="Générer tout", fg_color="#28a745", hover_color="#218838")
            self.stop_timer()
            return

        for i, pdf_path in enumerate(self.pdf_paths):
            if not self.is_generating: # Check for cancellation
                break

            filename = os.path.basename(pdf_path)
            extracted_text = self.extracted_texts.get(pdf_path, "")

            if "[Extraction Error]" in extracted_text:
                self.status_bar.configure(text=f"Ignorant {filename} due à une erreur d'extraction.")
                self.summaries.append({
                    "filename": filename,
                    "summary": f"Erreur d'extraction: {extracted_text}",
                    "date": None # No date if extraction failed
                })
                continue

            if not extracted_text:
                self.status_bar.configure(text=f"Aucun texte extrait pour {filename}. Passant au suivant.")
                self.summaries.append({
                    "filename": filename,
                    "summary": "Aucun texte extrait.",
                    "date": None
                })
                continue

            self.status_bar.configure(text=f"Génération du résumé pour {filename} ({i+1}/{len(self.pdf_paths)})...",text_color="white",fg_color="orange")
            self.update_idletasks()

            try:
                # This part is similar to your original _generate_response_thread_local_llama
                # but it processes one PDF at a time and stores the result
                llm = self.llm
                temperature = self.temperature
                
                prompt = f"Voici le texte à traiter: {extracted_text} /no_think"

                # Non-streaming for batch processing to get full summary at once
                response = llm.create_chat_completion(
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=temperature,
                    max_tokens=10000,
                    stream=False
                )

                if response and "choices" in response and len(response["choices"]) > 0:
                    summary_content = response["choices"][0]["message"]["content"]
                    summary_content = summary_content.split('</think>')[-1].lstrip('\n') if '</think>' in summary_content else summary_content
                    
                    # Try to extract date from the summary
                    extracted_date = self._extract_date_from_summary(summary_content)

                    self.summaries.append({
                        "filename": filename,
                        "summary": summary_content,
                        "date": extracted_date # Store parsed date or None
                    })
                    # Update UI with the *current* PDF's summary
                    self.llm_output.insert(tk.END, "\n\n---\n\n" + summary_content)
                    self.llm_output.see(tk.END)
                    self.update_idletasks()

                else:
                    self.status_bar.configure(text=f"Erreur de réponse du modèle pour {filename}.",text_color="white",fg_color="red")
                    self.summaries.append({
                        "filename": filename,
                        "summary": "Erreur: Réponse vide du modèle.",
                        "date": None
                    })

            except Exception as e:
                self.status_bar.configure(text=f"Erreur durant la génération pour {filename}: {str(e)}",text_color="white",fg_color="red")
                self.summaries.append({
                    "filename": filename,
                    "summary": f"Erreur de génération: {str(e)}",
                    "date": None
                })
        
        self.stop_timer()
        self.is_generating = False
        self.generate_all_button.configure(text="Générer tout",fg_color=ctk.ThemeManager.theme["CTkButton"]["fg_color"][0], hover_color=ctk.ThemeManager.theme["CTkButton"]["hover_color"][0])
        self.save_summaries_button.configure(state="normal", fg_color="#28a745", hover_color="#218838") # Enable save button after all generations

        if self.summaries:
            self.status_bar.configure(text="Génération de tous les résumés terminée ! Prêt à sauvegarder.",text_color="white",fg_color="green")
        else:
            self.status_bar.configure(text="Aucun résumé généré.")


    def _extract_date_from_summary(self, summary_text):
        """
        Extracts a date in JJ/MM/AAAA or JJ MOIS AAAA format from the summary text.
        Assumes the date is prefixed with "Date : ".
        """
        # --- Pattern for JJ/MM/AAAA (Group 1) OR JJ MOIS AAAA (Group 2 and 3) ---
        # The MOIS part captures all French month names (uppercase)
        # Map of French month names to their numerical values
        _FRENCH_MONTH_MAP = {
            'JANVIER': 1, 'FÉVRIER': 2, 'MARS': 3, 'AVRIL': 4,
            'MAI': 5, 'JUIN': 6, 'JUILLET': 7, 'AOÛT': 8,
            'SEPTEMBRE': 9, 'OCTOBRE': 10, 'NOVEMBRE': 11, 'DÉCEMBRE': 12
        }
        month_names = "|".join(_FRENCH_MONTH_MAP.keys())
        
        # Regex to capture either format. 
        # Group 1: DD/MM/YYYY
        # Group 2: DD (for French format)
        # Group 3: MOIS (for French format)
        # Group 4: YYYY (for French format)
        regex = rf"Date\s*:\s*(?:(\d{{2}}/\d{{2}}/\d{{4}})|(\d{{2}})\s*({month_names})\s*(\d{{4}}))"

        match = re.search(regex, summary_text.upper()) # Use .upper() to ensure month names match the map
        
        if not match:
            return None

        # 1. Handle JJ/MM/AAAA format (captured in Group 1)
        if match.group(1):
            date_str = match.group(1)
            date_format = "%d/%m/%Y"
            
        # 2. Handle JJ MOIS AAAA format (captured in Groups 2, 3, and 4)
        elif match.group(2) and match.group(3) and match.group(4):
            day = match.group(2)
            month_name = match.group(3)
            year = match.group(4)
            
            # Convert month name to number
            month = self._FRENCH_MONTH_MAP.get(month_name)
            
            if not month:
                print(f"Error: Unknown month name '{month_name}' encountered.")
                return None
            
            # Reconstruct the date string into a standard format (e.g., '14/10/2025')
            date_str = f"{day}/{month:02d}/{year}" # :02d ensures '01', '02', etc.
            date_format = "%d/%m/%Y"
        
        else:
             return None # Should not happen, but for safety

        # Parse the date string into a datetime object
        try:
            return datetime.strptime(date_str, date_format)
        except ValueError:
            print(f"Could not parse date: {date_str}")
            return None

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

    def save_summaries(self):
        """Save the generated summaries to a JSON file and a Word document."""
        if not self.summaries:
            self.status_bar.configure(text="Aucun résumé à sauvegarder.")
            return

        # Save to Word Document
        word_output_filename = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx")],
            title="Sauvegarder les résumés en Word"
        )
        if word_output_filename:
            try:
                self.manage_all_summaries_in_document(word_output_filename,self.summaries)
                self.status_bar.configure(text=f"Résumés sauvegardés dans le Word: {os.path.basename(word_output_filename)}",text_color="white",fg_color="green")
                self.status_bar.configure(text="Sauvegarde terminée.",text_color="white",fg_color="green")
                self.open_file(word_output_filename)
            except Exception as e:
                self.status_bar.configure(text=f"Erreur lors de la sauvegarde du Word: {str(e)}",text_color="white",fg_color="red")

        else:
            self.status_bar.configure(text="Sauvegarde annulée.",text_color="white",fg_color="red")

    def open_file(self, file_path):
        try:
            os.startfile(file_path)
        except Exception as e:
            self.status_bar.configure(text=f"Erreur lors de l'ouverture du fichier: {str(e)}",text_color="white",fg_color="red")

if __name__ == "__main__":
    app = PDFToLLMApp()
    app.mainloop()


