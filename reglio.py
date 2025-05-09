import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import pandas as pd
from datetime import datetime
from PIL import Image, ImageTk
import tempfile
import subprocess
import locale

class ChequeVirementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Générateur de Documents Bancaires")
        
        try:
            locale.setlocale(locale.LC_ALL, 'fr_FR.UTF-8')
        except:
            locale.setlocale(locale.LC_ALL, '')
        
        # Initialize data structures
        self.payee_db = None
        self.payee_list = []
        self.virements_db_path = None
        self.template_path = None
        self.cities = ["Témara", "Rabat", "Casablanca", "Autre"]
        
        # Initialize all variables
        self.initialize_variables()
        
        # Setup UI
        self.setup_ui()
        
        # Register fonts and load layouts
        self.register_fonts()
        self.cheque_layout = self.load_layout_config("cheque")
        self.letter_layout = self.load_layout_config("letter")
        self.virement_layout = self.load_layout_config("virement")

    def initialize_variables(self):
        """Initialize all Tkinter variables"""
        # Cheque Tab
        self.payee_var = tk.StringVar()
        self.amount_var = tk.StringVar()
        self.amount_words_var = tk.StringVar()
        self.city_var = tk.StringVar(value="Témara")
        self.date_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        
        # Virement Tab
        self.virement_payee_var = tk.StringVar()
        self.virement_amount_var = tk.StringVar()
        self.virement_amount_words_var = tk.StringVar()
        self.virement_type_var = tk.StringVar(value="Ordinaire")
        self.virement_motif_var = tk.StringVar()
        self.virement_rib_var = tk.StringVar()
        self.virement_bank_var = tk.StringVar()
        self.virement_city_var = tk.StringVar(value="Témara")
        
        # Letter Tab
        self.letter_payee_var = tk.StringVar()
        self.letter_amount_var = tk.StringVar()
        self.letter_amount_words_var = tk.StringVar()
        self.letter_due_date_var = tk.StringVar()
        self.letter_city_var = tk.StringVar(value="Témara")
        self.letter_edition_date_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        self.letter_label_var = tk.StringVar()
        
        # Settings Tab
        self.font_var = tk.StringVar(value="Arial")
        self.size_var = tk.IntVar(value=10)
        self.preview_text_var = tk.StringVar(value="Exemple de texte")

    def register_fonts(self):
        """Register required fonts"""
        try:
            pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))
            pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))
        except:
            print("Warning: Arial fonts not found - using defaults")

    def setup_ui(self):
        """Setup the main notebook interface"""
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        self.setup_cheque_tab()
        self.setup_virement_tab()
        self.setup_letter_tab()
        self.setup_settings_tab()

    def setup_cheque_tab(self):
        """Setup cheque tab widgets"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Chèque")
        
        # Payee field with autocomplete
        ttk.Label(tab, text="Bénéficiaire:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.payee_cb = ttk.Combobox(tab, textvariable=self.payee_var, values=self.payee_list)
        self.payee_cb.grid(row=0, column=1, sticky="ew")
        self.payee_cb.bind('<KeyRelease>', self.filter_payees)
        
        # Amount field with auto-formatting
        ttk.Label(tab, text="Montant (#000 000,00):").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.amount_entry = ttk.Entry(tab, textvariable=self.amount_var, width=20)
        self.amount_entry.grid(row=1, column=1, sticky="w")
        self.amount_entry.bind('<FocusOut>', self.update_amount_fields)
        
        # Amount in words (auto-generated)
        ttk.Label(tab, text="Montant en lettres:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(tab, textvariable=self.amount_words_var, width=40, state='readonly').grid(row=2, column=1, sticky="w")
        
        # City dropdown
        ttk.Label(tab, text="Ville:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
        ttk.Combobox(tab, textvariable=self.city_var, values=self.cities).grid(row=3, column=1, sticky="w")
        
        # Date field
        ttk.Label(tab, text="Date (jj/mm/aaaa):").grid(row=4, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(tab, textvariable=self.date_var, width=20).grid(row=4, column=1, sticky="w")
        
        # Generate button
        ttk.Button(tab, text="Générer Chèque", command=self.generate_cheque).grid(row=5, columnspan=2, pady=10)
        
        tab.grid_columnconfigure(1, weight=1)

    def setup_virement_tab(self):
        """Setup virement tab widgets"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Virement")
        
        # Payee field with autocomplete and auto-fill
        ttk.Label(tab, text="Bénéficiaire:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.virement_payee_cb = ttk.Combobox(tab, textvariable=self.virement_payee_var, values=self.payee_list)
        self.virement_payee_cb.grid(row=0, column=1, sticky="ew")
        self.virement_payee_cb.bind('<KeyRelease>', self.filter_payees)
        self.virement_payee_cb.bind('<<ComboboxSelected>>', self.fill_payee_details)
        
        # Amount field with auto-formatting
        ttk.Label(tab, text="Montant (#000 000,00):").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.virement_amount_entry = ttk.Entry(tab, textvariable=self.virement_amount_var, width=20)
        self.virement_amount_entry.grid(row=1, column=1, sticky="w")
        self.virement_amount_entry.bind('<FocusOut>', self.update_amount_fields)
        
        # Amount in words (auto-generated)
        ttk.Label(tab, text="Montant en lettres:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(tab, textvariable=self.virement_amount_words_var, width=40, state='readonly').grid(row=2, column=1, sticky="w")
        
        # Virement type dropdown
        ttk.Label(tab, text="Type:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
        ttk.Combobox(tab, textvariable=self.virement_type_var, values=["Ordinaire", "Instantané"], width=15).grid(row=3, column=1, sticky="w")
        
        # Motif field
        ttk.Label(tab, text="Motif:").grid(row=4, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(tab, textvariable=self.virement_motif_var, width=20).grid(row=4, column=1, sticky="w")
        
        # RIB field (auto-filled)
        ttk.Label(tab, text="RIB:").grid(row=5, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(tab, textvariable=self.virement_rib_var, width=20).grid(row=5, column=1, sticky="w")
        
        # Bank field (auto-filled)
        ttk.Label(tab, text="Banque:").grid(row=6, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(tab, textvariable=self.virement_bank_var, width=20).grid(row=6, column=1, sticky="w")
        
        # City dropdown
        ttk.Label(tab, text="Ville:").grid(row=7, column=0, sticky="e", padx=5, pady=5)
        ttk.Combobox(tab, textvariable=self.virement_city_var, values=self.cities).grid(row=7, column=1, sticky="w")
        
        # Generate button
        ttk.Button(tab, text="Générer Virement", command=self.generate_virement).grid(row=8, columnspan=2, pady=10)
        
        tab.grid_columnconfigure(1, weight=1)

    def setup_letter_tab(self):
        """Setup letter tab widgets"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Lettre de Change")
        
        # Payee field with autocomplete
        ttk.Label(tab, text="Bénéficiaire:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.letter_payee_cb = ttk.Combobox(tab, textvariable=self.letter_payee_var, values=self.payee_list)
        self.letter_payee_cb.grid(row=0, column=1, sticky="ew")
        self.letter_payee_cb.bind('<KeyRelease>', self.filter_payees)
        
        # Amount field with auto-formatting
        ttk.Label(tab, text="Montant (#000 000,00):").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.letter_amount_entry = ttk.Entry(tab, textvariable=self.letter_amount_var, width=20)
        self.letter_amount_entry.grid(row=1, column=1, sticky="w")
        self.letter_amount_entry.bind('<FocusOut>', self.update_amount_fields)
        
        # Amount in words (auto-generated)
        ttk.Label(tab, text="Montant en lettres:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(tab, textvariable=self.letter_amount_words_var, width=40, state='readonly').grid(row=2, column=1, sticky="w")
        
        # Due date field
        ttk.Label(tab, text="Date d'échéance:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(tab, textvariable=self.letter_due_date_var, width=20).grid(row=3, column=1, sticky="w")
        
        # City dropdown
        ttk.Label(tab, text="Ville:").grid(row=4, column=0, sticky="e", padx=5, pady=5)
        ttk.Combobox(tab, textvariable=self.letter_city_var, values=self.cities).grid(row=4, column=1, sticky="w")
        
        # Edition date field
        ttk.Label(tab, text="Date d'édition:").grid(row=5, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(tab, textvariable=self.letter_edition_date_var, width=20).grid(row=5, column=1, sticky="w")
        
        # Label field
        ttk.Label(tab, text="Libellé:").grid(row=6, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(tab, textvariable=self.letter_label_var, width=40).grid(row=6, column=1, sticky="w")
        
        # Generate button
        ttk.Button(tab, text="Générer Lettre", command=self.generate_letter).grid(row=7, columnspan=2, pady=10)
        
        tab.grid_columnconfigure(1, weight=1)

    def setup_settings_tab(self):
        """Setup settings tab widgets"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Paramètres")
        
        # File import buttons
        ttk.Button(tab, text="Importer Base Bénéficiaires", command=self.import_payee_db).pack(pady=5)
        ttk.Button(tab, text="Importer Fichier Virements", command=self.import_virements_db).pack(pady=5)
        ttk.Button(tab, text="Importer Modèle Word", command=self.import_template).pack(pady=5)
        
        # Font preview section
        self.setup_font_preview(tab)

    def setup_font_preview(self, parent):
        """Setup font preview panel"""
        preview_frame = ttk.LabelFrame(parent, text="Aperçu Police", padding=10)
        preview_frame.pack(fill=tk.X, pady=10)
        
        # Font selection
        ttk.Label(preview_frame, text="Police:").grid(row=0, column=0)
        font_cb = ttk.Combobox(preview_frame, textvariable=self.font_var, 
                              values=list(pdfmetrics.getRegisteredFontNames()))
        font_cb.grid(row=0, column=1)
        
        # Font size
        ttk.Label(preview_frame, text="Taille:").grid(row=1, column=0)
        ttk.Spinbox(preview_frame, from_=8, to=24, textvariable=self.size_var).grid(row=1, column=1)
        
        # Preview text
        ttk.Label(preview_frame, text="Texte:").grid(row=2, column=0)
        ttk.Entry(preview_frame, textvariable=self.preview_text_var).grid(row=2, column=1)
        
        # Update button
        ttk.Button(preview_frame, text="Actualiser", command=self.update_font_preview).grid(row=3, columnspan=2)
        
        # Preview canvas
        self.preview_canvas = tk.Canvas(preview_frame, width=400, height=100, bg='white')
        self.preview_canvas.grid(row=4, columnspan=2)

    def update_font_preview(self):
        """Update font preview canvas"""
        try:
            # Create temporary PDF
            temp_pdf = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            c = canvas.Canvas(temp_pdf.name, pagesize=(200, 100))
            
            # Draw sample text
            style = ParagraphStyle(
                name='Preview',
                fontName=self.font_var.get(),
                fontSize=self.size_var.get(),
                textColor=colors.black,
                leading=self.size_var.get() * 1.2
            )
            p = Paragraph(self.preview_text_var.get(), style)
            p.wrapOn(c, 180, 100)
            p.drawOn(c, 10, 80 - p.height)
            c.save()
            
            # Convert to image
            img = self.pdf_to_image(temp_pdf.name)
            os.unlink(temp_pdf.name)
            
            # Update canvas
            self.preview_canvas.delete("all")
            self.tk_img = ImageTk.PhotoImage(img)
            self.preview_canvas.create_image(0, 0, anchor=tk.NW, image=self.tk_img)
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Échec de l'aperçu:\n{str(e)}")

    def pdf_to_image(self, pdf_path):
        """Convert PDF to PIL Image"""
        try:
            from pdf2image import convert_from_path
            images = convert_from_path(pdf_path, dpi=100)
            return images[0] if images else Image.new('RGB', (400, 100), 'white')
        except:
            return Image.new('RGB', (400, 100), 'white')

    def filter_payees(self, event=None):
        """Filter payees based on combobox input"""
        typed = self.payee_var.get().lower()
        if not typed:
            self.payee_cb['values'] = self.payee_list
            return
        
        filtered = [p for p in self.payee_list if typed in p.lower()]
        self.payee_cb['values'] = filtered
        self.payee_cb.event_generate('<Down>')

    def fill_payee_details(self, event=None):
        """Fill RIB, bank, and city from payee database"""
        if self.payee_db is None:
            return
            
        payee = self.virement_payee_var.get()
        if not payee:
            return
            
        try:
            # Find payee in database (case insensitive)
            match = self.payee_db[self.payee_db.iloc[:, 0].str.lower() == payee.lower()]
            if not match.empty:
                # Assuming columns are: Payee, RIB, Bank, City
                if len(match.columns) > 1:
                    self.virement_rib_var.set(str(match.iloc[0, 1]))
                if len(match.columns) > 2:
                    self.virement_bank_var.set(str(match.iloc[0, 2]))
                if len(match.columns) > 3:
                    self.virement_city_var.set(str(match.iloc[0, 3]))
        except Exception as e:
            print(f"Error filling payee details: {e}")

    def format_amount(self, amount_str):
        """Format amount as #000 000,00"""
        try:
            # Remove any existing formatting
            clean_amount = amount_str.replace('#', '').replace(' ', '').replace(',', '.')
            amount = float(clean_amount)
            
            # Format with thousands separator and 2 decimals
            formatted = f"#{amount:,.2f}".replace(".", "X").replace(",", " ").replace("X", ",")
            return formatted
        except:
            return amount_str

    def amount_to_words(self, amount_str):
        """Convert numeric amount to French words"""
        try:
            # Remove formatting characters and handle empty input
            if not amount_str:
                return ""
                
            clean_amount = amount_str.replace('#', '').replace(' ', '').replace(',', '.')
            
            # Handle cases where amount is not a valid number
            try:
                amount = float(clean_amount)
            except ValueError:
                return ""
            
            # Handle zero case
            if amount == 0:
                return "Zéro dirham"
            
            # Split into dirhams and centimes
            dirhams = int(amount)
            centimes = int(round((amount - dirhams) * 100))
            
            # French number words
            units = ["", "un", "deux", "trois", "quatre", "cinq", 
                    "six", "sept", "huit", "neuf"]
            teens = ["dix", "onze", "douze", "treize", "quatorze", 
                    "quinze", "seize", "dix-sept", "dix-huit", "dix-neuf"]
            tens = ["", "dix", "vingt", "trente", "quarante", 
                   "cinquante", "soixante", "soixante", "quatre-vingt", "quatre-vingt"]
            
            def convert_less_than_one_thousand(n):
                if n < 10:
                    return units[n]
                elif n < 20:
                    return teens[n - 10]
                elif n < 100:
                    if n % 10 == 0:
                        return tens[n // 10]
                    elif n // 10 == 7 or n // 10 == 9:
                        return tens[n // 10] + "-" + teens[n % 10]
                    else:
                        return tens[n // 10] + "-" + units[n % 10]
                else:
                    if n % 100 == 0:
                        return units[n // 100] + " cent"
                    else:
                        return units[n // 100] + " cent " + convert_less_than_one_thousand(n % 100)
            
            # Convert dirhams
            if dirhams == 1:
                words = "un dirham"
            else:
                words = convert_less_than_one_thousand(dirhams) + " dirhams"
            
            # Convert centimes if any
            if centimes > 0:
                if centimes == 1:
                    words += " et un centime"
                else:
                    words += " et " + convert_less_than_one_thousand(centimes) + " centimes"
            
            # Capitalize first letter
            return words.capitalize()
            
        except Exception as e:
            print(f"Error converting amount to words: {e}")
            return ""

    def get_layout_config(self, field_name):
        """Get layout configuration for a specific field"""
        # First check cheque layout
        if field_name in self.cheque_layout:
            return self.cheque_layout[field_name]
        # Then check virement layout
        elif field_name in self.virement_layout:
            return self.virement_layout[field_name]
        # Finally check letter layout
        elif field_name in self.letter_layout:
            return self.letter_layout[field_name]
        return None

    def load_layout_config(self, doc_type):
        """Load layout configuration for document type"""
        if doc_type == "cheque":
            return {
                "payee": {"x": 35, "y": 45, "max_width": 130, "font": "Arial", "size": 10, "align": "left"},
                "amount": {"x": 130, "y": 73, "max_width": 31, "font": "Arial", "size": 10, "align": "left"},
                "amount in letters line 1": {"x": 40, "y": 56, "max_width": 100, "font": "Arial", "size": 10, "align": "center"},
                "amount in letters line 2": {"x": 10, "y": 51, "max_width": 160, "font": "Arial", "size": 10, "align": "center"},
                "ville": {"x": 75, "y": 38, "max_width": 35, "font": "Arial", "size": 10, "align": "left"},
                "date": {"x": 130, "y": 38, "max_width": 30, "font": "Arial", "size": 10, "align": "left"}
            }
        elif doc_type == "letter":
            return {
                "amount": {"x": 155, "y": 84, "max_width": 40, "font": "Arial", "size": 10, "align": "center"},
                "amount in letters line 1": {"x": 148, "y": 62, "max_width": 48, "font": "Arial", "size": 10, "align": "left"},
                "amount in letters line 2": {"x": 148, "y": 58, "max_width": 48, "font": "Arial", "size": 10, "align": "left"},
                "amount in letters line 3": {"x": 148, "y": 54, "max_width": 48, "font": "Arial", "size": 10, "align": "left"},
                "payee 1": {"x": 85, "y": 71, "max_width": 110, "font": "Arial", "size": 10, "align": "left"},
                "payee 2": {"x": 7, "y": 69, "max_width": 55, "font": "Arial", "size": 10, "align": "left"},
                "due date": {"x": 155, "y": 94, "max_width": 40, "font": "Arial", "size": 10, "align": "center"},
                "motif": {"x": 85, "y": 56, "max_width": 55, "font": "Arial", "size": 10, "align": "left"},
                "city and edition date": {"x": 85, "y": 62, "max_width": 55, "font": "Arial", "size": 10, "align": "left"}
            }
        elif doc_type == "virement":
            return {
                "virement_num": {"x": 50, "y": 250, "max_width": 80, "font": "Arial-Bold", "size": 12, "align": "left"},
                "amount": {"x": 50, "y": 230, "max_width": 60, "font": "Arial", "size": 12, "align": "left"},
                "amount in letters line 1": {"x": 50, "y": 210, "max_width": 140, "font": "Arial", "size": 10, "align": "left"},
                "amount in letters line 2": {"x": 50, "y": 200, "max_width": 140, "font": "Arial", "size": 10, "align": "left"},
                "payee": {"x": 50, "y": 180, "max_width": 120, "font": "Arial", "size": 10, "align": "left"},
                "type": {"x": 50, "y": 160, "max_width": 60, "font": "Arial", "size": 10, "align": "left"},
                "motif": {"x": 50, "y": 140, "max_width": 120, "font": "Arial", "size": 10, "align": "left"},
                "rib": {"x": 50, "y": 120, "max_width": 100, "font": "Arial", "size": 10, "align": "left"},
                "bank": {"x": 50, "y": 100, "max_width": 100, "font": "Arial", "size": 10, "align": "left"},
                "city": {"x": 50, "y": 80, "max_width": 80, "font": "Arial", "size": 10, "align": "left"}
            }
        return {}

    def update_amount_fields(self, event=None):
        """Update formatted amount and words when amount changes"""
        try:
            amount = self.amount_var.get()
            if not amount:
                return
                
            formatted = self.format_amount(amount)
            self.amount_var.set(formatted)
            
            words = self.amount_to_words(amount)
            if words:  # Only update if conversion succeeded
                self.amount_words_var.set(words)
                self.virement_amount_words_var.set(words)
                self.letter_amount_words_var.set(words)
        except Exception as e:
            print(f"Error updating amount fields: {e}")

    def generate_cheque(self):
        """Generate cheque PDF and show preview"""
        try:
            # Validate fields
            if not all([self.payee_var.get(), self.amount_var.get(), self.date_var.get()]):
                messagebox.showerror("Erreur", "Champs obligatoires manquants!")
                return
            
            # Create temporary PDF
            temp_pdf = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            temp_pdf_path = temp_pdf.name
            temp_pdf.close()
            
            c = canvas.Canvas(temp_pdf_path, pagesize=(210*mm, 99*mm))
            
            # Draw fields using layout config
            self.draw_field(c, self.payee_var.get(), "payee", 99)
            self.draw_field(c, f"#{self.amount_var.get()}", "amount", 99)
            
            # Amount in words (multi-line)
            amount_lines = self.amount_words_var.get().split(" et ", 1)
            self.draw_field(c, amount_lines[0], "amount in letters line 1", 99)
            if len(amount_lines) > 1:
                self.draw_field(c, "et " + amount_lines[1], "amount in letters line 2", 99)
            
            self.draw_field(c, self.city_var.get(), "ville", 99)
            self.draw_field(c, self.date_var.get(), "date", 99)
            
            c.save()
            
            # Show preview instead of saving
            self.show_pdf_preview(temp_pdf_path)
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Échec de génération:\n{str(e)}")

    def generate_virement(self):
        """Generate virement PDF and show preview"""
        try:
            # Auto-format the amount
            self.update_amount_fields()
            
            # Validate fields
            if not all([self.virement_payee_var.get(), self.virement_amount_var.get()]):
                messagebox.showerror("Erreur", "Bénéficiaire et montant sont obligatoires!")
                return
            
            # Auto-numbering
            current_year = datetime.now().year
            last_num = self.get_last_virement_number()
            next_num = f"{current_year}/{(int(last_num.split('/')[1]) + 1):03d}" if last_num and str(current_year) in last_num else f"{current_year}/001"
            
            # Create temporary PDF
            temp_pdf = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            temp_pdf_path = temp_pdf.name
            temp_pdf.close()
            
            c = canvas.Canvas(temp_pdf_path, pagesize=A4)
            
            # Draw fields
            self.draw_field(c, f"VIR {next_num}", "virement_num", 297)
            self.draw_field(c, f"#{self.virement_amount_var.get()}", "amount", 297)
            
            # Amount in words (multi-line)
            amount_lines = self.split_amount_text(self.virement_amount_words_var.get())
            for i, line in enumerate(amount_lines[:3]):  # Max 3 lines
                self.draw_field(c, line, f"amount in letters line {i+1}", 297)
            
            self.draw_field(c, self.virement_payee_var.get(), "payee", 297)
            self.draw_field(c, self.virement_type_var.get(), "type", 297)
            self.draw_field(c, self.virement_motif_var.get(), "motif", 297)
            self.draw_field(c, self.virement_rib_var.get(), "rib", 297)
            self.draw_field(c, self.virement_bank_var.get(), "bank", 297)
            self.draw_field(c, self.virement_city_var.get(), "city", 297)
            
            c.save()
            
            # Log virement and show preview
            self.log_virement(next_num)
            self.show_pdf_preview(temp_pdf_path)
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Échec de génération:\n{str(e)}")

    def generate_letter(self):
        """Generate letter PDF and show preview"""
        try:
            # Validate fields
            if not all([self.letter_payee_var.get(), self.letter_amount_var.get(), self.letter_due_date_var.get()]):
                messagebox.showerror("Erreur", "Champs obligatoires manquants!")
                return
            
            # Create temporary PDF
            temp_pdf = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            temp_pdf_path = temp_pdf.name
            temp_pdf.close()
            
            c = canvas.Canvas(temp_pdf_path, pagesize=A4)
            
            # Draw fields
            self.draw_field(c, f"#{self.letter_amount_var.get()}", "amount", 297)
            
            # Amount in words (multi-line)
            amount_lines = self.split_amount_text(self.letter_amount_words_var.get())
            for i, line in enumerate(amount_lines[:3]):  # Max 3 lines
                self.draw_field(c, line, f"amount in letters line {i+1}", 297)
            
            self.draw_field(c, self.letter_payee_var.get(), "payee 1", 297)
            self.draw_field(c, self.letter_due_date_var.get(), "due date", 297)
            self.draw_field(c, f"{self.letter_city_var.get()}, le {self.letter_edition_date_var.get()}", 
                          "city and edition date", 297)
            self.draw_field(c, self.letter_label_var.get(), "label", 297)
            
            c.save()
            
            # Show preview
            self.show_pdf_preview(temp_pdf_path)
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Échec de génération:\n{str(e)}")

    def draw_field(self, canvas, text, field_name, doc_height_mm):
        """Generic field drawing method"""
        config = self.get_layout_config(field_name)
        if not config:
            return
            
        style = ParagraphStyle(
            name='Field',
            fontName=config['font'],
            fontSize=config['size'],
            alignment=config['align'].upper(),
            leading=config['size'] * 1.2
        )
        p = Paragraph(text, style)
        p.wrapOn(canvas, config['max_width'] * mm, 1000)
        y_pos = doc_height_mm*mm - config['y'] * mm - p.height
        p.drawOn(canvas, config['x'] * mm, y_pos)

    def split_amount_text(self, text, max_line_length=60):
        """Smart text splitting for amount in words"""
        if len(text) <= max_line_length:
            return [text]
        
        # Try to split at natural breaks
        for delimiter in [" et ", ", "]:
            if delimiter in text:
                return text.split(delimiter)
        
        # Hard split if no natural break found
        return [text[i:i+max_line_length] for i in range(0, len(text), max_line_length)]

    def get_last_virement_number(self):
        """Get last virement number from Excel log"""
        if not hasattr(self, 'virements_db_path') or not self.virements_db_path:
            return None
            
        try:
            df = pd.read_excel(self.virements_db_path, sheet_name="VIREMENTS")
            if not df.empty and "ORDER_DE_VIR" in df.columns:
                return df["ORDER_DE_VIR"].iloc[-1]
            return None
        except:
            return None

    def log_virement(self, virement_num):
        """Record virement in Excel log"""
        if not hasattr(self, 'virements_db_path') or not self.virements_db_path:
            return
            
        try:
            new_data = {
                "DATE": [datetime.now().strftime("%d/%m/%Y")],
                "ORDER_DE_VIR": [virement_num],
                "FOURNISSEUR": [self.virement_payee_var.get()],
                "MONTANT": [self.virement_amount_var.get()],
                "MONTANT_EN_LETTRES": [self.virement_amount_words_var.get()],
                "TYPE_VIR": [self.virement_type_var.get()],
                "RIB": [self.virement_rib_var.get()],
                "BANQUE": [self.virement_bank_var.get()],
                "VILLE": [self.virement_city_var.get()]
            }
            
            df = pd.DataFrame(new_data)
            
            try:
                with pd.ExcelWriter(self.virements_db_path, engine='openpyxl', mode='a', 
                                   if_sheet_exists='overlay') as writer:
                    df.to_excel(writer, sheet_name="VIREMENTS", 
                               startrow=writer.sheets["VIREMENTS"].max_row, 
                               header=False, index=False)
            except:
                # Create new file if doesn't exist
                df.to_excel(self.virements_db_path, sheet_name="VIREMENTS", index=False)
                
        except Exception as e:
            messagebox.showwarning("Attention", f"Virement non enregistré:\n{str(e)}")

    def import_payee_db(self):
        """Import payee database from Excel"""
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if path:
            try:
                self.payee_db = pd.read_excel(path)
                self.payee_list = self.payee_db.iloc[:, 0].tolist()  # First column as payee names
                
                # Update all comboboxes
                self.payee_cb['values'] = self.payee_list
                self.virement_payee_cb['values'] = self.payee_list
                self.letter_payee_cb['values'] = self.payee_list
                
                messagebox.showinfo("Succès", f"Base chargée: {len(self.payee_db)} bénéficiaires")
            except Exception as e:
                messagebox.showerror("Erreur", f"Échec du chargement:\n{str(e)}")

    def import_virements_db(self):
        """Import virements database"""
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if path:
            self.virements_db_path = path
            messagebox.showinfo("Succès", "Fichier virements configuré")

    def import_template(self):
        """Import Word template"""
        path = filedialog.askopenfilename(filetypes=[("Word Files", "*.docm *.docx")])
        if path:
            self.template_path = path
            messagebox.showinfo("Succès", "Modèle importé")

    def load_layout_config(self, doc_type):
        """Load layout configuration for document type"""
        if doc_type == "cheque":
            return {
                "payee": {"x": 35, "y": 45, "max_width": 130, "font": "Arial", "size": 10, "align": "left"},
                "amount": {"x": 130, "y": 73, "max_width": 31, "font": "Arial", "size": 10, "align": "left"},
                "amount in letters line 1": {"x": 40, "y": 56, "max_width": 100, "font": "Arial", "size": 10, "align": "center"},
                "amount in letters line 2": {"x": 10, "y": 51, "max_width": 160, "font": "Arial", "size": 10, "align": "center"},
                "ville": {"x": 75, "y": 38, "max_width": 35, "font": "Arial", "size": 10, "align": "left"},
                "date": {"x": 130, "y": 38, "max_width": 30, "font": "Arial", "size": 10, "align": "left"}
            }
        elif doc_type == "letter":
            return {
                "amount": {"x": 155, "y": 84, "max_width": 40, "font": "Arial", "size": 10, "align": "center"},
                "amount in letters line 1": {"x": 148, "y": 62, "max_width": 48, "font": "Arial", "size": 10, "align": "left"},
                "amount in letters line 2": {"x": 148, "y": 58, "max_width": 48, "font": "Arial", "size": 10, "align": "left"},
                "amount in letters line 3": {"x": 148, "y": 54, "max_width": 48, "font": "Arial", "size": 10, "align": "left"},
                "payee 1": {"x": 85, "y": 71, "max_width": 110, "font": "Arial", "size": 10, "align": "left"},
                "payee 2": {"x": 7, "y": 69, "max_width": 55, "font": "Arial", "size": 10, "align": "left"},
                "due date": {"x": 155, "y": 94, "max_width": 40, "font": "Arial", "size": 10, "align": "center"},
                "motif": {"x": 85, "y": 56, "max_width": 55, "font": "Arial", "size": 10, "align": "left"},
                "city and edition date": {"x": 85, "y": 62, "max_width": 55, "font": "Arial", "size": 10, "align": "left"}
            }
        elif doc_type == "virement":
            return {
                "virement_num": {"x": 50, "y": 250, "max_width": 80, "font": "Arial-Bold", "size": 12, "align": "left"},
                "amount": {"x": 50, "y": 230, "max_width": 60, "font": "Arial", "size": 12, "align": "left"},
                "amount in letters line 1": {"x": 50, "y": 210, "max_width": 140, "font": "Arial", "size": 10, "align": "left"},
                "amount in letters line 2": {"x": 50, "y": 200, "max_width": 140, "font": "Arial", "size": 10, "align": "left"},
                "payee": {"x": 50, "y": 180, "max_width": 120, "font": "Arial", "size": 10, "align": "left"},
                "type": {"x": 50, "y": 160, "max_width": 60, "font": "Arial", "size": 10, "align": "left"},
                "motif": {"x": 50, "y": 140, "max_width": 120, "font": "Arial", "size": 10, "align": "left"},
                "rib": {"x": 50, "y": 120, "max_width": 100, "font": "Arial", "size": 10, "align": "left"},
                "bank": {"x": 50, "y": 100, "max_width": 100, "font": "Arial", "size": 10, "align": "left"},
                "city": {"x": 50, "y": 80, "max_width": 80, "font": "Arial", "size": 10, "align": "left"}
            }
        return {}

    def show_pdf_preview(self, pdf_path):
        """Display PDF preview in a new window"""
        try:
            from pdf2image import convert_from_path
            images = convert_from_path(pdf_path, dpi=100)
            
            preview_window = tk.Toplevel(self.root)
            preview_window.title("Aperçu du Document")
            
            # Convert first page to ImageTk
            img = images[0]
            img_tk = ImageTk.PhotoImage(img)
            
            # Display image
            label = tk.Label(preview_window, image=img_tk)
            label.image = img_tk  # Keep reference
            label.pack()
            
            # Add print button
            print_btn = ttk.Button(
                preview_window, 
                text="Imprimer", 
                command=lambda: self.print_pdf(pdf_path)
            )
            print_btn.pack(pady=10)
            
            # Add close button
            close_btn = ttk.Button(
                preview_window, 
                text="Fermer", 
                command=lambda: [preview_window.destroy(), os.unlink(pdf_path)]
            )
            close_btn.pack(pady=5)
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'afficher l'aperçu:\n{str(e)}")
            os.unlink(pdf_path)

    def print_pdf(self, pdf_path):
        """Print the PDF file"""
        try:
            if os.name == 'nt':  # Windows
                os.startfile(pdf_path, "print")
            else:  # MacOS and Linux
                subprocess.run(["lp", pdf_path])
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'imprimer:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ChequeVirementApp(root)
    root.mainloop()
