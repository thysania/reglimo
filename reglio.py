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

class ChequeVirementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Générateur de Documents Bancaires")
        
        # Initialize all StringVars FIRST
        self.initialize_variables()
        
        # Then setup UI and other components
        self.setup_ui()
        self.load_config()
        self.register_fonts()
        
        # Data storage
        self.payee_db = None
        self.payee_list = []  # Store payee names for combobox  # ADD THIS LINE
        self.virements_db_path = None
        self.template_path = None
        
        # Layout configurations
        self.cheque_layout = self.load_layout_config("cheque")
        self.letter_layout = self.load_layout_config("letter")
        self.virement_layout = self.load_layout_config("virement")
        
        # Default cities  # ADD THESE LINES
        self.cities = ["Témara", "Rabat", "Casablanca", "Autre"]
        self.city_var.set("Témara")
        self.virement_city_var.set("Témara")
        self.letter_city_var.set("Témara")

    def initialize_variables(self):
        """Initialize all Tkinter variables before UI setup"""
        # Cheque Tab Variables
        self.payee_var = tk.StringVar()
        self.amount_var = tk.StringVar()
        self.amount_words_var = tk.StringVar()
        self.city_var = tk.StringVar()
        self.date_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        
        # Virement Tab Variables
        self.virement_payee_var = tk.StringVar()
        self.virement_amount_var = tk.StringVar()
        self.virement_amount_words_var = tk.StringVar()
        self.virement_type_var = tk.StringVar(value="Ordinaire")
        self.virement_motif_var = tk.StringVar()
        self.virement_rib_var = tk.StringVar()
        self.virement_bank_var = tk.StringVar()
        self.virement_city_var = tk.StringVar()
        
        # Lettre de Change Tab Variables
        self.letter_payee_var = tk.StringVar()
        self.letter_amount_var = tk.StringVar()
        self.letter_amount_words_var = tk.StringVar()
        self.letter_due_date_var = tk.StringVar()
        self.letter_city_var = tk.StringVar()
        self.letter_edition_date_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        self.letter_label_var = tk.StringVar()
        
        # Settings Tab Variables
        self.font_var = tk.StringVar(value="Arial")
        self.size_var = tk.IntVar(value=10)
        self.preview_text_var = tk.StringVar(value="Exemple de texte")

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
            clean_amount = amount_str.replace('#', '').replace(' ', '').replace(',', '.')
            amount = float(clean_amount)
            
            units = ["", "un", "deux", "trois", "quatre", "cinq", "six", "sept", "huit", "neuf"]
            teens = ["dix", "onze", "douze", "treize", "quatorze", "quinze", "seize", 
                    "dix-sept", "dix-huit", "dix-neuf"]
            tens = ["", "dix", "vingt", "trente", "quarante", "cinquante", 
                   "soixante", "soixante", "quatre-vingt", "quatre-vingt"]
        
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
                        return units[n // 100] + " cents"
                    else:
                        return units[n // 100] + " cent " + convert_less_than_one_thousand(n % 100)
            
            dirhams = int(amount)
            centimes = int(round(amount % 1 * 100))
        
            if dirhams == 0:
                words = "zéro"
            elif dirhams == 1:
                words = "un dirham"
            else:
                words = convert_less_than_one_thousand(dirhams) + " dirhams"
                
            if centimes > 0:
                if centimes == 1:
                    words += " et un centime"
                else:
                    words += " et " + convert_less_than_one_thousand(centimes) + " centimes"
                    
            return words.capitalize()
        except:
            return ""
    
    def filter_payees(self, event=None):
        """Filter payees based on combobox input"""
        typed = self.payee_var.get().lower()
        if not typed:
            self.payee_cb['values'] = self.payee_list
            return
        
        filtered = [p for p in self.payee_list if typed in p.lower()]
        self.payee_cb['values'] = filtered
        self.payee_cb.event_generate('<Down>')
    
    def update_amount_fields(self, event=None):
        """Update formatted amount and words when amount changes"""
        amount = self.amount_var.get()
        formatted = self.format_amount(amount)
        self.amount_var.set(formatted)
        
        words = self.amount_to_words(amount)
        self.amount_words_var.set(words)
        self.virement_amount_words_var.set(words)
        self.letter_amount_words_var.set(words)
    
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

    # ----------------------------
    # UI SETUP
    # ----------------------------
    def setup_ui(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Cheque Tab
        self.setup_cheque_tab()
        
        # Virement Tab
        self.setup_virement_tab()
        
        # Lettre de Change Tab
        self.setup_letter_tab()
        
        # Settings Tab
        self.setup_settings_tab()

    def setup_cheque_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Chèque")
        
        # Fields
        ttk.Label(tab, text="Bénéficiaire:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.payee_cb = ttk.Combobox(tab, textvariable=self.payee_var, values=self.payee_list)
        self.payee_cb.grid(row=0, column=1, sticky="ew")
        self.payee_cb.bind('<KeyRelease>', self.filter_payees)
        
        ttk.Label(tab, text="Montant (#000 000,00):").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.amount_entry = ttk.Entry(tab, textvariable=self.amount_var, width=20)
        self.amount_entry.grid(row=1, column=1, sticky="w")
        self.amount_entry.bind('<FocusOut>', self.update_amount_fields)
        
        ttk.Label(tab, text="Montant en lettres:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(tab, textvariable=self.amount_words_var, width=40, state='readonly').grid(row=2, column=1, sticky="w")
        
        ttk.Label(tab, text="Ville:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
        ttk.Combobox(tab, textvariable=self.city_var, values=self.cities).grid(row=3, column=1, sticky="w")
        
        ttk.Label(tab, text="Date (jj/mm/aaaa):").grid(row=4, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(tab, textvariable=self.date_var, width=20).grid(row=4, column=1, sticky="w")
        
        ttk.Button(tab, text="Générer Chèque", command=self.generate_cheque).grid(row=5, columnspan=2, pady=10)
        
        # Configure grid weights
        tab.grid_columnconfigure(1, weight=1)

    def filter_payees(self, event=None):
        """Filter payees based on combobox input"""
        typed = self.payee_var.get().lower()
        if not typed:
            self.payee_cb['values'] = self.payee_list
            return
    
        filtered = [p for p in self.payee_list if typed in p.lower()]
        self.payee_cb['values'] = filtered
        self.payee_cb.event_generate('<Down>')

    def update_amount_fields(self, event=None):
        """Update formatted amount and words when amount changes"""
        amount = self.amount_var.get()
        formatted = self.format_amount(amount)
        self.amount_var.set(formatted)
        
        words = self.amount_to_words(amount)
        self.amount_words_var.set(words)
        self.virement_amount_words_var.set(words)
        self.letter_amount_words_var.set(words)

    def setup_virement_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Virement")
        
        # Fields
        ttk.Label(tab, text="Bénéficiaire:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.virement_payee_cb = ttk.Combobox(tab, textvariable=self.virement_payee_var, values=self.payee_list)
        self.virement_payee_cb.grid(row=0, column=1, sticky="ew")
        self.virement_payee_cb.bind('<KeyRelease>', self.filter_payees)
        self.virement_payee_cb.bind('<<ComboboxSelected>>', self.fill_payee_details)
        
        ttk.Label(tab, text="Montant (#000 000,00):").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.virement_amount_entry = ttk.Entry(tab, textvariable=self.virement_amount_var, width=20)
        self.virement_amount_entry.grid(row=1, column=1, sticky="w")
        self.virement_amount_entry.bind('<FocusOut>', self.update_amount_fields)
        
        ttk.Label(tab, text="Montant en lettres:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(tab, textvariable=self.virement_amount_words_var, width=40, state='readonly').grid(row=2, column=1, sticky="w")
        
        ttk.Label(tab, text="Type:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
        ttk.Combobox(tab, textvariable=self.virement_type_var, values=["Ordinaire", "Instantané"], width=15).grid(row=3, column=1, sticky="w")
        
        ttk.Label(tab, text="Motif:").grid(row=4, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(tab, textvariable=self.virement_motif_var, width=20).grid(row=4, column=1, sticky="w")
        
        ttk.Label(tab, text="RIB:").grid(row=5, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(tab, textvariable=self.virement_rib_var, width=20).grid(row=5, column=1, sticky="w")
        
        ttk.Label(tab, text="Banque:").grid(row=6, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(tab, textvariable=self.virement_bank_var, width=20).grid(row=6, column=1, sticky="w")
        
        ttk.Label(tab, text="Ville:").grid(row=7, column=0, sticky="e", padx=5, pady=5)
        ttk.Combobox(tab, textvariable=self.virement_city_var, values=self.cities).grid(row=7, column=1, sticky="w")
        
        ttk.Button(tab, text="Générer Virement", command=self.generate_virement).grid(row=8, columnspan=2, pady=10)
        
        # Configure grid weights
        tab.grid_columnconfigure(1, weight=1)

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

    def setup_letter_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Lettre de Change")
        
        # Fields
        ttk.Label(tab, text="Bénéficiaire:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.letter_payee_cb = ttk.Combobox(tab, textvariable=self.letter_payee_var, values=self.payee_list)
        self.letter_payee_cb.grid(row=0, column=1, sticky="ew")
        self.letter_payee_cb.bind('<KeyRelease>', self.filter_payees)
        
        ttk.Label(tab, text="Montant (#000 000,00):").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.letter_amount_entry = ttk.Entry(tab, textvariable=self.letter_amount_var, width=20)
        self.letter_amount_entry.grid(row=1, column=1, sticky="w")
        self.letter_amount_entry.bind('<FocusOut>', self.update_amount_fields)
        
        ttk.Label(tab, text="Montant en lettres:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(tab, textvariable=self.letter_amount_words_var, width=40, state='readonly').grid(row=2, column=1, sticky="w")
        
        ttk.Label(tab, text="Date d'échéance:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(tab, textvariable=self.letter_due_date_var, width=20).grid(row=3, column=1, sticky="w")
        
        ttk.Label(tab, text="Ville:").grid(row=4, column=0, sticky="e", padx=5, pady=5)
        ttk.Combobox(tab, textvariable=self.letter_city_var, values=self.cities).grid(row=4, column=1, sticky="w")
        
        ttk.Label(tab, text="Date d'édition:").grid(row=5, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(tab, textvariable=self.letter_edition_date_var, width=20).grid(row=5, column=1, sticky="w")
        
        ttk.Label(tab, text="Libellé:").grid(row=6, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(tab, textvariable=self.letter_label_var, width=40).grid(row=6, column=1, sticky="w")
        
        ttk.Button(tab, text="Générer Lettre", command=self.generate_letter).grid(row=7, columnspan=2, pady=10)
        
        # Configure grid weights
        tab.grid_columnconfigure(1, weight=1)
        
        # Set default edition date
        self.letter_edition_date_var.set(datetime.now().strftime("%d/%m/%Y"))

    def setup_settings_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Paramètres")
        
        # File Import
        ttk.Button(tab, text="Importer Base Bénéficiaires", command=self.import_payee_db).pack(pady=5)
        ttk.Button(tab, text="Importer Fichier Virements", command=self.import_virements_db).pack(pady=5)
        ttk.Button(tab, text="Importer Modèle Word", command=self.import_template).pack(pady=5)
        
        # Font Preview
        self.setup_font_preview(tab)
    def import_payee_db(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if path:
            try:
                self.payee_db = pd.read_excel(path)
                self.payee_list = self.payee_db.iloc[:, 0].tolist()  # First column as payee names
                self.payee_cb['values'] = self.payee_list
                self.virement_payee_cb['values'] = self.payee_list
                self.letter_payee_cb['values'] = self.payee_list
                self.save_config('payee_db_path', path)
                messagebox.showinfo("Succès", f"Base chargée: {len(self.payee_db)} bénéficiaires")
            except Exception as e:
                messagebox.showerror("Erreur", f"Échec du chargement:\n{str(e)}")
    
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
            clean_amount = amount_str.replace('#', '').replace(' ', '').replace(',', '.')
            amount = float(clean_amount)
            
            units = ["", "un", "deux", "trois", "quatre", "cinq", "six", "sept", "huit", "neuf"]
            teens = ["dix", "onze", "douze", "treize", "quatorze", "quinze", "seize", 
                    "dix-sept", "dix-huit", "dix-neuf"]
            tens = ["", "dix", "vingt", "trente", "quarante", "cinquante", 
                   "soixante", "soixante", "quatre-vingt", "quatre-vingt"]
            
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
                        return units[n // 100] + " cents"
                    else:
                        return units[n // 100] + " cent " + convert_less_than_one_thousand(n % 100)
            
            dirhams = int(amount)
            centimes = int(round(amount % 1 * 100))
            
            if dirhams == 0:
                words = "zéro"
            elif dirhams == 1:
                words = "un dirham"
            else:
                words = convert_less_than_one_thousand(dirhams) + " dirhams"
                
            if centimes > 0:
                if centimes == 1:
                    words += " et un centime"
                else:
                    words += " et " + convert_less_than_one_thousand(centimes) + " centimes"
                    
            return words.capitalize()
        except:
            return ""
    # ----------------------------
    # DOCUMENT GENERATION
    # ----------------------------
    def generate_cheque(self):
        try:
            # Validate fields
            if not all([self.payee_var.get(), self.amount_var.get(), self.date_var.get()]):
                messagebox.showerror("Erreur", "Champs obligatoires manquants!")
                return
            
            output_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf")],
                title="Enregistrer le chèque"
            )
            if not output_path:
                return
            
            c = canvas.Canvas(output_path, pagesize=(210*mm, 99*mm))
            
            # Draw fields
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
            messagebox.showinfo("Succès", f"Chèque généré:\n{output_path}")
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Échec de génération:\n{str(e)}")

    def generate_virement(self):
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
            
            output_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf")],
                initialfile=f"Virement_{next_num}.pdf"
            )
            if not output_path:
                return
            
            c = canvas.Canvas(output_path, pagesize=A4)
            
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
            self.log_virement(next_num)
            messagebox.showinfo("Succès", f"Virement {next_num} généré!")
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Échec de génération:\n{str(e)}")

    def generate_letter(self):
        try:
            # Validate fields
            if not all([self.letter_payee_var.get(), self.letter_amount_var.get(), self.letter_due_date_var.get()]):
                messagebox.showerror("Erreur", "Champs obligatoires manquants!")
                return
            
            output_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf")],
                title="Enregistrer la lettre de change"
            )
            if not output_path:
                return
            
            c = canvas.Canvas(output_path, pagesize=A4)
            
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
            messagebox.showinfo("Succès", f"Lettre de change générée:\n{output_path}")
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Échec de génération:\n{str(e)}")

    # ----------------------------
    # SUPPORTING METHODS
    # ----------------------------
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

    # ----------------------------
    # INITIALIZATION & UTILITIES
    # ----------------------------
    def register_fonts(self):
        """Register system and custom fonts"""
        try:
            pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))
            pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))
        except:
            print("Arial fonts not found - using defaults")

    def load_layout_config(self, doc_type):
        """Load layout configuration for document type"""
        # In a real implementation, this would load from your Excel config
        # This is a simplified version with sample coordinates
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

    def setup_font_preview(self, parent):
        """Font preview panel setup"""
        preview_frame = ttk.LabelFrame(parent, text="Aperçu Police", padding=10)
        preview_frame.pack(fill=tk.X, pady=10)
        
        # Font selection
        ttk.Label(preview_frame, text="Police:").grid(row=0, column=0)
        self.font_var = tk.StringVar(value="Arial")
        ttk.Combobox(preview_frame, textvariable=self.font_var, 
                    values=list(pdfmetrics.getRegisteredFontNames())).grid(row=0, column=1)
        
        # Size
        ttk.Label(preview_frame, text="Taille:").grid(row=1, column=0)
        self.size_var = tk.IntVar(value=10)
        ttk.Spinbox(preview_frame, from_=8, to=24, textvariable=self.size_var).grid(row=1, column=1)
        
        # Preview text
        ttk.Label(preview_frame, text="Texte:").grid(row=2, column=0)
        self.preview_text_var = tk.StringVar(value="Exemple de texte")
        ttk.Entry(preview_frame, textvariable=self.preview_text_var).grid(row=2, column=1)
        
        # Preview button
        ttk.Button(preview_frame, text="Actualiser", command=self.update_font_preview).grid(row=3, columnspan=2)
        
        # Canvas for preview
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
            images = convert_from_path(pdf_path, dpi=96)
            return images[0] if images else Image.new('RGB', (400, 100), 'white')
        except:
            return Image.new('RGB', (400, 100), 'white')

    # ----------------------------
    # FILE OPERATIONS
    # ----------------------------
    def import_payee_db(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if path:
            try:
                self.payee_db = pd.read_excel(path)
                self.save_config('payee_db_path', path)
                messagebox.showinfo("Succès", f"Base chargée: {len(self.payee_db)} bénéficiaires")
            except Exception as e:
                messagebox.showerror("Erreur", f"Échec du chargement:\n{str(e)}")

    def import_virements_db(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if path:
            self.virements_db_path = path
            self.save_config('virements_db_path', path)
            messagebox.showinfo("Succès", "Fichier virements configuré")

    def import_template(self):
        path = filedialog.askopenfilename(filetypes=[("Word Files", "*.docm *.docx")])
        if path:
            self.template_path = path
            self.save_config('template_path', path)
            messagebox.showinfo("Succès", "Modèle importé")

    # ----------------------------
    # CONFIGURATION
    # ----------------------------
    def load_config(self):
        """Initialize all StringVars and load config"""
        # Initialize variables
        self.payee_var = tk.StringVar()
        self.amount_var = tk.StringVar()
        self.amount_words_var = tk.StringVar()
        self.city_var = tk.StringVar()
        self.date_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        
        # Virement variables
        self.virement_payee_var = tk.StringVar()
        self.virement_amount_var = tk.StringVar()
        self.virement_amount_words_var = tk.StringVar()
        self.virement_type_var = tk.StringVar(value="Ordinaire")
        self.virement_motif_var = tk.StringVar()
        self.virement_rib_var = tk.StringVar()
        self.virement_bank_var = tk.StringVar()
        self.virement_city_var = tk.StringVar()
        
        # Lettre de change variables
        self.letter_payee_var = tk.StringVar()
        self.letter_amount_var = tk.StringVar()
        self.letter_amount_words_var = tk.StringVar()
        self.letter_due_date_var = tk.StringVar()
        self.letter_city_var = tk.StringVar()
        self.letter_edition_date_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        self.letter_label_var = tk.StringVar()

    def save_config(self, key, value):
        """Save configuration (simplified)"""
        # In a real implementation, this would save to a config file
        pass

# ----------------------------
# RUN THE APPLICATION
# ----------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = ChequeVirementApp(root)
    root.mainloop()
