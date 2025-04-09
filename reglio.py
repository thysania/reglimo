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

# ----------------------------
# 1. MAIN APPLICATION CLASS
# ----------------------------
class ChequeVirementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Générateur de Chèques et Virements")
        self.setup_ui()
        self.load_config()
        
        # Data
        self.payee_db = None
        self.virements_db_path = None
        self.template_path = None
        
        # Register default fonts
        self.register_fonts()
    
    # ----------------------------
    # 2. UI SETUP
    # ----------------------------
    def setup_ui(self):
        # Notebook for document types
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Cheque Tab
        self.setup_cheque_tab()
        
        # Virement Tab
        self.setup_virement_tab()
        
        # Settings Tab
        self.setup_settings_tab()
    
    def setup_cheque_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Chèque")
        
        # Fields
        ttk.Label(tab, text="Bénéficiaire:").grid(row=0, column=0, sticky="e")
        self.payee_var = tk.StringVar()
        ttk.Combobox(tab, textvariable=self.payee_var).grid(row=0, column=1)
        
        ttk.Label(tab, text="Montant (#000 000,00):").grid(row=1, column=0, sticky="e")
        self.amount_var = tk.StringVar()
        ttk.Entry(tab, textvariable=self.amount_var).grid(row=1, column=1)
        
        ttk.Label(tab, text="Montant en lettres:").grid(row=2, column=0, sticky="e")
        self.amount_words_var = tk.StringVar()
        ttk.Entry(tab, textvariable=self.amount_words_var, width=40).grid(row=2, column=1)
        
        ttk.Label(tab, text="Ville:").grid(row=3, column=0, sticky="e")
        self.city_var = tk.StringVar()
        ttk.Entry(tab, textvariable=self.city_var).grid(row=3, column=1)
        
        ttk.Label(tab, text="Date (jj/mm/aaaa):").grid(row=4, column=0, sticky="e")
        self.date_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        ttk.Entry(tab, textvariable=self.date_var).grid(row=4, column=1)
        
        # Buttons
        ttk.Button(tab, text="Générer Chèque", command=self.generate_cheque).grid(row=5, columnspan=2)

    def setup_virement_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Virement")
        
        # Fields (similar structure to cheque tab)
        # ... (omitted for brevity) ...
        
        ttk.Button(tab, text="Générer Virement", command=self.generate_virement).grid(row=10, columnspan=2)

    def setup_settings_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Paramètres")
        
        # File Import
        ttk.Button(tab, text="Importer Base Bénéficiaires", command=self.import_payee_db).pack(pady=5)
        ttk.Button(tab, text="Importer Fichier Virements", command=self.import_virements_db).pack(pady=5)
        ttk.Button(tab, text="Importer Modèle Word", command=self.import_template).pack(pady=5)
        
        # Font Preview
        self.setup_font_preview(tab)

    def setup_font_preview(self, parent):
        preview_frame = ttk.LabelFrame(parent, text="Aperçu Police", padding=10)
        preview_frame.pack(fill=tk.X, pady=10)
        
        # Font selection
        ttk.Label(preview_frame, text="Police:").grid(row=0, column=0)
        self.font_var = tk.StringVar(value="Arial")
        ttk.Combobox(preview_frame, textvariable=self.font_var, values=["Arial", "Helvetica"]).grid(row=0, column=1)
        
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

    # ----------------------------
    # 3. CORE FUNCTIONALITY
    # ----------------------------
    def generate_cheque(self):
        # Validate fields
        if not all([self.payee_var.get(), self.amount_var.get(), self.date_var.get()]):
            messagebox.showerror("Erreur", "Champs obligatoires manquants!")
            return
        
        # Generate PDF
        try:
            output_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf")]
            )
            if not output_path:
                return
            
            c = canvas.Canvas(output_path, pagesize=(210*mm, 99*mm))
            
            # Draw fields using your coordinates
            self.draw_cheque_field(c, self.payee_var.get(), "payee")
            self.draw_cheque_field(c, f"#{self.amount_var.get()}", "amount")
            
            # Handle multi-line amount
            amount_lines = self.amount_words_var.get().split(" et ", 1)
            self.draw_cheque_field(c, amount_lines[0], "amount in letters line 1")
            if len(amount_lines) > 1:
                self.draw_cheque_field(c, "et " + amount_lines[1], "amount in letters line 2")
            
            c.save()
            messagebox.showinfo("Succès", f"Chèque généré:\n{output_path}")
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Échec de génération:\n{str(e)}")

    def draw_cheque_field(self, canvas, text, field_name):
        """Draw text at precise coordinates from spreadsheet"""
        config = self.cheque_layout[field_name]
        style = ParagraphStyle(
            name='Field',
            fontName=config['font'],
            fontSize=config['size'],
            alignment=config['align'].upper(),
            leading=config['size'] * 1.2
        )
        p = Paragraph(text, style)
        p.wrapOn(canvas, config['max_width'] * mm, 1000)
        y_pos = 99*mm - config['y'] * mm - p.height
        p.drawOn(canvas, config['x'] * mm, y_pos)

    # ----------------------------
    # 4. FONT & TEMPLATE HANDLING
    # ----------------------------
    def register_fonts(self):
        """Register system and custom fonts"""
        try:
            pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))
            pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))
        except:
            messagebox.showwarning("Police", "Polices Arial non trouvées")

    def update_font_preview(self):
        """Update the font preview canvas"""
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
        """Convert PDF page to PIL Image (requires pdftoppm)"""
        try:
            temp_img = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
            subprocess.run(["pdftoppm", "-png", "-singlefile", pdf_path, temp_img.name[:-4]])
            img = Image.open(temp_img.name)
            os.unlink(temp_img.name)
            return img
        except:
            return Image.new('RGB', (400, 100), 'white')

    # ----------------------------
    # 5. FILE IMPORT FUNCTIONS
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
    # 6. CONFIGURATION HANDLING
    # ----------------------------
    def load_config(self):
        config = ConfigParser()
        if os.path.exists('config.ini'):
            config.read('config.ini')
            if config.has_option('DEFAULT', 'payee_db_path'):
                self.import_payee_db(config.get('DEFAULT', 'payee_db_path'))
            # Load other configs...

    def save_config(self, key, value):
        config = ConfigParser()
        if os.path.exists('config.ini'):
            config.read('config.ini')
        config.set('DEFAULT', key, value)
        with open('config.ini', 'w') as f:
            config.write(f)

# ----------------------------
# RUN THE APPLICATION
# ----------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = ChequeVirementApp(root)
    root.mainloop()
