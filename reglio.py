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
        
        # Set French locale for number to words conversion
        locale.setlocale(locale.LC_ALL, 'fr_FR.UTF-8')
        
        # Initialize data structures first
        self.payee_db = None
        self.payee_list = []  # Initialize empty payee list
        self.virements_db_path = None
        self.template_path = None
        self.cities = ["Témara", "Rabat", "Casablanca", "Autre"]
        
        # Initialize all StringVars
        self.initialize_variables()
        
        # Then setup UI and other components
        self.setup_ui()
        self.load_config()
        self.register_fonts()
        
        # Layout configurations
        self.cheque_layout = self.load_layout_config("cheque")
        self.letter_layout = self.load_layout_config("letter")
        self.virement_layout = self.load_layout_config("virement")

    # [Previous methods remain the same until generate_cheque]

    def generate_cheque(self):
        try:
            # Validate fields
            if not all([self.payee_var.get(), self.amount_var.get(), self.date_var.get()]):
                messagebox.showerror("Erreur", "Champs obligatoires manquants!")
                return
            
            # Create a temporary PDF file
            temp_pdf = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            temp_pdf_path = temp_pdf.name
            temp_pdf.close()
            
            c = canvas.Canvas(temp_pdf_path, pagesize=(210*mm, 99*mm))
            
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
            
            # Show PDF preview instead of saving
            self.show_pdf_preview(temp_pdf_path)
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Échec de génération:\n{str(e)}")

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
                command=preview_window.destroy
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

    def amount_to_words(self, amount_str):
        """Convert numeric amount to French words"""
        try:
            # Remove formatting characters
            clean_amount = amount_str.replace('#', '').replace(' ', '').replace(',', '.')
            amount = float(clean_amount)
            
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

# [Rest of the class implementation remains the same]

if __name__ == "__main__":
    root = tk.Tk()
    app = ChequeVirementApp(root)
    root.mainloop()
