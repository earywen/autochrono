"""
AutoChrono - Interface graphique CustomTkinter
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
from vba_generator import VBAGenerator


class AutoChronoApp(ctk.CTk):
    """Fenêtre principale de l'application AutoChrono."""
    
    def __init__(self):
        super().__init__()
        
        self.title("AutoChrono - Générateur VBA Outlook")
        self.geometry("600x700")
        self.minsize(500, 600)
        
        # Variables
        self.trigram_var = ctk.StringVar()
        self.chrono_folder_var = ctk.StringVar()
        self.chrono_file_var = ctk.StringVar()
        self.col_chrono_var = ctk.StringVar(value="A")
        self.col_client_var = ctk.StringVar(value="B")
        self.col_trigram_var = ctk.StringVar(value="C")
        
        self._create_widgets()
    
    def _create_widgets(self):
        """Crée tous les widgets de l'interface."""
        
        # Titre
        title_label = ctk.CTkLabel(
            self,
            text="AutoChrono",
            font=ctk.CTkFont(size=28, weight="bold")
        )
        title_label.pack(pady=(30, 5))
        
        subtitle_label = ctk.CTkLabel(
            self,
            text="Générateur de module VBA pour Outlook",
            font=ctk.CTkFont(size=14),
            text_color="gray"
        )
        subtitle_label.pack(pady=(0, 30))
        
        # Frame principal
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=30, pady=(0, 30))
        
        # Section: Informations utilisateur
        self._create_section_header(main_frame, "👤 Informations utilisateur")
        
        self._create_input_row(
            main_frame,
            "Trigramme :",
            self.trigram_var,
            placeholder="Ex: LABR"
        )
        
        # Section: Chemins Chrono
        self._create_section_header(main_frame, "📁 Chemins Chrono")
        
        self._create_path_row(
            main_frame,
            "Dossier Chrono :",
            self.chrono_folder_var,
            is_folder=True
        )
        
        self._create_path_row(
            main_frame,
            "Fichier Excel :",
            self.chrono_file_var,
            is_folder=False
        )
        
        # Section: Colonnes Excel
        self._create_section_header(main_frame, "📊 Colonnes Excel")
        
        columns_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        columns_frame.pack(fill="x", padx=20, pady=(5, 15))
        
        col_labels = [
            ("N° Chrono", self.col_chrono_var),
            ("Client", self.col_client_var),
            ("Trigramme", self.col_trigram_var)
        ]
        
        for i, (label_text, var) in enumerate(col_labels):
            col_frame = ctk.CTkFrame(columns_frame, fg_color="transparent")
            col_frame.pack(side="left", expand=True, fill="x", padx=5)
            
            ctk.CTkLabel(col_frame, text=label_text, font=ctk.CTkFont(size=12)).pack()
            ctk.CTkEntry(
                col_frame,
                textvariable=var,
                width=60,
                justify="center",
                font=ctk.CTkFont(size=14)
            ).pack(pady=5)
        
        # Boutons
        buttons_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        buttons_frame.pack(fill="x", padx=20, pady=30)
        
        generate_btn = ctk.CTkButton(
            buttons_frame,
            text="🚀 Générer le module VBA",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=50,
            command=self._generate_vba
        )
        generate_btn.pack(fill="x", pady=(0, 10))
        
        instructions_btn = ctk.CTkButton(
            buttons_frame,
            text="📖 Instructions d'import Outlook",
            font=ctk.CTkFont(size=14),
            height=40,
            fg_color="transparent",
            border_width=2,
            command=self._show_instructions
        )
        instructions_btn.pack(fill="x")
    
    def _create_section_header(self, parent, text):
        """Crée un en-tête de section."""
        header = ctk.CTkLabel(
            parent,
            text=text,
            font=ctk.CTkFont(size=16, weight="bold"),
            anchor="w"
        )
        header.pack(fill="x", padx=20, pady=(20, 10))
    
    def _create_input_row(self, parent, label_text, variable, placeholder=""):
        """Crée une ligne avec label et champ de saisie."""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", padx=20, pady=5)
        
        ctk.CTkLabel(frame, text=label_text, width=120, anchor="w").pack(side="left")
        ctk.CTkEntry(
            frame,
            textvariable=variable,
            placeholder_text=placeholder,
            width=300
        ).pack(side="left", fill="x", expand=True, padx=(10, 0))
    
    def _create_path_row(self, parent, label_text, variable, is_folder=True):
        """Crée une ligne avec label, champ et bouton parcourir."""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", padx=20, pady=5)
        
        ctk.CTkLabel(frame, text=label_text, width=120, anchor="w").pack(side="left")
        
        entry = ctk.CTkEntry(frame, textvariable=variable, width=250)
        entry.pack(side="left", fill="x", expand=True, padx=(10, 10))
        
        browse_btn = ctk.CTkButton(
            frame,
            text="Parcourir",
            width=80,
            command=lambda: self._browse_path(variable, is_folder)
        )
        browse_btn.pack(side="left")
    
    def _browse_path(self, variable, is_folder):
        """Ouvre le dialogue de sélection de fichier/dossier."""
        if is_folder:
            path = filedialog.askdirectory(title="Sélectionner le dossier Chrono")
        else:
            path = filedialog.askopenfilename(
                title="Sélectionner le fichier Excel Chrono",
                filetypes=[("Fichiers Excel", "*.xlsx *.xls"), ("Tous les fichiers", "*.*")]
            )
        
        if path:
            variable.set(path)
    
    def _validate_inputs(self):
        """Valide les entrées utilisateur."""
        errors = []
        
        if not self.trigram_var.get().strip():
            errors.append("Le trigramme est requis")
        
        if not self.chrono_folder_var.get().strip():
            errors.append("Le dossier Chrono est requis")
        
        if not self.chrono_file_var.get().strip():
            errors.append("Le fichier Excel est requis")
        
        if not self.col_chrono_var.get().strip():
            errors.append("La colonne N° Chrono est requise")
        
        if not self.col_client_var.get().strip():
            errors.append("La colonne Client est requise")
        
        if not self.col_trigram_var.get().strip():
            errors.append("La colonne Trigramme est requise")
        
        return errors
    
    def _generate_vba(self):
        """Génère le fichier VBA."""
        errors = self._validate_inputs()
        
        if errors:
            messagebox.showerror("Erreur de validation", "\n".join(errors))
            return
        
        # Demander où sauvegarder
        output_path = filedialog.asksaveasfilename(
            title="Enregistrer le module VBA",
            defaultextension=".bas",
            filetypes=[("Module VBA", "*.bas"), ("Tous les fichiers", "*.*")],
            initialfile="AutoChrono.bas"
        )
        
        if not output_path:
            return
        
        try:
            generator = VBAGenerator(
                trigram=self.trigram_var.get().strip(),
                chrono_folder=self.chrono_folder_var.get().strip(),
                chrono_file=self.chrono_file_var.get().strip(),
                col_chrono=self.col_chrono_var.get().strip().upper(),
                col_client=self.col_client_var.get().strip().upper(),
                col_trigram=self.col_trigram_var.get().strip().upper()
            )
            
            generator.generate(output_path)
            
            messagebox.showinfo(
                "Succès ! 🎉",
                f"Module VBA généré avec succès !\n\n{output_path}\n\nCliquez sur 'Instructions d'import' pour savoir comment l'importer dans Outlook."
            )
        
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de la génération :\n{str(e)}")
    
    def _show_instructions(self):
        """Affiche les instructions d'import."""
        instructions = """
📖 INSTRUCTIONS D'IMPORT DANS OUTLOOK

1. Ouvrir Outlook

2. Appuyer sur Alt + F11 pour ouvrir l'éditeur VBA

3. Dans le menu : Fichier → Importer un fichier...

4. Sélectionner le fichier AutoChrono.bas généré

5. Fermer l'éditeur VBA (Ctrl + Q)

6. C'est prêt ! À chaque envoi de mail contenant 
   "REF : ... - N°XXXXX", une fenêtre vous proposera 
   de ranger le mail automatiquement.

⚠️ IMPORTANT :
- Si Outlook bloque les macros, aller dans :
  Fichier → Options → Centre de gestion de la confidentialité
  → Paramètres → Paramètres des macros
  → Sélectionner "Notifications pour toutes les macros"
"""
        
        # Créer une fenêtre popup
        popup = ctk.CTkToplevel(self)
        popup.title("Instructions d'import")
        popup.geometry("500x450")
        popup.transient(self)
        popup.grab_set()
        
        text = ctk.CTkTextbox(popup, font=ctk.CTkFont(size=13))
        text.pack(fill="both", expand=True, padx=20, pady=20)
        text.insert("1.0", instructions)
        text.configure(state="disabled")
        
        close_btn = ctk.CTkButton(popup, text="Fermer", command=popup.destroy)
        close_btn.pack(pady=(0, 20))
