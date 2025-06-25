import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import win32com.client
import os
from pathlib import Path
import sqlite3
from datetime import datetime
from tkinter import ttk, messagebox
import locale
from flask import Flask, render_template, jsonify, send_from_directory
from flask_cors import CORS
import threading
import webbrowser
import time
import sys

# Imports pour la sauvegarde Excel
try:
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    OPENPYXL_AVAILABLE = True
except ImportError:
    print("Pour un meilleur formatage Excel, installez openpyxl: pip install openpyxl")
    OPENPYXL_AVAILABLE = False

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class WebServer:
    def __init__(self):
        self.app = Flask(__name__, 
            static_folder='static',    
            template_folder='templates'
        )
        
        cors = CORS(self.app)
        self.app.config['CORS_HEADERS'] = 'Content-Type'
        self.workbook = None
        self.cached_data = None 
        self.lock = threading.Lock()
        self.main_thread_id = threading.current_thread().ident
        
        self.setup_routes()

    def setup_routes(self):
        @self.app.route('/')
        def index():
            return render_template('index.html')

        @self.app.route('/api/data')
        def get_data():
            if not self.workbook:
                return jsonify({"error": "No workbook loaded"}), 400
            
            try:
                with self.lock:
                    if self.cached_data:
                        return jsonify(self.cached_data)
                    return jsonify({"error": "No data available"}), 404
                    
            except Exception as e:
                return jsonify({"error": str(e)}), 500
                
        @self.app.route('/api/refresh', methods=['POST'])
        def refresh_data():
            """Force le rafraîchissement des données"""
            try:
                self.update_cache()
                return jsonify({"status": "success", "message": "Données actualisées"})
            except Exception as e:
                return jsonify({"status": "error", "message": str(e)}), 500

        @self.app.after_request 
        def after_request(response):
            response.headers.add('Access-Control-Allow-Origin', '*')
            response.headers.add('Access-Control-Allow-Headers', 'Content-Type,Authorization')
            response.headers.add('Access-Control-Allow-Methods', 'GET,PUT,POST,DELETE,OPTIONS')
            return response

    def format_number(self, value):
        """Formatte et arrondit les nombres"""
        try:
            if isinstance(value, str) and value == "-":
                return "-"
            if value is None:
                return "0"
            num = float(value)
            rounded = round(num)
            return f"{rounded:,}".replace(",", " ")
        except:
            return str(value)

    def get_regime_values(self, regime, sheet):
        try:
            # Utilisation de la feuille web au lieu de synthese
            web_sheet = self.workbook.Sheets("web")
            
            # Mapping des régimes vers leurs lignes dans la feuille web
            regime_mapping = {
                "micro nu": 2,           # Ligne 2 : micro nu + meublé
                "micro meublé": 2,       # Ligne 2 : micro nu + meublé  
                "micro classé": 2,       # Ligne 2 : micro nu + meublé
                "SCI IS": 3,             # Ligne 3 : SCI IS
                "SCI IS PREL BONI": 4,   # Ligne 4 : SCI IS PREL BONI
                "SCI IR": 5,             # Ligne 5 : SCI IR
                "LMNP": 2,               # Ligne 2 : micro nu + meublé (approximation)
                "LMNP CGA": 2            # Ligne 2 : micro nu + meublé (approximation)
            }
            
            row_num = regime_mapping.get(regime.lower().replace(" ", ""), 2)
            
            # Lecture du coût global depuis la feuille web (colonne B)
            cost_global = web_sheet.Range(f"B{row_num}").Value or 0
            
            return [
                regime,
                cost_global / 40 if cost_global > 0 else 0,  # coût moyen annuel (40 ans)
                cost_global / 20 if cost_global > 0 else 0,  # coût moyen annuel (durée estimation)
                cost_global,  # coût global (40 ans)
                cost_global,  # coût global (durée de détention)
                0,  # fiscalité plus value (40 ans) - à ajuster selon vos besoins
                0,  # fiscalité plus value (durée de détention) - à ajuster selon vos besoins
                cost_global  # coût global total
            ]

        except Exception as e:
            print(f"Erreur dans get_regime_values pour {regime}: {str(e)}")
            return [regime] + [0] * 7

    def update_cache(self):
        try:
            if not self.workbook:
                return

            sheet = self.workbook.Sheets("feuil1")
            web_sheet = self.workbook.Sheets("web")

            # Récupération des données d'entrée
            with self.lock:
                input_data = {
                    "Prix d'acquisition": f"{self.format_number(sheet.Range('c4').Value)} €",
                    "Travaux": f"{self.format_number(sheet.Range('b51').Value)} €",
                    "TF + charges loc.": f"{self.format_number(sheet.Range('b25').Value)} €",
                    "Assurance": f"{self.format_number(sheet.Range('b26').Value)} €",
                    "Loyer mensuel": f"{self.format_number(sheet.Range('c1').Value)} €",
                    "Emprunt": f"{self.format_number(sheet.Range('b39').Value)} €",
                    "Durée emprunt": f"{self.format_number(sheet.Range('b40').Value)} ans",
                    "Taux emprunt": f"{self.format_number(sheet.Range('b43').Value)} %",
                    "TMI perso./physique": f"{self.format_number(sheet.Range('c6').Value)} %",
                    "Prix de cession": f"{self.format_number(sheet.Range('b47').Value)} €",
                    "Durée détention": f"{self.format_number(sheet.Range('c3').Value)} ans",
                    "Prél. prix de cession": sheet.Range("c2").Value,
                    "CGA": sheet.Range("f3").Value
                }

                fiscal_data = []
                regimes = [
                    "micro nu", "micro meublé", "micro classé", "SCI IS",
                    "SCI IS PREL BONI", "SCI IR", "LMNP", "LMNP CGA"
                ]

                for regime in regimes:
                    values = self.get_regime_values(regime, web_sheet)
                    fiscal_data.append({
                        "regime": regime,
                        "cout_moyen_40": f"{self.format_number(values[1])} €",
                        "cout_moyen_detention": f"{self.format_number(values[2])} €",
                        "cout_global_40": f"{self.format_number(values[3])} €",
                        "cout_global_detention": f"{self.format_number(values[4])} €",
                        "fisc_plus_value_40": f"{self.format_number(values[5])} €",
                        "fisc_plus_value_detention": f"{self.format_number(values[6])} €",
                        "cout_global": f"{self.format_number(values[7])} €"
                    })

                self.cached_data = {
                    "input_data": input_data,
                    "fiscal_data": fiscal_data
                }

            print("Cache mis à jour avec succès")
            
        except Exception as e:
            print(f"Erreur lors de la mise à jour du cache: {str(e)}")

    def set_workbook(self, workbook):
        print("Setting workbook")
        self.workbook = workbook
        self.update_cache()
        print("Workbook set successfully")

    def run(self):
        print("Démarrage du serveur web...")
        try:
            from werkzeug.serving import run_simple
            run_simple('0.0.0.0', 5000, self.app, use_reloader=False, threaded=True)
        except Exception as e:
            print(f"Erreur lors du démarrage du serveur web: {str(e)}")


class DataEntryForm(tk.Toplevel):
    def __init__(self, parent, workbook):
        super().__init__(parent)
        self.title("Saisie des Données - Analyse Fiscale Immobilière")
        self.workbook = workbook
        
        # Configuration de la fenêtre
        self.geometry("1000x800")  # Augmenté pour voir tous les éléments
        self.configure(bg="#f8f9fa")
        self.resizable(True, True)
        
        # Variables pour les couleurs
        self.COLORS = {
            'primary': "#2c3e50",
            'secondary': "#34495e", 
            'accent': "#3498db",
            'success': "#27ae60",
            'danger': "#e74c3c",
            'warning': "#f39c12",
            'info': "#17a2b8",
            'light': "#ecf0f1",
            'white': "#ffffff",
            'text_dark': "#2c3e50",
            'text_light': "#7f8c8d",
            'border': "#bdc3c7"
        }
        
        # Dictionnaire pour stocker les entrées
        self.entries = {}
        
        # Configuration principale de la fenêtre
        self.rowconfigure(0, weight=0)  # Header fixe
        self.rowconfigure(1, weight=1)  # Contenu principal extensible
        self.rowconfigure(2, weight=0)  # Footer fixe
        self.columnconfigure(0, weight=1)
        
        # Création de l'interface
        self.create_header()
        self.create_main_content()
        self.create_footer()
        
        # Centrer la fenêtre
        self.center_window()
        
        # Focus sur la fenêtre
        self.lift()
        self.focus_force()
        
    def create_header(self):
        """Création de l'en-tête professionnel - VERSION CORRIGÉE"""
        # Frame header avec hauteur fixe
        header_frame = tk.Frame(self, bg=self.COLORS['primary'], height=100)
        header_frame.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        header_frame.grid_propagate(False)  # Important pour maintenir la hauteur
        header_frame.columnconfigure(0, weight=1)
        
        # Conteneur interne pour le contenu
        header_content = tk.Frame(header_frame, bg=self.COLORS['primary'])
        header_content.pack(fill=tk.BOTH, expand=True, padx=30, pady=15)
        
        # Titre principal
        title_label = tk.Label(header_content,
                              text="📊 Formulaire de Saisie",
                              font=('Segoe UI', 22, 'bold'),
                              bg=self.COLORS['primary'],
                              fg=self.COLORS['white'])
        title_label.pack(anchor="w")
        
        # Sous-titre
        subtitle_label = tk.Label(header_content,
                                 text="Analyse Fiscale et Optimisation d'Investissement Immobilier",
                                 font=('Segoe UI', 12),
                                 bg=self.COLORS['primary'],
                                 fg=self.COLORS['light'])
        subtitle_label.pack(anchor="w", pady=(5, 0))
        
    def create_main_content(self):
        """Création du contenu principal avec sections organisées - VERSION CORRIGÉE"""
        # Frame principal pour le contenu
        main_frame = tk.Frame(self, bg="#f8f9fa")
        main_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=20)
        main_frame.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        # Canvas et scrollbar pour le défilement
        canvas = tk.Canvas(main_frame, bg="#f8f9fa", highlightthickness=0)
        scrollbar = tk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#f8f9fa")
        
        # Configuration du scrolling
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Placement du canvas et scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Configuration du scrollable_frame
        scrollable_frame.columnconfigure(0, weight=1)
        scrollable_frame.columnconfigure(1, weight=1)
        
        # Création des sections de formulaire
        self.create_form_sections(scrollable_frame)
        
        # Binding pour le scroll avec la molette
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        self.bind_all("<MouseWheel>", _on_mousewheel)
        
    def create_form_sections(self, parent):
        """Création des sections organisées du formulaire"""
        # Conteneur principal pour les sections
        sections_container = tk.Frame(parent, bg="#f8f9fa")
        sections_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        sections_container.columnconfigure(0, weight=1)
        sections_container.columnconfigure(1, weight=1)
        
        # Section 1: Informations sur le bien
        section1 = self.create_section(sections_container, "🏠 Informations sur le Bien", 0, 0)
        self.add_form_fields(section1, [
            ("Prix d'acquisition", "€", "c4", "💰"),
            ("Travaux", "€", "b51", "🔨"),
            ("TF + charges loc.", "€", "b25", "📋"),
            ("Assurance", "€", "b26", "🛡️")
        ])
        
        # Section 2: Revenus et financement
        section2 = self.create_section(sections_container, "💵 Revenus et Financement", 0, 1)
        self.add_form_fields(section2, [
            ("Loyer mensuel", "€", "c1", "🏠"),
            ("Emprunt", "€", "b39", "🏦"),
            ("Durée emprunt", "années", "b40", "📅"),
            ("Taux emprunt", "%", "b43", "📈")
        ])
        
        # Section 3: Fiscalité
        section3 = self.create_section(sections_container, "📊 Paramètres Fiscaux", 1, 0)
        self.add_form_fields(section3, [
            ("TMI perso./physique", "%", "c6", "💼"),
            ("Prix de cession", "€", "b47", "💰"),
            ("Durée détention", "années", "c3", "⏰")
        ])
        
        # Section 4: Options
        section4 = self.create_section(sections_container, "⚙️ Options Avancées", 1, 1)
        self.add_form_fields(section4, [
            ("Prél. prix de cession", "OUI/NON", "c2", "✅"),
            ("CGA", "OUI/NON", "f3", "📋")
        ])
    
    def create_section(self, parent, title, row, col):
        """Création d'une section avec titre"""
        # Frame principale de la section
        section_frame = tk.Frame(parent, bg=self.COLORS['white'], relief="solid", bd=1)
        section_frame.grid(row=row, column=col, sticky="nsew", padx=10, pady=10)
        section_frame.columnconfigure(0, weight=1)
        
        # En-tête de section
        header_frame = tk.Frame(section_frame, bg=self.COLORS['accent'], height=50)
        header_frame.grid(row=0, column=0, sticky="ew")
        header_frame.grid_propagate(False)
        header_frame.columnconfigure(0, weight=1)
        
        # Titre de la section
        title_label = tk.Label(header_frame,
                              text=title,
                              font=('Segoe UI', 14, 'bold'),
                              bg=self.COLORS['accent'],
                              fg=self.COLORS['white'])
        title_label.pack(pady=12)
        
        # Corps de la section
        body_frame = tk.Frame(section_frame, bg=self.COLORS['white'])
        body_frame.grid(row=1, column=0, sticky="ew", padx=15, pady=15)
        body_frame.columnconfigure(1, weight=1)
        
        return body_frame
    
    def add_form_fields(self, parent, fields):
        """Ajout des champs de formulaire dans une section"""
        for i, (label, unit, cell, icon) in enumerate(fields):
            # Frame pour chaque champ
            field_frame = tk.Frame(parent, bg=self.COLORS['white'])
            field_frame.grid(row=i, column=0, columnspan=3, sticky="ew", pady=5)
            field_frame.columnconfigure(1, weight=1)
            
            # Icône
            icon_label = tk.Label(field_frame,
                                 text=icon,
                                 font=('Segoe UI', 14),
                                 bg=self.COLORS['white'],
                                 fg=self.COLORS['accent'],
                                 width=3)
            icon_label.grid(row=0, column=0, padx=(0, 8))
            
            # Label du champ
            label_widget = tk.Label(field_frame,
                                   text=label,
                                   font=('Segoe UI', 10, 'bold'),
                                   bg=self.COLORS['white'],
                                   fg=self.COLORS['text_dark'],
                                   anchor="w")
            label_widget.grid(row=0, column=1, sticky="ew", padx=(0, 8))
            
            # Frame pour l'entrée et l'unité
            entry_frame = tk.Frame(field_frame, bg=self.COLORS['white'])
            entry_frame.grid(row=0, column=2, sticky="e")
            
            # Champ de saisie avec style moderne
            if unit in ["OUI/NON"]:
                # Combobox pour les champs OUI/NON
                entry = ttk.Combobox(entry_frame,
                                   values=["OUI", "NON"],
                                   font=('Segoe UI', 10),
                                   width=12,
                                   state="readonly")
                entry.set("NON")  # Valeur par défaut
            else:
                entry = tk.Entry(entry_frame,
                               font=('Segoe UI', 10),
                               bg=self.COLORS['white'],
                               fg=self.COLORS['text_dark'],
                               relief="solid",
                               bd=1,
                               width=12,
                               justify="right",
                               highlightthickness=1,
                               highlightcolor=self.COLORS['accent'],
                               highlightbackground=self.COLORS['border'])
            
            entry.grid(row=0, column=0, padx=(0, 5), ipady=5)
            self.entries[cell] = entry
            
            # Label d'unité
            unit_label = tk.Label(entry_frame,
                                 text=unit,
                                 font=('Segoe UI', 10, 'bold'),
                                 bg=self.COLORS['white'],
                                 fg=self.COLORS['text_light'],
                                 width=6)
            unit_label.grid(row=0, column=1)
            
            # Effet hover sur les champs
            self.add_hover_effect(entry)
    
    def add_hover_effect(self, widget):
        """Ajout d'effet hover sur les widgets"""
        def on_enter(event):
            if isinstance(widget, tk.Entry):
                widget.configure(bg="#f8f9fa")
        
        def on_leave(event):
            if isinstance(widget, tk.Entry):
                widget.configure(bg=self.COLORS['white'])
        
        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)
    
    def create_footer(self):
        """Création du pied de page avec boutons d'action - VERSION CORRIGÉE"""
        # Frame footer avec hauteur fixe
        footer_frame = tk.Frame(self, bg=self.COLORS['light'], height=100)
        footer_frame.grid(row=2, column=0, sticky="ew", padx=0, pady=0)
        footer_frame.grid_propagate(False)  # Important pour maintenir la hauteur
        footer_frame.columnconfigure(0, weight=1)
        
        # Conteneur des boutons centré
        button_container = tk.Frame(footer_frame, bg=self.COLORS['light'])
        button_container.pack(expand=True)  # Centrage vertical et horizontal
        
        # Bouton Valider
        validate_btn = tk.Button(button_container,
                               text="✅ Valider et Enregistrer",
                               command=self.validate_form,
                               font=('Segoe UI', 11, 'bold'),
                               bg=self.COLORS['success'],
                               fg=self.COLORS['white'],
                               relief="flat",
                               padx=25,
                               pady=10,
                               cursor="hand2",
                               activebackground="#229954",
                               activeforeground=self.COLORS['white'])
        validate_btn.pack(side=tk.LEFT, padx=10)
        
        # Bouton Sauvegarder Simulation
        save_sim_btn = tk.Button(button_container,
                               text="💾 Sauvegarder Simulation",
                               command=self.save_simulation,
                               font=('Segoe UI', 11, 'bold'),
                               bg=self.COLORS['info'],
                               fg=self.COLORS['white'],
                               relief="flat",
                               padx=25,
                               pady=10,
                               cursor="hand2",
                               activebackground="#138496",
                               activeforeground=self.COLORS['white'])
        save_sim_btn.pack(side=tk.LEFT, padx=10)
        
        # Bouton Annuler
        cancel_btn = tk.Button(button_container,
                             text="❌ Annuler",
                             command=self.destroy,
                             font=('Segoe UI', 11, 'bold'),
                             bg=self.COLORS['danger'],
                             fg=self.COLORS['white'],
                             relief="flat",
                             padx=25,
                             pady=10,
                             cursor="hand2",
                             activebackground="#c0392b",
                             activeforeground=self.COLORS['white'])
        cancel_btn.pack(side=tk.LEFT, padx=10)
        
        # Bouton Reset
        reset_btn = tk.Button(button_container,
                            text="🔄 Réinitialiser",
                            command=self.reset_form,
                            font=('Segoe UI', 11, 'bold'),
                            bg=self.COLORS['warning'],
                            fg=self.COLORS['white'],
                            relief="flat",
                            padx=25,
                            pady=10,
                            cursor="hand2",
                            activebackground="#e67e22",
                            activeforeground=self.COLORS['white'])
        reset_btn.pack(side=tk.LEFT, padx=10)
        
        # Ajout d'effets hover sur les boutons
        self.add_button_effects(validate_btn, "#229954", self.COLORS['success'])
        self.add_button_effects(save_sim_btn, "#138496", self.COLORS['info'])
        self.add_button_effects(cancel_btn, "#c0392b", self.COLORS['danger'])
        self.add_button_effects(reset_btn, "#e67e22", self.COLORS['warning'])
    
    def add_button_effects(self, button, hover_color, normal_color):
        """Ajout d'effets visuels sur les boutons"""
        def on_enter(event):
            button.configure(bg=hover_color)
        
        def on_leave(event):
            button.configure(bg=normal_color)
        
        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)
    
    def reset_form(self):
        """Réinitialisation du formulaire"""
        for entry in self.entries.values():
            if isinstance(entry, ttk.Combobox):
                entry.set("NON")
            else:
                entry.delete(0, tk.END)
    
    def validate_form(self):
        """Validation du formulaire avec mise à jour web"""
        try:
            sheet = self.workbook.Sheets("feuil1")
            
            # Validation des champs obligatoires
            required_fields = ["c4", "c1", "c3"]  # Prix acquisition, loyer, durée
            missing_fields = []
            
            for cell in required_fields:
                entry = self.entries[cell]
                value = entry.get().strip()
                if not value:
                    missing_fields.append(cell)
            
            if missing_fields:
                messagebox.showwarning("Champs obligatoires", 
                                     "Veuillez remplir tous les champs obligatoires:\n" +
                                     "- Prix d'acquisition\n- Loyer mensuel\n- Durée de détention")
                return
            
            # Mise à jour des cellules Excel
            for cell, entry in self.entries.items():
                value = entry.get().strip()
                if value:
                    try:
                        # Tentative de conversion en nombre
                        if value.upper() not in ["OUI", "NON"]:
                            value = float(value.replace(" ", "").replace(",", "."))
                    except ValueError:
                        pass
                    sheet.Range(cell).Value = value
            
            # Activation de la feuille synthèse
            self.workbook.Sheets("synthese").Activate()
            
            # Mise à jour du cache web via la méthode update_cache
            try:
                main_window = self.master
                if hasattr(main_window, 'web_server'):
                    main_window.web_server.update_cache()
                    print("Cache web mis à jour avec succès")
                
                # Mise à jour de l'interface principale
                if hasattr(main_window, 'update_data_tree'):
                    main_window.update_data_tree()
                if hasattr(main_window, 'refresh_input_summary'):
                    main_window.refresh_input_summary()
                if hasattr(main_window, 'refresh_fiscal_summary'):
                    main_window.refresh_fiscal_summary()
                    
            except Exception as e:
                print(f"Erreur lors de la mise à jour: {str(e)}")
            
            # Message de succès avec style
            self.show_success_message()
            
        except Exception as e:
            print(f"Erreur lors de la validation: {str(e)}")
            messagebox.showerror("Erreur", f"Erreur lors de la validation: {str(e)}")
    
    def show_success_message(self):
        """Affichage d'un message de succès stylé"""
        success_window = tk.Toplevel(self)
        success_window.title("Succès")
        success_window.geometry("400x200")
        success_window.configure(bg=self.COLORS['white'])
        success_window.resizable(False, False)
        
        # Centrer la fenêtre de succès
        success_window.transient(self)
        success_window.grab_set()
        
        # Contenu du message
        content_frame = tk.Frame(success_window, bg=self.COLORS['white'])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)
        
        # Icône de succès
        icon_label = tk.Label(content_frame,
                             text="✅",
                             font=('Segoe UI', 48),
                             bg=self.COLORS['white'],
                             fg=self.COLORS['success'])
        icon_label.pack(pady=(0, 20))
        
        # Message
        message_label = tk.Label(content_frame,
                               text="Données enregistrées avec succès!",
                               font=('Segoe UI', 14, 'bold'),
                               bg=self.COLORS['white'],
                               fg=self.COLORS['text_dark'])
        message_label.pack()
        
        # Bouton OK
        ok_btn = tk.Button(content_frame,
                          text="OK",
                          command=lambda: [success_window.destroy(), self.destroy()],
                          font=('Segoe UI', 12, 'bold'),
                          bg=self.COLORS['success'],
                          fg=self.COLORS['white'],
                          relief="flat",
                          padx=30,
                          pady=10,
                          cursor="hand2")
        ok_btn.pack(pady=(20, 0))
        
        # Auto-fermeture après 2 secondes
        success_window.after(2000, lambda: [success_window.destroy(), self.destroy()])

    # === MÉTHODES DE SAUVEGARDE EXCEL ===
    
    def save_simulation(self):
        """Sauvegarde dans un fichier Excel dédié aux simulations"""
        try:
            if not self.workbook:
                messagebox.showwarning("Attention", "Aucun classeur Excel ouvert!")
                return
                
            # Calcul des données
            data = self.calculate_simulation_data()
            
            if not data:
                messagebox.showwarning("Attention", "Impossible de calculer les données de simulation")
                return
            
            # Chemin du fichier d'historique
            history_file = "historique_simulations.xlsx"
            
            # Création ou ouverture du fichier
            try:
                if os.path.exists(history_file):
                    df_existing = pd.read_excel(history_file)
                else:
                    df_existing = pd.DataFrame()
            except Exception as e:
                print(f"Erreur lors de la lecture du fichier existant: {e}")
                df_existing = pd.DataFrame()
            
            # Nouvelle ligne de données
            new_row = {
                'ID': len(df_existing) + 1,
                'Date': datetime.now().strftime('%d/%m/%Y %H:%M'),
                'Prix acquisition (€)': self.format_currency(data.get('acquisition_price', 0)),
                'Travaux (€)': self.format_currency(data.get('works_cost', 0)),
                'Emprunt (€)': self.format_currency(data.get('loan_amount', 0)),
                'Prix cession (€)': self.format_currency(data.get('selling_price', 0)),
                'Loyer mensuel (€)': self.format_currency(data.get('rent', 0)),
                'Durée détention (ans)': data.get('detention_duration', 0),
                'Coût minimal (€)': self.format_currency(data.get('min_cost', 0)),
                'Option optimale': data.get('optimal_option', ''),
                'Rentabilité brute (%)': self.calculate_profitability(data),
                'ROI estimé (%)': self.calculate_roi(data),
                'Cash-flow annuel (€)': self.format_currency(self.calculate_cashflow(data))
            }
            
            # Ajout de la nouvelle ligne
            if df_existing.empty:
                df_new = pd.DataFrame([new_row])
            else:
                df_new = pd.concat([df_existing, pd.DataFrame([new_row])], ignore_index=True)
            
            # Sauvegarde simple
            df_new.to_excel(history_file, index=False)
            
            # Message de succès avec option d'ouvrir le fichier
            result = messagebox.askyesno("Succès", 
                                       f"Simulation #{new_row['ID']} sauvegardée!\n\n"
                                       f"Option optimale: {new_row['Option optimale']}\n"
                                       f"Coût minimal: {new_row['Coût minimal (€)']}\n\n"
                                       f"Voulez-vous ouvrir le fichier d'historique?")
            
            if result:
                try:
                    os.startfile(history_file)  # Windows
                except:
                    try:
                        os.system(f"open {history_file}")  # macOS
                    except:
                        os.system(f"xdg-open {history_file}")  # Linux
            
            return history_file
            
        except Exception as e:
            print(f"Erreur lors de la sauvegarde: {str(e)}")
            messagebox.showerror("Erreur", f"Erreur lors de la sauvegarde Excel: {str(e)}")

    def calculate_simulation_data(self):
        """Calcule les données de la simulation"""
        try:
            sheet = self.workbook.Sheets("feuil1")
            web_sheet = self.workbook.Sheets("web")
            
            # Recherche de l'option optimale en utilisant la feuille web
            regimes = [
                ("micro nu + meublé", 2),    # Ligne 2 dans la feuille web
                ("SCI IS", 3),               # Ligne 3 dans la feuille web
                ("SCI IS PREL BONI", 4),     # Ligne 4 dans la feuille web
                ("SCI IR", 5)                # Ligne 5 dans la feuille web
            ]
            
            min_cost = float('inf')
            optimal_option = ""
            
            for regime_name, row_num in regimes:
                try:
                    cost_value = web_sheet.Range(f"B{row_num}").Value
                    if cost_value is not None and cost_value < min_cost:
                        min_cost = cost_value
                        optimal_option = regime_name
                except:
                    continue
            
            return {
                'acquisition_price': sheet.Range("c4").Value or 0,
                'works_cost': sheet.Range("b51").Value or 0,
                'loan_amount': sheet.Range("b39").Value or 0,
                'selling_price': sheet.Range("b47").Value or 0,
                'rent': sheet.Range("c1").Value or 0,
                'detention_duration': sheet.Range("c3").Value or 0,
                'min_cost': min_cost if min_cost != float('inf') else 0,
                'optimal_option': optimal_option or "Non défini"
            }
            
        except Exception as e:
            print(f"Erreur dans calculate_simulation_data: {str(e)}")
            return None

    def calculate_profitability(self, data):
        """Calcule la rentabilité brute annuelle"""
        try:
            acquisition = data.get('acquisition_price', 0)
            annual_rent = (data.get('rent', 0) * 12)
            if acquisition > 0:
                return round((annual_rent / acquisition) * 100, 2)
            return 0
        except:
            return 0

    def calculate_roi(self, data):
        """Calcule le ROI estimé sur la durée de détention"""
        try:
            acquisition = data.get('acquisition_price', 0)
            works = data.get('works_cost', 0)
            total_investment = acquisition + works
            
            annual_rent = data.get('rent', 0) * 12
            duration = data.get('detention_duration', 1)
            total_rent = annual_rent * duration
            
            selling_price = data.get('selling_price', 0)
            min_cost = data.get('min_cost', 0)
            
            total_return = total_rent + selling_price - min_cost
            
            if total_investment > 0:
                roi = ((total_return - total_investment) / total_investment) * 100
                return round(roi, 2)
            return 0
        except:
            return 0

    def calculate_cashflow(self, data):
        """Calcule le cash-flow annuel estimé"""
        try:
            annual_rent = data.get('rent', 0) * 12
            # Estimation des charges annuelles (à ajuster selon vos besoins)
            estimated_annual_costs = data.get('min_cost', 0) / data.get('detention_duration', 1) if data.get('detention_duration', 0) > 0 else 0
            
            return annual_rent - estimated_annual_costs
        except:
            return 0

    def format_currency(self, value):
        """Formate une valeur en devise"""
        try:
            if value is None:
                return "0"
            return f"{int(round(float(value))):,}".replace(",", " ")
        except:
            return "0"
    
    def center_window(self):
        """Centrage de la fenêtre"""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')

class SummaryTable(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.configure(style='Main.TFrame')
        
        # Définition des colonnes pour les données saisies
        self.input_columns = [
            "Prix d'acquisition",
            "Travaux",
            "TF + charges loc.",
            "Assurance",
            "Loyer mensuel",
            "Emprunt",
            "Durée emprunt",
            "Taux emprunt",
            "TMI perso./physique",
            "Prix de cession",
            "Durée détention",
            "Prél. prix de cession",
            "CGA"
        ]
        
        # Définition des colonnes pour la synthèse fiscale
        self.fiscal_columns = [
            "Type Régime",
            "Coût Moyen Annuel (40 ans)",
            "Coût Moyen Annuel (durée de détention)",
            "Coût Global (40 ans)",
            "Coût Global (durée de détention)",
            "Fiscalité Plus Value (40 ans)",
            "Fiscalité Plus Value (durée de détention)",
            "Coût Global"
        ]
        
        self.setup_gui()
    
    def setup_gui(self):
        # Configuration des styles
        style = ttk.Style()
        style.configure("Summary.Treeview",
                       background="white",
                       fieldbackground="white",
                       font=('Arial', 10))
        style.configure("Summary.Treeview.Heading",
                       font=('Arial', 10, 'bold'))
        
        # Création des notebooks pour organiser les tableaux
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Frame pour les données saisies
        input_frame = ttk.Frame(notebook)
        notebook.add(input_frame, text="Données Saisies")
        
        # Frame pour la synthèse fiscale
        fiscal_frame = ttk.Frame(notebook)
        notebook.add(fiscal_frame, text="Synthèse Fiscale")
        
        # Création du tableau des données saisies
        self.input_tree = self.create_treeview(input_frame, self.input_columns)
        
        # Création du tableau de synthèse fiscale
        self.fiscal_tree = self.create_treeview(fiscal_frame, self.fiscal_columns)
    
    def create_treeview(self, parent, columns):
        # Création du frame avec scrollbars
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Création du treeview
        tree = ttk.Treeview(frame, columns=columns, show="headings", style="Summary.Treeview")
        
        # Configuration des scrollbars
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Configuration des colonnes
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=150, anchor="center")
        
        # Placement des éléments
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)
        
        return tree
    
    def update_input_data(self, data_dict):
        """Mise à jour des données saisies"""
        self.input_tree.delete(*self.input_tree.get_children())
        values = [data_dict.get(col, "") for col in self.input_columns]
        self.input_tree.insert("", "end", values=values)
    
    def update_fiscal_data(self, fiscal_data):
        """Mise à jour des données fiscales"""
        self.fiscal_tree.delete(*self.fiscal_tree.get_children())
        for regime_data in fiscal_data:
            self.fiscal_tree.insert("", "end", values=regime_data)
    
    def clear_all(self):
        """Effacement de toutes les données"""
        self.input_tree.delete(*self.input_tree.get_children())
        self.fiscal_tree.delete(*self.fiscal_tree.get_children())

class ExcelInterface:
    def __init__(self, root):
        self.root = root
        self.root.title("Interface de Gestion Immobilière")
        
        # Configuration initiale
        self.root.state('zoomed')
        self.excel_path = None
        self.excel = None
        self.workbook = None
        
        # Initialisation du serveur web
        self.web_server = WebServer()
        self.server_thread = None
        
        # Couleurs
        self.COLORS = {
            'bg': "#F0F8FF",
            'header': "#1E3F66",
            'btn_green': "#4CAF50",
            'btn_red': "#F44336",
            'highlight': "#E6F3FF"
        }
        
        # Configuration du style
        self.style = ttk.Style()
        # Utilisation d'un thème plus moderne
        self.style.theme_use("clam")
        self.configure_styles()
        
        # Construction de l'interface
        self.build_interface()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
            
    def on_closing(self):
        """Gestionnaire simple de fermeture de l'application"""
        try:
            print("Fermeture de l'application...")
            self.close_excel_properly()
            self.root.destroy()
        except Exception as e:
            print(f"Erreur lors de la fermeture: {str(e)}")
            self.root.destroy()

    def close_excel_properly(self):
        """Fermeture propre d'Excel"""
        try:
            print("Fermeture d'Excel...")
            
            # Arrêter le serveur web d'abord
            if hasattr(self, 'server_thread') and self.server_thread:
                print("Arrêt du serveur web...")
                self.server_thread = None
            
            # Fermer le classeur
            if self.workbook:
                print("Fermeture du classeur Excel...")
                self.workbook.Close()  # Laisse Excel décider de sauvegarder ou non
                self.workbook = None
            
            # Fermer l'application Excel
            if self.excel:
                print("Fermeture de l'application Excel...")
                self.excel.DisplayAlerts = False  # Éviter les pop-ups
                self.excel.Quit()
                self.excel = None
                
            print("Excel fermé avec succès")
            
        except Exception as e:
            print(f"Erreur lors de la fermeture d'Excel: {str(e)}")
            # En cas d'erreur, essayer de forcer la fermeture
            try:
                if self.excel:
                    self.excel.Quit()
                    self.excel = None
            except:
                pass

    def __del__(self):
        """Destructeur - fermeture en cas d'oubli"""
        try:
            self.close_excel_properly()
        except:
            pass
    def configure_styles(self):
        """Configuration des styles de l'interface"""
        self.style.configure("Main.TFrame", background=self.COLORS['bg'])
        self.style.configure("Header.TFrame", background=self.COLORS['header'])
        self.style.configure("Header.TLabel",
                            font=('Segoe UI', 11, 'bold'),
                            background=self.COLORS['header'],
                            foreground='white')

        self.style.configure("Green.TButton",
                            font=('Segoe UI', 10, 'bold'),
                            padding=10,
                            background=self.COLORS['btn_green'])

        self.style.configure("Red.TButton",
                            font=('Segoe UI', 10, 'bold'),
                            padding=10,
                            background=self.COLORS['btn_red'])

        self.style.configure("Treeview",
                            font=('Segoe UI', 10),
                            rowheight=28,
                            background="white",
                            fieldbackground="white",
                            foreground="black")

        self.style.configure("Treeview.Heading",
                            font=('Segoe UI', 10, 'bold'),
                            foreground="black",
                            background="white",
                            relief="solid",
                            padding=(5, 25))
        
        self.style.map("Treeview.Heading",
                      background=[('active', 'white')],
                      foreground=[('active', 'black')])

        self.style.map("Treeview",
                      background=[('selected', self.COLORS['highlight'])])

    def build_interface(self):
        """Construction de l'interface principale"""
        # Container principal
        self.main_container = ttk.Frame(self.root, style="Main.TFrame")
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # En-tête avec les boutons
        self.build_header(self.main_container)
        
        # Notebook principal pour organiser le contenu
        self.notebook = ttk.Notebook(self.main_container)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Création des onglets principaux
        self.create_data_tab(self.notebook)
        self.create_summary_tab(self.notebook)
    
    def build_header(self, parent):
        """Construction de l'en-tête avec les boutons"""
        header = ttk.Frame(parent, style="Header.TFrame")
        header.pack(fill=tk.X, padx=5, pady=5)
        
        # Bouton de sélection de fichier
        self.select_file_btn = ttk.Button(header,
                                        text="📂 Sélectionner Excel",
                                        command=self.select_file,
                                        style="Green.TButton")
        self.select_file_btn.pack(side=tk.LEFT, padx=5)
        
        # Label pour le nom du fichier
        self.file_label = ttk.Label(header,
                                  text="Aucun fichier sélectionné",
                                  style="Header.TLabel")
        self.file_label.pack(side=tk.LEFT, padx=20)
        
        # Boutons d'action
        buttons_frame = ttk.Frame(header, style="Header.TFrame")
        buttons_frame.pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(buttons_frame,
                  text="📝 Saisie",
                  command=self.show_data_entry,
                  style="Red.TButton").pack(side=tk.LEFT, padx=5)
        
        ttk.Button(buttons_frame,
                  text="📊 Synthèse",
                  command=self.show_fiscal_synthesis,
                  style="Red.TButton").pack(side=tk.LEFT, padx=5)
        
        ttk.Button(buttons_frame,
                  text="🔄 Actualiser",
                  command=self.refresh_data,
                  style="Green.TButton").pack(side=tk.LEFT, padx=5)
                  
        ttk.Button(buttons_frame,
                  text="🌐 Version Web",
                  command=self.open_web_version,
                  style="Green.TButton").pack(side=tk.LEFT, padx=5)
        
        # Bouton Historique
        ttk.Button(buttons_frame,
                  text="📊 Historique",
                  command=self.show_simulation_history,
                  style="Green.TButton").pack(side=tk.LEFT, padx=5)
    
    def create_data_tab(self, notebook):
        """Création de l'onglet des données Excel"""
        data_frame = ttk.Frame(notebook)
        notebook.add(data_frame, text="Données Excel")
        
        # Création du tableau
        self.data_tree = ttk.Treeview(data_frame)
        
        # Scrollbars
        y_scroll = ttk.Scrollbar(data_frame, orient="vertical", command=self.data_tree.yview)
        x_scroll = ttk.Scrollbar(data_frame, orient="horizontal", command=self.data_tree.xview)
        
        # Configuration
        self.data_tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        
        # Placement
        self.data_tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        
        data_frame.grid_columnconfigure(0, weight=1)
        data_frame.grid_rowconfigure(0, weight=1)
    
    def create_summary_tab(self, notebook):
        """Création de l'onglet récapitulatif"""
        summary_frame = ttk.Frame(notebook)
        notebook.add(summary_frame, text="Récapitulatif")
        
        # Sous-notebook pour les deux types de données
        sub_notebook = ttk.Notebook(summary_frame)
        sub_notebook.pack(fill=tk.BOTH, expand=True)
        
        # Création des tableaux récapitulatifs
        self.create_input_summary(sub_notebook)
        self.create_fiscal_summary(sub_notebook)
    
    def create_input_summary(self, notebook):
        """Création du tableau récapitulatif des données saisies"""
        input_frame = ttk.Frame(notebook)
        notebook.add(input_frame, text="Données Saisies")
        
        # Configuration des colonnes
        columns = ("Champ", "Valeur", "Unité")
        self.input_tree = self.create_treeview(input_frame, columns)
    
    def create_fiscal_summary(self, notebook):
        """Création du tableau récapitulatif fiscal"""
        fiscal_frame = ttk.Frame(notebook)
        notebook.add(fiscal_frame, text="Synthèse Fiscale")
        
        # Configuration des colonnes avec leurs titres
        self.fiscal_columns = {
            "regime": "Type de Régime",
            "cout_moyen_40": "Coût Moyen Annuel\n(40 ans)",
            "cout_moyen_detention": "Coût Moyen Annuel\n(durée de détention)",
            "cout_global_40": "Coût Global\n(40 ans)",
            "cout_global_detention": "Coût Global\n(durée de détention)",
            "fisc_plus_value_40": "Fiscalité Plus Value\n(40 ans)",
            "fisc_plus_value_detention": "Fiscalité Plus Value\n(durée de détention)",
            "cout_global": "Coût Global Total"
        }
        
        # Création du treeview avec les colonnes
        self.fiscal_tree = ttk.Treeview(fiscal_frame, columns=list(self.fiscal_columns.keys()), show="headings")
        
        # Configuration des colonnes avec leurs titres
        for col, title in self.fiscal_columns.items():
            self.fiscal_tree.heading(col, text=title)
            # Ajustement de la largeur en fonction de la longueur du titre
            width = max(150, len(title) * 8)  # 8 pixels par caractère approximativement
            self.fiscal_tree.column(col, width=width, anchor="center")
        
        # Scrollbars
        y_scroll = ttk.Scrollbar(fiscal_frame, orient="vertical", command=self.fiscal_tree.yview)
        x_scroll = ttk.Scrollbar(fiscal_frame, orient="horizontal", command=self.fiscal_tree.xview)
        
        self.fiscal_tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        
        # Placement
        self.fiscal_tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        
        fiscal_frame.grid_columnconfigure(0, weight=1)
        fiscal_frame.grid_rowconfigure(0, weight=1)
    
    def create_treeview(self, parent, columns):
        """Création d'un tableau avec scrollbars"""
        tree = ttk.Treeview(parent, columns=columns, show="headings")
        
        # Configuration des colonnes
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=150, anchor="center")
        
        # Scrollbars
        y_scroll = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        x_scroll = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        
        tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        
        # Placement
        tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        
        return tree
    
    def update_input_summary(self, data_dict):
        """Mise à jour du récapitulatif des données saisies"""
        self.input_tree.delete(*self.input_tree.get_children())
        for field, value in data_dict.items():
            self.input_tree.insert("", "end", values=(field, value, ""))
    
    def update_fiscal_summary(self, fiscal_data):
        """Mise à jour du récapitulatif fiscal"""
        self.fiscal_tree.delete(*self.fiscal_tree.get_children())
        for regime_data in fiscal_data:
            self.fiscal_tree.insert("", "end", values=regime_data)

    def format_number(self, value):
        """Formatage des nombres pour l'affichage"""
        try:
            if isinstance(value, str) and value == "-":
                return "-"
            if value is None:
                return "0"
                
            rounded_value = int(round(float(value)))
            if rounded_value < 0:
                return f"-{abs(rounded_value):,}".replace(",", " ")
            return f"{rounded_value:,}".replace(",", " ")
        except (ValueError, TypeError):
            return str(value)

    def select_file(self):
        """Sélection du fichier Excel et ouverture"""
        file_path = filedialog.askopenfilename(
            title="Sélectionner un fichier Excel",
            filetypes=[
                ("Fichiers Excel avec Macros", "*.xlsm"),
                ("Tous les fichiers Excel", "*.xlsx *.xls *.xlsm"),
                ("Fichier Excel standard", "*.xlsx"),
                ("Ancien format Excel", "*.xls"),
                ("Tous les fichiers", "*.*")
            ]
        )
        if file_path:
            self.excel_path = Path(file_path)
            self.file_label.configure(text=f"Fichier : {self.excel_path.name}")
            self.open_excel()

    def open_excel(self):
        """Ouverture du fichier Excel sélectionné"""
        try:
            if self.excel:
                try:
                    if self.workbook:
                        self.workbook.Close(SaveChanges=False)
                    self.excel.Quit()
                except:
                    pass
            
            self.excel = win32com.client.Dispatch("Excel.Application")
            self.excel.Visible = False
            self.excel.DisplayAlerts = False
            
            self.workbook = self.excel.Workbooks.Open(str(self.excel_path.absolute()))
            
            # Mise à jour immédiate du tableau avec les données de synthèse
            self.update_data_tree()
            
            # Mise à jour des autres données
            self.refresh_input_summary()
            self.refresh_fiscal_summary()
            
            # Configuration du serveur web
            self.web_server.set_workbook(self.workbook)
            self.start_web_server()
            
            # Activer les boutons une fois le fichier ouvert
            self.enable_buttons()
            
            messagebox.showinfo("Succès", "Fichier Excel ouvert avec succès!")
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de l'ouverture du fichier: {str(e)}")
            if self.excel:
                try:
                    self.excel.Quit()
                except:
                    pass
            self.excel = None
            self.workbook = None
            self.disable_buttons()

    def enable_buttons(self):
        """Active les boutons après ouverture du fichier"""
        for widget in self.root.winfo_children():
            if isinstance(widget, ttk.Button):
                widget.configure(state='normal')

    def disable_buttons(self):
        """Désactive les boutons si pas de fichier ouvert"""
        for widget in self.root.winfo_children():
            if isinstance(widget, ttk.Button) and widget != self.select_file_btn:
                widget.configure(state='disabled')
                
    def refresh_data(self):
        """Actualisation des données depuis Excel"""
        try:
            if self.workbook:
                # Mise à jour du tableau principal avec les données de synthèse
                self.update_data_tree()
                
                # Mise à jour des récapitulatifs
                self.refresh_input_summary()
                self.refresh_fiscal_summary()
                
                # Forcer la mise à jour du web server
                self.web_server.update_cache()
                
                messagebox.showinfo("Succès", "Données actualisées avec succès!")
                
            else:
                messagebox.showwarning("Attention", "Aucun fichier Excel ouvert!")
                
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de l'actualisation: {str(e)}")

    def update_data_tree(self, df=None):
        """Mise à jour du tableau principal des données avec la feuille synthèse"""
        try:
            # Nettoyage des données existantes
            for item in self.data_tree.get_children():
                self.data_tree.delete(item)
            
            if self.workbook:
                # Lecture de la feuille synthèse
                synthese_sheet = self.workbook.Sheets("synthese")
                
                # Configuration des colonnes basée sur la structure réelle
                columns = [
                    "Type de Régime",
                    "Coût Moyen Annuel (40 ans)",
                    "Coût Moyen Annuel (durée détention)",
                    "Coût Global (40 ans)",
                    "Coût Global (durée détention)",
                    "Fiscalité Plus Value (40 ans)",
                    "Fiscalité Plus Value (durée détention)", 
                    "Coût Global Total"
                ]
                
                self.data_tree["columns"] = columns
                self.data_tree["show"] = "headings"
                
                # Configuration des en-têtes avec largeurs adaptées
                column_widths = [200, 180, 180, 150, 150, 180, 180, 150]
                for i, col in enumerate(columns):
                    self.data_tree.heading(col, text=col)
                    self.data_tree.column(col, width=column_widths[i], anchor="center")
                
                # Données des régimes avec leurs positions dans la feuille synthèse
                # Basé sur l'analyse : lignes 4, 5, 6, 7 et colonnes A, B, C, D, E, F, G, H
                regime_rows = [
                    (4, "micro nu + meublé"),     # Ligne 4
                    (5, "SCI IS"),               # Ligne 5  
                    (6, "SCI IS PREL BONI"),     # Ligne 6
                    (7, "SCI IR")                # Ligne 7
                ]
                
                # Insertion des données
                for row_num, regime_name in regime_rows:
                    try:
                        # Lecture des valeurs des colonnes B à H (indices 1 à 7)
                        values = []
                        values.append(regime_name)  # Nom du régime
                        
                        # Colonnes B à H (coût moyen annuel 40 ans, durée détention, etc.)
                        for col_index in range(1, 8):  # B=1, C=2, D=3, E=4, F=5, G=6, H=7
                            cell_value = synthese_sheet.Cells(row_num, col_index + 1).Value  # +1 car Excel est 1-indexed
                            
                            if cell_value is not None:
                                formatted_value = self.format_number(cell_value)
                                values.append(f"{formatted_value} €")
                            else:
                                values.append("-")
                        
                        # Insertion de la ligne avec style alterné
                        tag = f"row_{len(self.data_tree.get_children()) % 2}"
                        self.data_tree.insert("", "end", values=values, tags=(tag,))
                        
                    except Exception as e:
                        print(f"Erreur pour le régime {regime_name}: {str(e)}")
                        # Insertion d'une ligne avec des valeurs par défaut en cas d'erreur
                        error_values = [regime_name] + ["Erreur"] * 7
                        self.data_tree.insert("", "end", values=error_values)
                
                # Configuration des couleurs alternées
                self.data_tree.tag_configure("row_0", background="#FFFFFF")
                self.data_tree.tag_configure("row_1", background="#F5F9FF")
                        
            else:
                # Si pas de classeur ouvert, affichage d'un message
                self.data_tree["columns"] = ["Message"]
                self.data_tree["show"] = "headings"
                self.data_tree.heading("Message", text="Statut")
                self.data_tree.column("Message", width=400, anchor="center")
                self.data_tree.insert("", "end", values=["Aucun fichier Excel ouvert"])
                
        except Exception as e:
            print(f"Erreur lors de la mise à jour du tableau: {str(e)}")
            # En cas d'erreur, affichage d'un message d'erreur
            self.data_tree["columns"] = ["Erreur"]
            self.data_tree["show"] = "headings" 
            self.data_tree.heading("Erreur", text="Erreur")
            self.data_tree.column("Erreur", width=400, anchor="center")
            self.data_tree.insert("", "end", values=[f"Erreur: {str(e)}"])

    def refresh_input_summary(self):
        """Actualisation du récapitulatif des données saisies"""
        try:
            if self.workbook:
                sheet = self.workbook.Sheets("feuil1")
                data = {
                    # Champ : [Valeur, Unité]
                    "Prix d'acquisition": [sheet.Range("c4").Value, "€"],
                    "Travaux": [sheet.Range("b51").Value, "€"],
                    "TF + charges loc.": [sheet.Range("b25").Value, "€"],
                    "Assurance": [sheet.Range("b26").Value, "€"],
                    "Loyer mensuel": [sheet.Range("c1").Value, "€"],
                    "Emprunt": [sheet.Range("b39").Value, "€"],
                    "Durée emprunt": [sheet.Range("b40").Value, "années"],
                    "Taux emprunt": [sheet.Range("b43").Value, "%"],
                    "TMI perso./physique": [sheet.Range("c6").Value, "%"],
                    "Prix de cession": [sheet.Range("b47").Value, "€"],
                    "Durée détention": [sheet.Range("c3").Value, "années"],
                    "Prél. prix de cession": [sheet.Range("c2").Value, "OUI/NON"],
                    "CGA": [sheet.Range("f3").Value, "OUI/NON"]
                }

                # Nettoyage du tableau existant
                self.input_tree.delete(*self.input_tree.get_children())
                
                # Insertion des données avec leurs unités
                for field, (value, unit) in data.items():
                    # Formatage des valeurs numériques
                    if isinstance(value, (int, float)) and unit not in ["OUI/NON"]:
                        formatted_value = self.format_number(value)
                    else:
                        formatted_value = str(value) if value is not None else "-"
                    
                    self.input_tree.insert("", "end", values=(field, formatted_value, unit))

        except Exception as e:
            print(f"Erreur lors de l'actualisation du récapitulatif des données: {str(e)}")

    def refresh_fiscal_summary(self):
        """Actualisation du récapitulatif fiscal - UTILISE LA FEUILLE WEB"""
        try:
            if not self.workbook:
                raise Exception("Aucun classeur Excel ouvert")

            # MODIFICATION: Utilisation de la feuille "web" au lieu de "fiscalité"
            web_sheet = self.workbook.Sheets("web")
            
            # Nettoyage du tableau existant
            self.fiscal_tree.delete(*self.fiscal_tree.get_children())
            
            # Définition des régimes avec leurs positions dans la feuille web
            regimes = [
                ("micro nu + meublé", 2),    # Ligne 2 dans la feuille web
                ("SCI IS", 3),               # Ligne 3 dans la feuille web
                ("SCI IS PREL BONI", 4),     # Ligne 4 dans la feuille web
                ("SCI IR", 5)                # Ligne 5 dans la feuille web
            ]
            
            # Configuration du style
            style = ttk.Style()
            style.configure("Treeview.Heading",
                          font=('Arial', 10, 'bold'),
                          foreground="black",
                          background="white",
                          padding=(5,15))
            
            # Récupération et insertion des données pour chaque régime
            for regime_name, row_num in regimes:
                try:
                    # Lecture du coût global depuis la colonne B de la feuille web
                    cost_global = web_sheet.Range(f"B{row_num}").Value or 0
                    
                    # Calcul des autres valeurs basées sur le coût global
                    values = [
                        regime_name,
                        self.format_number(cost_global / 40) if cost_global > 0 else "-",  # Coût moyen annuel (40 ans)
                        self.format_number(cost_global / 20) if cost_global > 0 else "-",  # Coût moyen annuel (durée détention - estimation)
                        self.format_number(cost_global),  # Coût global (40 ans)
                        self.format_number(cost_global),  # Coût global (durée détention)
                        "-",  # Fiscalité plus value (40 ans) - données non disponibles dans feuille web
                        "-",  # Fiscalité plus value (durée détention) - données non disponibles dans feuille web
                        self.format_number(cost_global)   # Coût global total
                    ]
                    
                    # Insertion avec style alterné
                    tag = f"row_{len(self.fiscal_tree.get_children()) % 2}"
                    self.fiscal_tree.insert("", "end", values=values, tags=(tag,))
                    
                except Exception as e:
                    print(f"Erreur pour le régime {regime_name}: {str(e)}")
                    self.fiscal_tree.insert("", "end", values=[regime_name] + ["-"] * 7)
            
            # Configuration des couleurs alternées
            self.fiscal_tree.tag_configure("row_0", background="#FFFFFF")
            self.fiscal_tree.tag_configure("row_1", background="#F5F9FF")
            
        except Exception as e:
            print(f"Erreur lors de l'actualisation du récapitulatif fiscal: {str(e)}")
            messagebox.showerror("Erreur", f"Erreur lors de l'actualisation : {str(e)}")

    def get_regime_values(self, regime, sheet):
        """Récupération des valeurs pour chaque régime - UTILISE LA FEUILLE WEB"""
        try:
            # Utilisation de la feuille web au lieu de synthese/fiscalité
            web_sheet = self.workbook.Sheets("web")
            
            # Mapping des régimes vers leurs lignes dans la feuille web
            regime_mapping = {
                "micro nu": 2,           # Ligne 2 : micro nu + meublé
                "micro meublé": 2,       # Ligne 2 : micro nu + meublé  
                "micro classé": 2,       # Ligne 2 : micro nu + meublé
                "SCI IS": 3,             # Ligne 3 : SCI IS
                "SCI IS PREL BONI": 4,   # Ligne 4 : SCI IS PREL BONI
                "SCI IR": 5,             # Ligne 5 : SCI IR
                "LMNP": 2,               # Ligne 2 : micro nu + meublé (approximation)
                "LMNP CGA": 2            # Ligne 2 : micro nu + meublé (approximation)
            }
            
            row_num = regime_mapping.get(regime.lower().replace(" ", ""), 2)
            
            # Lecture du coût global depuis la feuille web (colonne B)
            cost_global = web_sheet.Range(f"B{row_num}").Value or 0
            
            return [
                regime,
                cost_global / 40 if cost_global > 0 else 0,  # coût moyen annuel (40 ans)
                cost_global / 20 if cost_global > 0 else 0,  # coût moyen annuel (durée estimation)
                cost_global,  # coût global (40 ans)
                cost_global,  # coût global (durée de détention)
                0,  # fiscalité plus value (40 ans) - données non disponibles
                0,  # fiscalité plus value (durée de détention) - données non disponibles
                cost_global  # coût global total
            ]
            
        except Exception as e:
            print(f"Erreur dans get_regime_values pour {regime}: {str(e)}")
            return [regime] + [0] * 7

    def show_data_entry(self):
        """Affichage du formulaire de saisie"""
        if not self.workbook:
            messagebox.showerror("Erreur", "Veuillez d'abord ouvrir un fichier Excel!")
            return
        DataEntryForm(self.root, self.workbook)

    def show_fiscal_synthesis(self):
        """Affichage de la synthèse fiscale"""
        if not self.workbook:
            messagebox.showerror("Erreur", "Veuillez d'abord ouvrir un fichier Excel!")
            return
        synthesis_window = tk.Toplevel(self.root)
        FiscalSynthesisInterface(synthesis_window, self.excel, self.workbook)

    def start_web_server(self):
        if not self.server_thread:
            self.web_server.set_workbook(self.workbook)
            self.server_thread = threading.Thread(target=self.web_server.run)
            self.server_thread.daemon = True
            
            try:
                self.server_thread.start()
                time.sleep(2)
                print("Serveur web démarré avec succès")
            except Exception as e:
                print(f"Erreur lors du démarrage du serveur web: {str(e)}")

    def open_web_version(self):
        """Ouvre la version web dans le navigateur"""
        if not self.workbook:
            messagebox.showerror("Erreur", "Veuillez d'abord ouvrir un fichier Excel!")
            return
            
        try:
            self.start_web_server()
            time.sleep(1)
            webbrowser.open('http://localhost:5000')
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de l'ouverture de la version web: {str(e)}")

    # === MÉTHODES DE GESTION DE L'HISTORIQUE ===
    
    def show_simulation_history(self):
        """Affiche l'historique des simulations dans une nouvelle fenêtre"""
        try:
            history_file = "historique_simulations.xlsx"
            
            if not os.path.exists(history_file):
                messagebox.showinfo("Info", "Aucun historique de simulations trouvé.\n"
                                          "Effectuez et sauvegardez une simulation d'abord.")
                return
            
            # Lecture des données
            df = pd.read_excel(history_file)
            
            if df.empty:
                messagebox.showinfo("Info", "L'historique des simulations est vide.")
                return
            
            # Création de la fenêtre d'historique
            history_window = tk.Toplevel(self.root)
            history_window.title("📊 Historique des Simulations")
            history_window.geometry("1400x700")
            history_window.configure(bg="#f8f9fa")
            
            # Frame principal
            main_frame = tk.Frame(history_window, bg="#f8f9fa")
            main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
            
            # Titre
            title_label = tk.Label(main_frame,
                                 text="📊 Historique des Simulations Sauvegardées",
                                 font=('Segoe UI', 18, 'bold'),
                                 bg="#f8f9fa",
                                 fg="#2c3e50")
            title_label.pack(pady=(0, 20))
            
            # Frame pour le tableau
            table_frame = tk.Frame(main_frame, bg="white", relief="solid", bd=1)
            table_frame.pack(fill=tk.BOTH, expand=True)
            
            # Création du Treeview
            columns = list(df.columns)
            tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=20)
            
            # Configuration des colonnes
            for col in columns:
                tree.heading(col, text=col)
                # Ajustement de la largeur selon le contenu
                if 'Date' in col or 'ID' in col:
                    tree.column(col, width=120, anchor="center")
                elif '€' in col or '%' in col:
                    tree.column(col, width=140, anchor="center")
                elif 'Option' in col:
                    tree.column(col, width=180, anchor="center")
                else:
                    tree.column(col, width=100, anchor="center")
            
            # Scrollbars
            v_scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
            h_scrollbar = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
            tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
            
            # Insertion des données avec couleurs alternées
            for i, (_, row) in enumerate(df.iterrows()):
                tag = 'odd' if i % 2 else 'even'
                tree.insert("", "end", values=list(row), tags=(tag,))
            
            # Configuration des couleurs
            tree.tag_configure('even', background='#f8f9fa')
            tree.tag_configure('odd', background='white')
            
            # Placement du tableau et scrollbars
            tree.grid(row=0, column=0, sticky="nsew")
            v_scrollbar.grid(row=0, column=1, sticky="ns")
            h_scrollbar.grid(row=1, column=0, sticky="ew")
            
            table_frame.grid_rowconfigure(0, weight=1)
            table_frame.grid_columnconfigure(0, weight=1)
            
            # Frame pour les boutons d'action
            action_frame = tk.Frame(main_frame, bg="#f8f9fa")
            action_frame.pack(fill=tk.X, pady=(20, 0))
            
            # Boutons d'action
            export_btn = tk.Button(action_frame,
                                 text="📤 Exporter",
                                 command=lambda: self.export_history(df),
                                 font=('Segoe UI', 11, 'bold'),
                                 bg="#28a745",
                                 fg="white",
                                 relief="flat",
                                 padx=20,
                                 pady=8,
                                 cursor="hand2")
            export_btn.pack(side=tk.LEFT, padx=10)
            
            refresh_btn = tk.Button(action_frame,
                                  text="🔄 Actualiser",
                                  command=lambda: self.refresh_history_window(history_window),
                                  font=('Segoe UI', 11, 'bold'),
                                  bg="#17a2b8",
                                  fg="white",
                                  relief="flat",
                                  padx=20,
                                  pady=8,
                                  cursor="hand2")
            refresh_btn.pack(side=tk.LEFT, padx=10)
            
            close_btn = tk.Button(action_frame,
                                text="❌ Fermer",
                                command=history_window.destroy,
                                font=('Segoe UI', 11, 'bold'),
                                bg="#dc3545",
                                fg="white",
                                relief="flat",
                                padx=20,
                                pady=8,
                                cursor="hand2")
            close_btn.pack(side=tk.RIGHT, padx=10)
            
            # Statistiques rapides
            stats_text = f"Total des simulations: {len(df)}"
            if len(df) > 0:
                latest_date = df['Date'].iloc[-1] if 'Date' in df.columns else "N/A"
                stats_text += f" | Dernière simulation: {latest_date}"
            
            stats_label = tk.Label(action_frame,
                                 text=stats_text,
                                 font=('Segoe UI', 10),
                                 bg="#f8f9fa",
                                 fg="#6c757d")
            stats_label.pack(side=tk.LEFT, padx=20)
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de l'affichage de l'historique: {str(e)}")

    def export_history(self, df):
        """Exporte l'historique vers un fichier choisi par l'utilisateur"""
        try:
            file_path = filedialog.asksaveasfilename(
                title="Exporter l'historique",
                defaultextension=".xlsx",
                filetypes=[("Fichiers Excel", "*.xlsx"), ("Fichiers CSV", "*.csv")]
            )
            
            if file_path:
                if file_path.endswith('.xlsx'):
                    df.to_excel(file_path, index=False)
                else:
                    df.to_csv(file_path, index=False, encoding='utf-8-sig')
                
                messagebox.showinfo("Succès", f"Historique exporté vers:\n{file_path}")
                
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de l'export: {str(e)}")

    def refresh_history_window(self, window):
        """Actualise la fenêtre d'historique"""
        window.destroy()
        self.show_simulation_history()

    def __del__(self):
        """Nettoyage des ressources à la fermeture"""
        try:
            if self.workbook:
                self.workbook.Close(SaveChanges=True)
            if self.excel:
                self.excel.Quit()

            # Arrêter le serveur web si nécessaire
            if self.server_thread:
                self.server_thread = None
        except:
            pass

class FiscalSynthesisInterface:
    def __init__(self, root, excel_instance=None, workbook=None):
        self.root = root
        self.root.title("Synthèse de la Fiscalité")
        
        # Configuration plein écran et style
        self.root.state('zoomed')
        self.root.configure(bg="#F0F8FF")
        self.root.lift()
        self.root.focus_force()
        
        # Couleurs
        self.HEADER_COLOR = "#1E3F66"
        self.ROW_COLORS = ["#FFFFFF", "#F5F9FF"]
        self.HIGHLIGHT_COLOR = "#E6F3FF"
        
        # Variables Excel
        self.excel = excel_instance
        self.workbook = workbook
        
        # Configuration du style
        self.style = ttk.Style()
        self.configure_styles()
        
        # Création de l'interface
        self.setup_gui()
        
    def configure_styles(self):
        self.style.configure("Title.TLabel",
                            font=('Arial', 18, 'bold'),
                            background="#F0F8FF",
                            foreground=self.HEADER_COLOR)
        
        self.style.configure("Custom.Treeview",
                            font=('Arial', 11),
                            rowheight=40,
                            background="#FFFFFF",
                            fieldbackground="#FFFFFF",
                            foreground=self.HEADER_COLOR)
        
        self.style.configure("Custom.Treeview.Heading",
                            font=('Arial', 10, 'bold'),
                            background="#FFFFFF",
                            foreground="black",
                            relief="solid",
                            padding=25)
        
        self.style.map("Custom.Treeview.Heading",
                      background=[('active', '#FFFFFF')],
                      foreground=[('active', 'black')])
        
        self.style.map("Custom.Treeview",
                      background=[('selected', self.HIGHLIGHT_COLOR)],
                      foreground=[('selected', 'black')])
        
        self.style.configure("Custom.TButton",
                            font=('Arial', 11),
                            padding=10)

    def setup_gui(self):
        """Configuration complète de l'interface graphique"""
        # Frame principal avec scrollbar
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # Canvas et scrollbar pour le défilement vertical
        canvas = tk.Canvas(main_container, bg="#F0F8FF")
        scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        # Configuration du scrollable frame
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        # Création de la fenêtre dans le canvas
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=canvas.winfo_width())
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Configuration du redimensionnement
        main_container.grid_rowconfigure(0, weight=1)
        main_container.grid_columnconfigure(0, weight=1)
        
        # Placement des éléments principaux
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Frame avec padding pour le contenu
        content_frame = ttk.Frame(scrollable_frame, padding="20")
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Titre principal
        title_label = ttk.Label(content_frame,
                              text="Synthèse de la Fiscalité en Fonction du Type de Régime",
                              style="Title.TLabel")
        title_label.pack(pady=(0, 20))
        
        # Bouton pour afficher les résultats
        ttk.Button(content_frame,
                  text="Afficher les Résultats",
                  command=self.show_results,
                  style="Custom.TButton").pack(pady=(0, 20))
        
        # Premier tableau - Synthèse fiscale
        fiscal_frame = ttk.Frame(content_frame)
        fiscal_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        # Configuration des colonnes pour le premier tableau
        columns = (
            "regime",
            "cout_moyen_40",
            "cout_moyen_detention",
            "cout_global_40",
            "cout_global_detention",
            "fisc_plus_value_40",
            "fisc_plus_value_detention",
            "cout_global"
        )
        
        # Définition des titres des colonnes
        column_titles = {
            "regime": "Type Régime",
            "cout_moyen_40": "Coût Moyen Annuel\n(40 ans)",
            "cout_moyen_detention": "Coût Moyen Annuel\n(durée de détention)",
            "cout_global_40": "Coût Global\n(40 ans)",
            "cout_global_detention": "Coût Global\n(durée de détention)",
            "fisc_plus_value_40": "Fiscalité Plus Value\n(40 ans)",
            "fisc_plus_value_detention": "Fiscalité Plus Value\n(durée de détention)",
            "cout_global": "Coût Global"
        }
        
        # Définition des largeurs des colonnes
        column_widths = {
            "regime": 150,
            "cout_moyen_40": 200,
            "cout_moyen_detention": 200,
            "cout_global_40": 200,
            "cout_global_detention": 200,
            "fisc_plus_value_40": 200,
            "fisc_plus_value_detention": 200,
            "cout_global": 180
        }
        
        # Création du premier Treeview
        self.tree = ttk.Treeview(fiscal_frame,
                                columns=columns,
                                show="headings",
                                style="Custom.Treeview")
        
        # Configuration des colonnes du premier tableau
        for col in columns:
            self.tree.heading(col, text=column_titles[col])
            self.tree.column(col, width=column_widths[col], anchor="center")
        
        # Scrollbars pour le premier tableau
        y_scrollbar1 = ttk.Scrollbar(fiscal_frame, orient="vertical", command=self.tree.yview)
        x_scrollbar1 = ttk.Scrollbar(fiscal_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scrollbar1.set, xscrollcommand=x_scrollbar1.set)
        
        # Placement du premier tableau et ses scrollbars
        y_scrollbar1.pack(side=tk.RIGHT, fill=tk.Y)
        x_scrollbar1.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Deuxième tableau - Synthèse des revenus
        revenue_frame = ttk.Frame(content_frame)
        revenue_frame.pack(fill=tk.BOTH, expand=True, pady=(20, 0))
        
        # Titre pour le second tableau
        revenue_title = ttk.Label(revenue_frame,
                               text="Synthèse des Revenus par Régime",
                               style="Title.TLabel")
        revenue_title.pack(pady=(0, 20))
        
        # Configuration des colonnes pour le second tableau
        revenue_columns = ("regime", "revenu_global", "resultat")
        revenue_titles = {
            "regime": "Type Régime",
            "revenu_global": "Revenu Global",
            "resultat": "Résultat Net"
        }
        revenue_widths = {
            "regime": 150,
            "revenu_global": 200,
            "resultat": 200
        }
        
        # Création du second Treeview
        self.revenue_tree = ttk.Treeview(revenue_frame,
                                        columns=revenue_columns,
                                        show="headings",
                                        style="Custom.Treeview")
        
        # Configuration des colonnes du second tableau
        for col in revenue_columns:
            self.revenue_tree.heading(col, text=revenue_titles[col])
            self.revenue_tree.column(col, width=revenue_widths[col], anchor="center")
        
        # Scrollbars pour le second tableau
        y_scrollbar2 = ttk.Scrollbar(revenue_frame, orient="vertical", command=self.revenue_tree.yview)
        x_scrollbar2 = ttk.Scrollbar(revenue_frame, orient="horizontal", command=self.revenue_tree.xview)
        self.revenue_tree.configure(yscrollcommand=y_scrollbar2.set, xscrollcommand=x_scrollbar2.set)
        
        # Placement du second tableau et ses scrollbars
        y_scrollbar2.pack(side=tk.RIGHT, fill=tk.Y)
        x_scrollbar2.pack(side=tk.BOTTOM, fill=tk.X)
        self.revenue_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Configuration du redimensionnement du canvas
        def configure_canvas(event):
            canvas.configure(width=main_container.winfo_width()-scrollbar.winfo_width())
            canvas.itemconfig(canvas.find_withtag("all")[0], width=main_container.winfo_width()-scrollbar.winfo_width())
        
        # Binding pour le redimensionnement
        self.root.bind("<Configure>", configure_canvas)
        
        # Chargement automatique des résultats
        self.show_results()

    def format_number(self, value):
        try:
            if isinstance(value, str) and value == "-":
                return "-"
            if value is None:
                return "0"
                
            # Arrondir à l'entier le plus proche
            rounded_value = int(round(float(value)))
            
            # Formater avec séparateur de milliers
            if rounded_value < 0:
                return f"-{abs(rounded_value):,}".replace(",", " ")
            return f"{rounded_value:,}".replace(",", " ")
        except (ValueError, TypeError):
            return str(value)

    def show_results(self):
        """MODIFICATION: Utilise la feuille web au lieu de fiscalité"""
        try:
            if self.workbook:
                # MODIFICATION: Utilisation de la feuille "web" au lieu de "fiscalité"
                web_sheet = self.workbook.Sheets("web")
                
                # Nettoyer le tableau existant
                for item in self.tree.get_children():
                    self.tree.delete(item)
                    
                # Données des régimes avec leurs positions dans la feuille web
                regimes = [
                    ("micro nu + meublé", 2),    # Ligne 2 dans la feuille web
                    ("SCI IS", 3),               # Ligne 3 dans la feuille web
                    ("SCI IS PREL BONI", 4),     # Ligne 4 dans la feuille web
                    ("SCI IR", 5)                # Ligne 5 dans la feuille web
                ]
                
                # Remplir le tableau avec les données
                for idx, (regime_name, row_num) in enumerate(regimes):
                    try:
                        # Lecture du coût global depuis la colonne B de la feuille web
                        cost_global = web_sheet.Range(f"B{row_num}").Value or 0
                        
                        # Calcul des autres valeurs basées sur le coût global
                        values = [
                            regime_name,
                            self.format_number(cost_global / 40) if cost_global > 0 else "-",  # Coût moyen annuel (40 ans)
                            self.format_number(cost_global / 20) if cost_global > 0 else "-",  # Coût moyen annuel (durée détention - estimation)
                            self.format_number(cost_global),  # Coût global (40 ans)
                            self.format_number(cost_global),  # Coût global (durée détention)
                            "-",  # Fiscalité plus value (40 ans) - données non disponibles dans feuille web
                            "-",  # Fiscalité plus value (durée détention) - données non disponibles dans feuille web
                            self.format_number(cost_global)   # Coût global total
                        ]
                        
                        tag = f"row_{idx % 2}"
                        self.tree.insert("", "end", values=values, tags=(tag,))
                        
                    except Exception as e:
                        print(f"Erreur pour le régime {regime_name}: {str(e)}")
                        self.tree.insert("", "end", values=[regime_name] + ["-"] * 7)
                    
                # Configurer les couleurs alternées
                self.tree.tag_configure("row_0", background=self.ROW_COLORS[0])
                self.tree.tag_configure("row_1", background=self.ROW_COLORS[1])
                
                # Mise à jour du tableau des revenus
                self.update_revenue_table()
                
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du chargement des données: {str(e)}")

    def get_regime_values(self, regime, sheet):
        """MODIFICATION: Utilise la feuille web pour récupérer les valeurs"""
        try:
            # Utilisation de la feuille web au lieu de synthese/fiscalité
            web_sheet = self.workbook.Sheets("web")
            
            # Mapping des régimes vers leurs lignes dans la feuille web
            regime_mapping = {
                "micro nu": 2,           # Ligne 2 : micro nu + meublé
                "micro meublé": 2,       # Ligne 2 : micro nu + meublé  
                "micro classé": 2,       # Ligne 2 : micro nu + meublé
                "SCI IS": 3,             # Ligne 3 : SCI IS
                "SCI IS PREL BONI": 4,   # Ligne 4 : SCI IS PREL BONI
                "SCI IR": 5,             # Ligne 5 : SCI IR
                "LMNP": 2,               # Ligne 2 : micro nu + meublé (approximation)
                "LMNP CGA": 2            # Ligne 2 : micro nu + meublé (approximation)
            }
            
            row_num = regime_mapping.get(regime.lower().replace(" ", ""), 2)
            
            # Lecture du coût global depuis la feuille web (colonne B)
            cost_global = web_sheet.Range(f"B{row_num}").Value or 0
            
            return [
                regime,
                cost_global / 40 if cost_global > 0 else 0,  # coût moyen annuel (40 ans)
                cost_global / 20 if cost_global > 0 else 0,  # coût moyen annuel (durée estimation)
                cost_global,  # coût global (40 ans)
                cost_global,  # coût global (durée de détention)
                0,  # fiscalité plus value (40 ans) - données non disponibles
                0,  # fiscalité plus value (durée de détention) - données non disponibles
                cost_global  # coût global total
            ]
            
        except Exception as e:
            print(f"Erreur pour le régime {regime}: {str(e)}")
            return [regime, "Erreur", "Erreur", "Erreur", "Erreur", "Erreur", "Erreur", "Erreur"]

    def calculate_revenue_data(self):
        """Calcul des données de revenus pour chaque régime"""
        try:
            if not self.workbook:
                return []
                
            sheet = self.workbook.Sheets("feuil1")
            web_sheet = self.workbook.Sheets("web")
            
            # Récupération des données de base
            loyer_mensuel = sheet.Range("c1").Value
            duree_detention = sheet.Range("c3").Value
            
            # Calcul du revenu global
            revenu_global = loyer_mensuel * 12 * duree_detention
            
            # Liste des régimes et leurs coûts globaux depuis la feuille web
            regimes_data = [
                ("micro nu + meublé", web_sheet.Range("B2").Value or 0),
                ("SCI IS", web_sheet.Range("B3").Value or 0),
                ("SCI IS PREL BONI", web_sheet.Range("B4").Value or 0),
                ("SCI IR", web_sheet.Range("B5").Value or 0)
            ]
            
            # Calcul des résultats
            results = []
            for regime, cout_global in regimes_data:
                resultat = revenu_global - cout_global
                results.append((regime, revenu_global, resultat))
                
            return results
            
        except Exception as e:
            print(f"Erreur lors du calcul des revenus: {str(e)}")
            return []

    def update_revenue_table(self):
        """Mise à jour du tableau des revenus"""
        try:
            # Nettoyer le tableau existant
            for item in self.revenue_tree.get_children():
                self.revenue_tree.delete(item)
                
            # Calculer et insérer les nouvelles données
            revenue_data = self.calculate_revenue_data()
            
            for idx, (regime, revenu_global, resultat) in enumerate(revenue_data):
                formatted_values = [
                    regime,
                    self.format_number(revenu_global),
                    self.format_number(resultat)
                ]
                tag = f"row_{idx % 2}"
                self.revenue_tree.insert("", "end", values=formatted_values, tags=(tag,))
                
            # Configurer les couleurs alternées
            self.revenue_tree.tag_configure("row_0", background=self.ROW_COLORS[0])
            self.revenue_tree.tag_configure("row_1", background=self.ROW_COLORS[1])
            
        except Exception as e:
            print(f"Erreur lors de la mise à jour du tableau des revenus: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelInterface(root)
    root.mainloop()