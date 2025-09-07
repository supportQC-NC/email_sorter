"""
Email Manager pour Thunderbird V3 - Gestionnaire d'emails avec tri automatique avancé
Version améliorée avec chaînes de règles et gestion des dossiers existants
Compatible avec tous les serveurs IMAP (Orange, Free, Hostinger, etc.)
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import imaplib
import email
from email.header import decode_header
import threading
from datetime import datetime, timedelta
import json
import os
import re
import copy
import platform
from pathlib import Path

class EmailManager:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("🦅 Email Manager pour Thunderbird - V3")
        self.root.geometry("1300x850")
        self.root.configure(bg='#2c3e50')
        
        # Créer le dossier de données
        self.setup_data_directory()
        
        # Variables
        self.connection = None
        self.rules = []
        self.rule_chains = []
        self.is_running = False
        self.folder_separator = "."
        self.use_inbox_prefix = True
        self.processed_emails = set()
        self.existing_folders = []
        
        # Interface
        self.setup_ui()
        
        # Charger configuration
        self.load_settings()
        
        # Fermeture
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def setup_data_directory(self):
        """Créer le dossier de données au premier lancement"""
        # Déterminer le chemin selon l'OS
        if platform.system() == 'Windows':
            # Sur Windows, utiliser le disque C:\Users\[username]\
            base_path = Path.home() / "support_data_email_sort"
        else:
            # Sur Linux/Mac, utiliser le home directory
            base_path = Path.home() / ".support_data_email_sort"
        
        # Créer le dossier s'il n'existe pas
        base_path.mkdir(parents=True, exist_ok=True)
        
        # Définir le chemin du fichier de configuration
        self.config_file = base_path / "email_manager_settings.json"
        self.rules_backup_file = base_path / "rules_backup.json"
        self.chains_file = base_path / "rule_chains.json"
        
        # Log du chemin
        print(f"📁 Dossier de données: {base_path}")
        print(f"📄 Fichier de configuration: {self.config_file}")
    
    def setup_ui(self):
        """Créer l'interface utilisateur complète"""
        
        # Style moderne
        style = ttk.Style()
        style.theme_use('clam')
        
        # Container principal
        main_frame = tk.Frame(self.root, bg='#ecf0f1')
        main_frame.pack(fill='both', expand=True, padx=2, pady=2)
        
        # === HEADER ===
        header = tk.Frame(main_frame, bg='#0a84ff', height=100)
        header.pack(fill='x')
        header.pack_propagate(False)
        
        # Titre avec icône
        title_frame = tk.Frame(header, bg='#0a84ff')
        title_frame.pack(expand=True)
        
        tk.Label(title_frame, text="🦅", font=("Arial", 40), 
                bg='#0a84ff', fg='white').pack(side='left', padx=10)
        
        title_text = tk.Frame(title_frame, bg='#0a84ff')
        title_text.pack(side='left')
        
        tk.Label(title_text, text="Email Manager pour Thunderbird V3", 
                font=("Arial", 26, "bold"), 
                bg='#0a84ff', fg='white').pack(anchor='w')
        
        tk.Label(title_text, text="Tri automatique avancé avec chaînes de règles et gestion des dossiers", 
                font=("Arial", 12), 
                bg='#0a84ff', fg='#ecf0f1').pack(anchor='w')
        
        # === NOTEBOOK (Onglets) ===
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # --- ONGLET 1: CONNEXION ---
        self.setup_connection_tab(notebook)
        
        # --- ONGLET 2: GESTION CC ---
        self.setup_cc_tab(notebook)
        
        # --- ONGLET 3: RÈGLES PERSONNALISÉES ---
        self.setup_rules_tab(notebook)
        
        # --- ONGLET 4: CHAÎNES DE RÈGLES ---
        self.setup_rule_chains_tab(notebook)
        
        # --- ONGLET 5: DOSSIERS EXISTANTS ---
        self.setup_folders_tab(notebook)
        
        # --- ONGLET 6: OPTIONS AVANCÉES ---
        self.setup_advanced_tab(notebook)
        
        # --- ONGLET 7: EXÉCUTION ---
        self.setup_execution_tab(notebook)
        
        # Barre de statut
        self.status_var = tk.StringVar(value="✅ Prêt - Email Manager V3")
        status_bar = tk.Label(self.root, textvariable=self.status_var,
                             bd=1, relief=tk.SUNKEN, anchor='w',
                             bg='#34495e', fg='white',
                             font=("Arial", 10))
        status_bar.pack(side='bottom', fill='x')
        
        # Initialisation
        self.log("✨ Email Manager V3 démarré!", "success")
        self.log(f"📁 Dossier de données: {self.config_file.parent}", "info")
        self.log("🔒 Les emails ne seront PAS marqués comme lus lors du tri", "warning")
        self.log("📌 Compatible avec tous les serveurs IMAP", "info")
    
    def setup_connection_tab(self, notebook):
        """Onglet de connexion"""
        conn_frame = tk.Frame(notebook, bg='white')
        notebook.add(conn_frame, text=' 🔌 Connexion ')
        
        conn_content = tk.Frame(conn_frame, bg='white')
        conn_content.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Configuration serveur
        server_frame = tk.LabelFrame(conn_content, text=" Configuration serveur IMAP ", 
                                    font=("Arial", 12, "bold"),
                                    bg='white', fg='#2c3e50', relief=tk.FLAT)
        server_frame.pack(fill='x', pady=(0, 20))
        
        # Exemples de serveurs
        examples_frame = tk.Frame(server_frame, bg='#e8f4f8')
        examples_frame.pack(fill='x', padx=10, pady=10)
        
        tk.Label(examples_frame, 
                text="📌 Serveurs IMAP courants:",
                font=("Arial", 11, "bold"), bg='#e8f4f8', fg='#0a84ff').pack(anchor='w', padx=10, pady=(5, 5))
        
        examples_text = """• Orange: imap.orange.fr (port 993)          • Free: imap.free.fr (port 993)
• SFR: imap.sfr.fr (port 993)               • Bbox: imap.bbox.fr (port 993)
• LaPoste: imap.laposte.net (port 993)      • Hostinger: imap.hostinger.com (port 993)
• OVH: ssl0.ovh.net (port 993)              • Gmail: imap.gmail.com (port 993)"""
        
        tk.Label(examples_frame, text=examples_text,
                font=("Arial", 9), bg='#e8f4f8', fg='#2c3e50',
                justify='left').pack(anchor='w', padx=20, pady=(0, 10))
        
        server_grid = tk.Frame(server_frame, bg='white')
        server_grid.pack(pady=15, padx=10)
        
        tk.Label(server_grid, text="Serveur IMAP:", font=("Arial", 11), 
                bg='white', width=15, anchor='e').grid(row=0, column=0, padx=5, pady=5)
        
        self.server_var = tk.StringVar(value="")
        self.server_entry = tk.Entry(server_grid, textvariable=self.server_var, 
                                    font=("Arial", 11), width=25)
        self.server_entry.grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(server_grid, text="Port:", font=("Arial", 11), 
                bg='white', width=8, anchor='e').grid(row=0, column=2, padx=5, pady=5)
        
        self.port_var = tk.StringVar(value="993")
        self.port_entry = tk.Entry(server_grid, textvariable=self.port_var, 
                                  font=("Arial", 11), width=8)
        self.port_entry.grid(row=0, column=3, padx=5, pady=5)
        
        # Identifiants
        creds_frame = tk.LabelFrame(conn_content, text=" Identifiants ", 
                                   font=("Arial", 12, "bold"),
                                   bg='white', fg='#2c3e50', relief=tk.FLAT)
        creds_frame.pack(fill='x', pady=(0, 20))
        
        creds_grid = tk.Frame(creds_frame, bg='white')
        creds_grid.pack(pady=15, padx=10)
        
        tk.Label(creds_grid, text="Email:", font=("Arial", 11), 
                bg='white', width=15, anchor='e').grid(row=0, column=0, padx=5, pady=10)
        
        self.email_var = tk.StringVar()
        self.email_entry = tk.Entry(creds_grid, textvariable=self.email_var, 
                                   font=("Arial", 11), width=35)
        self.email_entry.grid(row=0, column=1, padx=5, pady=10)
        
        tk.Label(creds_grid, text="Mot de passe:", font=("Arial", 11), 
                bg='white', width=15, anchor='e').grid(row=1, column=0, padx=5, pady=10)
        
        self.password_var = tk.StringVar()
        self.password_entry = tk.Entry(creds_grid, textvariable=self.password_var, 
                                      font=("Arial", 11), width=35, show="•")
        self.password_entry.grid(row=1, column=1, padx=5, pady=10)
        
        # Option importante : Préservation du statut non-lu
        preserve_frame = tk.Frame(conn_content, bg='#fff3cd', relief=tk.RIDGE, bd=2)
        preserve_frame.pack(fill='x', pady=10)
        
        self.preserve_unread_var = tk.BooleanVar(value=True)
        tk.Checkbutton(preserve_frame, 
                      text=" 🔒 TOUJOURS préserver le statut non-lu des emails lors du tri",
                      variable=self.preserve_unread_var,
                      font=("Arial", 11, "bold"),
                      bg='#fff3cd', fg='#856404',
                      activebackground='#fff3cd').pack(anchor='w', padx=10, pady=8)
        
        # Boutons
        buttons_frame = tk.Frame(conn_content, bg='white')
        buttons_frame.pack(pady=20)
        
        self.test_btn = tk.Button(buttons_frame, text=" 🔧 Tester la connexion ",
                                 font=("Arial", 12, "bold"),
                                 bg='#0a84ff', fg='white',
                                 padx=30, pady=10,
                                 command=self.test_connection,
                                 cursor='hand2')
        self.test_btn.pack(side='left', padx=10)
        
        self.load_folders_btn = tk.Button(buttons_frame, text=" 📁 Charger les dossiers ",
                                         font=("Arial", 12, "bold"),
                                         bg='#27ae60', fg='white',
                                         padx=30, pady=10,
                                         command=self.load_existing_folders,
                                         cursor='hand2')
        self.load_folders_btn.pack(side='left', padx=10)
    
    def setup_cc_tab(self, notebook):
        """Onglet de gestion des CC"""
        cc_frame = tk.Frame(notebook, bg='white')
        notebook.add(cc_frame, text=' 📋 Gestion CC ')
        
        cc_content = tk.Frame(cc_frame, bg='white')
        cc_content.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Configuration CC
        cc_config = tk.LabelFrame(cc_content, 
                                 text=" ⚙️ Configuration des emails en copie (CC) ", 
                                 font=("Arial", 12, "bold"),
                                 bg='white', fg='#2c3e50', relief=tk.FLAT)
        cc_config.pack(fill='x', pady=(0, 20))
        
        cc_inner = tk.Frame(cc_config, bg='white')
        cc_inner.pack(pady=20, padx=20)
        
        # Activation CC
        activation_frame = tk.Frame(cc_inner, bg='#e8f8f5', relief=tk.RIDGE, bd=2)
        activation_frame.pack(fill='x', pady=10)
        
        self.cc_enabled_var = tk.BooleanVar(value=True)
        cc_check = tk.Checkbutton(activation_frame, 
                                 text=" ✅ Activer le tri automatique des emails où je suis en copie (CC)",
                                 variable=self.cc_enabled_var,
                                 font=("Arial", 12, "bold"),
                                 bg='#e8f8f5', fg='#27ae60',
                                 activebackground='#e8f8f5',
                                 command=self.toggle_cc_options)
        cc_check.pack(anchor='w', padx=10, pady=10)
        
        # Options CC
        self.cc_options_frame = tk.Frame(cc_inner, bg='#ecf0f1')
        self.cc_options_frame.pack(fill='x', pady=10)
        
        folder_frame = tk.Frame(self.cc_options_frame, bg='#ecf0f1')
        folder_frame.pack(pady=10)
        
        tk.Label(folder_frame, text="📁 Dossier de destination pour les emails en CC:", 
                font=("Arial", 11), bg='#ecf0f1').pack(side='left', padx=10)
        
        self.cc_folder_var = tk.StringVar(value="EN_COPIE")
        self.cc_folder_entry = tk.Entry(folder_frame, 
                                       textvariable=self.cc_folder_var,
                                       font=("Arial", 11, "bold"), width=25,
                                       bg='#ffffff', fg='#2c3e50')
        self.cc_folder_entry.pack(side='left', padx=5)
        
        # Options supplémentaires CC
        extra_options = tk.Frame(self.cc_options_frame, bg='#ecf0f1')
        extra_options.pack(fill='x', pady=10, padx=20)
        
        self.cc_mark_read_after_var = tk.BooleanVar(value=False)
        tk.Checkbutton(extra_options, 
                      text=" 📖 Marquer comme lu APRÈS déplacement (optionnel)",
                      variable=self.cc_mark_read_after_var,
                      font=("Arial", 10), bg='#ecf0f1').pack(anchor='w', pady=5)
        
        self.cc_skip_important_var = tk.BooleanVar(value=True)
        tk.Checkbutton(extra_options, 
                      text=" ⭐ Ne pas déplacer les emails marqués comme importants",
                      variable=self.cc_skip_important_var,
                      font=("Arial", 10), bg='#ecf0f1').pack(anchor='w', pady=5)
        
        self.cc_skip_recent_var = tk.BooleanVar(value=False)
        tk.Checkbutton(extra_options, 
                      text=" 🕐 Ne pas déplacer les emails de moins de 24h",
                      variable=self.cc_skip_recent_var,
                      font=("Arial", 10), bg='#ecf0f1').pack(anchor='w', pady=5)
        
        # Info box
        info_frame = tk.Frame(cc_content, bg='#d4edda', relief=tk.RIDGE, bd=2)
        info_frame.pack(fill='x', pady=10)
        
        tk.Label(info_frame, text="ℹ️ Comment ça marche ?", 
                font=("Arial", 11, "bold"), 
                bg='#d4edda', fg='#155724').pack(anchor='w', padx=10, pady=(10, 5))
        
        info_text = """✅ Les emails où vous êtes le destinataire principal → RESTENT dans la boîte de réception
✅ Les emails où vous êtes uniquement en copie (CC) → DÉPLACÉS vers le dossier configuré
🔒 Le statut non-lu est TOUJOURS préservé (sauf si option explicite)
✅ Le dossier sera créé automatiquement s'il n'existe pas"""
        
        tk.Label(info_frame, text=info_text,
                font=("Arial", 10), bg='#d4edda', 
                fg='#155724', justify='left').pack(anchor='w', padx=25, pady=(0, 10))
    
    def setup_rules_tab(self, notebook):
        """Onglet des règles personnalisées amélioré avec priorités"""
        rules_frame = tk.Frame(notebook, bg='white')
        notebook.add(rules_frame, text=' 🎯 Règles ')
        
        rules_content = tk.Frame(rules_frame, bg='white')
        rules_content.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Créer règle
        create_rule_frame = tk.LabelFrame(rules_content, 
                                         text=" ➕ Créer une nouvelle règle de tri ", 
                                         font=("Arial", 12, "bold"),
                                         bg='white', fg='#2c3e50', relief=tk.FLAT)
        create_rule_frame.pack(fill='x', pady=(0, 20))
        
        rule_builder = tk.Frame(create_rule_frame, bg='white')
        rule_builder.pack(pady=15, padx=15)
        
        # Ligne 0 - Nom et priorité
        line0 = tk.Frame(rule_builder, bg='white')
        line0.pack(fill='x', pady=5)
        
        tk.Label(line0, text="Nom de la règle:", font=("Arial", 11, "bold"), bg='white').pack(side='left', padx=5)
        
        self.rule_name_var = tk.StringVar()
        name_entry = tk.Entry(line0, textvariable=self.rule_name_var, 
                             width=30, font=("Arial", 11))
        name_entry.pack(side='left', padx=5)
        
        tk.Label(line0, text="Priorité:", font=("Arial", 11, "bold"), bg='white').pack(side='left', padx=10)
        
        self.rule_priority_var = tk.StringVar(value="50")
        priority_spinbox = tk.Spinbox(line0, from_=1, to=100, 
                                      textvariable=self.rule_priority_var,
                                      width=8, font=("Arial", 11))
        priority_spinbox.pack(side='left', padx=5)
        
        tk.Label(line0, text="(1=haute, 100=basse)", 
                font=("Arial", 9, "italic"), bg='white', fg='gray').pack(side='left')
        
        # Ligne 1 - Condition principale
        line1 = tk.Frame(rule_builder, bg='white')
        line1.pack(fill='x', pady=5)
        
        tk.Label(line1, text="Si", font=("Arial", 11, "bold"), bg='white').pack(side='left', padx=5)
        
        self.rule_field_var = tk.StringVar(value="Sujet")
        field_menu = ttk.Combobox(line1, textvariable=self.rule_field_var,
                                  values=["Sujet", "Expéditeur", "Corps", "Destinataire", 
                                         "Sujet ou Corps", "Domaine expéditeur"],
                                  width=18, state='readonly')
        field_menu.pack(side='left', padx=5)
        
        self.rule_condition_var = tk.StringVar(value="contient")
        condition_menu = ttk.Combobox(line1, textvariable=self.rule_condition_var,
                                      values=["contient", "ne contient pas", "commence par", 
                                              "finit par", "est exactement", "n'est pas", 
                                              "correspond à (regex)", "contient un de (liste)"],
                                      width=20, state='readonly')
        condition_menu.pack(side='left', padx=5)
        
        self.rule_keyword_var = tk.StringVar()
        keyword_entry = tk.Entry(line1, textvariable=self.rule_keyword_var, 
                                width=30, font=("Arial", 11))
        keyword_entry.pack(side='left', padx=5)
        
        # Ligne 2 - Conditions additionnelles
        line2 = tk.Frame(rule_builder, bg='white')
        line2.pack(fill='x', pady=5)
        
        tk.Label(line2, text="ET (optionnel):", font=("Arial", 11), bg='white').pack(side='left', padx=5)
        
        self.rule_and_field_var = tk.StringVar(value="")
        and_field_menu = ttk.Combobox(line2, textvariable=self.rule_and_field_var,
                                      values=["", "Sujet", "Expéditeur", "Corps", "Destinataire"],
                                      width=15, state='readonly')
        and_field_menu.pack(side='left', padx=5)
        
        self.rule_and_condition_var = tk.StringVar(value="contient")
        and_condition_menu = ttk.Combobox(line2, textvariable=self.rule_and_condition_var,
                                          values=["contient", "ne contient pas", "commence par", 
                                                  "finit par", "est exactement"],
                                          width=18, state='readonly')
        and_condition_menu.pack(side='left', padx=5)
        
        self.rule_and_keyword_var = tk.StringVar()
        and_keyword_entry = tk.Entry(line2, textvariable=self.rule_and_keyword_var, 
                                     width=25, font=("Arial", 11))
        and_keyword_entry.pack(side='left', padx=5)
        
        # Ligne 3 - Options
        line3 = tk.Frame(rule_builder, bg='white')
        line3.pack(fill='x', pady=5)
        
        tk.Label(line3, text="Options:", font=("Arial", 11, "bold"), bg='white').pack(side='left', padx=5)
        
        self.rule_case_sensitive_var = tk.BooleanVar(value=False)
        tk.Checkbutton(line3, text="Sensible à la casse",
                      variable=self.rule_case_sensitive_var,
                      font=("Arial", 10), bg='white').pack(side='left', padx=5)
        
        self.rule_continue_chain_var = tk.BooleanVar(value=False)
        tk.Checkbutton(line3, text="Peut continuer vers d'autres règles",
                      variable=self.rule_continue_chain_var,
                      font=("Arial", 10), bg='white').pack(side='left', padx=10)
        
        # Ligne 4 - Action
        line4 = tk.Frame(rule_builder, bg='white')
        line4.pack(fill='x', pady=5)
        
        tk.Label(line4, text="Alors", font=("Arial", 11, "bold"), bg='white').pack(side='left', padx=5)
        
        self.rule_action_var = tk.StringVar(value="Déplacer vers")
        action_menu = ttk.Combobox(line4, textvariable=self.rule_action_var,
                                   values=["Déplacer vers", "Copier vers", "Marquer comme lu", 
                                          "Marquer comme important", "Supprimer", "Étiqueter"],
                                   width=18, state='readonly')
        action_menu.pack(side='left', padx=5)
        action_menu.bind('<<ComboboxSelected>>', self.on_action_changed)
        
        self.rule_folder_var = tk.StringVar()
        self.folder_entry = tk.Entry(line4, textvariable=self.rule_folder_var, 
                                     width=25, font=("Arial", 11))
        self.folder_entry.pack(side='left', padx=5)
        
        # Dropdown pour les dossiers existants
        self.folder_dropdown = ttk.Combobox(line4, values=self.existing_folders,
                                           width=25, state='readonly')
        self.folder_dropdown.pack(side='left', padx=5)
        self.folder_dropdown.bind('<<ComboboxSelected>>', self.on_folder_selected)
        
        # Ligne 5 - Options d'action
        line5 = tk.Frame(rule_builder, bg='white')
        line5.pack(fill='x', pady=5)
        
        self.rule_stop_processing_var = tk.BooleanVar(value=False)
        tk.Checkbutton(line5, text=" 🛑 Arrêter le traitement après cette règle",
                      variable=self.rule_stop_processing_var,
                      font=("Arial", 10), bg='white').pack(side='left', padx=5)
        
        self.rule_mark_after_move_var = tk.BooleanVar(value=False)
        self.mark_checkbox = tk.Checkbutton(line5, 
                                           text=" 📖 Marquer comme lu après action",
                                           variable=self.rule_mark_after_move_var,
                                           font=("Arial", 10), bg='white')
        self.mark_checkbox.pack(side='left', padx=15)
        
        # Bouton ajouter
        add_btn_frame = tk.Frame(rule_builder, bg='white')
        add_btn_frame.pack(fill='x', pady=10)
        
        add_rule_btn = tk.Button(add_btn_frame, text=" ➕ Ajouter la règle ",
                                bg='#27ae60', fg='white',
                                font=("Arial", 11, "bold"),
                                command=self.add_custom_rule,
                                cursor='hand2')
        add_rule_btn.pack()
        
        # Liste des règles
        rules_list_frame = tk.LabelFrame(rules_content, 
                                        text=" 📋 Règles actives (triées par priorité) ", 
                                        font=("Arial", 12, "bold"),
                                        bg='white', fg='#2c3e50', relief=tk.FLAT)
        rules_list_frame.pack(fill='both', expand=True)
        
        # Frame avec scrollbar
        list_container = tk.Frame(rules_list_frame, bg='white')
        list_container.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Treeview pour afficher les règles
        columns = ('Nom', 'Priorité', 'Champ', 'Condition', 'Valeur', 'Action', 'Options')
        self.rules_tree = ttk.Treeview(list_container, columns=columns, show='tree headings', height=8)
        
        # Configuration des colonnes
        self.rules_tree.heading('#0', text='#')
        self.rules_tree.column('#0', width=40)
        
        for col in columns:
            self.rules_tree.heading(col, text=col)
            width = 120 if col in ['Valeur', 'Options'] else 100
            if col == 'Nom':
                width = 150
            self.rules_tree.column(col, width=width)
        
        # Scrollbars
        vsb = ttk.Scrollbar(list_container, orient="vertical", command=self.rules_tree.yview)
        hsb = ttk.Scrollbar(list_container, orient="horizontal", command=self.rules_tree.xview)
        self.rules_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.rules_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        list_container.grid_rowconfigure(0, weight=1)
        list_container.grid_columnconfigure(0, weight=1)
        
        # Boutons de gestion des règles
        rules_btns = tk.Frame(rules_list_frame, bg='white')
        rules_btns.pack(fill='x', padx=10, pady=10)
        
        tk.Button(rules_btns, text=" ✏️ Modifier ",
                 bg='#f39c12', fg='white',
                 font=("Arial", 10, "bold"),
                 command=self.edit_rule).pack(side='left', padx=5)
        
        tk.Button(rules_btns, text=" 📋 Dupliquer ",
                 bg='#3498db', fg='white',
                 font=("Arial", 10, "bold"),
                 command=self.duplicate_rule).pack(side='left', padx=5)
        
        tk.Button(rules_btns, text=" 🗑️ Supprimer ",
                 bg='#e74c3c', fg='white',
                 font=("Arial", 10, "bold"),
                 command=self.delete_rule).pack(side='left', padx=5)
        
        tk.Button(rules_btns, text=" 📥 Importer ",
                 bg='#95a5a6', fg='white',
                 font=("Arial", 10, "bold"),
                 command=self.import_rules).pack(side='left', padx=20)
        
        tk.Button(rules_btns, text=" 📤 Exporter ",
                 bg='#95a5a6', fg='white',
                 font=("Arial", 10, "bold"),
                 command=self.export_rules).pack(side='left', padx=5)
    
    def setup_rule_chains_tab(self, notebook):
        """Onglet pour créer des chaînes de règles"""
        chains_frame = tk.Frame(notebook, bg='white')
        notebook.add(chains_frame, text=' ⛓️ Chaînes ')
        
        chains_content = tk.Frame(chains_frame, bg='white')
        chains_content.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Créer une chaîne
        create_chain_frame = tk.LabelFrame(chains_content, 
                                          text=" ⛓️ Créer une chaîne de règles ", 
                                          font=("Arial", 12, "bold"),
                                          bg='white', fg='#2c3e50', relief=tk.FLAT)
        create_chain_frame.pack(fill='x', pady=(0, 20))
        
        chain_builder = tk.Frame(create_chain_frame, bg='white')
        chain_builder.pack(pady=15, padx=15)
        
        # Nom de la chaîne
        name_frame = tk.Frame(chain_builder, bg='white')
        name_frame.pack(fill='x', pady=5)
        
        tk.Label(name_frame, text="Nom de la chaîne:", 
                font=("Arial", 11, "bold"), bg='white').pack(side='left', padx=5)
        
        self.chain_name_var = tk.StringVar()
        chain_name_entry = tk.Entry(name_frame, textvariable=self.chain_name_var,
                                    width=30, font=("Arial", 11))
        chain_name_entry.pack(side='left', padx=5)
        
        tk.Label(name_frame, text="Priorité globale:", 
                font=("Arial", 11, "bold"), bg='white').pack(side='left', padx=10)
        
        self.chain_priority_var = tk.StringVar(value="50")
        chain_priority_spinbox = tk.Spinbox(name_frame, from_=1, to=100,
                                           textvariable=self.chain_priority_var,
                                           width=8, font=("Arial", 11))
        chain_priority_spinbox.pack(side='left', padx=5)
        
        # Sélection des règles
        selection_frame = tk.Frame(chain_builder, bg='white')
        selection_frame.pack(fill='both', expand=True, pady=10)
        
        # Règles disponibles
        tk.Label(selection_frame, text="Règles disponibles:", 
                font=("Arial", 11, "bold"), bg='white').pack(anchor='w')
        
        available_frame = tk.Frame(selection_frame, bg='white')
        available_frame.pack(fill='both', expand=True)
        
        self.available_rules_listbox = tk.Listbox(available_frame, height=6,
                                                  font=("Arial", 10))
        self.available_rules_listbox.pack(side='left', fill='both', expand=True)
        
        # Boutons de transfert
        transfer_btns = tk.Frame(selection_frame, bg='white')
        transfer_btns.pack(side='left', padx=10)
        
        tk.Button(transfer_btns, text=" → ",
                 command=self.add_rule_to_chain,
                 font=("Arial", 12, "bold")).pack(pady=5)
        
        tk.Button(transfer_btns, text=" ← ",
                 command=self.remove_rule_from_chain,
                 font=("Arial", 12, "bold")).pack(pady=5)
        
        # Règles dans la chaîne
        tk.Label(selection_frame, text="Règles dans la chaîne (ordre d'exécution):", 
                font=("Arial", 11, "bold"), bg='white').pack(anchor='w')
        
        chain_frame = tk.Frame(selection_frame, bg='white')
        chain_frame.pack(fill='both', expand=True)
        
        self.chain_rules_listbox = tk.Listbox(chain_frame, height=6,
                                              font=("Arial", 10))
        self.chain_rules_listbox.pack(side='left', fill='both', expand=True)
        
        # Boutons d'ordre
        order_btns = tk.Frame(chain_frame, bg='white')
        order_btns.pack(side='left', padx=5)
        
        tk.Button(order_btns, text=" ↑ ",
                 command=self.move_chain_rule_up,
                 font=("Arial", 10)).pack(pady=2)
        
        tk.Button(order_btns, text=" ↓ ",
                 command=self.move_chain_rule_down,
                 font=("Arial", 10)).pack(pady=2)
        
        # Options de la chaîne
        chain_options = tk.Frame(chain_builder, bg='white')
        chain_options.pack(fill='x', pady=10)
        
        self.chain_stop_on_match_var = tk.BooleanVar(value=True)
        tk.Checkbutton(chain_options, 
                      text=" 🛑 Arrêter la chaîne à la première règle correspondante",
                      variable=self.chain_stop_on_match_var,
                      font=("Arial", 10), bg='white').pack(anchor='w')
        
        self.chain_enabled_var = tk.BooleanVar(value=True)
        tk.Checkbutton(chain_options, 
                      text=" ✅ Chaîne active",
                      variable=self.chain_enabled_var,
                      font=("Arial", 10), bg='white').pack(anchor='w')
        
        # Bouton créer chaîne
        create_chain_btn = tk.Button(chain_builder, text=" ⛓️ Créer la chaîne ",
                                    bg='#27ae60', fg='white',
                                    font=("Arial", 11, "bold"),
                                    command=self.create_rule_chain,
                                    cursor='hand2')
        create_chain_btn.pack(pady=10)
        
        # Liste des chaînes
        chains_list_frame = tk.LabelFrame(chains_content, 
                                         text=" 📋 Chaînes de règles actives ", 
                                         font=("Arial", 12, "bold"),
                                         bg='white', fg='#2c3e50', relief=tk.FLAT)
        chains_list_frame.pack(fill='both', expand=True)
        
        # Treeview pour les chaînes
        self.chains_tree = ttk.Treeview(chains_list_frame, height=6)
        self.chains_tree.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Boutons de gestion des chaînes
        chains_btns = tk.Frame(chains_list_frame, bg='white')
        chains_btns.pack(fill='x', padx=10, pady=10)
        
        tk.Button(chains_btns, text=" ✏️ Modifier ",
                 bg='#f39c12', fg='white',
                 font=("Arial", 10, "bold"),
                 command=self.edit_chain).pack(side='left', padx=5)
        
        tk.Button(chains_btns, text=" 🗑️ Supprimer ",
                 bg='#e74c3c', fg='white',
                 font=("Arial", 10, "bold"),
                 command=self.delete_chain).pack(side='left', padx=5)
        
        tk.Button(chains_btns, text=" 🔄 Activer/Désactiver ",
                 bg='#3498db', fg='white',
                 font=("Arial", 10, "bold"),
                 command=self.toggle_chain).pack(side='left', padx=5)
    
    def setup_folders_tab(self, notebook):
        """Onglet pour gérer les dossiers existants"""
        folders_frame = tk.Frame(notebook, bg='white')
        notebook.add(folders_frame, text=' 📁 Dossiers ')
        
        folders_content = tk.Frame(folders_frame, bg='white')
        folders_content.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Configuration des dossiers à trier
        folders_config = tk.LabelFrame(folders_content, 
                                      text=" 📁 Gestion des dossiers existants ", 
                                      font=("Arial", 12, "bold"),
                                      bg='white', fg='#2c3e50', relief=tk.FLAT)
        folders_config.pack(fill='both', expand=True)
        
        # Instructions
        instructions = tk.Frame(folders_config, bg='#e8f4f8')
        instructions.pack(fill='x', padx=10, pady=10)
        
        tk.Label(instructions, 
                text="📌 Configurez ici les dossiers existants que vous souhaitez organiser automatiquement",
                font=("Arial", 11, "bold"), bg='#e8f4f8', fg='#0a84ff').pack(anchor='w', padx=10, pady=5)
        
        tk.Label(instructions,
                text="• Sélectionnez les dossiers à analyser en plus de la boîte de réception\n"
                     "• Les règles définies s'appliqueront aussi à ces dossiers\n"
                     "• Utile pour réorganiser des emails déjà triés manuellement",
                font=("Arial", 10), bg='#e8f4f8', fg='#2c3e50',
                justify='left').pack(anchor='w', padx=20, pady=5)
        
        # Frame principal
        main_folders_frame = tk.Frame(folders_config, bg='white')
        main_folders_frame.pack(fill='both', expand=True, padx=15, pady=15)
        
        # Liste des dossiers disponibles
        available_label = tk.Label(main_folders_frame, 
                                  text="Dossiers disponibles sur le serveur:",
                                  font=("Arial", 11, "bold"), bg='white')
        available_label.pack(anchor='w', pady=5)
        
        # Frame pour la liste avec scrollbar
        list_frame = tk.Frame(main_folders_frame, bg='white')
        list_frame.pack(fill='both', expand=True, pady=10)
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side='right', fill='y')
        
        self.folders_listbox = tk.Listbox(list_frame, 
                                         selectmode='multiple',
                                         yscrollcommand=scrollbar.set,
                                         font=("Arial", 10),
                                         height=10)
        self.folders_listbox.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=self.folders_listbox.yview)
        
        # Options
        options_frame = tk.Frame(main_folders_frame, bg='white')
        options_frame.pack(fill='x', pady=10)
        
        self.include_inbox_var = tk.BooleanVar(value=True)
        tk.Checkbutton(options_frame, 
                      text=" ✅ Inclure la boîte de réception (INBOX)",
                      variable=self.include_inbox_var,
                      font=("Arial", 11, "bold"),
                      bg='white', fg='#27ae60').pack(anchor='w', pady=5)
        
        self.scan_subfolders_var = tk.BooleanVar(value=False)
        tk.Checkbutton(options_frame, 
                      text=" 📂 Analyser aussi les sous-dossiers",
                      variable=self.scan_subfolders_var,
                      font=("Arial", 10),
                      bg='white').pack(anchor='w', pady=5)
        
        self.exclude_special_var = tk.BooleanVar(value=True)
        tk.Checkbutton(options_frame, 
                      text=" 🚫 Exclure les dossiers spéciaux (Brouillons, Envoyés, Corbeille, Spam)",
                      variable=self.exclude_special_var,
                      font=("Arial", 10),
                      bg='white').pack(anchor='w', pady=5)
        
        # Boutons d'action
        buttons_frame = tk.Frame(main_folders_frame, bg='white')
        buttons_frame.pack(fill='x', pady=15)
        
        tk.Button(buttons_frame, text=" 🔄 Actualiser la liste ",
                 bg='#3498db', fg='white',
                 font=("Arial", 11, "bold"),
                 command=self.refresh_folders_list,
                 cursor='hand2').pack(side='left', padx=5)
        
        tk.Button(buttons_frame, text=" ✅ Sélectionner tout ",
                 bg='#27ae60', fg='white',
                 font=("Arial", 11, "bold"),
                 command=self.select_all_folders,
                 cursor='hand2').pack(side='left', padx=5)
        
        tk.Button(buttons_frame, text=" ❌ Désélectionner tout ",
                 bg='#e74c3c', fg='white',
                 font=("Arial", 11, "bold"),
                 command=self.deselect_all_folders,
                 cursor='hand2').pack(side='left', padx=5)
        
        # Dossiers sélectionnés
        selected_frame = tk.Frame(main_folders_frame, bg='#d4edda')
        selected_frame.pack(fill='x', pady=10)
        
        tk.Label(selected_frame, text="📋 Dossiers sélectionnés pour le tri:", 
                font=("Arial", 11, "bold"), 
                bg='#d4edda', fg='#155724').pack(anchor='w', padx=10, pady=5)
        
        self.selected_folders_label = tk.Label(selected_frame, 
                                              text="Aucun dossier sélectionné",
                                              font=("Arial", 10), 
                                              bg='#d4edda', fg='#155724',
                                              wraplength=600,
                                              justify='left')
        self.selected_folders_label.pack(anchor='w', padx=20, pady=5)
    
    def setup_advanced_tab(self, notebook):
        """Onglet des options avancées"""
        adv_frame = tk.Frame(notebook, bg='white')
        notebook.add(adv_frame, text=' ⚙️ Avancé ')
        
        adv_content = tk.Frame(adv_frame, bg='white')
        adv_content.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Options de traitement
        processing_frame = tk.LabelFrame(adv_content, 
                                        text=" 🔧 Options de traitement ", 
                                        font=("Arial", 12, "bold"),
                                        bg='white', fg='#2c3e50', relief=tk.FLAT)
        processing_frame.pack(fill='x', pady=(0, 20))
        
        proc_inner = tk.Frame(processing_frame, bg='white')
        proc_inner.pack(pady=15, padx=15)
        
        # Mode de traitement
        mode_frame = tk.Frame(proc_inner, bg='white')
        mode_frame.pack(fill='x', pady=5)
        
        tk.Label(mode_frame, text="Mode de traitement:", 
                font=("Arial", 11), bg='white').pack(side='left', padx=5)
        
        self.processing_mode_var = tk.StringVar(value="peek")
        tk.Radiobutton(mode_frame, text="PEEK (Ne pas marquer comme lu)",
                      variable=self.processing_mode_var, value="peek",
                      font=("Arial", 10), bg='white').pack(side='left', padx=10)
        
        tk.Radiobutton(mode_frame, text="BODY (Peut marquer comme lu)",
                      variable=self.processing_mode_var, value="body",
                      font=("Arial", 10), bg='white').pack(side='left', padx=10)
        
        # Options de filtrage
        filter_frame = tk.Frame(proc_inner, bg='white')
        filter_frame.pack(fill='x', pady=10)
        
        tk.Label(filter_frame, text="Filtrer les emails:", 
                font=("Arial", 11), bg='white').pack(anchor='w', pady=5)
        
        self.filter_unread_only_var = tk.BooleanVar(value=False)
        tk.Checkbutton(filter_frame, text=" 📧 Traiter uniquement les emails non lus",
                      variable=self.filter_unread_only_var,
                      font=("Arial", 10), bg='white').pack(anchor='w', padx=20, pady=2)
        
        self.filter_date_var = tk.BooleanVar(value=False)
        date_check = tk.Checkbutton(filter_frame, text=" 📅 Traiter uniquement les emails des",
                                   variable=self.filter_date_var,
                                   font=("Arial", 10), bg='white',
                                   command=self.toggle_date_filter)
        date_check.pack(anchor='w', padx=20, pady=2)
        
        date_frame = tk.Frame(filter_frame, bg='white')
        date_frame.pack(anchor='w', padx=40, pady=2)
        
        self.filter_days_var = tk.StringVar(value="7")
        self.days_spinbox = tk.Spinbox(date_frame, from_=1, to=365, 
                                       textvariable=self.filter_days_var,
                                       width=5, font=("Arial", 10), state='disabled')
        self.days_spinbox.pack(side='left', padx=5)
        
        tk.Label(date_frame, text="derniers jours", 
                font=("Arial", 10), bg='white').pack(side='left')
        
        # Options de sécurité
        safety_frame = tk.LabelFrame(adv_content, 
                                    text=" 🔒 Options de sécurité ", 
                                    font=("Arial", 12, "bold"),
                                    bg='white', fg='#2c3e50', relief=tk.FLAT)
        safety_frame.pack(fill='x', pady=(0, 20))
        
        safety_inner = tk.Frame(safety_frame, bg='white')
        safety_inner.pack(pady=15, padx=15)
        
        self.dry_run_var = tk.BooleanVar(value=False)
        tk.Checkbutton(safety_inner, 
                      text=" 🧪 Mode test (simuler sans déplacer les emails)",
                      variable=self.dry_run_var,
                      font=("Arial", 11, "bold"),
                      bg='white', fg='#e74c3c').pack(anchor='w', pady=5)
        
        self.backup_before_move_var = tk.BooleanVar(value=False)
        tk.Checkbutton(safety_inner, 
                      text=" 💾 Créer une copie de sauvegarde avant déplacement",
                      variable=self.backup_before_move_var,
                      font=("Arial", 10),
                      bg='white').pack(anchor='w', pady=5)
        
        self.confirm_actions_var = tk.BooleanVar(value=False)
        tk.Checkbutton(safety_inner, 
                      text=" ❓ Demander confirmation pour chaque action",
                      variable=self.confirm_actions_var,
                      font=("Arial", 10),
                      bg='white').pack(anchor='w', pady=5)
        
        # Performances
        perf_frame = tk.LabelFrame(adv_content, 
                                  text=" ⚡ Options de performance ", 
                                  font=("Arial", 12, "bold"),
                                  bg='white', fg='#2c3e50', relief=tk.FLAT)
        perf_frame.pack(fill='x', pady=(0, 20))
        
        perf_inner = tk.Frame(perf_frame, bg='white')
        perf_inner.pack(pady=15, padx=15)
        
        batch_frame = tk.Frame(perf_inner, bg='white')
        batch_frame.pack(fill='x', pady=5)
        
        tk.Label(batch_frame, text="Traiter par lots de:", 
                font=("Arial", 11), bg='white').pack(side='left', padx=5)
        
        self.batch_size_var = tk.StringVar(value="50")
        tk.Spinbox(batch_frame, from_=10, to=500, increment=10,
                  textvariable=self.batch_size_var,
                  width=8, font=("Arial", 11)).pack(side='left', padx=5)
        
        tk.Label(batch_frame, text="emails", 
                font=("Arial", 11), bg='white').pack(side='left')
        
        self.parallel_processing_var = tk.BooleanVar(value=False)
        tk.Checkbutton(perf_inner, 
                      text=" 🚀 Activer le traitement parallèle (expérimental)",
                      variable=self.parallel_processing_var,
                      font=("Arial", 10),
                      bg='white').pack(anchor='w', pady=5)
    
    def setup_execution_tab(self, notebook):
        """Onglet d'exécution"""
        run_frame = tk.Frame(notebook, bg='white')
        notebook.add(run_frame, text=' ▶️ Exécution ')
        
        run_content = tk.Frame(run_frame, bg='white')
        run_content.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Options d'exécution
        exec_options = tk.LabelFrame(run_content, 
                                    text=" ⚙️ Options d'analyse ", 
                                    font=("Arial", 12, "bold"),
                                    bg='white', fg='#2c3e50', relief=tk.FLAT)
        exec_options.pack(fill='x', pady=(0, 20))
        
        options_grid = tk.Frame(exec_options, bg='white')
        options_grid.pack(pady=15, padx=15)
        
        tk.Label(options_grid, text="Nombre max d'emails à traiter:", 
                font=("Arial", 11), bg='white').pack(side='left', padx=5)
        
        self.max_emails_var = tk.StringVar(value="100")
        tk.Spinbox(options_grid, from_=1, to=10000, 
                  textvariable=self.max_emails_var,
                  width=10, font=("Arial", 11)).pack(side='left', padx=5)
        
        tk.Label(options_grid, text="(0 = tous les emails)", 
                font=("Arial", 9, "italic"), bg='white', fg='gray').pack(side='left')
        
        # Statistiques actuelles
        stats_frame = tk.Frame(run_content, bg='#e8f4f8', relief=tk.RIDGE, bd=2)
        stats_frame.pack(fill='x', pady=10)
        
        tk.Label(stats_frame, text="📊 Statistiques de session", 
                font=("Arial", 11, "bold"), 
                bg='#e8f4f8', fg='#0a84ff').pack(anchor='w', padx=10, pady=(10, 5))
        
        self.stats_label = tk.Label(stats_frame, 
                                   text="Aucune analyse effectuée pour le moment",
                                   font=("Arial", 10), 
                                   bg='#e8f4f8', fg='#2c3e50')
        self.stats_label.pack(anchor='w', padx=25, pady=(0, 10))
        
        # Gros bouton analyser
        analyze_frame = tk.Frame(run_content, bg='white')
        analyze_frame.pack(fill='x', pady=20)
        
        self.analyze_btn = tk.Button(analyze_frame, 
                                     text="🚀 ANALYSER ET TRIER LES EMAILS",
                                     font=("Arial", 18, "bold"),
                                     bg='#e74c3c', fg='white',
                                     padx=40, pady=20,
                                     command=self.start_analysis,
                                     cursor='hand2',
                                     relief=tk.RAISED, bd=3)
        self.analyze_btn.pack()
        
        # Console de logs
        console_frame = tk.LabelFrame(run_content, 
                                     text=" 📊 Console d'exécution ", 
                                     font=("Arial", 12, "bold"),
                                     bg='white', fg='#2c3e50', relief=tk.FLAT)
        console_frame.pack(fill='both', expand=True)
        
        self.console = scrolledtext.ScrolledText(console_frame, 
                                                 height=10,
                                                 bg='#1e1e1e', fg='#00ff00',
                                                 font=("Consolas", 10),
                                                 insertbackground='#00ff00')
        self.console.pack(fill='both', expand=True, padx=10, pady=10)
    
    def toggle_cc_options(self):
        """Activer/désactiver les options CC"""
        if self.cc_enabled_var.get():
            self.cc_folder_entry.config(state='normal')
            self.log("✅ Tri automatique des CC activé", "success")
        else:
            self.cc_folder_entry.config(state='disabled')
            self.log("❌ Tri automatique des CC désactivé", "warning")
    
    def toggle_date_filter(self):
        """Activer/désactiver le filtre de date"""
        if self.filter_date_var.get():
            self.days_spinbox.config(state='normal')
        else:
            self.days_spinbox.config(state='disabled')
    
    def on_action_changed(self, event=None):
        """Gérer le changement d'action dans les règles"""
        action = self.rule_action_var.get()
        if action in ["Déplacer vers", "Copier vers"]:
            self.folder_entry.config(state='normal')
            self.folder_dropdown.config(state='readonly')
            self.mark_checkbox.config(state='normal')
        else:
            self.folder_entry.config(state='disabled')
            self.folder_dropdown.config(state='disabled')
            if action == "Marquer comme lu":
                self.mark_checkbox.config(state='disabled')
                self.rule_mark_after_move_var.set(False)
    
    def on_folder_selected(self, event=None):
        """Quand un dossier est sélectionné dans le dropdown"""
        selected = self.folder_dropdown.get()
        if selected:
            self.rule_folder_var.set(selected)
    
    def load_existing_folders(self):
        """Charger la liste des dossiers existants depuis le serveur"""
        try:
            if not self.server_var.get() or not self.email_var.get() or not self.password_var.get():
                messagebox.showwarning("Attention", "Configurez d'abord la connexion!")
                return
            
            self.log("📁 Chargement des dossiers...", "info")
            
            connection = imaplib.IMAP4_SSL(self.server_var.get(), int(self.port_var.get()))
            connection.login(self.email_var.get(), self.password_var.get())
            
            result, folders = connection.list()
            
            if result == 'OK':
                self.existing_folders = []
                for folder in folders:
                    if folder:
                        folder_str = folder.decode('utf-8') if isinstance(folder, bytes) else str(folder)
                        
                        # Parser le format IMAP pour extraire le nom du dossier
                        # Format typique: '(\\HasNoChildren) "." "INBOX.Dossier"'
                        # ou: '(\\HasChildren) "/" "Folder Name"'
                        
                        # Méthode améliorée pour extraire le nom
                        folder_name = None
                        
                        # Essayer d'abord avec des guillemets
                        if '"' in folder_str:
                            parts = folder_str.split('"')
                            if len(parts) >= 2:
                                # Le nom est généralement le dernier élément entre guillemets
                                folder_name = parts[-2]
                        
                        # Si pas de guillemets ou échec, essayer avec des espaces
                        if not folder_name or folder_name in ['.', '/', '|']:
                            parts = folder_str.split()
                            if len(parts) >= 3:
                                # Le nom est généralement le dernier élément
                                folder_name = parts[-1].strip('"')
                        
                        # Si toujours pas de nom valide, méthode simple sans regex
                        if not folder_name or folder_name in ['.', '/', '|']:
                            # Prendre le dernier élément après un espace
                            parts = folder_str.rsplit(' ', 1)
                            if len(parts) > 1:
                                folder_name = parts[-1].strip('"\'')
                            else:
                                # Si pas d'espace, prendre toute la chaîne nettoyée
                                folder_name = folder_str.strip('"\' ')
                        
                        # Nettoyer le nom
                        if folder_name:
                            folder_name = folder_name.strip().strip('"')
                            if folder_name and folder_name not in ['.', '/', '|', '']:
                                self.existing_folders.append(folder_name)
                                self.log(f"  • Dossier trouvé: {folder_name}", "info")
                
                # Retirer les doublons tout en préservant l'ordre
                seen = set()
                unique_folders = []
                for folder in self.existing_folders:
                    if folder not in seen:
                        seen.add(folder)
                        unique_folders.append(folder)
                self.existing_folders = unique_folders
                
                # Mettre à jour les widgets
                self.folder_dropdown['values'] = self.existing_folders
                self.folders_listbox.delete(0, tk.END)
                for folder in self.existing_folders:
                    self.folders_listbox.insert(tk.END, folder)
                
                # Actualiser la liste des règles disponibles pour les chaînes
                self.refresh_available_rules()
                
                self.log(f"✅ {len(self.existing_folders)} dossiers chargés", "success")
                messagebox.showinfo("Succès", f"{len(self.existing_folders)} dossiers chargés avec succès!")
            
            connection.logout()
            
        except Exception as e:
            self.log(f"❌ Erreur: {str(e)}", "error")
            messagebox.showerror("Erreur", f"Impossible de charger les dossiers:\n{str(e)}")
    
    def refresh_folders_list(self):
        """Actualiser la liste des dossiers"""
        self.load_existing_folders()
    
    def select_all_folders(self):
        """Sélectionner tous les dossiers"""
        self.folders_listbox.select_set(0, tk.END)
        self.update_selected_folders_label()
    
    def deselect_all_folders(self):
        """Désélectionner tous les dossiers"""
        self.folders_listbox.select_clear(0, tk.END)
        self.update_selected_folders_label()
    
    def update_selected_folders_label(self):
        """Mettre à jour le label des dossiers sélectionnés"""
        selected_indices = self.folders_listbox.curselection()
        if selected_indices:
            selected = [self.folders_listbox.get(i) for i in selected_indices]
            self.selected_folders_label.config(text=", ".join(selected))
        else:
            self.selected_folders_label.config(text="Aucun dossier sélectionné")
    
    def add_custom_rule(self):
        """Ajouter une règle personnalisée améliorée"""
        name = self.rule_name_var.get().strip()
        keyword = self.rule_keyword_var.get().strip()
        action = self.rule_action_var.get()
        
        if not name:
            messagebox.showwarning("Attention", "Veuillez donner un nom à la règle!")
            return
        
        if not keyword:
            messagebox.showwarning("Attention", "Veuillez entrer un mot-clé ou une expression!")
            return
        
        if action in ["Déplacer vers", "Copier vers"] and not self.rule_folder_var.get().strip():
            messagebox.showwarning("Attention", "Veuillez spécifier le dossier de destination!")
            return
        
        try:
            priority = int(self.rule_priority_var.get())
        except:
            priority = 50
        
        rule = {
            "name": name,
            "field": self.rule_field_var.get(),
            "condition": self.rule_condition_var.get(),
            "keyword": keyword,
            "and_field": self.rule_and_field_var.get(),
            "and_condition": self.rule_and_condition_var.get(),
            "and_keyword": self.rule_and_keyword_var.get(),
            "action": action,
            "folder": self.rule_folder_var.get().strip() if action in ["Déplacer vers", "Copier vers"] else "",
            "case_sensitive": self.rule_case_sensitive_var.get(),
            "priority": priority,
            "stop_processing": self.rule_stop_processing_var.get(),
            "continue_chain": self.rule_continue_chain_var.get(),
            "mark_after_action": self.rule_mark_after_move_var.get()
        }
        
        self.rules.append(rule)
        self.sort_rules_by_priority()
        self.refresh_rules_tree()
        self.refresh_available_rules()
        
        # Réinitialiser les champs
        self.rule_name_var.set("")
        self.rule_keyword_var.set("")
        self.rule_and_keyword_var.set("")
        self.rule_folder_var.set("")
        self.rule_case_sensitive_var.set(False)
        self.rule_priority_var.set("50")
        self.rule_stop_processing_var.set(False)
        self.rule_continue_chain_var.set(False)
        self.rule_mark_after_move_var.set(False)
        
        self.log(f"✅ Règle ajoutée: {name}", "success")
        self.save_settings()
    
    def sort_rules_by_priority(self):
        """Trier les règles par priorité (nombre croissant = priorité décroissante)"""
        self.rules.sort(key=lambda x: x.get('priority', 50))
    
    def refresh_rules_tree(self):
        """Actualiser l'affichage des règles dans le TreeView"""
        # Effacer l'arbre
        for item in self.rules_tree.get_children():
            self.rules_tree.delete(item)
        
        # Ajouter les règles
        for i, rule in enumerate(self.rules, 1):
            options = []
            if rule.get('case_sensitive'):
                options.append("Casse")
            if rule.get('stop_processing'):
                options.append("Stop")
            if rule.get('continue_chain'):
                options.append("Chaîne")
            if rule.get('mark_after_action'):
                options.append("Marquer")
            
            # Condition complète
            condition = f"{rule['field']} {rule['condition']} '{rule['keyword'][:20]}'"
            if rule.get('and_field') and rule.get('and_keyword'):
                condition += f" ET {rule['and_field']} {rule['and_condition']} '{rule['and_keyword'][:15]}'"
            
            values = (
                rule.get('name', 'Sans nom'),
                rule.get('priority', 50),
                rule['field'],
                rule['condition'],
                rule['keyword'][:30] + ('...' if len(rule['keyword']) > 30 else ''),
                f"{rule['action']} {rule.get('folder', '')}".strip(),
                ', '.join(options)
            )
            
            self.rules_tree.insert('', 'end', text=str(i), values=values)
    
    def refresh_available_rules(self):
        """Actualiser la liste des règles disponibles pour les chaînes"""
        if hasattr(self, 'available_rules_listbox'):
            self.available_rules_listbox.delete(0, tk.END)
            for rule in self.rules:
                self.available_rules_listbox.insert(tk.END, f"{rule.get('name', 'Sans nom')} (P:{rule.get('priority', 50)})")
    
    def edit_rule(self):
        """Modifier une règle existante"""
        selection = self.rules_tree.selection()
        if selection:
            item = selection[0]
            index = self.rules_tree.index(item)
            rule = self.rules[index]
            
            # Remplir les champs avec les valeurs de la règle
            self.rule_name_var.set(rule.get('name', ''))
            self.rule_field_var.set(rule['field'])
            self.rule_condition_var.set(rule['condition'])
            self.rule_keyword_var.set(rule['keyword'])
            self.rule_and_field_var.set(rule.get('and_field', ''))
            self.rule_and_condition_var.set(rule.get('and_condition', 'contient'))
            self.rule_and_keyword_var.set(rule.get('and_keyword', ''))
            self.rule_action_var.set(rule.get('action', 'Déplacer vers'))
            self.rule_folder_var.set(rule.get('folder', ''))
            self.rule_case_sensitive_var.set(rule.get('case_sensitive', False))
            self.rule_priority_var.set(str(rule.get('priority', 50)))
            self.rule_stop_processing_var.set(rule.get('stop_processing', False))
            self.rule_continue_chain_var.set(rule.get('continue_chain', False))
            self.rule_mark_after_move_var.set(rule.get('mark_after_action', False))
            
            # Supprimer la règle de la liste
            del self.rules[index]
            self.refresh_rules_tree()
            self.refresh_available_rules()
            
            self.log("✏️ Règle en cours de modification", "info")
        else:
            messagebox.showinfo("Info", "Sélectionnez une règle à modifier")
    
    def duplicate_rule(self):
        """Dupliquer une règle"""
        selection = self.rules_tree.selection()
        if selection:
            item = selection[0]
            index = self.rules_tree.index(item)
            rule = copy.deepcopy(self.rules[index])
            rule['name'] = rule.get('name', 'Sans nom') + " (copie)"
            self.rules.append(rule)
            self.sort_rules_by_priority()
            self.refresh_rules_tree()
            self.refresh_available_rules()
            self.save_settings()
            self.log("📋 Règle dupliquée", "success")
        else:
            messagebox.showinfo("Info", "Sélectionnez une règle à dupliquer")
    
    def delete_rule(self):
        """Supprimer la règle sélectionnée"""
        selection = self.rules_tree.selection()
        if selection:
            item = selection[0]
            index = self.rules_tree.index(item)
            
            if messagebox.askyesno("Confirmation", "Supprimer cette règle ?"):
                del self.rules[index]
                self.refresh_rules_tree()
                self.refresh_available_rules()
                self.save_settings()
                self.log("🗑️ Règle supprimée", "warning")
        else:
            messagebox.showinfo("Info", "Sélectionnez une règle à supprimer")
    
    def import_rules(self):
        """Importer des règles depuis un fichier"""
        filename = filedialog.askopenfilename(
            title="Importer des règles",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    imported_rules = json.load(f)
                
                if isinstance(imported_rules, list):
                    self.rules.extend(imported_rules)
                    self.sort_rules_by_priority()
                    self.refresh_rules_tree()
                    self.refresh_available_rules()
                    self.save_settings()
                    self.log(f"✅ {len(imported_rules)} règles importées", "success")
                    messagebox.showinfo("Succès", f"{len(imported_rules)} règles importées avec succès!")
                else:
                    messagebox.showerror("Erreur", "Format de fichier invalide")
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible d'importer les règles:\n{str(e)}")
    
    def export_rules(self):
        """Exporter les règles vers un fichier"""
        if not self.rules:
            messagebox.showinfo("Info", "Aucune règle à exporter")
            return
        
        filename = filedialog.asksaveasfilename(
            title="Exporter les règles",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(self.rules, f, indent=4, ensure_ascii=False)
                
                self.log(f"✅ {len(self.rules)} règles exportées", "success")
                messagebox.showinfo("Succès", f"{len(self.rules)} règles exportées avec succès!")
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible d'exporter les règles:\n{str(e)}")
    
    # Méthodes pour les chaînes de règles
    def add_rule_to_chain(self):
        """Ajouter une règle à la chaîne en cours"""
        selection = self.available_rules_listbox.curselection()
        if selection:
            rule_text = self.available_rules_listbox.get(selection[0])
            self.chain_rules_listbox.insert(tk.END, rule_text)
    
    def remove_rule_from_chain(self):
        """Retirer une règle de la chaîne"""
        selection = self.chain_rules_listbox.curselection()
        if selection:
            self.chain_rules_listbox.delete(selection[0])
    
    def move_chain_rule_up(self):
        """Monter une règle dans la chaîne"""
        selection = self.chain_rules_listbox.curselection()
        if selection and selection[0] > 0:
            index = selection[0]
            item = self.chain_rules_listbox.get(index)
            self.chain_rules_listbox.delete(index)
            self.chain_rules_listbox.insert(index - 1, item)
            self.chain_rules_listbox.select_set(index - 1)
    
    def move_chain_rule_down(self):
        """Descendre une règle dans la chaîne"""
        selection = self.chain_rules_listbox.curselection()
        if selection and selection[0] < self.chain_rules_listbox.size() - 1:
            index = selection[0]
            item = self.chain_rules_listbox.get(index)
            self.chain_rules_listbox.delete(index)
            self.chain_rules_listbox.insert(index + 1, item)
            self.chain_rules_listbox.select_set(index + 1)
    
    def create_rule_chain(self):
        """Créer une chaîne de règles"""
        name = self.chain_name_var.get().strip()
        if not name:
            messagebox.showwarning("Attention", "Donnez un nom à la chaîne!")
            return
        
        # Récupérer les règles de la chaîne
        chain_rules = []
        for i in range(self.chain_rules_listbox.size()):
            rule_text = self.chain_rules_listbox.get(i)
            # Extraire le nom de la règle
            rule_name = rule_text.split(" (P:")[0]
            # Trouver la règle correspondante
            for rule in self.rules:
                if rule.get('name') == rule_name:
                    chain_rules.append(rule)
                    break
        
        if not chain_rules:
            messagebox.showwarning("Attention", "Ajoutez au moins une règle à la chaîne!")
            return
        
        try:
            priority = int(self.chain_priority_var.get())
        except:
            priority = 50
        
        chain = {
            "name": name,
            "priority": priority,
            "rules": chain_rules,
            "stop_on_match": self.chain_stop_on_match_var.get(),
            "enabled": self.chain_enabled_var.get()
        }
        
        self.rule_chains.append(chain)
        self.refresh_chains_tree()
        self.save_chains()
        
        # Réinitialiser
        self.chain_name_var.set("")
        self.chain_priority_var.set("50")
        self.chain_rules_listbox.delete(0, tk.END)
        
        self.log(f"⛓️ Chaîne créée: {name} avec {len(chain_rules)} règles", "success")
    
    def refresh_chains_tree(self):
        """Actualiser l'affichage des chaînes"""
        # Effacer l'arbre
        for item in self.chains_tree.get_children():
            self.chains_tree.delete(item)
        
        # Ajouter les chaînes
        for chain in self.rule_chains:
            status = "✅ Active" if chain.get('enabled', True) else "❌ Inactive"
            rules_count = len(chain.get('rules', []))
            text = f"{chain['name']} - {status} - {rules_count} règles - Priorité: {chain.get('priority', 50)}"
            
            parent = self.chains_tree.insert('', 'end', text=text, open=False)
            
            # Ajouter les règles de la chaîne comme enfants
            for rule in chain.get('rules', []):
                rule_text = f"→ {rule.get('name', 'Sans nom')}: {rule['field']} {rule['condition']} '{rule['keyword'][:30]}'"
                self.chains_tree.insert(parent, 'end', text=rule_text)
    
    def edit_chain(self):
        """Modifier une chaîne"""
        selection = self.chains_tree.selection()
        if selection:
            # À implémenter si nécessaire
            messagebox.showinfo("Info", "Fonction en cours de développement")
    
    def delete_chain(self):
        """Supprimer une chaîne"""
        selection = self.chains_tree.selection()
        if selection:
            item = selection[0]
            # Trouver l'index de la chaîne
            text = self.chains_tree.item(item)['text']
            chain_name = text.split(" - ")[0]
            
            if messagebox.askyesno("Confirmation", f"Supprimer la chaîne '{chain_name}' ?"):
                self.rule_chains = [c for c in self.rule_chains if c['name'] != chain_name]
                self.refresh_chains_tree()
                self.save_chains()
                self.log(f"🗑️ Chaîne supprimée: {chain_name}", "warning")
    
    def toggle_chain(self):
        """Activer/désactiver une chaîne"""
        selection = self.chains_tree.selection()
        if selection:
            item = selection[0]
            text = self.chains_tree.item(item)['text']
            chain_name = text.split(" - ")[0]
            
            for chain in self.rule_chains:
                if chain['name'] == chain_name:
                    chain['enabled'] = not chain.get('enabled', True)
                    break
            
            self.refresh_chains_tree()
            self.save_chains()
            self.log(f"🔄 Chaîne basculée: {chain_name}", "info")
    
    def save_chains(self):
        """Sauvegarder les chaînes de règles"""
        try:
            with open(self.chains_file, 'w', encoding='utf-8') as f:
                json.dump(self.rule_chains, f, indent=4, ensure_ascii=False)
        except Exception as e:
            self.log(f"⚠️ Impossible de sauvegarder les chaînes: {str(e)}", "warning")
    
    def load_chains(self):
        """Charger les chaînes de règles"""
        try:
            if self.chains_file.exists():
                with open(self.chains_file, 'r', encoding='utf-8') as f:
                    self.rule_chains = json.load(f)
                self.refresh_chains_tree()
        except Exception as e:
            self.log(f"⚠️ Impossible de charger les chaînes: {str(e)}", "warning")
    
    def test_connection(self):
        """Tester la connexion au serveur avec mode PEEK"""
        self.log("\n" + "="*50, "separator")
        self.log("🔌 TEST DE CONNEXION", "header")
        self.log("="*50, "separator")
        
        try:
            server = self.server_var.get()
            if not server:
                messagebox.showerror("Erreur", "Veuillez entrer un serveur IMAP!")
                return
                
            self.log(f"Connexion à {server}:{self.port_var.get()}...", "info")
            
            connection = imaplib.IMAP4_SSL(server, int(self.port_var.get()))
            connection.login(self.email_var.get(), self.password_var.get())
            
            # Obtenir info
            result, folders = connection.list()
            
            self.log("✅ Connexion réussie!", "success")
            self.log(f"📁 {len(folders)} dossiers trouvés", "info")
            
            # Test du mode PEEK
            connection.select('INBOX', readonly=True)
            result, data = connection.search(None, 'ALL')
            if result == 'OK':
                email_count = len(data[0].split())
                self.log(f"📧 {email_count} emails dans la boîte de réception", "info")
                
                # Tester PEEK sur le premier email
                if email_count > 0:
                    first_id = data[0].split()[0]
                    result, msg_data = connection.fetch(first_id, '(BODY.PEEK[HEADER])')
                    if result == 'OK':
                        self.log("✅ Mode PEEK supporté - Les emails ne seront PAS marqués comme lus", "success")
                    else:
                        self.log("⚠️ Mode PEEK peut ne pas être supporté complètement", "warning")
            
            connection.logout()
            
            messagebox.showinfo("Succès", 
                              f"✅ Connexion réussie!\n\n"
                              f"Serveur: {server}\n"
                              f"Email: {self.email_var.get()}\n"
                              f"{len(folders)} dossiers disponibles\n"
                              f"Mode PEEK supporté ✓")
            
        except Exception as e:
            self.log(f"❌ Erreur: {str(e)}", "error")
            error_msg = str(e)
            
            if "authentication failed" in error_msg.lower():
                error_msg += "\n\n💡 Vérifiez:\n• Le serveur IMAP est correct\n• L'email et le mot de passe\n• Pour Gmail: utilisez un mot de passe d'application"
            
            messagebox.showerror("Erreur de connexion", error_msg)
    
    def get_full_folder_name(self, folder_name):
        """Obtenir le nom complet du dossier - ne pas modifier si déjà complet"""
        # Si le dossier existe déjà dans la liste, le retourner tel quel
        if folder_name in self.existing_folders:
            return folder_name
        
        # Si c'est déjà un chemin complet (contient INBOX ou commence par un séparateur)
        if "INBOX" in folder_name or folder_name.startswith(("/", ".", "\\")):
            return folder_name
        
        # Sinon, essayer avec INBOX. (pour les nouveaux dossiers)
        # Mais seulement si on n'a pas de dossiers existants pour vérifier
        if not self.existing_folders:
            return f"INBOX.{folder_name}"
        
        # Si on a des dossiers existants, analyser leur format
        for existing in self.existing_folders:
            if "INBOX." in existing:
                return f"INBOX.{folder_name}"
            elif "INBOX/" in existing:
                return f"INBOX/{folder_name}"
        
        # Par défaut, retourner tel quel
        return folder_name
    
    def create_folder_if_needed(self, connection, folder_name):
        """Créer un dossier IMAP s'il n'existe pas"""
        if not folder_name:
            return True
            
        try:
            # Utiliser le nom tel quel si c'est un dossier existant
            if folder_name in self.existing_folders:
                self.log(f"📁 Dossier '{folder_name}' déjà existant", "info")
                return True
            
            # Pour un nouveau dossier, essayer de le créer
            full_folder_name = self.get_full_folder_name(folder_name)
            
            # Lister tous les dossiers existants
            result, folders = connection.list()
            
            # Vérifier si le dossier existe sous différentes formes
            folder_exists = False
            if result == 'OK':
                for folder in folders:
                    if folder:
                        folder_str = folder.decode('utf-8') if isinstance(folder, bytes) else str(folder)
                        # Vérifier si le nom apparaît dans la chaîne
                        if folder_name.lower() in folder_str.lower() or full_folder_name.lower() in folder_str.lower():
                            folder_exists = True
                            self.log(f"📁 Dossier '{folder_name}' trouvé", "info")
                            break
            
            if not folder_exists:
                # Essayer de créer le dossier
                self.log(f"📁 Création du dossier '{full_folder_name}'...", "info")
                result = connection.create(full_folder_name)
                if result[0] == 'OK':
                    self.log(f"✅ Dossier '{full_folder_name}' créé avec succès", "success")
                    connection.subscribe(full_folder_name)
                    # Ajouter à la liste des dossiers existants
                    if full_folder_name not in self.existing_folders:
                        self.existing_folders.append(full_folder_name)
                    return True
                else:
                    # Si échec avec INBOX., essayer sans
                    if "INBOX." in full_folder_name:
                        simple_name = folder_name
                        result = connection.create(simple_name)
                        if result[0] == 'OK':
                            self.log(f"✅ Dossier '{simple_name}' créé avec succès", "success")
                            connection.subscribe(simple_name)
                            if simple_name not in self.existing_folders:
                                self.existing_folders.append(simple_name)
                            return True
                    
                    self.log(f"❌ Impossible de créer '{full_folder_name}': {result}", "error")
                    return False
            return True
                    
        except Exception as e:
            self.log(f"⚠️ Erreur avec le dossier '{folder_name}': {str(e)[:100]}", "warning")
            return False
    
    def start_analysis(self):
        """Démarrer l'analyse des emails"""
        if not self.email_var.get() or not self.password_var.get():
            messagebox.showerror("Erreur", "Configurez d'abord vos identifiants dans l'onglet Connexion!")
            return
        
        if not self.server_var.get():
            messagebox.showerror("Erreur", "Le serveur IMAP n'est pas configuré!")
            return
        
        if self.is_running:
            messagebox.showinfo("Info", "Une analyse est déjà en cours!")
            return
        
        # Avertissement mode test
        if self.dry_run_var.get():
            if not messagebox.askyesno("Mode test", 
                                       "Le mode test est activé.\n\n"
                                       "Les emails seront analysés mais PAS déplacés.\n"
                                       "Continuer ?"):
                return
        
        self.is_running = True
        self.analyze_btn.config(state='disabled', text="⏳ ANALYSE EN COURS...")
        self.status_var.set("🔄 Analyse en cours...")
        
        # Réinitialiser les emails traités
        self.processed_emails.clear()
        
        # Sauvegarder avant l'analyse
        self.save_settings()
        
        # Thread pour ne pas bloquer l'interface
        thread = threading.Thread(target=self.analysis_worker, daemon=True)
        thread.start()
    
    def analysis_worker(self):
        """Worker pour l'analyse des emails avec support des chaînes et dossiers multiples"""
        stats = {
            'total': 0,
            'processed': 0,
            'cc_moved': 0,
            'rules_applied': 0,
            'chains_applied': 0,
            'errors': 0
        }
        
        try:
            # Connexion
            self.log("\n" + "="*60, "separator")
            self.log("🚀 DÉMARRAGE DE L'ANALYSE V3", "header")
            self.log("="*60, "separator")
            
            if self.dry_run_var.get():
                self.log("🧪 MODE TEST ACTIVÉ - Aucun email ne sera déplacé", "warning")
            
            if self.preserve_unread_var.get():
                self.log("🔒 Préservation du statut non-lu activée", "success")
            
            self.log(f"🔌 Connexion à {self.server_var.get()}...", "info")
            
            connection = imaplib.IMAP4_SSL(self.server_var.get(), int(self.port_var.get()))
            connection.login(self.email_var.get(), self.password_var.get())
            
            self.log(f"✅ Connecté avec succès!", "success")
            
            # Créer les dossiers nécessaires
            folders_to_create = set()
            
            if self.cc_enabled_var.get() and self.cc_folder_var.get():
                folders_to_create.add(self.cc_folder_var.get())
            
            for rule in self.rules:
                if rule.get('action') in ['Déplacer vers', 'Copier vers'] and rule.get('folder'):
                    folders_to_create.add(rule['folder'])
            
            for folder in folders_to_create:
                self.create_folder_if_needed(connection, folder)
            
            # Déterminer les dossiers à analyser
            folders_to_process = []
            
            # Toujours inclure INBOX si configuré
            if self.include_inbox_var.get():
                folders_to_process.append('INBOX')
            
            # Ajouter les dossiers sélectionnés
            selected_indices = self.folders_listbox.curselection()
            for index in selected_indices:
                folder = self.folders_listbox.get(index)
                if folder not in folders_to_process:
                    folders_to_process.append(folder)
            
            if not folders_to_process:
                folders_to_process = ['INBOX']  # Par défaut
            
            self.log(f"📁 Dossiers à analyser: {', '.join(folders_to_process)}", "info")
            
            # Traiter chaque dossier
            for folder in folders_to_process:
                self.log(f"\n📂 Analyse du dossier: {folder}", "header")
                self.process_folder(connection, folder, stats)
            
            # Déconnexion
            connection.close()
            connection.logout()
            
            # Résumé final
            self.display_summary(stats)
            
        except Exception as e:
            self.log(f"❌ Erreur critique: {str(e)}", "error")
            self.status_var.set("❌ Erreur - Vérifiez la connexion")
            
            error_msg = str(e)
            if "authentication" in error_msg.lower():
                error_msg += "\n\n💡 Vérifiez:\n• Le serveur IMAP\n• L'email et le mot de passe"
            
            messagebox.showerror("Erreur", f"Erreur lors de l'analyse:\n\n{error_msg}")
        
        finally:
            self.is_running = False
            self.root.after(0, lambda: self.analyze_btn.config(
                state='normal',
                text="🚀 ANALYSER ET TRIER LES EMAILS"
            ))
            
            # Mettre à jour les statistiques
            self.update_stats(stats)
    
    def process_folder(self, connection, folder, stats):
        """Traiter un dossier spécifique"""
        try:
            # Sélectionner le dossier - toujours en mode normal pour pouvoir effectuer les actions
            # Le mode PEEK sera utilisé uniquement pour la récupération des emails
            connection.select(folder)
            self.log(f"📖 {folder} ouvert pour traitement", "info")
            
            # Construire la requête de recherche
            search_criteria = self.build_search_criteria()
            result, data = connection.search(None, search_criteria)
            
            if result != 'OK':
                self.log(f"❌ Erreur lors de la recherche dans {folder}", "error")
                return
            
            email_ids = data[0].split()
            folder_total = len(email_ids)
            
            if folder_total == 0:
                self.log(f"📭 Aucun email dans {folder}", "warning")
                return
            
            # Limiter si nécessaire
            try:
                max_emails = int(self.max_emails_var.get())
                if max_emails > 0 and folder_total > max_emails:
                    email_ids = email_ids[-max_emails:]
                    folder_total = max_emails
            except:
                pass
            
            self.log(f"📬 {folder_total} emails à analyser dans {folder}", "info")
            
            stats['total'] += folder_total
            
            # Traiter par lots
            batch_size = int(self.batch_size_var.get())
            
            for i in range(0, len(email_ids), batch_size):
                batch = email_ids[i:i+batch_size]
                
                for num in batch:
                    if not self.is_running:
                        self.log("⏹️ Analyse interrompue", "warning")
                        return
                    
                    stats['processed'] += 1
                    
                    # Mise à jour du statut
                    if stats['processed'] % 10 == 0:
                        self.status_var.set(f"🔄 {folder}: {stats['processed']}/{stats['total']} emails")
                    
                    try:
                        # Récupérer l'email avec PEEK pour ne pas le marquer comme lu
                        if self.preserve_unread_var.get():
                            fetch_command = '(BODY.PEEK[] FLAGS)'
                        else:
                            fetch_command = '(RFC822 FLAGS)'
                        
                        result, msg_data = connection.fetch(num, fetch_command)
                        
                        if result != 'OK':
                            stats['errors'] += 1
                            continue
                        
                        # Parser l'email
                        raw_email = msg_data[0][1]
                        msg = email.message_from_bytes(raw_email)
                        
                        # Récupérer les flags
                        current_flags = self.extract_flags(msg_data)
                        is_unread = b'\\Seen' not in current_flags
                        
                        # Décoder les headers
                        subject = self.decode_header(msg.get("Subject", ""))[:100]
                        from_addr = self.decode_header(msg.get("From", ""))
                        to_addr = self.decode_header(msg.get("To", ""))
                        cc_addr = self.decode_header(msg.get("Cc", ""))
                        date = msg.get("Date", "")
                        
                        # Analyser avec le nouveau système
                        action = self.analyze_email_v3(msg, subject, from_addr, to_addr, 
                                                       cc_addr, date, current_flags, stats)
                        
                        if action:
                            if not self.dry_run_var.get():
                                # Exécuter l'action immédiatement
                                success = self.execute_action(connection, num, action, subject, is_unread)
                                if not success:
                                    stats['errors'] += 1
                            else:
                                self.log(f"🧪 [TEST] {subject[:50]}... → {action.get('folder', action.get('action'))}", "test")
                        
                    except Exception as e:
                        stats['errors'] += 1
                        self.log(f"⚠️ Erreur sur un email: {str(e)[:100]}", "error")
            
            # Expurger les messages marqués pour suppression
            if not self.dry_run_var.get():
                try:
                    result = connection.expunge()
                    if result[0] == 'OK':
                        self.log(f"🗑️ Messages supprimés expurgés dans {folder}", "info")
                except Exception as e:
                    self.log(f"⚠️ Erreur lors de l'expunge: {str(e)}", "warning")
                
        except Exception as e:
            self.log(f"⚠️ Erreur dans le dossier {folder}: {str(e)}", "error")
            stats['errors'] += 1
    
    def analyze_email_v3(self, msg, subject, from_addr, to_addr, cc_addr, date, flags, stats):
        """Analyser un email avec le système de chaînes et priorités"""
        user_email = self.email_var.get().lower()
        
        # Vérifier d'abord les chaînes de règles actives
        for chain in sorted(self.rule_chains, key=lambda x: x.get('priority', 50)):
            if not chain.get('enabled', True):
                continue
            
            for rule in chain.get('rules', []):
                if self.check_rule_v3(msg, subject, from_addr, to_addr, cc_addr, rule):
                    self.log(f"⛓️ Chaîne '{chain['name']}' → Règle '{rule.get('name')}'", "info")
                    stats['chains_applied'] += 1
                    
                    action = self.create_action_from_rule(rule)
                    
                    if chain.get('stop_on_match', True):
                        return action
                    
                    if not rule.get('continue_chain', False):
                        return action
        
        # Ensuite les règles individuelles par priorité
        for rule in self.rules:
            if self.check_rule_v3(msg, subject, from_addr, to_addr, cc_addr, rule):
                self.log(f"📍 Règle: {rule.get('name', 'Sans nom')}", "info")
                stats['rules_applied'] += 1
                
                action = self.create_action_from_rule(rule)
                
                if rule.get('stop_processing'):
                    return action
                
                if not rule.get('continue_chain'):
                    return action
        
        # Enfin la gestion CC
        is_in_cc = cc_addr and user_email in cc_addr.lower()
        is_primary = to_addr and user_email in to_addr.lower()
        
        if self.cc_enabled_var.get() and is_in_cc and not is_primary:
            if self.cc_skip_important_var.get() and b'\\Flagged' in flags:
                return None
            
            if self.cc_skip_recent_var.get():
                try:
                    from email.utils import parsedate_to_datetime
                    email_date = parsedate_to_datetime(date)
                    if (datetime.now(email_date.tzinfo) - email_date).days < 1:
                        return None
                except:
                    pass
            
            stats['cc_moved'] += 1
            return {
                'type': 'move',
                'folder': self.cc_folder_var.get(),
                'mark_read': self.cc_mark_read_after_var.get()
            }
        
        return None
    
    def check_rule_v3(self, msg, subject, from_addr, to_addr, cc_addr, rule):
        """Vérifier si un email correspond à une règle avec conditions multiples"""
        # Première condition
        if not self.check_single_condition(msg, subject, from_addr, to_addr, cc_addr, 
                                          rule.get('field'), rule.get('condition'), 
                                          rule.get('keyword'), rule.get('case_sensitive')):
            return False
        
        # Condition ET (optionnelle)
        if rule.get('and_field') and rule.get('and_keyword'):
            if not self.check_single_condition(msg, subject, from_addr, to_addr, cc_addr,
                                              rule.get('and_field'), rule.get('and_condition'),
                                              rule.get('and_keyword'), rule.get('case_sensitive')):
                return False
        
        return True
    
    def check_single_condition(self, msg, subject, from_addr, to_addr, cc_addr, 
                              field, condition, keyword, case_sensitive):
        """Vérifier une condition unique"""
        # Obtenir le texte à vérifier
        if field == "Sujet":
            text = subject
        elif field == "Expéditeur":
            text = from_addr
        elif field == "Destinataire":
            text = to_addr
        elif field == "Corps":
            text = self.get_email_body(msg)
        elif field == "Sujet ou Corps":
            text = subject + " " + self.get_email_body(msg)
        elif field == "Domaine expéditeur":
            # Extraire le domaine
            import re
            match = re.search(r'@([^\s>]+)', from_addr)
            text = match.group(1) if match else ""
        else:
            text = subject
        
        # Gestion de la casse
        if not case_sensitive:
            text = text.lower()
            keyword = keyword.lower()
        
        # Vérifier la condition
        if condition == "contient":
            return keyword in text
        elif condition == "ne contient pas":
            return keyword not in text
        elif condition == "commence par":
            return text.startswith(keyword)
        elif condition == "finit par":
            return text.endswith(keyword)
        elif condition == "est exactement":
            return text == keyword
        elif condition == "n'est pas":
            return text != keyword
        elif condition == "correspond à (regex)":
            try:
                return bool(re.search(keyword, text))
            except:
                return False
        elif condition == "contient un de (liste)":
            # Séparer par virgules
            keywords = [k.strip() for k in keyword.split(',')]
            return any(k in text for k in keywords)
        
        return False
    
    def create_action_from_rule(self, rule):
        """Créer une action depuis une règle"""
        return {
            'type': rule.get('action', 'move'),
            'action': rule.get('action'),
            'folder': rule.get('folder', ''),
            'mark_read': rule.get('mark_after_action', False)
        }
    
    def build_search_criteria(self):
        """Construire les critères de recherche IMAP"""
        criteria = []
        
        if self.filter_unread_only_var.get():
            criteria.append('UNSEEN')
        
        if self.filter_date_var.get():
            try:
                days = int(self.filter_days_var.get())
                since_date = (datetime.now() - timedelta(days=days)).strftime("%d-%b-%Y")
                criteria.append(f'SINCE {since_date}')
            except:
                pass
        
        return ' '.join(criteria) if criteria else 'ALL'
    
    def extract_flags(self, msg_data):
        """Extraire les flags d'un message"""
        for response in msg_data:
            if isinstance(response, tuple) and len(response) >= 2:
                if b'FLAGS' in response[0]:
                    return response[0]
        return b''
    
    def decode_header(self, header):
        """Décoder un header d'email"""
        if not header:
            return ""
        
        try:
            decoded = decode_header(header)[0][0]
            if isinstance(decoded, bytes):
                return decoded.decode('utf-8', errors='ignore')
            return str(decoded)
        except:
            return str(header)
    
    def get_email_body(self, msg):
        """Extraire le corps du message"""
        body = ""
        
        try:
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    if "text/plain" in content_type:
                        try:
                            body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                            if body:
                                break
                        except:
                            continue
            else:
                try:
                    body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
                except:
                    body = str(msg.get_payload())
        except:
            body = ""
        
        return body[:1000]
    
    def execute_action(self, connection, num, action, subject, was_unread):
        """Exécuter une action sur un email"""
        try:
            action_type = action.get('action', action.get('type', 'move'))
            
            if action_type in ['move', 'Déplacer vers']:
                # Utiliser le nom de dossier tel quel s'il existe, sinon essayer de le formater
                folder_name = action['folder']
                if folder_name not in self.existing_folders:
                    folder_name = self.get_full_folder_name(folder_name)
                
                self.log(f"📦 Déplacement vers: {folder_name}", "info")
                
                if self.backup_before_move_var.get():
                    # Créer une copie de sauvegarde
                    backup_folder = "BACKUP"
                    if backup_folder not in self.existing_folders:
                        backup_folder = self.get_full_folder_name("BACKUP")
                    self.create_folder_if_needed(connection, "BACKUP")
                    connection.copy(num, backup_folder)
                
                # Copier vers le nouveau dossier
                result = connection.copy(num, folder_name)
                
                if result[0] == 'OK':
                    # Marquer pour suppression dans le dossier source
                    connection.store(num, '+FLAGS', '\\Deleted')
                    
                    # Gérer le statut lu/non-lu après déplacement si demandé
                    if not self.preserve_unread_var.get() and action.get('mark_read'):
                        # Note: cela ne fonctionnera que sur l'email source, pas la copie
                        connection.store(num, '+FLAGS', '\\Seen')
                    
                    self.log(f"✅ {subject[:50]}... → {folder_name}", "success")
                    return True
                else:
                    self.log(f"⚠️ Échec du déplacement vers {folder_name}: {result}", "warning")
                    # Essayer avec un nom alternatif si échec
                    if "INBOX." not in folder_name and folder_name != "INBOX":
                        alt_folder = f"INBOX.{folder_name}"
                        self.log(f"🔄 Tentative avec: {alt_folder}", "info")
                        result = connection.copy(num, alt_folder)
                        if result[0] == 'OK':
                            connection.store(num, '+FLAGS', '\\Deleted')
                            self.log(f"✅ {subject[:50]}... → {alt_folder}", "success")
                            return True
            
            elif action_type in ['copy', 'Copier vers']:
                folder_name = action['folder']
                if folder_name not in self.existing_folders:
                    folder_name = self.get_full_folder_name(folder_name)
                
                result = connection.copy(num, folder_name)
                
                if result[0] == 'OK':
                    if action.get('mark_read') and not self.preserve_unread_var.get():
                        connection.store(num, '+FLAGS', '\\Seen')
                    self.log(f"📄 {subject[:50]}... copié vers {folder_name}", "info")
                    return True
                else:
                    self.log(f"⚠️ Échec de la copie vers {folder_name}", "warning")
            
            elif action_type == 'Marquer comme lu':
                if not self.preserve_unread_var.get():
                    connection.store(num, '+FLAGS', '\\Seen')
                    self.log(f"📖 {subject[:50]}... marqué comme lu", "info")
                    return True
            
            elif action_type == 'Marquer comme important':
                connection.store(num, '+FLAGS', '\\Flagged')
                self.log(f"⭐ {subject[:50]}... marqué comme important", "info")
                return True
            
            elif action_type == 'Supprimer':
                connection.store(num, '+FLAGS', '\\Deleted')
                self.log(f"🗑️ {subject[:50]}... supprimé", "warning")
                return True
            
            elif action_type == 'Étiqueter':
                if action.get('folder'):
                    connection.store(num, '+FLAGS', f'({action["folder"]})')
                    self.log(f"🏷️ {subject[:50]}... étiqueté: {action['folder']}", "info")
                    return True
            
        except Exception as e:
            self.log(f"❌ Erreur lors de l'action sur '{subject[:30]}': {str(e)}", "error")
            return False
    
    def display_summary(self, stats):
        """Afficher le résumé de l'analyse"""
        self.log("\n" + "="*60, "separator")
        self.log("📊 RÉSUMÉ DE L'ANALYSE V3", "header")
        self.log("="*60, "separator")
        
        if self.dry_run_var.get():
            self.log("🧪 MODE TEST - Aucun email n'a été réellement déplacé", "warning")
        
        self.log(f"✅ Emails analysés: {stats['processed']}/{stats['total']}", "success")
        self.log(f"📋 Emails en CC déplacés: {stats['cc_moved']}", "info")
        self.log(f"🎯 Règles appliquées: {stats['rules_applied']}", "info")
        self.log(f"⛓️ Chaînes appliquées: {stats['chains_applied']}", "info")
        
        if stats['errors'] > 0:
            self.log(f"⚠️ Erreurs rencontrées: {stats['errors']}", "warning")
        
        total_moved = stats['cc_moved'] + stats['rules_applied'] + stats['chains_applied']
        self.log(f"📧 TOTAL traité: {total_moved} actions", "success")
        
        if self.preserve_unread_var.get():
            self.log("🔒 Statut non-lu préservé pour tous les emails", "success")
        
        self.status_var.set(f"✅ Terminé - {total_moved} actions sur {stats['processed']} emails")
        
        # Message de fin
        if not self.dry_run_var.get():
            if total_moved > 0:
                messagebox.showinfo("Analyse terminée", 
                                   f"✅ Analyse terminée avec succès!\n\n"
                                   f"📊 Résultats:\n"
                                   f"• {stats['processed']} emails analysés\n"
                                   f"• {stats['cc_moved']} emails en CC déplacés\n"
                                   f"• {stats['rules_applied']} règles appliquées\n"
                                   f"• {stats['chains_applied']} chaînes appliquées\n"
                                   f"• Total: {total_moved} actions effectuées\n\n"
                                   f"{'🔒 Statut non-lu préservé' if self.preserve_unread_var.get() else ''}")
            else:
                messagebox.showinfo("Analyse terminée", 
                                   f"Analyse terminée.\n\n"
                                   f"📊 {stats['processed']} emails analysés\n"
                                   f"Aucun email à traiter selon les critères.")
    
    def update_stats(self, stats):
        """Mettre à jour les statistiques affichées"""
        stats_text = (f"Dernière analyse: {datetime.now().strftime('%d/%m/%Y %H:%M')}\n"
                     f"• Emails traités: {stats['processed']}\n"
                     f"• CC déplacés: {stats['cc_moved']}\n"
                     f"• Règles appliquées: {stats['rules_applied']}\n"
                     f"• Chaînes appliquées: {stats.get('chains_applied', 0)}")
        self.stats_label.config(text=stats_text)
    
    def log(self, message, tag="info"):
        """Ajouter un message au log avec coloration"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        if tag not in ["separator", "header"]:
            message = f"[{timestamp}] {message}"
        
        self.console.insert(tk.END, message + "\n")
        
        # Colorer selon le type
        colors = {
            "error": "#ff6b6b",
            "success": "#51cf66",
            "warning": "#ffd43b",
            "info": "#74c0fc",
            "cc": "#a9e34b",
            "rule": "#ff8cc3",
            "header": "#ffffff",
            "test": "#ffa94d"
        }
        
        if tag in colors:
            self.console.tag_add(tag, f"end-2l", "end-1l")
            self.console.tag_config(tag, foreground=colors[tag])
            if tag == "header":
                self.console.tag_config(tag, font=("Consolas", 11, "bold"))
        
        self.console.see(tk.END)
        self.root.update_idletasks()
    
    def save_settings(self):
        """Sauvegarder les paramètres"""
        settings = {
            "email": self.email_var.get(),
            "server": self.server_var.get(),
            "port": self.port_var.get(),
            "preserve_unread": self.preserve_unread_var.get(),
            "cc_enabled": self.cc_enabled_var.get(),
            "cc_folder": self.cc_folder_var.get(),
            "cc_mark_read_after": self.cc_mark_read_after_var.get(),
            "cc_skip_important": self.cc_skip_important_var.get(),
            "cc_skip_recent": self.cc_skip_recent_var.get(),
            "max_emails": self.max_emails_var.get(),
            "processing_mode": self.processing_mode_var.get(),
            "filter_unread_only": self.filter_unread_only_var.get(),
            "filter_date": self.filter_date_var.get(),
            "filter_days": self.filter_days_var.get(),
            "dry_run": self.dry_run_var.get(),
            "backup_before_move": self.backup_before_move_var.get(),
            "confirm_actions": self.confirm_actions_var.get(),
            "batch_size": self.batch_size_var.get(),
            "parallel_processing": self.parallel_processing_var.get(),
            "include_inbox": self.include_inbox_var.get(),
            "scan_subfolders": self.scan_subfolders_var.get(),
            "exclude_special": self.exclude_special_var.get(),
            "rules": self.rules,
            "existing_folders": self.existing_folders
        }
        
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=4, ensure_ascii=False)
            
            # Sauvegarder aussi les chaînes
            self.save_chains()
            
        except Exception as e:
            self.log(f"⚠️ Impossible de sauvegarder les paramètres: {str(e)}", "warning")
    
    def load_settings(self):
        """Charger les paramètres sauvegardés"""
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                
                # Charger les paramètres
                self.email_var.set(settings.get("email", ""))
                self.server_var.set(settings.get("server", ""))
                self.port_var.set(settings.get("port", "993"))
                self.preserve_unread_var.set(settings.get("preserve_unread", True))
                self.cc_enabled_var.set(settings.get("cc_enabled", True))
                self.cc_folder_var.set(settings.get("cc_folder", "EN_COPIE"))
                self.cc_mark_read_after_var.set(settings.get("cc_mark_read_after", False))
                self.cc_skip_important_var.set(settings.get("cc_skip_important", True))
                self.cc_skip_recent_var.set(settings.get("cc_skip_recent", False))
                self.max_emails_var.set(settings.get("max_emails", "100"))
                self.processing_mode_var.set(settings.get("processing_mode", "peek"))
                self.filter_unread_only_var.set(settings.get("filter_unread_only", False))
                self.filter_date_var.set(settings.get("filter_date", False))
                self.filter_days_var.set(settings.get("filter_days", "7"))
                self.dry_run_var.set(settings.get("dry_run", False))
                self.backup_before_move_var.set(settings.get("backup_before_move", False))
                self.confirm_actions_var.set(settings.get("confirm_actions", False))
                self.batch_size_var.set(settings.get("batch_size", "50"))
                self.parallel_processing_var.set(settings.get("parallel_processing", False))
                self.include_inbox_var.set(settings.get("include_inbox", True))
                self.scan_subfolders_var.set(settings.get("scan_subfolders", False))
                self.exclude_special_var.set(settings.get("exclude_special", True))
                self.rules = settings.get("rules", [])
                self.existing_folders = settings.get("existing_folders", [])
                
                # Mettre à jour les widgets
                if self.existing_folders:
                    self.folder_dropdown['values'] = self.existing_folders
                    self.folders_listbox.delete(0, tk.END)
                    for folder in self.existing_folders:
                        self.folders_listbox.insert(tk.END, folder)
                
                self.sort_rules_by_priority()
                self.refresh_rules_tree()
                self.refresh_available_rules()
                
                if self.filter_date_var.get():
                    self.days_spinbox.config(state='normal')
                
                # Charger les chaînes
                self.load_chains()
                
                self.log("✅ Paramètres chargés depuis la dernière session", "success")
        except Exception as e:
            self.log(f"ℹ️ Première utilisation ou erreur de chargement: {str(e)[:50]}", "info")
    
    def on_closing(self):
        """Gestion de la fermeture"""
        if self.is_running:
            if messagebox.askokcancel("Quitter", "Une analyse est en cours. Voulez-vous vraiment quitter?"):
                self.is_running = False
                self.save_settings()
                self.root.destroy()
        else:
            self.save_settings()
            self.root.destroy()
    
    def run(self):
        """Lancer l'application"""
        self.root.mainloop()

# === POINT D'ENTRÉE PRINCIPAL ===
def main():
    """Fonction principale pour lancer Email Manager V3"""
    app = EmailManager()
    app.run()

if __name__ == "__main__":
    main()