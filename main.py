"""
Email Manager pour Thunderbird - Gestionnaire d'emails avec tri automatique
Version améliorée avec préservation du statut non-lu
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

class EmailManager:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("🦅 Email Manager pour Thunderbird - V2")
        self.root.geometry("1200x800")
        self.root.configure(bg='#2c3e50')
        
        # Variables
        self.connection = None
        self.config_file = "email_manager_settings.json"
        self.rules = []
        self.is_running = False
        self.folder_separator = "."
        self.use_inbox_prefix = True
        self.processed_emails = set()  # Pour éviter les doublons
        
        # Interface
        self.setup_ui()
        
        # Charger configuration
        self.load_settings()
        
        # Fermeture
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
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
        
        tk.Label(title_text, text="Email Manager pour Thunderbird V2", 
                font=("Arial", 26, "bold"), 
                bg='#0a84ff', fg='white').pack(anchor='w')
        
        tk.Label(title_text, text="Tri automatique sans modifier le statut de lecture", 
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
        
        # --- ONGLET 4: OPTIONS AVANCÉES ---
        self.setup_advanced_tab(notebook)
        
        # --- ONGLET 5: EXÉCUTION ---
        self.setup_execution_tab(notebook)
        
        # Barre de statut
        self.status_var = tk.StringVar(value="✅ Prêt - Email Manager V2")
        status_bar = tk.Label(self.root, textvariable=self.status_var,
                             bd=1, relief=tk.SUNKEN, anchor='w',
                             bg='#34495e', fg='white',
                             font=("Arial", 10))
        status_bar.pack(side='bottom', fill='x')
        
        # Initialisation
        self.log("✨ Email Manager V2 démarré!", "success")
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
        
        # Bouton test
        test_btn_frame = tk.Frame(conn_content, bg='white')
        test_btn_frame.pack(pady=20)
        
        self.test_btn = tk.Button(test_btn_frame, text=" 🔧 Tester la connexion ",
                                 font=("Arial", 12, "bold"),
                                 bg='#0a84ff', fg='white',
                                 padx=30, pady=10,
                                 command=self.test_connection,
                                 cursor='hand2')
        self.test_btn.pack()
    
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
        """Onglet des règles personnalisées amélioré"""
        rules_frame = tk.Frame(notebook, bg='white')
        notebook.add(rules_frame, text=' 🎯 Règles personnalisées ')
        
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
        
        # Ligne 1 - Condition principale
        line1 = tk.Frame(rule_builder, bg='white')
        line1.pack(fill='x', pady=5)
        
        tk.Label(line1, text="Si", font=("Arial", 11, "bold"), bg='white').pack(side='left', padx=5)
        
        self.rule_field_var = tk.StringVar(value="Sujet")
        field_menu = ttk.Combobox(line1, textvariable=self.rule_field_var,
                                  values=["Sujet", "Expéditeur", "Corps", "Destinataire", "Sujet ou Corps"],
                                  width=15, state='readonly')
        field_menu.pack(side='left', padx=5)
        
        self.rule_condition_var = tk.StringVar(value="contient")
        condition_menu = ttk.Combobox(line1, textvariable=self.rule_condition_var,
                                      values=["contient", "ne contient pas", "commence par", 
                                              "finit par", "est exactement", "n'est pas", 
                                              "correspond à (regex)"],
                                      width=18, state='readonly')
        condition_menu.pack(side='left', padx=5)
        
        self.rule_keyword_var = tk.StringVar()
        keyword_entry = tk.Entry(line1, textvariable=self.rule_keyword_var, 
                                width=30, font=("Arial", 11))
        keyword_entry.pack(side='left', padx=5)
        
        # Ligne 2 - Options de la règle
        line2 = tk.Frame(rule_builder, bg='white')
        line2.pack(fill='x', pady=5)
        
        tk.Label(line2, text="Options:", font=("Arial", 11, "bold"), bg='white').pack(side='left', padx=5)
        
        self.rule_case_sensitive_var = tk.BooleanVar(value=False)
        tk.Checkbutton(line2, text="Sensible à la casse",
                      variable=self.rule_case_sensitive_var,
                      font=("Arial", 10), bg='white').pack(side='left', padx=5)
        
        self.rule_priority_var = tk.StringVar(value="Normal")
        tk.Label(line2, text="Priorité:", font=("Arial", 10), bg='white').pack(side='left', padx=10)
        priority_menu = ttk.Combobox(line2, textvariable=self.rule_priority_var,
                                     values=["Haute", "Normal", "Basse"],
                                     width=10, state='readonly')
        priority_menu.pack(side='left', padx=5)
        
        # Ligne 3 - Action
        line3 = tk.Frame(rule_builder, bg='white')
        line3.pack(fill='x', pady=5)
        
        tk.Label(line3, text="Alors", font=("Arial", 11, "bold"), bg='white').pack(side='left', padx=5)
        
        self.rule_action_var = tk.StringVar(value="Déplacer vers")
        action_menu = ttk.Combobox(line3, textvariable=self.rule_action_var,
                                   values=["Déplacer vers", "Copier vers", "Marquer comme lu", 
                                          "Marquer comme important", "Supprimer"],
                                   width=18, state='readonly')
        action_menu.pack(side='left', padx=5)
        action_menu.bind('<<ComboboxSelected>>', self.on_action_changed)
        
        self.rule_folder_var = tk.StringVar()
        self.folder_entry = tk.Entry(line3, textvariable=self.rule_folder_var, 
                                     width=25, font=("Arial", 11))
        self.folder_entry.pack(side='left', padx=5)
        
        # Ligne 4 - Options d'action
        line4 = tk.Frame(rule_builder, bg='white')
        line4.pack(fill='x', pady=5)
        
        self.rule_stop_processing_var = tk.BooleanVar(value=False)
        tk.Checkbutton(line4, text=" 🛑 Arrêter le traitement après cette règle",
                      variable=self.rule_stop_processing_var,
                      font=("Arial", 10), bg='white').pack(side='left', padx=5)
        
        self.rule_mark_after_move_var = tk.BooleanVar(value=False)
        self.mark_checkbox = tk.Checkbutton(line4, 
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
                                        text=" 📋 Règles actives (par ordre de priorité) ", 
                                        font=("Arial", 12, "bold"),
                                        bg='white', fg='#2c3e50', relief=tk.FLAT)
        rules_list_frame.pack(fill='both', expand=True)
        
        # Frame avec scrollbar
        list_container = tk.Frame(rules_list_frame, bg='white')
        list_container.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Treeview pour afficher les règles
        columns = ('Priorité', 'Champ', 'Condition', 'Valeur', 'Action', 'Options')
        self.rules_tree = ttk.Treeview(list_container, columns=columns, show='tree headings', height=8)
        
        # Configuration des colonnes
        self.rules_tree.heading('#0', text='#')
        self.rules_tree.column('#0', width=40)
        
        for col in columns:
            self.rules_tree.heading(col, text=col)
            width = 120 if col in ['Valeur', 'Options'] else 100
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
        
        tk.Button(rules_btns, text=" ⬆️ Monter ",
                 bg='#3498db', fg='white',
                 font=("Arial", 10, "bold"),
                 command=self.move_rule_up).pack(side='left', padx=5)
        
        tk.Button(rules_btns, text=" ⬇️ Descendre ",
                 bg='#3498db', fg='white',
                 font=("Arial", 10, "bold"),
                 command=self.move_rule_down).pack(side='left', padx=5)
        
        tk.Button(rules_btns, text=" ✏️ Modifier ",
                 bg='#f39c12', fg='white',
                 font=("Arial", 10, "bold"),
                 command=self.edit_rule).pack(side='left', padx=5)
        
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
    
    def setup_advanced_tab(self, notebook):
        """Onglet des options avancées"""
        adv_frame = tk.Frame(notebook, bg='white')
        notebook.add(adv_frame, text=' ⚙️ Options avancées ')
        
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
        
        # Note importante
        note_frame = tk.Frame(adv_content, bg='#fff3cd', relief=tk.RIDGE, bd=2)
        note_frame.pack(fill='x', pady=10)
        
        tk.Label(note_frame, text="⚠️ Note importante sur le mode PEEK", 
                font=("Arial", 11, "bold"), 
                bg='#fff3cd', fg='#856404').pack(anchor='w', padx=10, pady=(10, 5))
        
        note_text = """Le mode PEEK est recommandé car il utilise BODY.PEEK[] qui garantit que les emails ne seront PAS marqués 
comme lus lors de l'analyse. Les emails seront correctement déplacés tout en préservant leur statut.
Utilisez le mode BODY uniquement si vous acceptez que les emails puissent être marqués comme lus."""
        
        tk.Label(note_frame, text=note_text,
                font=("Arial", 9), bg='#fff3cd', 
                fg='#856404', justify='left', wraplength=650).pack(anchor='w', padx=25, pady=(0, 10))
    
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
            self.mark_checkbox.config(state='normal')
        else:
            self.folder_entry.config(state='disabled')
            if action == "Marquer comme lu":
                self.mark_checkbox.config(state='disabled')
                self.rule_mark_after_move_var.set(False)
    
    def add_custom_rule(self):
        """Ajouter une règle personnalisée améliorée"""
        keyword = self.rule_keyword_var.get().strip()
        action = self.rule_action_var.get()
        
        if not keyword:
            messagebox.showwarning("Attention", "Veuillez entrer un mot-clé ou une expression!")
            return
        
        if action in ["Déplacer vers", "Copier vers"] and not self.rule_folder_var.get().strip():
            messagebox.showwarning("Attention", "Veuillez spécifier le dossier de destination!")
            return
        
        # Déterminer la priorité numérique
        priority_map = {"Haute": 1, "Normal": 2, "Basse": 3}
        
        rule = {
            "field": self.rule_field_var.get(),
            "condition": self.rule_condition_var.get(),
            "keyword": keyword,
            "action": action,
            "folder": self.rule_folder_var.get().strip() if action in ["Déplacer vers", "Copier vers"] else "",
            "case_sensitive": self.rule_case_sensitive_var.get(),
            "priority": priority_map[self.rule_priority_var.get()],
            "priority_text": self.rule_priority_var.get(),
            "stop_processing": self.rule_stop_processing_var.get(),
            "mark_after_action": self.rule_mark_after_move_var.get()
        }
        
        self.rules.append(rule)
        self.sort_rules_by_priority()
        self.refresh_rules_tree()
        
        # Réinitialiser les champs
        self.rule_keyword_var.set("")
        self.rule_folder_var.set("")
        self.rule_case_sensitive_var.set(False)
        self.rule_priority_var.set("Normal")
        self.rule_stop_processing_var.set(False)
        self.rule_mark_after_move_var.set(False)
        
        self.log(f"✅ Règle ajoutée: {rule['field']} {rule['condition']} '{keyword}'", "success")
        self.save_settings()
    
    def sort_rules_by_priority(self):
        """Trier les règles par priorité"""
        self.rules.sort(key=lambda x: x.get('priority', 2))
    
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
            if rule.get('mark_after_action'):
                options.append("Marquer lu")
            
            values = (
                rule.get('priority_text', 'Normal'),
                rule['field'],
                rule['condition'],
                rule['keyword'][:30] + ('...' if len(rule['keyword']) > 30 else ''),
                f"{rule['action']} {rule.get('folder', '')}".strip(),
                ', '.join(options)
            )
            
            self.rules_tree.insert('', 'end', text=str(i), values=values)
    
    def move_rule_up(self):
        """Déplacer une règle vers le haut"""
        selection = self.rules_tree.selection()
        if selection:
            item = selection[0]
            index = self.rules_tree.index(item)
            if index > 0:
                self.rules[index], self.rules[index-1] = self.rules[index-1], self.rules[index]
                self.refresh_rules_tree()
                self.save_settings()
                self.log("⬆️ Règle déplacée vers le haut", "info")
    
    def move_rule_down(self):
        """Déplacer une règle vers le bas"""
        selection = self.rules_tree.selection()
        if selection:
            item = selection[0]
            index = self.rules_tree.index(item)
            if index < len(self.rules) - 1:
                self.rules[index], self.rules[index+1] = self.rules[index+1], self.rules[index]
                self.refresh_rules_tree()
                self.save_settings()
                self.log("⬇️ Règle déplacée vers le bas", "info")
    
    def edit_rule(self):
        """Modifier une règle existante"""
        selection = self.rules_tree.selection()
        if selection:
            item = selection[0]
            index = self.rules_tree.index(item)
            rule = self.rules[index]
            
            # Remplir les champs avec les valeurs de la règle
            self.rule_field_var.set(rule['field'])
            self.rule_condition_var.set(rule['condition'])
            self.rule_keyword_var.set(rule['keyword'])
            self.rule_action_var.set(rule.get('action', 'Déplacer vers'))
            self.rule_folder_var.set(rule.get('folder', ''))
            self.rule_case_sensitive_var.set(rule.get('case_sensitive', False))
            self.rule_priority_var.set(rule.get('priority_text', 'Normal'))
            self.rule_stop_processing_var.set(rule.get('stop_processing', False))
            self.rule_mark_after_move_var.set(rule.get('mark_after_action', False))
            
            # Supprimer la règle de la liste
            del self.rules[index]
            self.refresh_rules_tree()
            
            self.log("✏️ Règle en cours de modification", "info")
        else:
            messagebox.showinfo("Info", "Sélectionnez une règle à modifier")
    
    def delete_rule(self):
        """Supprimer la règle sélectionnée"""
        selection = self.rules_tree.selection()
        if selection:
            item = selection[0]
            index = self.rules_tree.index(item)
            
            if messagebox.askyesno("Confirmation", "Supprimer cette règle ?"):
                del self.rules[index]
                self.refresh_rules_tree()
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
            connection.select('INBOX', readonly=True)  # Mode lecture seule pour le test
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
        """Obtenir le nom complet du dossier avec le préfixe INBOX. si nécessaire"""
        if not folder_name.startswith("INBOX.") and not folder_name.upper() == "INBOX":
            return f"INBOX.{folder_name}"
        return folder_name
    
    def create_folder_if_needed(self, connection, folder_name):
        """Créer un dossier IMAP s'il n'existe pas"""
        if not folder_name:
            return True
            
        try:
            full_folder_name = self.get_full_folder_name(folder_name)
            
            # Lister tous les dossiers existants
            result, folders = connection.list()
            
            # Vérifier si le dossier existe
            folder_exists = False
            if result == 'OK':
                for folder in folders:
                    if folder:
                        folder_str = folder.decode('utf-8') if isinstance(folder, bytes) else str(folder)
                        if full_folder_name.lower() in folder_str.lower():
                            folder_exists = True
                            self.log(f"📁 Dossier '{full_folder_name}' déjà existant", "info")
                            break
            
            if not folder_exists:
                # Créer le dossier
                result = connection.create(full_folder_name)
                if result[0] == 'OK':
                    self.log(f"✅ Dossier '{full_folder_name}' créé avec succès", "success")
                    connection.subscribe(full_folder_name)
                    return True
                else:
                    self.log(f"❌ Impossible de créer '{full_folder_name}': {result[1]}", "error")
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
        """Worker pour l'analyse des emails avec préservation du statut non-lu"""
        stats = {
            'total': 0,
            'processed': 0,
            'cc_moved': 0,
            'rules_applied': 0,
            'errors': 0
        }
        
        try:
            # Connexion
            self.log("\n" + "="*60, "separator")
            self.log("🚀 DÉMARRAGE DE L'ANALYSE", "header")
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
            
            # Sélectionner INBOX avec mode approprié
            if self.processing_mode_var.get() == "peek" and self.preserve_unread_var.get():
                # Mode lecture seule pour préserver absolument le statut
                connection.select('INBOX', readonly=True)
                self.log("📖 INBOX ouvert en mode lecture seule (PEEK)", "info")
            else:
                connection.select('INBOX')
                self.log("📖 INBOX ouvert en mode normal", "info")
            
            # Construire la requête de recherche
            search_criteria = self.build_search_criteria()
            result, data = connection.search(None, search_criteria)
            
            if result != 'OK':
                self.log("❌ Erreur lors de la recherche des emails", "error")
                return
            
            email_ids = data[0].split()
            stats['total'] = len(email_ids)
            
            if stats['total'] == 0:
                self.log("📭 Aucun email correspondant aux critères", "warning")
                return
            
            # Limiter si nécessaire
            try:
                max_emails = int(self.max_emails_var.get())
                if max_emails > 0 and stats['total'] > max_emails:
                    email_ids = email_ids[-max_emails:]
                    stats['total'] = max_emails
            except:
                pass
            
            self.log(f"📬 {stats['total']} emails à analyser", "info")
            self.log("-" * 50, "separator")
            
            # Liste des actions à effectuer (pour le mode readonly)
            actions_to_perform = []
            
            for num in email_ids:
                if not self.is_running:
                    self.log("⏹️ Analyse interrompue par l'utilisateur", "warning")
                    break
                
                stats['processed'] += 1
                
                try:
                    # Mise à jour du statut
                    if stats['processed'] % 10 == 0:
                        self.status_var.set(f"🔄 Analyse... {stats['processed']}/{stats['total']} emails traités")
                    
                    # Récupérer l'email avec PEEK pour ne pas le marquer comme lu
                    if self.processing_mode_var.get() == "peek":
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
                    
                    # Récupérer les flags actuels
                    current_flags = self.extract_flags(msg_data)
                    is_unread = b'\\Seen' not in current_flags
                    
                    # Décoder les headers
                    subject = self.decode_header(msg.get("Subject", ""))[:100]
                    from_addr = self.decode_header(msg.get("From", ""))
                    to_addr = self.decode_header(msg.get("To", ""))
                    cc_addr = self.decode_header(msg.get("Cc", ""))
                    date = msg.get("Date", "")
                    
                    # Analyser et appliquer les règles
                    action = self.analyze_email(msg, subject, from_addr, to_addr, cc_addr, 
                                               date, current_flags, stats)
                    
                    if action and not self.dry_run_var.get():
                        # Stocker l'action pour l'exécuter plus tard si en mode readonly
                        if self.processing_mode_var.get() == "peek" and self.preserve_unread_var.get():
                            actions_to_perform.append((num, action, subject, is_unread))
                        else:
                            # Exécuter l'action immédiatement
                            self.execute_action(connection, num, action, subject, is_unread)
                    
                    elif action and self.dry_run_var.get():
                        self.log(f"🧪 [TEST] {subject[:50]}... → {action.get('folder', 'action')}", "test")
                    
                except Exception as e:
                    stats['errors'] += 1
                    self.log(f"⚠️ Erreur sur un email: {str(e)[:100]}", "error")
                    continue
            
            # Si on était en mode readonly, reconnecter pour effectuer les actions
            if actions_to_perform and self.processing_mode_var.get() == "peek" and self.preserve_unread_var.get():
                self.log("-" * 50, "separator")
                self.log("🔄 Reconnexion pour effectuer les déplacements...", "info")
                
                connection.close()
                connection.logout()
                
                # Reconnecter en mode normal
                connection = imaplib.IMAP4_SSL(self.server_var.get(), int(self.port_var.get()))
                connection.login(self.email_var.get(), self.password_var.get())
                connection.select('INBOX')
                
                for num, action, subject, was_unread in actions_to_perform:
                    try:
                        self.execute_action(connection, num, action, subject, was_unread)
                    except Exception as e:
                        self.log(f"⚠️ Erreur lors du déplacement: {str(e)[:50]}", "error")
            
            # Expurger les messages supprimés
            if not self.dry_run_var.get() and (stats['cc_moved'] > 0 or stats['rules_applied'] > 0):
                self.log("-" * 50, "separator")
                self.log("🗑️ Nettoyage des emails déplacés...", "info")
                connection.expunge()
            
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
    
    def analyze_email(self, msg, subject, from_addr, to_addr, cc_addr, date, flags, stats):
        """Analyser un email et déterminer l'action à effectuer"""
        user_email = self.email_var.get().lower()
        
        # Vérifier si l'email est en CC
        is_in_cc = cc_addr and user_email in cc_addr.lower()
        is_primary = to_addr and user_email in to_addr.lower()
        
        # Confirmation pour actions si demandé
        if self.confirm_actions_var.get():
            if not self.should_process_email(subject, from_addr):
                return None
        
        # Gestion CC (priorité sur les règles)
        if self.cc_enabled_var.get() and is_in_cc and not is_primary:
            # Vérifications supplémentaires
            if self.cc_skip_important_var.get() and b'\\Flagged' in flags:
                self.log(f"⏭️ Email CC important ignoré: {subject[:50]}", "info")
                return None
            
            if self.cc_skip_recent_var.get():
                # Vérifier si l'email a moins de 24h
                try:
                    from email.utils import parsedate_to_datetime
                    email_date = parsedate_to_datetime(date)
                    if (datetime.now(email_date.tzinfo) - email_date).days < 1:
                        self.log(f"⏭️ Email CC récent ignoré: {subject[:50]}", "info")
                        return None
                except:
                    pass
            
            stats['cc_moved'] += 1
            return {
                'type': 'move',
                'folder': self.cc_folder_var.get(),
                'mark_read': self.cc_mark_read_after_var.get()
            }
        
        # Appliquer les règles personnalisées
        body = self.get_email_body(msg) if any(r['field'] in ['Corps', 'Sujet ou Corps'] for r in self.rules) else ""
        
        for rule in self.rules:
            if self.check_rule(msg, subject, from_addr, to_addr, body, rule):
                self.log(f"📍 Règle correspondante: {rule['field']} {rule['condition']} '{rule['keyword']}'", "info")
                
                # Créer l'action basée sur la règle
                action = {
                    'type': 'rule',
                    'action': rule.get('action', 'Déplacer vers'),
                    'folder': rule.get('folder', ''),
                    'mark_read': rule.get('mark_after_action', False)
                }
                
                # Incrémenter le compteur selon le type d'action
                if rule.get('action') in ['Déplacer vers', 'Copier vers']:
                    stats['rules_applied'] += 1
                
                # Si la règle dit d'arrêter le traitement
                if rule.get('stop_processing'):
                    self.log(f"🛑 Arrêt du traitement après cette règle", "info")
                    return action
                
                return action
        
        return None
    
    def should_process_email(self, subject, from_addr):
        """Demander confirmation pour traiter un email"""
        response = messagebox.askyesno(
            "Confirmation",
            f"Traiter cet email ?\n\nSujet: {subject[:50]}...\nDe: {from_addr[:50]}..."
        )
        return response
    
    def execute_action(self, connection, num, action, subject, was_unread):
        """Exécuter une action sur un email"""
        try:
            action_type = action.get('action', action.get('type', 'move'))
            
            if action_type in ['move', 'Déplacer vers']:
                folder_name = self.get_full_folder_name(action['folder'])
                
                if self.backup_before_move_var.get():
                    # Créer une copie de sauvegarde
                    backup_folder = self.get_full_folder_name("BACKUP")
                    self.create_folder_if_needed(connection, "BACKUP")
                    connection.copy(num, backup_folder)
                
                # Copier vers le nouveau dossier
                result = connection.copy(num, folder_name)
                
                if result[0] == 'OK':
                    # Marquer pour suppression
                    connection.store(num, '+FLAGS', '\\Deleted')
                    
                    # Gérer le statut lu/non-lu
                    if not self.preserve_unread_var.get() and action.get('mark_read'):
                        connection.store(num, '+FLAGS', '\\Seen')
                    
                    self.log(f"✅ {subject[:50]}... → {folder_name}", "success")
                else:
                    self.log(f"⚠️ Impossible de déplacer vers {folder_name}", "warning")
            
            elif action_type in ['copy', 'Copier vers']:
                folder_name = self.get_full_folder_name(action['folder'])
                result = connection.copy(num, folder_name)
                
                if result[0] == 'OK':
                    if action.get('mark_read') and not self.preserve_unread_var.get():
                        connection.store(num, '+FLAGS', '\\Seen')
                    self.log(f"📄 {subject[:50]}... copié vers {folder_name}", "info")
            
            elif action_type == 'Marquer comme lu':
                if not self.preserve_unread_var.get():
                    connection.store(num, '+FLAGS', '\\Seen')
                    self.log(f"📖 {subject[:50]}... marqué comme lu", "info")
            
            elif action_type == 'Marquer comme important':
                connection.store(num, '+FLAGS', '\\Flagged')
                self.log(f"⭐ {subject[:50]}... marqué comme important", "info")
            
            elif action_type == 'Supprimer':
                connection.store(num, '+FLAGS', '\\Deleted')
                self.log(f"🗑️ {subject[:50]}... supprimé", "warning")
            
        except Exception as e:
            self.log(f"⚠️ Erreur lors de l'action: {str(e)[:100]}", "error")
    
    def check_rule(self, msg, subject, from_addr, to_addr, body, rule):
        """Vérifier si un email correspond à une règle"""
        # Obtenir le champ à vérifier
        field = rule.get('field', 'Sujet')
        if field == "Sujet":
            text = subject
        elif field == "Expéditeur":
            text = from_addr
        elif field == "Destinataire":
            text = to_addr
        elif field == "Corps":
            text = body
        elif field == "Sujet ou Corps":
            text = subject + " " + body
        else:
            text = subject
        
        keyword = rule.get('keyword', '')
        condition = rule.get('condition', 'contient')
        case_sensitive = rule.get('case_sensitive', False)
        
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
        
        return False
    
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
    
    def display_summary(self, stats):
        """Afficher le résumé de l'analyse"""
        self.log("\n" + "="*60, "separator")
        self.log("📊 RÉSUMÉ DE L'ANALYSE", "header")
        self.log("="*60, "separator")
        
        if self.dry_run_var.get():
            self.log("🧪 MODE TEST - Aucun email n'a été réellement déplacé", "warning")
        
        self.log(f"✅ Emails analysés: {stats['processed']}/{stats['total']}", "success")
        self.log(f"📋 Emails en CC déplacés: {stats['cc_moved']}", "info")
        self.log(f"🎯 Règles appliquées: {stats['rules_applied']}", "info")
        
        if stats['errors'] > 0:
            self.log(f"⚠️ Erreurs rencontrées: {stats['errors']}", "warning")
        
        total_moved = stats['cc_moved'] + stats['rules_applied']
        self.log(f"📧 TOTAL traité: {total_moved} emails", "success")
        
        if self.preserve_unread_var.get():
            self.log("🔒 Statut non-lu préservé pour tous les emails", "success")
        
        self.status_var.set(f"✅ Terminé - {total_moved} emails triés sur {stats['processed']} analysés")
        
        # Message de fin
        if not self.dry_run_var.get():
            if total_moved > 0:
                messagebox.showinfo("Analyse terminée", 
                                   f"✅ Analyse terminée avec succès!\n\n"
                                   f"📊 Résultats:\n"
                                   f"• {stats['processed']} emails analysés\n"
                                   f"• {stats['cc_moved']} emails en CC déplacés\n"
                                   f"• {stats['rules_applied']} règles appliquées\n"
                                   f"• Total: {total_moved} emails organisés\n\n"
                                   f"{'🔒 Statut non-lu préservé' if self.preserve_unread_var.get() else ''}")
            else:
                messagebox.showinfo("Analyse terminée", 
                                   f"Analyse terminée.\n\n"
                                   f"📊 {stats['processed']} emails analysés\n"
                                   f"Aucun email à déplacer selon les critères configurés.")
        else:
            messagebox.showinfo("Mode test terminé",
                              f"🧪 Simulation terminée\n\n"
                              f"• {stats['processed']} emails analysés\n"
                              f"• {total_moved} emails auraient été triés\n\n"
                              f"Désactivez le mode test pour effectuer les actions.")
    
    def update_stats(self, stats):
        """Mettre à jour les statistiques affichées"""
        stats_text = (f"Dernière analyse: {datetime.now().strftime('%d/%m/%Y %H:%M')}\n"
                     f"• Emails traités: {stats['processed']}\n"
                     f"• CC déplacés: {stats['cc_moved']}\n"
                     f"• Règles appliquées: {stats['rules_applied']}")
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
            "rules": self.rules
        }
        
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=4, ensure_ascii=False)
        except Exception as e:
            self.log(f"⚠️ Impossible de sauvegarder les paramètres: {str(e)}", "warning")
    
    def load_settings(self):
        """Charger les paramètres sauvegardés"""
        try:
            if os.path.exists(self.config_file):
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
                self.rules = settings.get("rules", [])
                
                # Compatibilité avec anciennes versions
                for rule in self.rules:
                    if 'priority' not in rule:
                        rule['priority'] = 2
                        rule['priority_text'] = 'Normal'
                    if 'action' not in rule:
                        rule['action'] = 'Déplacer vers'
                
                self.sort_rules_by_priority()
                self.refresh_rules_tree()
                
                if self.filter_date_var.get():
                    self.days_spinbox.config(state='normal')
                
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
    """Fonction principale pour lancer Email Manager V2"""
    app = EmailManager()
    app.run()

if __name__ == "__main__":
    main()