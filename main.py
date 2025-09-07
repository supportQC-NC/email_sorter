

"""
Email Manager pour Thunderbird - Gestionnaire d'emails avec tri automatique
Compatible avec tous les serveurs IMAP (Orange, Free, Hostinger, etc.)
Gestion spéciale des emails en copie (CC)
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import imaplib
import email
from email.header import decode_header
import threading
from datetime import datetime
import json
import os
import re

class EmailManager:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("🦅 Email Manager pour Thunderbird")
        self.root.geometry("1100x750")
        self.root.configure(bg='#2c3e50')
        
        # Icône de l'application
        try:
            self.root.iconbitmap('email.ico')
        except:
            pass
        
        # Variables
        self.connection = None
        self.config_file = "email_manager_settings.json"
        self.rules = []
        self.is_running = False
        self.folder_separator = "."
        self.use_inbox_prefix = True  # Toujours utiliser INBOX. pour Thunderbird
        
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
        
        tk.Label(title_text, text="Email Manager pour Thunderbird", 
                font=("Arial", 26, "bold"), 
                bg='#0a84ff', fg='white').pack(anchor='w')
        
        tk.Label(title_text, text="Tri automatique des emails - Compatible tous serveurs IMAP", 
                font=("Arial", 12), 
                bg='#0a84ff', fg='#ecf0f1').pack(anchor='w')
        
        # === NOTEBOOK (Onglets) ===
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # --- ONGLET 1: CONNEXION ---
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
        
        # Note d'information
        info_note = tk.Label(conn_content, 
                           text="💡 Entrez votre serveur IMAP et vos identifiants Thunderbird\nPour Gmail: utilisez un mot de passe d'application",
                           font=("Arial", 10, "italic"),
                           bg='#fff3cd', fg='#856404',
                           wraplength=600, pady=10)
        info_note.pack(pady=10)
        
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
        
        # --- ONGLET 2: GESTION CC ---
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
        
        # Options supplémentaires
        extra_options = tk.Frame(self.cc_options_frame, bg='#ecf0f1')
        extra_options.pack(fill='x', pady=10, padx=20)
        
        self.cc_mark_read_var = tk.BooleanVar(value=False)
        tk.Checkbutton(extra_options, text=" 📖 Marquer comme lu après déplacement",
                      variable=self.cc_mark_read_var,
                      font=("Arial", 10), bg='#ecf0f1').pack(anchor='w', pady=5)
        
        self.cc_skip_important_var = tk.BooleanVar(value=True)
        tk.Checkbutton(extra_options, 
                      text=" ⭐ Ne pas déplacer les emails marqués comme importants",
                      variable=self.cc_skip_important_var,
                      font=("Arial", 10), bg='#ecf0f1').pack(anchor='w', pady=5)
        
        # Info box
        info_frame = tk.Frame(cc_content, bg='#d4edda', relief=tk.RIDGE, bd=2)
        info_frame.pack(fill='x', pady=10)
        
        tk.Label(info_frame, text="ℹ️ Comment ça marche ?", 
                font=("Arial", 11, "bold"), 
                bg='#d4edda', fg='#155724').pack(anchor='w', padx=10, pady=(10, 5))
        
        info_text = """✅ Les emails où vous êtes le destinataire principal → RESTENT dans la boîte de réception
✅ Les emails où vous êtes uniquement en copie (CC) → DÉPLACÉS vers le dossier configuré
✅ Le dossier sera créé automatiquement s'il n'existe pas (sous INBOX.)
✅ Compatible avec tous les serveurs IMAP Thunderbird"""
        
        tk.Label(info_frame, text=info_text,
                font=("Arial", 10), bg='#d4edda', 
                fg='#155724', justify='left').pack(anchor='w', padx=25, pady=(0, 10))
        
        # --- ONGLET 3: RÈGLES PERSONNALISÉES ---
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
        
        # Ligne 1
        line1 = tk.Frame(rule_builder, bg='white')
        line1.pack(fill='x', pady=5)
        
        tk.Label(line1, text="Si", font=("Arial", 11, "bold"), bg='white').pack(side='left', padx=5)
        
        self.rule_field_var = tk.StringVar(value="Sujet")
        field_menu = ttk.Combobox(line1, textvariable=self.rule_field_var,
                                  values=["Sujet", "Expéditeur", "Corps", "Destinataire"],
                                  width=15, state='readonly')
        field_menu.pack(side='left', padx=5)
        
        self.rule_condition_var = tk.StringVar(value="contient")
        condition_menu = ttk.Combobox(line1, textvariable=self.rule_condition_var,
                                      values=["contient", "commence par", "finit par", "est exactement"],
                                      width=15, state='readonly')
        condition_menu.pack(side='left', padx=5)
        
        self.rule_keyword_var = tk.StringVar()
        keyword_entry = tk.Entry(line1, textvariable=self.rule_keyword_var, 
                                width=30, font=("Arial", 11))
        keyword_entry.pack(side='left', padx=5)
        
        # Ligne 2
        line2 = tk.Frame(rule_builder, bg='white')
        line2.pack(fill='x', pady=5)
        
        tk.Label(line2, text="Alors déplacer vers le dossier", 
                font=("Arial", 11, "bold"), bg='white').pack(side='left', padx=5)
        
        self.rule_folder_var = tk.StringVar()
        folder_entry = tk.Entry(line2, textvariable=self.rule_folder_var, 
                               width=25, font=("Arial", 11))
        folder_entry.pack(side='left', padx=5)
        
        add_rule_btn = tk.Button(line2, text=" ➕ Ajouter la règle ",
                                bg='#27ae60', fg='white',
                                font=("Arial", 11, "bold"),
                                command=self.add_custom_rule,
                                cursor='hand2')
        add_rule_btn.pack(side='left', padx=20)
        
        # Note sur les dossiers
        tk.Label(rules_content, 
                text="💡 Les dossiers seront créés automatiquement sous INBOX. (ex: INBOX.compta)",
                font=("Arial", 9, "italic"), bg='white', fg='#7f8c8d').pack(pady=5)
        
        # Liste des règles
        rules_list_frame = tk.LabelFrame(rules_content, 
                                        text=" 📋 Règles actives ", 
                                        font=("Arial", 12, "bold"),
                                        bg='white', fg='#2c3e50', relief=tk.FLAT)
        rules_list_frame.pack(fill='both', expand=True)
        
        # Scrollable list
        list_container = tk.Frame(rules_list_frame, bg='white')
        list_container.pack(fill='both', expand=True, padx=10, pady=10)
        
        scrollbar = ttk.Scrollbar(list_container)
        scrollbar.pack(side='right', fill='y')
        
        self.rules_listbox = tk.Listbox(list_container, 
                                        yscrollcommand=scrollbar.set,
                                        font=("Courier", 10),
                                        height=8,
                                        bg='#f8f9fa')
        self.rules_listbox.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=self.rules_listbox.yview)
        
        # Boutons règles
        rules_btns = tk.Frame(rules_list_frame, bg='white')
        rules_btns.pack(fill='x', padx=10, pady=10)
        
        tk.Button(rules_btns, text=" 🗑️ Supprimer ",
                 bg='#e74c3c', fg='white',
                 font=("Arial", 10, "bold"),
                 command=self.delete_rule).pack(side='left', padx=5)
        
        tk.Button(rules_btns, text=" 🔄 Actualiser ",
                 bg='#3498db', fg='white',
                 font=("Arial", 10, "bold"),
                 command=self.refresh_rules).pack(side='left', padx=5)
        
        # --- ONGLET 4: EXÉCUTION ---
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
        tk.Spinbox(options_grid, from_=1, to=1000, 
                  textvariable=self.max_emails_var,
                  width=10, font=("Arial", 11)).pack(side='left', padx=5)
        
        tk.Label(options_grid, text="(0 = tous les emails)", 
                font=("Arial", 9, "italic"), bg='white', fg='gray').pack(side='left')
        
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
        
        # Barre de statut
        self.status_var = tk.StringVar(value="✅ Prêt - Email Manager pour Thunderbird")
        status_bar = tk.Label(self.root, textvariable=self.status_var,
                             bd=1, relief=tk.SUNKEN, anchor='w',
                             bg='#34495e', fg='white',
                             font=("Arial", 10))
        status_bar.pack(side='bottom', fill='x')
        
        # Initialisation
        self.log("✨ Email Manager pour Thunderbird démarré!", "success")
        self.log("📌 Compatible avec tous les serveurs IMAP", "info")
        self.log("🔧 Les dossiers seront créés sous INBOX.", "info")
        self.log("📚 Consultez l'onglet Documentation pour l'aide", "warning")
    
    def toggle_cc_options(self):
        """Activer/désactiver les options CC"""
        if self.cc_enabled_var.get():
            self.cc_folder_entry.config(state='normal')
            self.log("✅ Tri automatique des CC activé", "success")
        else:
            self.cc_folder_entry.config(state='disabled')
            self.log("❌ Tri automatique des CC désactivé", "warning")
    
    def test_connection(self):
        """Tester la connexion au serveur"""
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
            self.log(f"📁 {len(folders)} dossiers trouvés:", "info")
            
            # Afficher quelques dossiers
            for folder in folders[:10]:
                try:
                    folder_name = folder.decode('utf-8').split('"')[-2]
                    self.log(f"   • {folder_name}", "info")
                except:
                    pass
            
            if len(folders) > 10:
                self.log(f"   ... et {len(folders)-10} autres", "info")
            
            connection.logout()
            
            messagebox.showinfo("Succès", f"✅ Connexion réussie!\n\nServeur: {server}\nEmail: {self.email_var.get()}\n{len(folders)} dossiers disponibles")
            
        except Exception as e:
            self.log(f"❌ Erreur: {str(e)}", "error")
            error_msg = str(e)
            
            # Messages d'aide spécifiques
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
        try:
            # Toujours utiliser le préfixe INBOX. pour les sous-dossiers
            full_folder_name = self.get_full_folder_name(folder_name)
            
            # Lister tous les dossiers existants
            result, folders = connection.list()
            
            # Vérifier si le dossier existe
            folder_exists = False
            if result == 'OK':
                for folder in folders:
                    if folder:
                        folder_str = folder.decode('utf-8') if isinstance(folder, bytes) else str(folder)
                        # Vérifier si le dossier existe déjà
                        if full_folder_name.lower() in folder_str.lower():
                            folder_exists = True
                            self.log(f"📁 Dossier '{full_folder_name}' déjà existant", "info")
                            break
            
            if not folder_exists:
                # Créer le dossier avec le nom complet
                result = connection.create(full_folder_name)
                if result[0] == 'OK':
                    self.log(f"✅ Dossier '{full_folder_name}' créé avec succès", "success")
                    # S'abonner au nouveau dossier
                    connection.subscribe(full_folder_name)
                    return True
                else:
                    self.log(f"❌ Impossible de créer '{full_folder_name}': {result[1]}", "error")
                    return False
            return True
                    
        except Exception as e:
            self.log(f"⚠️ Erreur avec le dossier '{folder_name}': {str(e)[:100]}", "warning")
            return False
    
    def add_custom_rule(self):
        """Ajouter une règle personnalisée"""
        keyword = self.rule_keyword_var.get().strip()
        folder = self.rule_folder_var.get().strip()
        
        if not keyword or not folder:
            messagebox.showwarning("Attention", "Remplissez le mot-clé et le dossier de destination!")
            return
        
        rule = {
            "field": self.rule_field_var.get(),
            "condition": self.rule_condition_var.get(),
            "keyword": keyword,
            "folder": folder
        }
        
        self.rules.append(rule)
        self.refresh_rules()
        
        # Clear fields
        self.rule_keyword_var.set("")
        self.rule_folder_var.set("")
        
        self.log(f"✅ Règle ajoutée: {rule['field']} {rule['condition']} '{keyword}' → INBOX.{folder}", "success")
        self.save_settings()
    
    def refresh_rules(self):
        """Actualiser la liste des règles"""
        self.rules_listbox.delete(0, tk.END)
        
        for i, rule in enumerate(self.rules, 1):
            text = f"{i}. {rule['field']} {rule['condition']} '{rule['keyword']}' → INBOX.{rule['folder']}"
            self.rules_listbox.insert(tk.END, text)
    
    def delete_rule(self):
        """Supprimer la règle sélectionnée"""
        selection = self.rules_listbox.curselection()
        if selection:
            index = selection[0]
            del self.rules[index]
            self.refresh_rules()
            self.log("🗑️ Règle supprimée", "warning")
            self.save_settings()
        else:
            messagebox.showinfo("Info", "Sélectionnez une règle à supprimer")
    
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
        
        self.is_running = True
        self.analyze_btn.config(state='disabled', text="⏳ ANALYSE EN COURS...")
        self.status_var.set("🔄 Analyse en cours...")
        
        # Sauvegarder avant l'analyse
        self.save_settings()
        
        # Thread pour ne pas bloquer l'interface
        thread = threading.Thread(target=self.analysis_worker, daemon=True)
        thread.start()
    
    def analysis_worker(self):
        """Worker pour l'analyse des emails"""
        try:
            # Connexion
            self.log("\n" + "="*60, "separator")
            self.log("🚀 DÉMARRAGE DE L'ANALYSE", "header")
            self.log("="*60, "separator")
            
            self.log(f"📧 Service: Thunderbird", "info")
            self.log(f"🔌 Connexion à {self.server_var.get()}...", "info")
            
            connection = imaplib.IMAP4_SSL(self.server_var.get(), int(self.port_var.get()))
            connection.login(self.email_var.get(), self.password_var.get())
            
            self.log(f"✅ Connecté avec succès!", "success")
            
            # Détecter le séparateur de hiérarchie
            try:
                result, hierarchy = connection.list('""', '""')
                if result == 'OK' and hierarchy[0]:
                    match = re.search(r'\(.*?\) "(.*?)" ', hierarchy[0].decode())
                    if match:
                        self.folder_separator = match.group(1)
                        self.log(f"📌 Séparateur de dossiers: '{self.folder_separator}'", "info")
            except:
                self.log("⚠️ Utilisation du séparateur par défaut: '.'", "warning")
            
            # Créer les dossiers si nécessaire
            folders_created = True
            
            if self.cc_enabled_var.get():
                if not self.create_folder_if_needed(connection, self.cc_folder_var.get()):
                    folders_created = False
            
            for rule in self.rules:
                if not self.create_folder_if_needed(connection, rule['folder']):
                    folders_created = False
            
            if not folders_created:
                self.log("⚠️ Certains dossiers n'ont pas pu être créés", "warning")
            
            # Sélectionner INBOX
            connection.select('INBOX')
            
            # Rechercher tous les emails
            result, data = connection.search(None, 'ALL')
            
            if result != 'OK':
                self.log("❌ Erreur lors de la recherche des emails", "error")
                return
            
            email_ids = data[0].split()
            total = len(email_ids)
            
            if total == 0:
                self.log("📭 Aucun email dans la boîte de réception", "warning")
                return
            
            # Limiter si nécessaire
            try:
                max_emails_str = self.max_emails_var.get().strip()
                max_emails = int(max_emails_str) if max_emails_str and max_emails_str.isdigit() else 0
            except:
                max_emails = 0
                
            if max_emails > 0 and total > max_emails:
                email_ids = email_ids[-max_emails:]
                total = max_emails
            
            self.log(f"📬 {total} emails à analyser", "info")
            self.log("-" * 50, "separator")
            
            cc_moved = 0
            rules_moved = 0
            processed = 0
            
            for num in email_ids:
                if not self.is_running:
                    self.log("⏹️ Analyse interrompue par l'utilisateur", "warning")
                    break
                
                processed += 1
                
                try:
                    # Progress
                    if processed % 10 == 0:
                        self.status_var.set(f"🔄 Analyse... {processed}/{total} emails traités")
                    
                    # Récupérer l'email
                    result, msg_data = connection.fetch(num, '(RFC822)')
                    raw_email = msg_data[0][1]
                    msg = email.message_from_bytes(raw_email)
                    
                    # Décoder les headers
                    subject = self.decode_header(msg.get("Subject", ""))[:100]
                    from_addr = self.decode_header(msg.get("From", ""))
                    to_addr = self.decode_header(msg.get("To", ""))
                    cc_addr = self.decode_header(msg.get("Cc", ""))
                    
                    # Vérifier si l'utilisateur est en CC
                    user_email = self.email_var.get().lower()
                    is_in_cc = False
                    is_primary = False
                    
                    if cc_addr and user_email in cc_addr.lower():
                        is_in_cc = True
                    
                    if to_addr and user_email in to_addr.lower():
                        is_primary = True
                    
                    moved = False
                    
                    # Gestion CC (priorité sur les règles)
                    if self.cc_enabled_var.get() and is_in_cc and not is_primary:
                        # Vérifier si pas marqué important
                        skip = False
                        if self.cc_skip_important_var.get():
                            result, flag_data = connection.fetch(num, '(FLAGS)')
                            if flag_data and flag_data[0]:
                                flags_str = flag_data[0] if isinstance(flag_data[0], bytes) else flag_data[0].encode()
                                if b'\\Flagged' in flags_str:
                                    skip = True
                        
                        if not skip:
                            try:
                                folder_name = self.get_full_folder_name(self.cc_folder_var.get())
                                result = connection.copy(num, folder_name)
                                
                                if result[0] == 'OK':
                                    connection.store(num, '+FLAGS', '\\Deleted')
                                    if self.cc_mark_read_var.get():
                                        connection.store(num, '+FLAGS', '\\Seen')
                                    cc_moved += 1
                                    moved = True
                                    self.log(f"📋 [CC] {subject[:50]}... → {folder_name}", "cc")
                                else:
                                    self.log(f"⚠️ Impossible de déplacer vers {folder_name}", "warning")
                            except Exception as e:
                                self.log(f"⚠️ Erreur déplacement CC: {str(e)[:100]}", "error")
                    
                    # Appliquer les règles personnalisées si pas déjà déplacé
                    if not moved and self.rules:
                        body = self.get_email_body(msg)
                        
                        for rule in self.rules:
                            if self.check_rule(msg, subject, from_addr, to_addr, body, rule):
                                try:
                                    folder_name = self.get_full_folder_name(rule['folder'])
                                    result = connection.copy(num, folder_name)
                                    
                                    if result[0] == 'OK':
                                        connection.store(num, '+FLAGS', '\\Deleted')
                                        rules_moved += 1
                                        self.log(f"🎯 [Règle] {subject[:50]}... → {folder_name}", "rule")
                                        break
                                    else:
                                        self.log(f"⚠️ Impossible de déplacer vers {folder_name}", "warning")
                                except Exception as e:
                                    self.log(f"⚠️ Erreur déplacement règle: {str(e)[:100]}", "error")
                    
                except Exception as e:
                    self.log(f"⚠️ Erreur sur un email: {str(e)[:100]}", "error")
                    continue
            
            # Expurger les messages supprimés
            self.log("-" * 50, "separator")
            self.log("🗑️ Nettoyage des emails déplacés...", "info")
            connection.expunge()
            
            # Déconnexion
            connection.logout()
            
            # Résumé final
            self.log("\n" + "="*60, "separator")
            self.log("📊 RÉSUMÉ DE L'ANALYSE", "header")
            self.log("="*60, "separator")
            self.log(f"✅ Emails analysés: {processed}/{total}", "success")
            self.log(f"📋 Emails en CC déplacés: {cc_moved}", "info")
            self.log(f"🎯 Emails triés par règles: {rules_moved}", "info")
            self.log(f"📧 TOTAL déplacé: {cc_moved + rules_moved} emails", "success")
            
            self.status_var.set(f"✅ Terminé - {cc_moved + rules_moved} emails triés sur {processed} analysés")
            
            if cc_moved + rules_moved > 0:
                messagebox.showinfo("Analyse terminée", 
                                   f"✅ Analyse terminée avec succès!\n\n"
                                   f"📊 Résultats:\n"
                                   f"• {processed} emails analysés\n"
                                   f"• {cc_moved} emails en CC déplacés\n"
                                   f"• {rules_moved} emails triés par règles\n"
                                   f"• Total: {cc_moved + rules_moved} emails organisés")
            else:
                messagebox.showinfo("Analyse terminée", 
                                   f"Analyse terminée.\n\n"
                                   f"📊 {processed} emails analysés\n"
                                   f"Aucun email à déplacer selon les critères configurés.")
            
        except Exception as e:
            self.log(f"❌ Erreur critique: {str(e)}", "error")
            self.status_var.set("❌ Erreur - Vérifiez la connexion")
            
            error_msg = str(e)
            if "authentication" in error_msg.lower():
                error_msg += "\n\n💡 Vérifiez:\n• Le serveur IMAP est correct\n• L'email et le mot de passe\n• Pour Gmail: utilisez un mot de passe d'application"
            
            messagebox.showerror("Erreur", f"Erreur lors de l'analyse:\n\n{error_msg}")
        
        finally:
            self.is_running = False
            self.root.after(0, lambda: self.analyze_btn.config(
                state='normal',
                text="🚀 ANALYSER ET TRIER LES EMAILS"
            ))
    
    def check_rule(self, msg, subject, from_addr, to_addr, body, rule):
        """Vérifier si un email correspond à une règle avec plus de conditions"""
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
                import re
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
    
    def log(self, message, tag="info"):
        """Ajouter un message au log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        if tag not in ["separator", "header"]:
            message = f"[{timestamp}] {message}"
        
        self.console.insert(tk.END, message + "\n")
        
        # Colorer selon le type
        if tag == "error":
            self.console.tag_add("error", f"end-2l", "end-1l")
            self.console.tag_config("error", foreground="#ff6b6b")
        elif tag == "success":
            self.console.tag_add("success", f"end-2l", "end-1l")
            self.console.tag_config("success", foreground="#51cf66")
        elif tag == "warning":
            self.console.tag_add("warning", f"end-2l", "end-1l")
            self.console.tag_config("warning", foreground="#ffd43b")
        elif tag == "cc":
            self.console.tag_add("cc", f"end-2l", "end-1l")
            self.console.tag_config("cc", foreground="#74c0fc")
        elif tag == "rule":
            self.console.tag_add("rule", f"end-2l", "end-1l")
            self.console.tag_config("rule", foreground="#ff8cc3")
        elif tag == "header":
            self.console.tag_add("header", f"end-2l", "end-1l")
            self.console.tag_config("header", foreground="#ffffff", font=("Consolas", 11, "bold"))
        
        self.console.see(tk.END)
        self.root.update_idletasks()
    
    def save_settings(self):
        """Sauvegarder les paramètres"""
        settings = {
            "email": self.email_var.get(),
            "server": self.server_var.get(),
            "port": self.port_var.get(),
            "cc_enabled": self.cc_enabled_var.get(),
            "cc_folder": self.cc_folder_var.get(),
            "cc_mark_read": self.cc_mark_read_var.get(),
            "cc_skip_important": self.cc_skip_important_var.get(),
            "max_emails": self.max_emails_var.get(),
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
                
                self.email_var.set(settings.get("email", ""))
                self.server_var.set(settings.get("server", ""))
                self.port_var.set(settings.get("port", "993"))
                self.cc_enabled_var.set(settings.get("cc_enabled", True))
                self.cc_folder_var.set(settings.get("cc_folder", "EN_COPIE"))
                self.cc_mark_read_var.set(settings.get("cc_mark_read", False))
                self.cc_skip_important_var.set(settings.get("cc_skip_important", True))
                self.max_emails_var.set(settings.get("max_emails", "100"))
                self.rules = settings.get("rules", [])
                
                self.refresh_rules()
                self.log("✅ Paramètres chargés depuis la dernière session", "success")
        except:
            self.log("ℹ️ Première utilisation - Paramètres par défaut", "info")
    
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
    """Fonction principale pour lancer Email Manager"""
    app = EmailManager()
    app.run()

if __name__ == "__main__":
    main()