# -*- coding: utf-8 -*-
import os
import re
import html
import tempfile
import subprocess
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from datetime import datetime
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime

try:
    import extract_msg
except ImportError:
    extract_msg = None


SUSPICIOUS_EXTENSIONS = {
    ".exe", ".bat", ".cmd", ".com", ".scr", ".js", ".jse", ".vbs", ".vbe",
    ".ps1", ".psm1", ".hta", ".msi", ".dll", ".jar", ".reg", ".iso",
    ".zip", ".rar", ".7z", ".cab", ".lnk", ".chm"
}

SUSPICIOUS_KEYWORDS = [
    "urgent", "immediately", "verify", "verification", "login", "reset password",
    "payment failed", "invoice attached", "open attachment", "enable macros",
    "click here", "confirm account", "security alert", "your account",
    "κάντε login", "επαλήθευση", "επιβεβαίωση", "επείγον", "άμεσα",
    "κωδικός", "λογαριασμός", "συνδεθείτε", "πατήστε εδώ", "τιμολόγιο",
    "επισυναπτόμενο", "άνοιξε το συνημμένο", "ενεργοποίηση μακροεντολών"
]

SHORTENER_DOMAINS = {
    "bit.ly", "tinyurl.com", "t.co", "goo.gl", "ow.ly", "is.gd", "rb.gy", "buff.ly"
}

TRUST_BRANDS = ["microsoft", "google", "paypal", "amazon", "bank", "apple", "dhl", "acs"]

URL_REGEX = re.compile(r'https?://[^\s<>"\']+|www\.[^\s<>"\']+', re.IGNORECASE)
EMAIL_REGEX = re.compile(r'[\w\.-]+@[\w\.-]+\.\w+', re.IGNORECASE)
IP_HOST_REGEX = re.compile(r'^(?:\d{1,3}\.){3}\d{1,3}$')

TEXTS = {
    "el": {
        "app_title": "Mail Viewer (.MSG / .EML) & Basic Security Check",
        "open_files": "Άνοιγμα email αρχείων",
        "open_folder": "Άνοιγμα φακέλου",
        "clear": "Καθαρισμός",
        "load_hint": "Φόρτωσε .msg ή .eml αρχεία",
        "files_list": "Αρχεία email",
        "file": "Αρχείο:",
        "type": "Τύπος:",
        "subject": "Θέμα:",
        "from": "Από:",
        "to": "Προς:",
        "date": "Ημερομηνία:",
        "tab_message": "Μήνυμα",
        "tab_security": "Έλεγχος Ασφάλειας",
        "message_body": "Περιεχόμενο μηνύματος",
        "alerts": "Alerts / Ευρήματα",
        "links": "Links",
        "attachments": "Συνημμένα",
        "flags": "Λέξεις / Flags",
        "ready": "Έτοιμο",
        "risk_score": "Risk Score",
        "level": "Επίπεδο",
        "defender_scan": "Defender Scan",
        "lang_label": "Γλώσσα:",
        "select_files_title": "Επιλογή email αρχείων",
        "select_folder_title": "Επιλογή φακέλου με .msg / .eml αρχεία",
        "missing_files": "Δεν βρέθηκαν αρχεία",
        "missing_files_msg": "Δεν βρέθηκαν .msg ή .eml αρχεία στον φάκελο.",
        "read_error_title": "Σφάλμα ανάγνωσης",
        "loaded_files": "Φορτώθηκαν {count} αρχεία.",
        "loaded_folder_files": "Φορτώθηκαν {count} αρχεία από τον φάκελο.",
        "total_loaded": "Σύνολο φορτωμένων μηνυμάτων: {count}",
        "showing_message": "Εμφανίζεται μήνυμα {index} από {total}",
        "cleared": "Καθαρίστηκαν όλα τα δεδομένα.",
        "missing_lib_title": "Λείπει βιβλιοθήκη",
        "missing_extract_msg": "Για αρχεία .msg χρειάζεται η βιβλιοθήκη extract_msg.\n\nΤρέξε:\npip install extract_msg",
        "unsupported_type": "Μη υποστηριζόμενος τύπος αρχείου",
        "sender_missing": "Λείπει αποστολέας.",
        "display_name_spoof": "Πιθανό display-name spoofing: εμφανίζεται '{brand}' αλλά domain '{domain}'.",
        "sender_short_domain": "Ο αποστολέας χρησιμοποιεί shortened/ύποπτο domain: {domain}",
        "ip_link": "Link με IP αντί για domain: {url}",
        "shortened_link": "Shortened link: {url}",
        "brand_mismatch_link": "Πιθανό brand mismatch σε link: {url}",
        "suspicious_attachment": "Ύποπτο συνημμένο: {name}",
        "many_links": "Πολλά links στο μήνυμα: {count}",
        "found_keywords": "Βρέθηκαν phishing/πίεσης keywords.",
        "uppercase_subject": "Το θέμα είναι όλο με κεφαλαία.",
        "defender_detected": "Ο Windows Defender εντόπισε απειλή σε συνημμένο.",
        "defender_error": "Defender scan δεν ολοκληρώθηκε: {msg}",
        "no_obvious_findings": "Δεν βρέθηκαν εμφανή ύποπτα στοιχεία με τον βασικό heuristic έλεγχο.",
        "no_links": "(Δεν βρέθηκαν links)",
        "no_attachments": "(Δεν υπάρχουν συνημμένα)",
        "no_keywords": "(Δεν βρέθηκαν flagged keywords)",
        "suspicious_tag": " [ΥΠΟΠΤΟ]",
        "no_name": "(χωρίς όνομα)",
        "no_subject": "(Χωρίς θέμα)",
        "unknown_sender": "(Άγνωστος αποστολέας)",
        "no_content": "(Δεν βρέθηκε περιεχόμενο μηνύματος)",
        "not_available": "not_available",
        "no_attachments_status": "Δεν υπάρχουν συνημμένα.",
        "no_extractable": "Δεν ήταν δυνατή η εξαγωγή συνημμένων για scan.",
        "clean": "Δεν εντοπίστηκε απειλή από τον Defender.",
        "defender_not_found": "Δεν βρέθηκε το MpCmdRun.exe του Windows Defender.",
        "defender_timeout": "Ο Defender scan έληξε λόγω timeout.",
        "risk_low": "Χαμηλό Ρίσκο",
        "risk_medium": "Μέτριο Ρίσκο",
        "risk_high": "Υψηλό Ρίσκο",
        "risk_very_high": "Πολύ Ύποπτο",
    },
    "en": {
        "app_title": "Mail Viewer (.MSG / .EML) & Basic Security Check",
        "open_files": "Open email files",
        "open_folder": "Open folder",
        "clear": "Clear",
        "load_hint": "Load .msg or .eml files",
        "files_list": "Email files",
        "file": "File:",
        "type": "Type:",
        "subject": "Subject:",
        "from": "From:",
        "to": "To:",
        "date": "Date:",
        "tab_message": "Message",
        "tab_security": "Security Check",
        "message_body": "Message body",
        "alerts": "Alerts / Findings",
        "links": "Links",
        "attachments": "Attachments",
        "flags": "Keywords / Flags",
        "ready": "Ready",
        "risk_score": "Risk Score",
        "level": "Level",
        "defender_scan": "Defender Scan",
        "lang_label": "Language:",
        "select_files_title": "Select email files",
        "select_folder_title": "Select folder with .msg / .eml files",
        "missing_files": "No files found",
        "missing_files_msg": "No .msg or .eml files were found in the selected folder.",
        "read_error_title": "Read error",
        "loaded_files": "Loaded {count} files.",
        "loaded_folder_files": "Loaded {count} files from folder.",
        "total_loaded": "Total loaded messages: {count}",
        "showing_message": "Showing message {index} of {total}",
        "cleared": "All data cleared.",
        "missing_lib_title": "Missing library",
        "missing_extract_msg": "The extract_msg library is required for .msg files.\n\nRun:\npip install extract_msg",
        "unsupported_type": "Unsupported file type",
        "sender_missing": "Sender is missing.",
        "display_name_spoof": "Possible display-name spoofing: '{brand}' shown but domain is '{domain}'.",
        "sender_short_domain": "Sender uses shortened/suspicious domain: {domain}",
        "ip_link": "Link uses IP instead of domain: {url}",
        "shortened_link": "Shortened link: {url}",
        "brand_mismatch_link": "Possible brand mismatch in link: {url}",
        "suspicious_attachment": "Suspicious attachment: {name}",
        "many_links": "Many links in message: {count}",
        "found_keywords": "Phishing / pressure keywords were found.",
        "uppercase_subject": "Subject is all uppercase.",
        "defender_detected": "Windows Defender detected a threat in an attachment.",
        "defender_error": "Defender scan did not complete: {msg}",
        "no_obvious_findings": "No obvious suspicious findings were detected by the basic heuristic check.",
        "no_links": "(No links found)",
        "no_attachments": "(No attachments)",
        "no_keywords": "(No flagged keywords found)",
        "suspicious_tag": " [SUSPICIOUS]",
        "no_name": "(no name)",
        "no_subject": "(No subject)",
        "unknown_sender": "(Unknown sender)",
        "no_content": "(No message content found)",
        "not_available": "not_available",
        "no_attachments_status": "No attachments.",
        "no_extractable": "Attachments could not be extracted for scanning.",
        "clean": "No threat detected by Defender.",
        "defender_not_found": "Windows Defender MpCmdRun.exe was not found.",
        "defender_timeout": "Defender scan timed out.",
        "risk_low": "Low Risk",
        "risk_medium": "Medium Risk",
        "risk_high": "High Risk",
        "risk_very_high": "Highly Suspicious",
    }
}


class MailViewerApp:
    def __init__(self, root):
        self.root = root
        self.current_lang = "el"
        self.messages = []
        self.current_index = None

        self.root.geometry("1360x840")
        self.root.minsize(1080, 700)

        self._build_ui()
        self.apply_language(refresh_current=False)

    def t(self, key, **kwargs):
        text = TEXTS[self.current_lang][key]
        if kwargs:
            return text.format(**kwargs)
        return text

    def _build_ui(self):
        self.file_var = tk.StringVar()
        self.type_var = tk.StringVar()
        self.subject_var = tk.StringVar()
        self.from_var = tk.StringVar()
        self.to_var = tk.StringVar()
        self.date_var = tk.StringVar()

        self.top_frame = ttk.Frame(self.root, padding=10)
        self.top_frame.pack(fill="x")

        self.btn_open_files = ttk.Button(self.top_frame, command=self.open_files)
        self.btn_open_files.pack(side="left")

        self.btn_open_folder = ttk.Button(self.top_frame, command=self.open_folder)
        self.btn_open_folder.pack(side="left", padx=(8, 0))

        self.btn_clear = ttk.Button(self.top_frame, command=self.clear_all)
        self.btn_clear.pack(side="left", padx=(8, 0))

        self.top_info_label = ttk.Label(self.top_frame)
        self.top_info_label.pack(side="left", padx=(14, 0))

        self.lang_label = ttk.Label(self.top_frame)
        self.lang_label.pack(side="right", padx=(8, 0))

        self.lang_var = tk.StringVar(value="Ελληνικά")
        self.lang_combo = ttk.Combobox(
            self.top_frame,
            textvariable=self.lang_var,
            values=["Ελληνικά", "English"],
            state="readonly",
            width=12
        )
        self.lang_combo.pack(side="right")
        self.lang_combo.bind("<<ComboboxSelected>>", self.on_language_change)

        self.main = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.main.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.left_frame = ttk.Frame(self.main, padding=8)
        self.main.add(self.left_frame, weight=1)

        self.files_label = ttk.Label(self.left_frame)
        self.files_label.pack(anchor="w")

        self.listbox = tk.Listbox(self.left_frame, activestyle="dotbox", exportselection=False, font=("Segoe UI", 10))
        self.listbox.pack(fill="both", expand=True, pady=(6, 0))
        self.listbox.bind("<<ListboxSelect>>", self.on_select_message)

        self.right_frame = ttk.Frame(self.main, padding=8)
        self.main.add(self.right_frame, weight=3)

        self.form = ttk.Frame(self.right_frame)
        self.form.pack(fill="x")

        self.lbl_file = ttk.Label(self.form, width=12)
        self.lbl_file.grid(row=0, column=0, sticky="w", pady=4)
        self.ent_file = ttk.Entry(self.form, textvariable=self.file_var, state="readonly")
        self.ent_file.grid(row=0, column=1, sticky="ew", pady=4)

        self.lbl_type = ttk.Label(self.form, width=12)
        self.lbl_type.grid(row=1, column=0, sticky="w", pady=4)
        self.ent_type = ttk.Entry(self.form, textvariable=self.type_var, state="readonly")
        self.ent_type.grid(row=1, column=1, sticky="ew", pady=4)

        self.lbl_subject = ttk.Label(self.form, width=12)
        self.lbl_subject.grid(row=2, column=0, sticky="w", pady=4)
        self.ent_subject = ttk.Entry(self.form, textvariable=self.subject_var, state="readonly")
        self.ent_subject.grid(row=2, column=1, sticky="ew", pady=4)

        self.lbl_from = ttk.Label(self.form, width=12)
        self.lbl_from.grid(row=3, column=0, sticky="w", pady=4)
        self.ent_from = ttk.Entry(self.form, textvariable=self.from_var, state="readonly")
        self.ent_from.grid(row=3, column=1, sticky="ew", pady=4)

        self.lbl_to = ttk.Label(self.form, width=12)
        self.lbl_to.grid(row=4, column=0, sticky="w", pady=4)
        self.ent_to = ttk.Entry(self.form, textvariable=self.to_var, state="readonly")
        self.ent_to.grid(row=4, column=1, sticky="ew", pady=4)

        self.lbl_date = ttk.Label(self.form, width=12)
        self.lbl_date.grid(row=5, column=0, sticky="w", pady=4)
        self.ent_date = ttk.Entry(self.form, textvariable=self.date_var, state="readonly")
        self.ent_date.grid(row=5, column=1, sticky="ew", pady=4)

        self.form.columnconfigure(1, weight=1)

        self.notebook = ttk.Notebook(self.right_frame)
        self.notebook.pack(fill="both", expand=True, pady=(12, 0))

        self.preview_tab = ttk.Frame(self.notebook, padding=8)
        self.security_tab = ttk.Frame(self.notebook, padding=8)
        self.notebook.add(self.preview_tab, text="")
        self.notebook.add(self.security_tab, text="")

        self.message_body_label = ttk.Label(self.preview_tab)
        self.message_body_label.pack(anchor="w", pady=(0, 4))

        self.body_text = ScrolledText(self.preview_tab, wrap="word", font=("Segoe UI", 11))
        self.body_text.pack(fill="both", expand=True)

        self._build_security_tab()

        self.status_var = tk.StringVar(value="")
        self.status_label = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w", padding=6)
        self.status_label.pack(fill="x", side="bottom")

    def _build_security_tab(self):
        top = ttk.Frame(self.security_tab)
        top.pack(fill="x")

        self.risk_var = tk.StringVar(value="")
        self.level_var = tk.StringVar(value="")
        self.defender_var = tk.StringVar(value="")

        ttk.Label(top, textvariable=self.risk_var, font=("Segoe UI", 11, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 16))
        ttk.Label(top, textvariable=self.level_var, font=("Segoe UI", 11, "bold")).grid(row=0, column=1, sticky="w", padx=(0, 16))
        ttk.Label(top, textvariable=self.defender_var, font=("Segoe UI", 11, "bold")).grid(row=0, column=2, sticky="w")

        lists_frame = ttk.Frame(self.security_tab)
        lists_frame.pack(fill="both", expand=True, pady=(10, 0))

        left = ttk.Frame(lists_frame)
        left.pack(side="left", fill="both", expand=True, padx=(0, 5))

        right = ttk.Frame(lists_frame)
        right.pack(side="left", fill="both", expand=True, padx=(5, 0))

        self.alerts_label = ttk.Label(left)
        self.alerts_label.pack(anchor="w")
        self.alerts_text = ScrolledText(left, height=12, wrap="word", font=("Consolas", 10))
        self.alerts_text.pack(fill="both", expand=True, pady=(4, 8))

        self.links_label = ttk.Label(left)
        self.links_label.pack(anchor="w")
        self.links_text = ScrolledText(left, height=8, wrap="word", font=("Consolas", 10))
        self.links_text.pack(fill="both", expand=True, pady=(4, 0))

        self.attachments_label = ttk.Label(right)
        self.attachments_label.pack(anchor="w")
        self.attachments_text = ScrolledText(right, height=12, wrap="word", font=("Consolas", 10))
        self.attachments_text.pack(fill="both", expand=True, pady=(4, 8))

        self.flags_label = ttk.Label(right)
        self.flags_label.pack(anchor="w")
        self.flags_text = ScrolledText(right, height=8, wrap="word", font=("Consolas", 10))
        self.flags_text.pack(fill="both", expand=True, pady=(4, 0))

    def on_language_change(self, event=None):
        selected = self.lang_var.get()
        self.current_lang = "en" if selected == "English" else "el"
        self.apply_language(refresh_current=True)

    def apply_language(self, refresh_current=True):
        self.root.title(self.t("app_title"))
        self.btn_open_files.config(text=self.t("open_files"))
        self.btn_open_folder.config(text=self.t("open_folder"))
        self.btn_clear.config(text=self.t("clear"))
        self.top_info_label.config(text=self.t("load_hint") if not self.messages else self.t("total_loaded", count=len(self.messages)))
        self.lang_label.config(text=self.t("lang_label"))
        self.files_label.config(text=self.t("files_list"))
        self.lbl_file.config(text=self.t("file"))
        self.lbl_type.config(text=self.t("type"))
        self.lbl_subject.config(text=self.t("subject"))
        self.lbl_from.config(text=self.t("from"))
        self.lbl_to.config(text=self.t("to"))
        self.lbl_date.config(text=self.t("date"))
        self.notebook.tab(0, text=self.t("tab_message"))
        self.notebook.tab(1, text=self.t("tab_security"))
        self.message_body_label.config(text=self.t("message_body"))
        self.alerts_label.config(text=self.t("alerts"))
        self.links_label.config(text=self.t("links"))
        self.attachments_label.config(text=self.t("attachments"))
        self.flags_label.config(text=self.t("flags"))

        if self.status_var.get() in ("", TEXTS["el"]["ready"], TEXTS["en"]["ready"]):
            self.status_var.set(self.t("ready"))

        self._refresh_list()

        if refresh_current and self.current_index is not None and 0 <= self.current_index < len(self.messages):
            self._display_message(self.current_index)
        else:
            self.risk_var.set(f"{self.t('risk_score')}: -")
            self.level_var.set(f"{self.t('level')}: -")
            self.defender_var.set(f"{self.t('defender_scan')}: -")

    def open_files(self):
        file_paths = filedialog.askopenfilenames(
            title=self.t("select_files_title"),
            filetypes=[("Email files", "*.msg *.eml"), ("MSG files", "*.msg"), ("EML files", "*.eml")]
        )
        if not file_paths:
            return

        loaded = 0
        for path in file_paths:
            if self._load_message(path):
                loaded += 1

        self._refresh_list()
        self.status_var.set(self.t("loaded_files", count=loaded))
        self.top_info_label.config(text=self.t("total_loaded", count=len(self.messages)))

        if self.messages and self.current_index is None:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(0)
            self.listbox.event_generate("<<ListboxSelect>>")

    def open_folder(self):
        folder = filedialog.askdirectory(title=self.t("select_folder_title"))
        if not folder:
            return

        mail_files = [
            os.path.join(folder, name)
            for name in os.listdir(folder)
            if name.lower().endswith((".msg", ".eml"))
        ]
        if not mail_files:
            messagebox.showinfo(self.t("missing_files"), self.t("missing_files_msg"))
            return

        loaded = 0
        for path in sorted(mail_files):
            if self._load_message(path):
                loaded += 1

        self._refresh_list()
        self.status_var.set(self.t("loaded_folder_files", count=loaded))
        self.top_info_label.config(text=self.t("total_loaded", count=len(self.messages)))

        if self.messages and self.current_index is None:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(0)
            self.listbox.event_generate("<<ListboxSelect>>")

    def _load_message(self, path):
        for item in self.messages:
            if os.path.abspath(item["path"]) == os.path.abspath(path):
                return False

        try:
            ext = os.path.splitext(path)[1].lower()
            if ext == ".msg":
                if extract_msg is None:
                    raise RuntimeError(self.t("missing_extract_msg"))
                data = self._read_msg(path)
            elif ext == ".eml":
                data = self._read_eml(path)
            else:
                raise RuntimeError(f"{self.t('unsupported_type')}: {ext}")

            data["analysis"] = self._analyze_message(
                data["subject"], data["sender"], data["to"], data["body"], data["attachments"]
            )
            self.messages.append(data)
            return True
        except Exception as e:
            messagebox.showwarning(self.t("read_error_title"), f"{path}\n\n{e}")
            return False

    def _read_msg(self, path):
        msg = extract_msg.Message(path)
        return {
            "path": path,
            "mail_type": "MSG",
            "subject": self._safe_text(msg.subject),
            "sender": self._safe_text(msg.sender),
            "to": self._safe_text(msg.to),
            "date": self._format_date(msg.date),
            "body": self._extract_msg_body(msg),
            "attachments": self._extract_msg_attachments_info(msg)
        }

    def _read_eml(self, path):
        with open(path, "rb") as f:
            eml = BytesParser(policy=policy.default).parse(f)

        return {
            "path": path,
            "mail_type": "EML",
            "subject": self._safe_text(eml.get("subject", "")),
            "sender": self._safe_text(eml.get("from", "")),
            "to": self._safe_text(eml.get("to", "")),
            "date": self._format_date(self._parse_email_date(eml.get("date", ""))),
            "body": self._extract_eml_body(eml),
            "attachments": self._extract_eml_attachments_info(eml)
        }

    def _extract_msg_body(self, msg):
        body = self._safe_text(getattr(msg, "body", ""))
        if body:
            return body

        html_body = getattr(msg, "htmlBody", b"")
        if isinstance(html_body, bytes):
            decoded = ""
            for enc in ("utf-8", "cp1253", "latin1"):
                try:
                    decoded = html_body.decode(enc, errors="replace")
                    break
                except Exception:
                    decoded = ""
            html_body = decoded

        html_body = self._safe_text(html_body)
        if html_body:
            text = re.sub(r"<[^>]+>", " ", html_body)
            text = html.unescape(text)
            text = re.sub(r"\s+", " ", text)
            return text.strip()

        return self.t("no_content")

    def _extract_eml_body(self, eml):
        plain_parts = []
        html_parts = []

        if eml.is_multipart():
            for part in eml.walk():
                content_disposition = str(part.get_content_disposition() or "").lower()
                if content_disposition == "attachment":
                    continue

                content_type = str(part.get_content_type() or "").lower()
                try:
                    payload = part.get_content()
                except Exception:
                    try:
                        raw = part.get_payload(decode=True) or b""
                        charset = part.get_content_charset() or "utf-8"
                        payload = raw.decode(charset, errors="replace")
                    except Exception:
                        payload = ""

                payload = self._safe_text(payload)
                if not payload:
                    continue

                if content_type == "text/plain":
                    plain_parts.append(payload)
                elif content_type == "text/html":
                    html_parts.append(payload)
        else:
            try:
                payload = eml.get_content()
            except Exception:
                raw = eml.get_payload(decode=True) or b""
                charset = eml.get_content_charset() or "utf-8"
                payload = raw.decode(charset, errors="replace")

            payload = self._safe_text(payload)
            ctype = str(eml.get_content_type() or "").lower()

            if ctype == "text/html":
                html_parts.append(payload)
            else:
                plain_parts.append(payload)

        if plain_parts:
            return "\n\n".join(p for p in plain_parts if p).strip()

        if html_parts:
            text = "\n\n".join(html_parts)
            text = re.sub(r"<[^>]+>", " ", text)
            text = html.unescape(text)
            text = re.sub(r"\s+", " ", text)
            return text.strip()

        return self.t("no_content")

    def _extract_msg_attachments_info(self, msg):
        results = []
        for att in getattr(msg, "attachments", []) or []:
            name = self._safe_text(getattr(att, "longFilename", "")) or self._safe_text(getattr(att, "filename", "")) or self.t("no_name")
            ext = os.path.splitext(name)[1].lower()
            data = getattr(att, "data", None)
            size_text = f"{len(data)} bytes" if isinstance(data, (bytes, bytearray)) else "-"
            results.append({"name": name, "ext": ext, "size": size_text, "data": data if isinstance(data, (bytes, bytearray)) else None})
        return results

    def _extract_eml_attachments_info(self, eml):
        results = []
        for part in eml.iter_attachments():
            name = self._safe_text(part.get_filename()) or self.t("no_name")
            ext = os.path.splitext(name)[1].lower()
            try:
                raw = part.get_payload(decode=True)
            except Exception:
                raw = None
            size_text = f"{len(raw)} bytes" if isinstance(raw, (bytes, bytearray)) else "-"
            results.append({"name": name, "ext": ext, "size": size_text, "data": raw if isinstance(raw, (bytes, bytearray)) else None})
        return results

    def _analyze_message(self, subject, sender, to_, body, attachments):
        alerts = []
        score = 0

        links = self._extract_urls(body)
        found_keywords = self._find_keywords(" ".join([subject, body]))
        sender_email = self._extract_email(sender)
        sender_domain = self._domain_from_email(sender_email)
        sender_text_lower = sender.lower()

        if not sender.strip():
            alerts.append(self.t("sender_missing"))
            score += 2

        if sender_domain:
            for brand in TRUST_BRANDS:
                if brand in sender_text_lower and brand not in sender_domain:
                    alerts.append(self.t("display_name_spoof", brand=brand, domain=sender_domain))
                    score += 3
                    break
            if sender_domain in SHORTENER_DOMAINS:
                alerts.append(self.t("sender_short_domain", domain=sender_domain))
                score += 2

        for url in links:
            host = self._host_from_url(url)
            if not host:
                continue
            if IP_HOST_REGEX.match(host):
                alerts.append(self.t("ip_link", url=url))
                score += 3
            if host.lower() in SHORTENER_DOMAINS:
                alerts.append(self.t("shortened_link", url=url))
                score += 2
            for brand in TRUST_BRANDS:
                if brand in body.lower() and brand not in host.lower() and any(b in url.lower() for b in ["login", "verify", "secure", "account"]):
                    alerts.append(self.t("brand_mismatch_link", url=url))
                    score += 2
                    break

        suspicious_attachments = []
        for att in attachments:
            if att["ext"] in SUSPICIOUS_EXTENSIONS:
                suspicious_attachments.append(att["name"])
                alerts.append(self.t("suspicious_attachment", name=att["name"]))
                score += 3

        if len(links) >= 3:
            alerts.append(self.t("many_links", count=len(links)))
            score += 1

        if found_keywords:
            alerts.append(self.t("found_keywords"))
            score += min(3, len(found_keywords))

        if subject and subject.isupper() and len(subject) > 8:
            alerts.append(self.t("uppercase_subject"))
            score += 1

        defender_result = self._scan_attachments_with_defender(attachments)
        if defender_result["status"] == "malicious":
            alerts.append(self.t("defender_detected"))
            score += 6
        elif defender_result["status"] == "error":
            alerts.append(self.t("defender_error", msg=defender_result["message"]))
            score += 1

        return {
            "score": score,
            "level": self._risk_level(score),
            "alerts": alerts or [self.t("no_obvious_findings")],
            "links": links,
            "keywords": found_keywords,
            "suspicious_attachments": suspicious_attachments,
            "defender": defender_result
        }

    def _scan_attachments_with_defender(self, attachments):
        if not attachments:
            return {"status": "no_attachments", "message": self.t("no_attachments_status")}

        mpcmdrun = r"C:\Program Files\Windows Defender\MpCmdRun.exe"
        if not os.path.exists(mpcmdrun):
            return {"status": self.t("not_available"), "message": self.t("defender_not_found")}

        temp_dir = tempfile.mkdtemp(prefix="mail_attach_scan_")
        try:
            written = 0
            for i, att in enumerate(attachments, start=1):
                if att["data"] is None:
                    continue
                safe_name = re.sub(r'[\\/:*?"<>|]+', "_", att["name"]) or f"attachment_{i}"
                out_path = os.path.join(temp_dir, safe_name)
                with open(out_path, "wb") as f:
                    f.write(att["data"])
                written += 1

            if written == 0:
                return {"status": "no_extractable", "message": self.t("no_extractable")}

            cmd = [mpcmdrun, "-Scan", "-ScanType", "3", "-File", temp_dir]
            proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=180)
            combined = "\n".join([proc.stdout or "", proc.stderr or ""]).strip().lower()

            if any(x in combined for x in ["threat", "infected", "malware", "virus", "found"]):
                return {"status": "malicious", "message": combined[:1500] or "Threat found"}
            if proc.returncode == 0:
                return {"status": "clean", "message": self.t("clean")}
            return {"status": "unknown", "message": combined[:1500] or f"Return code: {proc.returncode}"}
        except subprocess.TimeoutExpired:
            return {"status": "error", "message": self.t("defender_timeout")}
        except Exception as e:
            return {"status": "error", "message": str(e)}
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def _refresh_list(self):
        self.listbox.delete(0, tk.END)
        for item in self.messages:
            subject = item["subject"] if item["subject"] else self.t("no_subject")
            sender = item["sender"] if item["sender"] else self.t("unknown_sender")
            self.listbox.insert(tk.END, f"[{item['mail_type']}] {subject}  |  {sender}")

    def on_select_message(self, event=None):
        selection = self.listbox.curselection()
        if not selection:
            return
        self._display_message(selection[0])

    def _display_message(self, index):
        self.current_index = index
        item = self.messages[index]

        self.file_var.set(item["path"])
        self.type_var.set(item["mail_type"])
        self.subject_var.set(item["subject"])
        self.from_var.set(item["sender"])
        self.to_var.set(item["to"])
        self.date_var.set(item["date"])

        self.body_text.delete("1.0", tk.END)
        self.body_text.insert("1.0", item["body"])

        self._populate_security(item["attachments"], item["analysis"])
        self.status_var.set(self.t("showing_message", index=index + 1, total=len(self.messages)))

    def _populate_security(self, attachments, analysis):
        self.risk_var.set(f"{self.t('risk_score')}: {analysis['score']}/10+")
        self.level_var.set(f"{self.t('level')}: {analysis['level']}")
        self.defender_var.set(f"{self.t('defender_scan')}: {analysis['defender']['status']}")

        self._set_text(self.alerts_text, "\n".join(f"- {a}" for a in analysis["alerts"]))
        self._set_text(self.links_text, "\n".join(analysis["links"]) if analysis["links"] else self.t("no_links"))

        if attachments:
            lines = []
            for att in attachments:
                tag = self.t("suspicious_tag") if att["ext"] in SUSPICIOUS_EXTENSIONS else ""
                lines.append(f"{att['name']} | {att['size']}{tag}")
            self._set_text(self.attachments_text, "\n".join(lines))
        else:
            self._set_text(self.attachments_text, self.t("no_attachments"))

        self._set_text(self.flags_text, "\n".join(analysis["keywords"]) if analysis["keywords"] else self.t("no_keywords"))

    def _set_text(self, widget, text):
        widget.delete("1.0", tk.END)
        widget.insert("1.0", text)

    def clear_all(self):
        self.messages.clear()
        self.current_index = None
        self.listbox.delete(0, tk.END)

        for var in [self.file_var, self.type_var, self.subject_var, self.from_var, self.to_var, self.date_var]:
            var.set("")

        self.body_text.delete("1.0", tk.END)
        self._set_text(self.alerts_text, "")
        self._set_text(self.links_text, "")
        self._set_text(self.attachments_text, "")
        self._set_text(self.flags_text, "")
        self.risk_var.set(f"{self.t('risk_score')}: -")
        self.level_var.set(f"{self.t('level')}: -")
        self.defender_var.set(f"{self.t('defender_scan')}: -")
        self.top_info_label.config(text=self.t("load_hint"))
        self.status_var.set(self.t("cleared"))

    @staticmethod
    def _extract_urls(text):
        return list(dict.fromkeys(URL_REGEX.findall(text or "")))

    @staticmethod
    def _find_keywords(text):
        text_lower = (text or "").lower()
        return [kw for kw in SUSPICIOUS_KEYWORDS if kw.lower() in text_lower]

    @staticmethod
    def _extract_email(text):
        match = EMAIL_REGEX.search(text or "")
        return match.group(0).lower() if match else ""

    @staticmethod
    def _domain_from_email(email):
        return email.split("@", 1)[1].lower() if "@" in email else ""

    @staticmethod
    def _host_from_url(url):
        cleaned = url.lower().strip()
        cleaned = re.sub(r"^https?://", "", cleaned)
        cleaned = re.sub(r"^www\.", "", cleaned)
        return cleaned.split("/", 1)[0].split(":", 1)[0]

    @staticmethod
    def _parse_email_date(date_str):
        if not date_str:
            return ""
        try:
            return parsedate_to_datetime(date_str)
        except Exception:
            return date_str

    def _risk_level(self, score):
        if score >= 9:
            return self.t("risk_very_high")
        if score >= 6:
            return self.t("risk_high")
        if score >= 3:
            return self.t("risk_medium")
        return self.t("risk_low")

    @staticmethod
    def _safe_text(value):
        if value is None:
            return ""
        try:
            return str(value).strip()
        except Exception:
            return ""

    @staticmethod
    def _format_date(value):
        if value is None or value == "":
            return ""
        try:
            if isinstance(value, datetime):
                return value.strftime("%d/%m/%Y %H:%M:%S")
            return str(value)
        except Exception:
            return str(value)


def main():
    root = tk.Tk()
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    MailViewerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
