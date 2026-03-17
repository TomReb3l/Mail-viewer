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


class MsgViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Mail Viewer (.MSG / .EML) & Basic Security Check")
        self.root.geometry("1320x820")
        self.root.minsize(1024, 680)

        self.messages = []
        self.current_index = None

        self._build_ui()

    def _build_ui(self):
        self._build_topbar()
        self._build_main_area()
        self._build_statusbar()

    def _build_topbar(self):
        top = ttk.Frame(self.root, padding=10)
        top.pack(fill="x")

        ttk.Button(top, text="Άνοιγμα email αρχείων", command=self.open_files).pack(side="left")
        ttk.Button(top, text="Άνοιγμα φακέλου", command=self.open_folder).pack(side="left", padx=(8, 0))
        ttk.Button(top, text="Καθαρισμός", command=self.clear_all).pack(side="left", padx=(8, 0))

        self.top_info_label = ttk.Label(top, text="Φόρτωσε .msg ή .eml αρχεία")
        self.top_info_label.pack(side="left", padx=(14, 0))

    def _build_main_area(self):
        main = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        left_frame = ttk.Frame(main, padding=8)
        main.add(left_frame, weight=1)

        ttk.Label(left_frame, text="Αρχεία email").pack(anchor="w")

        self.listbox = tk.Listbox(left_frame, activestyle="dotbox", exportselection=False, font=("Segoe UI", 10))
        self.listbox.pack(fill="both", expand=True, pady=(6, 0))
        self.listbox.bind("<<ListboxSelect>>", self.on_select_message)

        right_frame = ttk.Frame(main, padding=8)
        main.add(right_frame, weight=3)

        form = ttk.Frame(right_frame)
        form.pack(fill="x")

        labels = ["Αρχείο:", "Τύπος:", "Θέμα:", "Από:", "Προς:", "Ημερομηνία:"]
        self.file_var = tk.StringVar()
        self.type_var = tk.StringVar()
        self.subject_var = tk.StringVar()
        self.from_var = tk.StringVar()
        self.to_var = tk.StringVar()
        self.date_var = tk.StringVar()
        vars_ = [self.file_var, self.type_var, self.subject_var, self.from_var, self.to_var, self.date_var]

        for i, (label_text, var) in enumerate(zip(labels, vars_)):
            ttk.Label(form, text=label_text, width=12).grid(row=i, column=0, sticky="w", pady=4)
            ttk.Entry(form, textvariable=var, state="readonly").grid(row=i, column=1, sticky="ew", pady=4)

        form.columnconfigure(1, weight=1)

        self.notebook = ttk.Notebook(right_frame)
        self.notebook.pack(fill="both", expand=True, pady=(12, 0))

        preview_tab = ttk.Frame(self.notebook, padding=8)
        security_tab = ttk.Frame(self.notebook, padding=8)

        self.notebook.add(preview_tab, text="Μήνυμα")
        self.notebook.add(security_tab, text="Έλεγχος Ασφάλειας")

        ttk.Label(preview_tab, text="Περιεχόμενο μηνύματος").pack(anchor="w", pady=(0, 4))
        self.body_text = ScrolledText(preview_tab, wrap="word", font=("Segoe UI", 11))
        self.body_text.pack(fill="both", expand=True)

        self._build_security_tab(security_tab)

    def _build_security_tab(self, parent):
        top = ttk.Frame(parent)
        top.pack(fill="x")

        self.risk_var = tk.StringVar(value="Risk Score: -")
        self.level_var = tk.StringVar(value="Επίπεδο: -")
        self.defender_var = tk.StringVar(value="Defender Scan: -")

        ttk.Label(top, textvariable=self.risk_var, font=("Segoe UI", 11, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 16))
        ttk.Label(top, textvariable=self.level_var, font=("Segoe UI", 11, "bold")).grid(row=0, column=1, sticky="w", padx=(0, 16))
        ttk.Label(top, textvariable=self.defender_var, font=("Segoe UI", 11, "bold")).grid(row=0, column=2, sticky="w")

        lists_frame = ttk.Frame(parent)
        lists_frame.pack(fill="both", expand=True, pady=(10, 0))

        left = ttk.Frame(lists_frame)
        left.pack(side="left", fill="both", expand=True, padx=(0, 5))

        right = ttk.Frame(lists_frame)
        right.pack(side="left", fill="both", expand=True, padx=(5, 0))

        ttk.Label(left, text="Alerts / Ευρήματα").pack(anchor="w")
        self.alerts_text = ScrolledText(left, height=12, wrap="word", font=("Consolas", 10))
        self.alerts_text.pack(fill="both", expand=True, pady=(4, 8))

        ttk.Label(left, text="Links").pack(anchor="w")
        self.links_text = ScrolledText(left, height=8, wrap="word", font=("Consolas", 10))
        self.links_text.pack(fill="both", expand=True, pady=(4, 0))

        ttk.Label(right, text="Συνημμένα").pack(anchor="w")
        self.attachments_text = ScrolledText(right, height=12, wrap="word", font=("Consolas", 10))
        self.attachments_text.pack(fill="both", expand=True, pady=(4, 8))

        ttk.Label(right, text="Λέξεις / Flags").pack(anchor="w")
        self.flags_text = ScrolledText(right, height=8, wrap="word", font=("Consolas", 10))
        self.flags_text.pack(fill="both", expand=True, pady=(4, 0))

    def _build_statusbar(self):
        self.status_var = tk.StringVar(value="Έτοιμο")
        ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w", padding=6).pack(fill="x", side="bottom")

    def open_files(self):
        if not self._precheck_libraries():
            return

        file_paths = filedialog.askopenfilenames(
            title="Επιλογή email αρχείων",
            filetypes=[("Email files", "*.msg *.eml"), ("MSG files", "*.msg"), ("EML files", "*.eml")]
        )
        if not file_paths:
            return

        loaded = 0
        for path in file_paths:
            if self._load_message(path):
                loaded += 1

        self._refresh_list()
        self.status_var.set(f"Φορτώθηκαν {loaded} αρχεία.")
        self.top_info_label.config(text=f"Σύνολο φορτωμένων μηνυμάτων: {len(self.messages)}")

        if self.messages and self.current_index is None:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(0)
            self.listbox.event_generate("<<ListboxSelect>>")

    def open_folder(self):
        if not self._precheck_libraries():
            return

        folder = filedialog.askdirectory(title="Επιλογή φακέλου με .msg / .eml αρχεία")
        if not folder:
            return

        mail_files = [
            os.path.join(folder, name)
            for name in os.listdir(folder)
            if name.lower().endswith((".msg", ".eml"))
        ]
        if not mail_files:
            messagebox.showinfo("Δεν βρέθηκαν αρχεία", "Δεν βρέθηκαν .msg ή .eml αρχεία στον φάκελο.")
            return

        loaded = 0
        for path in sorted(mail_files):
            if self._load_message(path):
                loaded += 1

        self._refresh_list()
        self.status_var.set(f"Φορτώθηκαν {loaded} αρχεία από τον φάκελο.")
        self.top_info_label.config(text=f"Σύνολο φορτωμένων μηνυμάτων: {len(self.messages)}")

        if self.messages and self.current_index is None:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(0)
            self.listbox.event_generate("<<ListboxSelect>>")

    def _precheck_libraries(self):
        return True

    def _load_message(self, path):
        for item in self.messages:
            if os.path.abspath(item["path"]) == os.path.abspath(path):
                return False

        try:
            ext = os.path.splitext(path)[1].lower()

            if ext == ".msg":
                if extract_msg is None:
                    raise RuntimeError("Λείπει η βιβλιοθήκη extract_msg. Τρέξε: pip install extract_msg")
                data = self._read_msg(path)
            elif ext == ".eml":
                data = self._read_eml(path)
            else:
                raise RuntimeError(f"Μη υποστηριζόμενος τύπος αρχείου: {ext}")

            analysis = self._analyze_message(
                data["subject"], data["sender"], data["to"], data["body"], data["attachments"]
            )

            data["analysis"] = analysis
            self.messages.append(data)
            return True

        except Exception as e:
            messagebox.showwarning("Σφάλμα ανάγνωσης", f"Το αρχείο δεν διαβάστηκε σωστά:\n{path}\n\n{e}")
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

        subject = self._safe_text(eml.get("subject", ""))
        sender = self._safe_text(eml.get("from", ""))
        to_ = self._safe_text(eml.get("to", ""))
        date_raw = eml.get("date", "")
        date_text = self._format_date(self._parse_email_date(date_raw))

        body = self._extract_eml_body(eml)
        attachments = self._extract_eml_attachments_info(eml)

        return {
            "path": path,
            "mail_type": "EML",
            "subject": subject,
            "sender": sender,
            "to": to_,
            "date": date_text,
            "body": body,
            "attachments": attachments
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

        return "(Δεν βρέθηκε περιεχόμενο μηνύματος)"

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

        return "(Δεν βρέθηκε περιεχόμενο μηνύματος)"

    def _extract_msg_attachments_info(self, msg):
        results = []
        for att in getattr(msg, "attachments", []) or []:
            name = self._safe_text(getattr(att, "longFilename", "")) or self._safe_text(getattr(att, "filename", "")) or "(χωρίς όνομα)"
            ext = os.path.splitext(name)[1].lower()
            data = getattr(att, "data", None)
            size_text = f"{len(data)} bytes" if isinstance(data, (bytes, bytearray)) else "-"
            results.append({
                "name": name,
                "ext": ext,
                "size": size_text,
                "data": data if isinstance(data, (bytes, bytearray)) else None
            })
        return results

    def _extract_eml_attachments_info(self, eml):
        results = []
        for part in eml.iter_attachments():
            name = self._safe_text(part.get_filename()) or "(χωρίς όνομα)"
            ext = os.path.splitext(name)[1].lower()

            try:
                raw = part.get_payload(decode=True)
            except Exception:
                raw = None

            size_text = f"{len(raw)} bytes" if isinstance(raw, (bytes, bytearray)) else "-"
            results.append({
                "name": name,
                "ext": ext,
                "size": size_text,
                "data": raw if isinstance(raw, (bytes, bytearray)) else None
            })
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
            alerts.append("Λείπει αποστολέας.")
            score += 2

        if sender_domain:
            for brand in TRUST_BRANDS:
                if brand in sender_text_lower and brand not in sender_domain:
                    alerts.append(f"Πιθανό display-name spoofing: εμφανίζεται '{brand}' αλλά domain '{sender_domain}'.")
                    score += 3
                    break

            if sender_domain in SHORTENER_DOMAINS:
                alerts.append(f"Ο αποστολέας χρησιμοποιεί shortened/ύποπτο domain: {sender_domain}")
                score += 2

        for url in links:
            host = self._host_from_url(url)
            if not host:
                continue
            if IP_HOST_REGEX.match(host):
                alerts.append(f"Link με IP αντί για domain: {url}")
                score += 3
            if host.lower() in SHORTENER_DOMAINS:
                alerts.append(f"Shortened link: {url}")
                score += 2
            for brand in TRUST_BRANDS:
                if brand in body.lower() and brand not in host.lower() and any(b in url.lower() for b in ["login", "verify", "secure", "account"]):
                    alerts.append(f"Πιθανό brand mismatch σε link: {url}")
                    score += 2
                    break

        suspicious_attachments = []
        for att in attachments:
            if att["ext"] in SUSPICIOUS_EXTENSIONS:
                suspicious_attachments.append(att["name"])
                alerts.append(f"Ύποπτο συνημμένο: {att['name']}")
                score += 3

        if len(links) >= 3:
            alerts.append(f"Πολλά links στο μήνυμα: {len(links)}")
            score += 1

        if found_keywords:
            alerts.append("Βρέθηκαν phishing/πίεσης keywords.")
            score += min(3, len(found_keywords))

        if subject and subject.isupper() and len(subject) > 8:
            alerts.append("Το θέμα είναι όλο με κεφαλαία.")
            score += 1

        defender_result = self._scan_attachments_with_defender(attachments)
        if defender_result["status"] == "malicious":
            alerts.append("Ο Windows Defender εντόπισε απειλή σε συνημμένο.")
            score += 6
        elif defender_result["status"] == "error":
            alerts.append(f"Defender scan δεν ολοκληρώθηκε: {defender_result['message']}")
            score += 1

        return {
            "score": score,
            "level": self._risk_level(score),
            "alerts": alerts or ["Δεν βρέθηκαν εμφανή ύποπτα στοιχεία με τον βασικό heuristic έλεγχο."],
            "links": links,
            "keywords": found_keywords,
            "suspicious_attachments": suspicious_attachments,
            "defender": defender_result
        }

    def _scan_attachments_with_defender(self, attachments):
        if not attachments:
            return {"status": "no_attachments", "message": "Δεν υπάρχουν συνημμένα."}

        mpcmdrun = r"C:\Program Files\Windows Defender\MpCmdRun.exe"
        if not os.path.exists(mpcmdrun):
            return {"status": "not_available", "message": "Δεν βρέθηκε το MpCmdRun.exe του Windows Defender."}

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
                return {"status": "no_extractable", "message": "Δεν ήταν δυνατή η εξαγωγή συνημμένων για scan."}

            cmd = [mpcmdrun, "-Scan", "-ScanType", "3", "-File", temp_dir]
            proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=180)

            combined = "\n".join([proc.stdout or "", proc.stderr or ""]).strip().lower()

            if any(x in combined for x in ["threat", "infected", "malware", "virus", "found"]):
                return {"status": "malicious", "message": combined[:1500] or "Βρέθηκε πιθανή απειλή."}

            if proc.returncode == 0:
                return {"status": "clean", "message": "Δεν εντοπίστηκε απειλή από τον Defender."}

            return {"status": "unknown", "message": combined[:1500] or f"Άγνωστο αποτέλεσμα Defender. Return code: {proc.returncode}"}
        except subprocess.TimeoutExpired:
            return {"status": "error", "message": "Ο Defender scan έληξε λόγω timeout."}
        except Exception as e:
            return {"status": "error", "message": str(e)}
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def _refresh_list(self):
        self.listbox.delete(0, tk.END)
        for item in self.messages:
            subject = item["subject"] if item["subject"] else "(Χωρίς θέμα)"
            sender = item["sender"] if item["sender"] else "(Άγνωστος αποστολέας)"
            self.listbox.insert(tk.END, f"[{item['mail_type']}] {subject}  |  {sender}")

    def on_select_message(self, event=None):
        selection = self.listbox.curselection()
        if not selection:
            return

        index = selection[0]
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
        self.status_var.set(f"Εμφανίζεται μήνυμα {index + 1} από {len(self.messages)}")

    def _populate_security(self, attachments, analysis):
        self.risk_var.set(f"Risk Score: {analysis['score']}/10+")
        self.level_var.set(f"Επίπεδο: {analysis['level']}")
        self.defender_var.set(f"Defender Scan: {analysis['defender']['status']}")

        self._set_text(self.alerts_text, "\n".join(f"- {a}" for a in analysis["alerts"]))
        self._set_text(self.links_text, "\n".join(analysis["links"]) if analysis["links"] else "(Δεν βρέθηκαν links)")

        if attachments:
            lines = []
            for att in attachments:
                tag = " [ΥΠΟΠΤΟ]" if att["ext"] in SUSPICIOUS_EXTENSIONS else ""
                lines.append(f"{att['name']} | {att['size']}{tag}")
            self._set_text(self.attachments_text, "\n".join(lines))
        else:
            self._set_text(self.attachments_text, "(Δεν υπάρχουν συνημμένα)")

        self._set_text(self.flags_text, "\n".join(analysis["keywords"]) if analysis["keywords"] else "(Δεν βρέθηκαν flagged keywords)")

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
        self.risk_var.set("Risk Score: -")
        self.level_var.set("Επίπεδο: -")
        self.defender_var.set("Defender Scan: -")

        self.top_info_label.config(text="Φόρτωσε .msg ή .eml αρχεία")
        self.status_var.set("Καθαρίστηκαν όλα τα δεδομένα.")

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

    @staticmethod
    def _risk_level(score):
        if score >= 9:
            return "Πολύ Ύποπτο"
        if score >= 6:
            return "Υψηλό Ρίσκο"
        if score >= 3:
            return "Μέτριο Ρίσκο"
        return "Χαμηλό Ρίσκο"

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
    MsgViewerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
