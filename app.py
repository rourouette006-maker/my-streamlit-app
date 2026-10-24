# =========================================================
# SMQ Dashboard — INPACKT (Médenine · Koutine)
# Nécessite Streamlit >= 1.31
# =========================================================
import io
import os
import re
import difflib
import unicodedata
import html
import base64
from datetime import datetime, date
from functools import lru_cache
from urllib.parse import quote

import pandas as pd
import streamlit as st

st.set_page_config(page_title="SMQ Dashboard · INPACKT", layout="wide")

cols = getattr(st, "columns", None) or st.beta_columns
rerun = getattr(st, "rerun", None) or st.experimental_rerun
expander = getattr(st, "expander", None) or st.beta_expander

BASE = os.path.dirname(os.path.abspath(__file__))
FILE = os.path.join(BASE, "database_smq.xlsx")

# Comptes : identifiant -> (mot de passe, rôle, nom affiché)
USERS = {
    "admin": ("qualite2026", "Administrateur", "Admin"),
    "qualiticien": ("qualite123", "Qualiticien", "Qualiticien"),
    "lecteur": ("lecture", "Lecteur", "Invité"),
}

# ---------------------------------------------------------
# SESSION
# ---------------------------------------------------------
for _k, _v in (("connected", False), ("lang", "Français"), ("dark_mode", False),
               ("show_settings", False), ("flash", ""), ("bg", "grad"),
               ("role", "Lecteur"), ("user", ""), ("display", "")):
    if _k not in st.session_state:
        st.session_state[_k] = _v

# ---------------------------------------------------------
# TRADUCTIONS
# ---------------------------------------------------------
I18N = {
    "Français": {
        "title": "SMQ Dashboard", "navigation": "Navigation",
        "home": "Vue d'ensemble", "nc": "Non-Conformités", "audits": "Audits",
        "reclamations": "Réclamations", "actions": "Actions Correctives",
        "fournisseurs": "Fournisseurs", "logistique": "Flux Logistiques",
        "views": "Vues par statut", "evaluation": "Évaluation",
        "stats": "Statistiques", "history": "Historique", "logout": "Se déconnecter",
        "hello_day": "Bonjour", "hello_night": "Bonsoir",
        "desc": "Voici un aperçu général de la qualité et des flux de votre entreprise.",
        "alerts": "Alertes", "summary": "Récapitulatif",
        "status_chart": "Statuts des dossiers", "recent": "Activités récentes",
        "settings": "Paramètres", "dark": "Mode sombre", "lang": "Langue",
        "mode_dark": "Sombre", "mode_light": "Clair",
    },
    "العربية": {
        "title": "SMQ Dashboard", "navigation": "التنقل",
        "home": "لوحة القيادة", "nc": "عدم المطابقات", "audits": "التدقيق",
        "reclamations": "الشكاوى", "actions": "الإجراءات التصحيحية",
        "fournisseurs": "الموردون", "logistique": "التدفقات اللوجستية",
        "views": "العروض حسب الحالة", "evaluation": "التقييم",
        "stats": "الإحصائيات", "history": "السجل", "logout": "تسجيل الخروج",
        "hello_day": "صباح الخير", "hello_night": "مساء الخير",
        "desc": "إليكم نظرة عامة على الجودة والتدفقات في مؤسستكم.",
        "alerts": "تنبيهات", "summary": "ملخص",
        "status_chart": "حالات الملفات", "recent": "آخر الأنشطة",
        "settings": "الإعدادات", "dark": "الوضع الليلي", "lang": "اللغة",
        "mode_dark": "ليلي", "mode_light": "نهاري",
    },
    "English": {
        "title": "SMQ Dashboard", "navigation": "Navigation",
        "home": "Overview", "nc": "Non-Conformities", "audits": "Audits",
        "reclamations": "Complaints", "actions": "Corrective Actions",
        "fournisseurs": "Suppliers", "logistique": "Logistics Flows",
        "views": "Status views", "evaluation": "Evaluation",
        "stats": "Statistics", "history": "History", "logout": "Log out",
        "hello_day": "Good morning", "hello_night": "Good evening",
        "desc": "Here is an overview of your company's quality and flows.",
        "alerts": "Alerts", "summary": "Summary",
        "status_chart": "File statuses", "recent": "Recent activities",
        "settings": "Settings", "dark": "Dark mode", "lang": "Language",
        "mode_dark": "Dark", "mode_light": "Light",
    },
}
t = I18N.get(st.session_state.lang, I18N["Français"])

# Dictionnaire des mots/valeurs affichés : clé française -> (English, العربية)
# Les données restent stockées en français dans Excel ; seule l'AFFICHAGE est traduit.
TR = {
    # --- colonnes ---
    "Réf": ("Ref", "المرجع"), "Date": ("Date", "التاريخ"),
    "Description": ("Description", "الوصف"), "Service": ("Department", "المصلحة"),
    "Responsable": ("Owner", "المسؤول"), "Gravité": ("Severity", "الخطورة"),
    "Cause": ("Cause", "السبب"), "Action liée": ("Linked action", "الإجراء المرتبط"),
    "Statut": ("Status", "الحالة"), "Date clôture": ("Closing date", "تاريخ الإغلاق"),
    "Type": ("Type", "النوع"), "Auditeur": ("Auditor", "المدقق"),
    "Date prévue": ("Planned date", "التاريخ المخطط"), "Réalisation": ("Completion", "الإنجاز"),
    "Résultat": ("Result", "النتيجة"), "Client": ("Customer", "الحريف"),
    "Délai": ("Target (days)", "الأجل (أيام)"), "Échéance": ("Due date", "الموعد النهائي"),
    "Action": ("Action", "الإجراء"), "Nom": ("Name", "الاسم"),
    "Catégorie": ("Category", "الفئة"), "Évaluation": ("Rating", "التقييم"),
    "Type de Flux": ("Flow type", "نوع الحركة"), "Article": ("Item", "المنتج"),
    "Quantité": ("Quantity", "الكمية"), "Tiers": ("Third party", "الطرف الثالث"),
    "Contrôle": ("Quality check", "المراقبة"), "Date/heure": ("Date/time", "التاريخ/الوقت"),
    "Utilisateur": ("User", "المستخدم"), "Module": ("Module", "الوحدة"),
    "Détail": ("Detail", "التفاصيل"), "Intitulé": ("Title", "العنوان"),
    "Nombre": ("Count", "العدد"), "Critère": ("Criterion", "المعيار"),
    "Valeur": ("Value", "القيمة"), "Objectif": ("Target", "الهدف"),
    # --- valeurs ---
    "Ouverte": ("Open", "مفتوحة"), "Nouvelle": ("New", "جديدة"), "À faire": ("To do", "للإنجاز"),
    "En cours": ("In progress", "قيد التنفيذ"), "Clôturée": ("Closed", "مغلقة"),
    "Traitée": ("Handled", "معالجة"), "Terminée": ("Done", "منجزة"),
    "Planifié": ("Planned", "مخطط"), "Réalisé": ("Completed", "منجز"),
    "Conforme": ("Compliant", "مطابق"), "Non conforme": ("Non-compliant", "غير مطابق"),
    "À améliorer": ("To improve", "للتحسين"),
    "Mineure": ("Minor", "طفيفة"), "Majeure": ("Major", "كبيرة"), "Critique": ("Critical", "حرجة"),
    "Matière": ("Material", "المادة"), "Méthode": ("Method", "الطريقة"),
    "Main-d'œuvre": ("Manpower", "اليد العاملة"), "Machine": ("Machine", "الآلة"),
    "Milieu": ("Environment", "المحيط"),
    "Actif": ("Active", "نشط"), "Sous observation": ("Under observation", "تحت المراقبة"),
    "Suspendu": ("Suspended", "موقوف"),
    "A - Excellent": ("A - Excellent", "A - ممتاز"), "B - Conforme": ("B - Compliant", "B - مطابق"),
    "C - À surveiller": ("C - To monitor", "C - للمراقبة"), "D - Critique": ("D - Critical", "D - حرج"),
    "Entrée (Réception Marchandise)": ("Inbound (Goods receipt)", "دخول (استلام بضاعة)"),
    "Sortie (Expédition / Vente)": ("Outbound (Shipping / Sale)", "خروج (شحن / بيع)"),
    "Retour Interne": ("Internal return", "إرجاع داخلي"),
    "Validé / Conforme": ("Approved / Compliant", "مقبول / مطابق"),
    "Réservé / Non-conforme": ("On hold / Non-compliant", "محجوز / غير مطابق"),
    "Ajout": ("Added", "إضافة"), "Modification": ("Edited", "تعديل"), "Suppression": ("Deleted", "حذف"),
    "Ouvert": ("Open", "مفتوح"), "Clôturé": ("Closed", "مغلق"),
    # --- modules ---
    "Non-Conformités": ("Non-Conformities", "عدم المطابقات"), "Audits": ("Audits", "التدقيق"),
    "Réclamations": ("Complaints", "الشكاوى"), "Actions Correctives": ("Corrective Actions", "الإجراءات التصحيحية"),
    "Fournisseurs": ("Suppliers", "الموردون"), "Flux Logistiques": ("Logistics Flows", "التدفقات اللوجستية"),
    "Non-Conformité": ("Non-Conformity", "عدم مطابقة"), "Audit": ("Audit", "تدقيق"),
    "Réclamation": ("Complaint", "شكوى"), "Fournisseur": ("Supplier", "مورد"),
    "Logistique": ("Logistics", "لوجستيك"),
    # --- classification ---
    "Tous": ("All", "الكل"), "Pas encore commencé": ("Not started yet", "لم يبدأ بعد"),
    "Terminé / Conforme": ("Done / OK", "منجز / مطابق"),
    # --- libellés de formulaires ---
    "Service concerné": ("Department concerned", "المصلحة المعنية"),
    "Cause (5M)": ("Cause (5M)", "السبب (5M)"),
    "Action corrective liée": ("Linked corrective action", "الإجراء التصحيحي المرتبط"),
    "Type d'audit": ("Audit type", "نوع التدقيق"), "Service audité": ("Audited department", "المصلحة المدققة"),
    "Date prévue (planning)": ("Planned date (schedule)", "التاريخ المخطط (البرنامج)"),
    "Délai cible (jours)": ("Target time (days)", "الأجل المستهدف (أيام)"),
    "Date d'évaluation": ("Evaluation date", "تاريخ التقييم"),
    "Nom du Fournisseur": ("Supplier name", "اسم المورد"),
    "Matière / Service fourni": ("Material / Service supplied", "المادة / الخدمة الموردة"),
    "Évaluation Qualité": ("Quality rating", "تقييم الجودة"),
    "Statut Partenariat": ("Partnership status", "حالة الشراكة"),
    "Date du mouvement": ("Movement date", "تاريخ الحركة"),
    "Référence / Nom de l'Article": ("Item reference / name", "مرجع / اسم المنتج"),
    "Fournisseur (si entrée) ou Client (si sortie)": ("Supplier (inbound) or Customer (outbound)", "المورد (دخول) أو الحريف (خروج)"),
    "Contrôle Qualité Réception/Expédition": ("Receiving / shipping quality check", "مراقبة الجودة عند الاستلام/الشحن"),
    # --- titres, boutons, messages ---
    "＋ Ajouter une Non-Conformité": ("＋ Add a Non-Conformity", "＋ إضافة عدم مطابقة"),
    "Liste des Non-Conformités": ("Non-Conformities list", "قائمة عدم المطابقات"),
    "Aucune donnée disponible.": ("No data available.", "لا توجد بيانات."),
    "Non-Conformité enregistrée !": ("Non-Conformity saved!", "تم تسجيل عدم المطابقة!"),
    "Enregistrer": ("Save", "حفظ"),
    "＋ Ajouter un Audit": ("＋ Add an Audit", "＋ إضافة تدقيق"),
    "Liste des Audits": ("Audits list", "قائمة التدقيق"),
    "Aucun audit enregistré.": ("No audit recorded.", "لا يوجد تدقيق مسجل."),
    "Audit enregistré !": ("Audit saved!", "تم تسجيل التدقيق!"),
    "＋ Ajouter une Réclamation": ("＋ Add a Complaint", "＋ إضافة شكوى"),
    "Liste des Réclamations": ("Complaints list", "قائمة الشكاوى"),
    "Aucune réclamation enregistrée.": ("No complaint recorded.", "لا توجد شكاوى مسجلة."),
    "Réclamation ajoutée !": ("Complaint added!", "تمت إضافة الشكوى!"),
    "＋ Ajouter une Action": ("＋ Add an Action", "＋ إضافة إجراء"),
    "Plan d'Actions": ("Action plan", "خطة الإجراءات"),
    "Aucune action enregistrée.": ("No action recorded.", "لا توجد إجراءات مسجلة."),
    "Action enregistrée !": ("Action saved!", "تم تسجيل الإجراء!"),
    "＋ Référencer un Fournisseur / Évaluation": ("＋ Register a Supplier / Rating", "＋ تسجيل مورد / تقييم"),
    "Annuaire et Évaluation des Fournisseurs": ("Supplier directory & rating", "دليل وتقييم الموردين"),
    "Aucun fournisseur enregistré.": ("No supplier recorded.", "لا يوجد موردون مسجلون."),
    "Fournisseur enregistré !": ("Supplier saved!", "تم تسجيل المورد!"),
    "Enregistrer le Fournisseur": ("Save supplier", "حفظ المورد"),
    "＋ Enregistrer un Mouvement de Stock": ("＋ Record a Stock Movement", "＋ تسجيل حركة مخزون"),
    "Historique des Flux Logistiques": ("Logistics flows history", "سجل التدفقات اللوجستية"),
    "Aucun flux enregistré.": ("No flow recorded.", "لا توجد حركات مسجلة."),
    "Mouvement enregistré !": ("Movement saved!", "تم تسجيل الحركة!"),
    "Valider le Mouvement": ("Confirm movement", "تأكيد الحركة"),
    # --- critères d'évaluation ---
    "Taux de clôture": ("Closing rate", "نسبة الإغلاق"),
    "Délai moyen de clôture": ("Average closing time", "متوسط مدة الإغلاق"),
    "NC critiques non clôturées": ("Critical NCs not closed", "عدم مطابقات حرجة غير مغلقة"),
    "NC avec action corrective liée": ("NCs with a linked corrective action", "عدم مطابقات لها إجراء تصحيحي"),
    "Taux de réalisation": ("Completion rate", "نسبة الإنجاز"),
    "Taux de conformité": ("Compliance rate", "نسبة المطابقة"),
    "Audits planifiés en retard": ("Planned audits overdue", "تدقيقات مخططة متأخرة"),
    "Taux de traitement": ("Handling rate", "نسبة المعالجة"),
    "Réclamations traitées dans les délais": ("Complaints handled on time", "الشكاوى المعالجة في الوقت"),
    "Réclamations non prises en charge": ("Complaints not yet taken over", "شكاوى لم يتم التكفل بها"),
    "Actions en retard": ("Overdue actions", "إجراءات متأخرة"),
    "Fournisseurs conformes (A/B)": ("Compliant suppliers (A/B)", "الموردون المطابقون (A/B)"),
    "Fournisseurs à risque": ("Suppliers at risk", "موردون في خطر"),
    "Taux de lots conformes": ("Compliant lots rate", "نسبة الدفعات المطابقة"),
    "Lots réservés / non conformes": ("Held / non-conforming lots", "دفعات غير مطابقة"),
    "Atteint": ("Met", "محقق"), "Proche de l'objectif": ("Close to target", "قريب من الهدف"),
    "Non atteint": ("Not met", "غير محقق"), "Pas de données": ("No data", "لا بيانات"),
    # --- arrière-plans ---
    "Dégradé doux": ("Soft gradient", "تدرج ناعم"), "Points": ("Dots", "نقاط"),
    "Grille qualité": ("Quality grid", "شبكة الجودة"), "Vagues": ("Waves", "أمواج"),
    "Filigrane logo": ("Logo watermark", "علامة مائية للشعار"), "Aucun": ("None", "بدون"),
}


def tr(s):
    s = "" if s is None else str(s)
    lang = st.session_state.lang
    if lang == "Français":
        return s
    e = TR.get(s)
    if not e:
        return s
    return e[0] if lang == "English" else e[1]


def L(fr, en, ar):
    return {"Français": fr, "English": en, "العربية": ar}[st.session_state.lang]


MODULE_FR = {"NonConformites": "Non-Conformités", "Audits": "Audits", "Reclamations": "Réclamations",
             "Actions": "Actions Correctives", "Fournisseurs": "Fournisseurs", "FluxLogistique": "Flux Logistiques"}


def mod_name(sheet):
    return tr(MODULE_FR.get(sheet, sheet))


# ---------------------------------------------------------
# LOGO INPACKT (logo.png à côté de app.py)
# ---------------------------------------------------------
@lru_cache(maxsize=1)
def logo_data():
    for name, mime in (("logo.png", "image/png"), ("logo.jpg", "image/jpeg"),
                       ("logo.jpeg", "image/jpeg"), ("logo.webp", "image/webp")):
        p = os.path.join(BASE, name)
        if os.path.exists(p):
            with open(p, "rb") as f:
                return f"data:{mime};base64," + base64.b64encode(f.read()).decode()
    return ""


# ---------------------------------------------------------
# ICÔNES SVG (viewBox 24x24)
# ---------------------------------------------------------
ICONS = {
    "home": '<path d="M3 10.5 12 3l9 7.5V20a1 1 0 0 1-1 1h-5v-6H9v6H4a1 1 0 0 1-1-1z"/>',
    "nc": '<path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/><path d="M12 9v4M12 17h.01"/>',
    "audits": '<path d="M9 5H5a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2h-4"/><rect x="9" y="3" width="6" height="4" rx="1"/><path d="m9 14 2 2 4-4"/>',
    "reclamations": '<path d="M21 11.5a8.4 8.4 0 0 1-.9 3.8 8.5 8.5 0 0 1-7.6 4.7 8.4 8.4 0 0 1-3.8-.9L3 21l1.9-5.7a8.4 8.4 0 0 1-.9-3.8A8.5 8.5 0 0 1 8.7 3.9a8.4 8.4 0 0 1 3.8-.9h.5a8.5 8.5 0 0 1 8 8z"/>',
    "actions": '<circle cx="12" cy="12" r="3"/><path d="M12 2v3M12 19v3M2 12h3M19 12h3M4.9 4.9 7 7M17 17l2.1 2.1M4.9 19.1 7 17M17 7l2.1-2.1"/>',
    "fournisseurs": '<path d="M1 3h15v13H1zM16 8h4l3 3v5h-7z"/><circle cx="5.5" cy="18.5" r="2"/><circle cx="18.5" cy="18.5" r="2"/>',
    "logistique": '<path d="M21 8 12 3 3 8v8l9 5 9-5z"/><path d="m3 8 9 5 9-5M12 13v8"/>',
    "views": '<rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/>',
    "evaluation": '<path d="m12 2 3.1 6.3 6.9 1-5 4.9 1.2 6.8L12 17.8 5.8 21l1.2-6.8-5-4.9 6.9-1z"/>',
    "stats": '<path d="M4 20V10M10 20V4M16 20v-8M22 20H2"/>',
    "history": '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
    "logout": '<path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4M16 17l5-5-5-5M21 12H9"/>',
    "shield": '<path d="M12 2 4 5v6c0 5 3.4 9.4 8 11 4.6-1.6 8-6 8-11V5z"/><path d="m9 12 2 2 4-4"/>',
}


def svg(name, color, size=20, width=1.8, fill="none"):
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" '
        f'viewBox="0 0 24 24" fill="{fill}" stroke="{color}" stroke-width="{width}" '
        f'stroke-linecap="round" stroke-linejoin="round">{ICONS[name]}</svg>'
    )


def svg_url(name, color):
    return "data:image/svg+xml;utf8," + quote(svg(name, color, 24))


NAV = ["home", "nc", "audits", "reclamations", "actions",
       "fournisseurs", "logistique", "views", "evaluation", "stats", "history"]

# ---------------------------------------------------------
# COULEURS DES STATUTS  (rouge = pas encore / problème, jaune = en cours, vert = terminé / OK)
# Pour changer une couleur : modifiez simplement VALUE_COLOR ci-dessous.
# ---------------------------------------------------------
COLORS = {  # clé -> (fond, texte)
    "red": ("#DC2626", "#FFFFFF"), "yellow": ("#FACC15", "#422006"),
    "green": ("#16A34A", "#FFFFFF"), "orange": ("#F97316", "#FFFFFF"),
    "blue": ("#3B82F6", "#FFFFFF"), "gray": ("#94A3B8", "#FFFFFF"),
}
VALUE_COLOR = {
    # rouge : pas encore commencé / problème
    "Ouverte": "red", "Nouvelle": "red", "À faire": "red", "Planifié": "red",
    "Suspendu": "red", "Non conforme": "red", "Réservé / Non-conforme": "red",
    "D - Critique": "red", "Critique": "red",
    # jaune : en cours
    "En cours": "yellow", "Sous observation": "yellow",
    # vert : terminé / conforme
    "Clôturée": "green", "Traitée": "green", "Terminée": "green", "Réalisé": "green",
    "Conforme": "green", "Actif": "green", "Validé / Conforme": "green",
    "A - Excellent": "green", "B - Conforme": "green",
    # autres
    "À améliorer": "orange", "C - À surveiller": "orange", "Majeure": "orange", "Mineure": "blue",
}
COLOR_COLS = {"Statut", "Réalisation", "Résultat", "Gravité", "Évaluation", "Contrôle"}
DATE_COLS = {"Date", "Date clôture", "Date prévue", "Échéance"}
# Colonne qui décide de la "classification" de chaque module
STATUS_COLS = {"NonConformites": "Statut", "Audits": "Réalisation", "Reclamations": "Statut",
               "Actions": "Statut", "Fournisseurs": "Statut", "FluxLogistique": "Contrôle"}
TITLE_COL = {"NonConformites": "Description", "Audits": "Type", "Reclamations": "Description",
             "Actions": "Action", "Fournisseurs": "Nom", "FluxLogistique": "Article"}
GROUPS = [("red", "🔴", "Pas encore commencé"), ("yellow", "🟡", "En cours"), ("green", "🟢", "Terminé / Conforme")]
GROUP_META = {k: (e, lab) for k, e, lab in GROUPS}
ALL = "__all__"


def radio_h(label, options, **kw):
    """st.radio horizontal ; compatible avec les anciennes versions de Streamlit."""
    try:
        return st.radio(label, options, horizontal=True, **kw)
    except TypeError:
        return st.radio(label, options, **kw)  # le CSS ci-dessous le rend horizontal


def chip(text, key):
    bg, fg = COLORS[key]
    return f'<span class="chip" style="background:{bg};color:{fg};">{html.escape(str(text))}</span>'


def cell_html(col, v):
    try:
        if pd.isna(v):
            return '<span style="opacity:.5">—</span>'
    except (TypeError, ValueError):
        pass
    if col in DATE_COLS:
        return html.escape(str(v)[:10])
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v)
    if col == "Module":
        return html.escape(mod_name(s))
    key = VALUE_COLOR.get(s) if col in COLOR_COLS else None
    if key:
        return chip(tr(s), key)
    return html.escape(tr(s))


def html_table(df, status_col=None):
    """Tableau HTML traduit + coloré (remplace st.table)."""
    if df is None or df.empty:
        return ""
    head = "".join(f"<th>{html.escape(tr(c))}</th>" for c in df.columns)
    body = ""
    for rec in df.to_dict("records"):
        rc = "transparent"
        if status_col and status_col in rec:
            k = VALUE_COLOR.get(str(rec[status_col]))
            if k in ("red", "yellow", "green"):
                rc = COLORS[k][0]
        body += f'<tr style="--rc:{rc}">' + "".join(f"<td>{cell_html(c, rec[c])}</td>" for c in df.columns) + "</tr>"
    return f'<div class="dt-wrap"><table class="dt"><tr>{head}</tr>{body}</table></div>'


# ---------------------------------------------------------
# DONNÉES
# ---------------------------------------------------------
REFS = {"NonConformites": "NC", "Audits": "AUD", "Reclamations": "REC",
        "Actions": "ACT", "Fournisseurs": "FOU", "FluxLogistique": "FLX"}


def next_number(df):
    nums = [0]
    if not df.empty and "Réf" in df.columns:
        for r in df["Réf"].dropna().astype(str):
            m = re.findall(r"(\d+)$", r)
            if m:
                nums.append(int(m[0]))
    return max(nums) + 1


def add_refs(df, prefix):
    """Attribue une référence (NC-001...) aux lignes qui n'en ont pas encore."""
    if df.empty:
        return df
    df = df.copy()
    if "Réf" not in df.columns:
        df.insert(0, "Réf", None)
    df["Réf"] = df["Réf"].astype(object)
    n = next_number(df) - 1
    for i in df.index:
        v = df.at[i, "Réf"]
        if pd.isna(v) or str(v).strip() == "":
            n += 1
            df.at[i, "Réf"] = f"{prefix}-{n:03d}"
    return df


def load_sheet(sheet_name):
    if os.path.exists(FILE):
        try:
            df = pd.read_excel(FILE, sheet_name=sheet_name)
            return add_refs(df, REFS[sheet_name]) if sheet_name in REFS else df
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()


def log_event(module, ref, action, detail=""):
    h = load_sheet("Historique")
    row = pd.DataFrame([{"Date/heure": datetime.now().strftime("%Y-%m-%d %H:%M"),
                         "Utilisateur": st.session_state["user"], "Module": module,
                         "Réf": ref, "Action": action, "Détail": detail}])
    save_sheet(row if h.empty else pd.concat([h, row], ignore_index=True), "Historique")


def save_sheet(df, sheet_name):
    dfs = {}
    if os.path.exists(FILE):
        try:
            xls = pd.ExcelFile(FILE)
            for sheet in xls.sheet_names:
                if sheet != sheet_name:
                    dfs[sheet] = pd.read_excel(FILE, sheet_name=sheet)
        except Exception:
            pass
    dfs[sheet_name] = df
    with pd.ExcelWriter(FILE, engine="openpyxl") as writer:
        for s_name, s_df in dfs.items():
            s_df.to_excel(writer, sheet_name=s_name, index=False)


def get_resolution_rate(df, closed_values):
    if df.empty or "Statut" not in df.columns:
        return 0.0
    return round(df["Statut"].isin(closed_values).sum() / len(df) * 100, 1)


ACTIVITY_SOURCES = [
    ("NonConformites", "Non-Conformité", "Description", "#EF4444"),
    ("Audits", "Audit", "Type", "#10B981"),
    ("Reclamations", "Réclamation", "Client", "#F59E0B"),
    ("Actions", "Action", "Action", "#8B5CF6"),
    ("Fournisseurs", "Fournisseur", "Nom", "#3B82F6"),
    ("FluxLogistique", "Logistique", "Article", "#14B8A6"),
]


def get_recent_activities(limit=5):
    rows = []
    for sheet, cat, col, color in ACTIVITY_SOURCES:
        df = load_sheet(sheet)
        if not df.empty and "Date" in df.columns:
            for _, r in df.iterrows():
                rows.append((str(r["Date"])[:10], cat, "#" + str(r.get("Réf", "")) + " " + str(r.get(col, "-")), color))
    rows.sort(key=lambda x: x[0], reverse=True)
    return rows[:limit]


STATUS_GROUPS = [
    ("Ouvert", "#EF4444", {"Ouverte", "Nouvelle", "À faire"}),
    ("En cours", "#FBBF24", {"En cours"}),
    ("Clôturé", "#3B82F6", {"Clôturée", "Traitée"}),
    ("Terminée", "#10B981", {"Terminée"}),
]


def count_statuses(dfs):
    counts = []
    values = []
    for df in dfs:
        if not df.empty and "Statut" in df.columns:
            values += df["Statut"].dropna().tolist()
    for label, color, members in STATUS_GROUPS:
        counts.append((label, color, sum(1 for v in values if v in members)))
    return counts


def virtual_view():
    """Table VIRTUELLE : toutes les lignes de tous les modules dans un même tableau."""
    rows = []
    for sh in REFS:
        df = load_sheet(sh)
        if df.empty:
            continue
        sc = STATUS_COLS[sh]
        for _, r in df.iterrows():
            rows.append({"Module": sh, "Réf": r.get("Réf", ""), "Intitulé": r.get(TITLE_COL[sh], "-"),
                         "Date": str(r.get("Date", ""))[:10], "Statut": str(r.get(sc, ""))})
    out = pd.DataFrame(rows, columns=["Module", "Réf", "Intitulé", "Date", "Statut"])
    return out.sort_values("Date", ascending=False).reset_index(drop=True) if not out.empty else out


# ---------------------------------------------------------
# CSS BARRE DU HAUT (mode + langues) — partagé connexion / application
# ---------------------------------------------------------
TOPBAR_CSS = """
.topbar-marker{display:none}
.element-container:has(.topbar-marker)+div button{background:transparent !important;
  border:none !important;box-shadow:none !important;color:@muted@ !important;
  border-radius:999px !important;font-weight:600;font-size:13px;min-height:36px;padding:0 14px}
.element-container:has(.topbar-marker)+div button:hover{background:@active@ !important;
  color:#2563EB !important}
.lang-on{display:flex;align-items:center;justify-content:center;height:36px;
  border-radius:999px;background:#2563EB;color:#fff;font-size:13px;font-weight:700}
"""
LANG_CODES = {"Français": "FR", "العربية": "AR", "English": "EN"}


def lang_switch(columns):
    for col, (name, code) in zip(columns, LANG_CODES.items()):
        with col:
            if name == st.session_state.lang:
                st.markdown(f'<div class="lang-on">{code}</div>', unsafe_allow_html=True)
            elif st.button(code, key="lang_" + code):
                st.session_state.lang = name
                rerun()


# ---------------------------------------------------------
# PAGE DE CONNEXION
# ---------------------------------------------------------
LOGIN_CSS = """
.stApp{background:linear-gradient(135deg,#EFF6FF 0%,#DBEAFE 55%,#E0F2FE 100%)}
header[data-testid="stHeader"]{background:transparent}
.block-container{max-width:1000px !important;margin:auto;padding-top:1.5rem}
[data-testid="stHorizontalBlock"]{gap:0 !important}
.login-left{position:relative;overflow:hidden;border-radius:22px 0 0 22px;min-height:540px;
  padding:40px 30px;background:linear-gradient(160deg,#1E3A8A 0%,#2563EB 55%,#38BDF8 100%);
  display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;
  box-shadow:0 18px 40px rgba(37,99,235,.25)}
.login-left *{color:#fff !important}
.blob{position:absolute;border-radius:50%;background:rgba(255,255,255,.10);animation:float 9s ease-in-out infinite}
.b1{width:220px;height:220px;top:-70px;left:-60px}
.b2{width:300px;height:300px;bottom:-120px;right:-90px;animation-delay:-4s}
@keyframes float{50%{transform:translateY(18px) scale(1.05)}}
@keyframes spin{to{transform:rotate(360deg)}}
.logo-wrap{position:relative;width:168px;height:168px;margin-bottom:6px}
.logo-wrap::before{content:"";position:absolute;inset:0;border-radius:50%;
  background:conic-gradient(from 0deg,#93C5FD,#22C55E,#FACC15,#F97316,#93C5FD);animation:spin 7s linear infinite}
.logo-in{position:absolute;inset:6px;border-radius:50%;background:#fff;display:flex;
  align-items:center;justify-content:center;overflow:hidden;box-shadow:inset 0 0 18px rgba(37,99,235,.15)}
.logo-in img{max-width:76%;max-height:76%;object-fit:contain}
.logo-in .wm{color:#1E3A8A !important;font-weight:800;letter-spacing:2px;font-size:20px}
.login-left h2{letter-spacing:7px;margin:16px 0 0 0;font-size:30px;padding:0}
.login-left .tag{font-size:14px;margin:4px 0 22px 0;opacity:.95}
.login-left ul{list-style:none;padding:0;margin:0 0 22px 0;font-size:13.5px;text-align:left}
.login-left li{margin:7px 0}
.login-left .loc{font-size:12.5px;opacity:.85;background:rgba(255,255,255,.16);padding:5px 14px;border-radius:999px}
[data-testid="column"]:has(.rm),[data-testid="stColumn"]:has(.rm){background:#fff;border-radius:0 22px 22px 0;
  padding:44px 38px;min-height:540px;box-shadow:0 18px 40px rgba(15,23,42,.10)}
.rm{display:none}
.rm-title{font-size:28px;font-weight:800;color:#0F172A;margin:6px 0 2px 0}
.rm-sub{color:#64748B;font-size:14px;margin:0 0 18px 0}
[data-testid="stForm"]{border:none !important;padding:0 !important}
[data-testid="stTextInput"] label p{color:#0F172A !important;font-weight:600}
[data-testid="stTextInput"] input{background:#F8FAFC;color:#0F172A}
[data-testid="stFormSubmitButton"] button{width:100%;background:linear-gradient(90deg,#2563EB,#38BDF8);
  color:#fff;border:none;font-weight:700;border-radius:10px;height:46px;margin-top:6px}
[data-testid="stFormSubmitButton"] button:hover{filter:brightness(1.08);color:#fff}
"""

if not st.session_state.connected:
    css_login = LOGIN_CSS + TOPBAR_CSS.replace("@muted@", "#475569").replace("@active@", "#DBEAFE")
    if st.session_state.lang == "العربية":
        css_login += ".login-left ul,.rm-title,.rm-sub{direction:rtl;text-align:right}"
    st.markdown(f"<style>{css_login}</style>", unsafe_allow_html=True)

    st.markdown('<div class="topbar-marker"></div>', unsafe_allow_html=True)
    top = cols([8, 0.7, 0.7, 0.7])
    lang_switch(top[1:])
    st.write("")

    lg = logo_data()
    logo_html = f'<img src="{lg}" alt="INPACKT">' if lg else '<span class="wm">INPACKT</span>'
    left, right = cols([1.15, 1])
    with left:
        st.markdown(
            f"""<div class="login-left"><div class="blob b1"></div><div class="blob b2"></div>
            <div class="logo-wrap"><div class="logo-in">{logo_html}</div></div>
            <h2>INPACKT</h2>
            <p class="tag">{L("Système de Management de la Qualité", "Quality Management System", "نظام إدارة الجودة")}</p>
            <ul>
              <li>✔ {L("Conformité ISO 9001", "ISO 9001 compliance", "المطابقة لمعيار ISO 9001")}</li>
              <li>✔ {L("Suivi des non-conformités et des audits", "Non-conformity and audit tracking", "متابعة عدم المطابقات والتدقيق")}</li>
              <li>✔ {L("Indicateurs et rapports en temps réel", "Real-time indicators and reports", "مؤشرات وتقارير في الوقت الفعلي")}</li>
            </ul>
            <div class="loc">📍 Médenine · Koutine</div></div>""",
            unsafe_allow_html=True,
        )
    with right:
        st.markdown(
            f'<div class="rm"></div><div>{svg("shield", "#2563EB", 40)}</div>'
            f'<div class="rm-title">{L("Bienvenue", "Welcome", "مرحبًا")}</div>'
            f'<div class="rm-sub">{L("Authentification sécurisée", "Secure authentication", "تسجيل دخول آمن")}</div>',
            unsafe_allow_html=True,
        )
        # st.form : la touche ENTRÉE et le bouton déclenchent tous les deux la connexion.
        with st.form("login_form"):
            username = st.text_input(L("Identifiant", "Username", "اسم المستخدم"), key="login_username")
            password = st.text_input(L("Mot de passe", "Password", "كلمة المرور"), type="password", key="login_password")
            submitted = st.form_submit_button(L("Connexion au portail", "Sign in to the portal", "الدخول إلى البوابة"))
        if submitted:
            account = USERS.get(username.strip())
            if account and password == account[0]:
                st.session_state.connected = True
                st.session_state["user"], st.session_state["role"], st.session_state["display"] = username.strip(), account[1], account[2]
                st.session_state["nav_menu"] = "home"
                rerun()
            else:
                st.error(L("Identifiant ou mot de passe incorrect", "Incorrect username or password", "اسم المستخدم أو كلمة المرور غير صحيحة"))
    st.stop()

# ---------------------------------------------------------
# THÈME (clair / sombre) + ARRIÈRE-PLAN
# ---------------------------------------------------------
if st.session_state.dark_mode:
    P = dict(bg="#0B1120", card="#111827", text="#F1F5F9", muted="#94A3B8",
             border="#1F2937", sb="#0F172A", active="rgba(59,130,246,.22)",
             thead="#1E293B", hero="linear-gradient(135deg,#1D4ED8 0%,#3B82F6 100%)")
else:
    P = dict(bg="#F8FAFC", card="#FFFFFF", text="#0F172A", muted="#64748B",
             border="#E2E8F0", sb="#FFFFFF", active="#DBEAFE",
             thead="#F1F5F9", hero="linear-gradient(135deg,#2563EB 0%,#60A5FA 100%)")

BG_KEYS = ["grad", "dots", "grid", "waves", "logo", "none"]
BG_FR = {"grad": "Dégradé doux", "dots": "Points", "grid": "Grille qualité",
         "waves": "Vagues", "logo": "Filigrane logo", "none": "Aucun"}


def bg_css(choice, dark):
    base = P["bg"]
    ln = "rgba(148,163,184,.10)" if dark else "rgba(37,99,235,.07)"
    dt = "rgba(148,163,184,.18)" if dark else "rgba(37,99,235,.18)"
    lg = logo_data()
    if choice == "logo" and not lg:
        choice = "grad"
    if choice == "grad":
        g = ("linear-gradient(135deg,#0B1120 0%,#13203A 55%,#0B1120 100%)" if dark
             else "linear-gradient(135deg,#F8FAFC 0%,#E3EEFF 50%,#F5F9FF 100%)")
        return ".stApp{background:" + g + ";background-attachment:fixed}"
    if choice == "dots":
        return (".stApp{background-color:" + base + ";background-image:radial-gradient(circle," + dt +
                " 1.3px,transparent 1.3px);background-size:22px 22px}")
    if choice == "grid":
        return (".stApp{background-color:" + base + ";background-image:linear-gradient(" + ln +
                " 1px,transparent 1px),linear-gradient(90deg," + ln + " 1px,transparent 1px);background-size:32px 32px}")
    if choice == "waves":
        c1 = "rgba(59,130,246,.18)" if dark else "rgba(37,99,235,.10)"
        c2 = "rgba(56,189,248,.14)" if dark else "rgba(56,189,248,.12)"
        s = ("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 1440 320' preserveAspectRatio='none'>"
             "<path fill='" + c1 + "' d='M0,224L48,208C96,192,192,160,288,170.7C384,181,480,235,576,234.7C672,235,768,181,864,165.3C960,149,1056,171,1152,192C1248,213,1344,235,1392,245.3L1440,256L1440,320L0,320Z'/>"
             "<path fill='" + c2 + "' d='M0,288L60,277.3C120,267,240,245,360,250.7C480,256,600,288,720,293.3C840,299,960,277,1080,256C1200,235,1320,213,1380,202.7L1440,192L1440,320L0,320Z'/></svg>")
        return (".stApp{background-color:" + base + ";background-image:url(\"data:image/svg+xml;utf8," + quote(s) +
                "\");background-repeat:no-repeat;background-position:bottom;background-size:100% 340px;background-attachment:fixed}")
    if choice == "logo":
        return (".stApp{background-color:" + base + "}.stApp::before{content:'';position:fixed;inset:0;background:url('" + lg +
                "') center/360px no-repeat;opacity:" + ("0.06" if dark else "0.07") + ";pointer-events:none;z-index:0}")
    return ".stApp{background-color:" + base + "}"


CSS = """
.stApp{color:@text@}
[data-testid="stAppViewContainer"],[data-testid="stMain"],.main{background:transparent}
header[data-testid="stHeader"]{background:transparent}
.block-container{padding-top:1rem;max-width:100%}
.stApp h1,.stApp h2,.stApp h3,.stApp p,.stApp label,.stApp li{color:@text@}

/* ----- Barre latérale ----- */
[data-testid="stSidebar"]{background:@sb@;border-right:1px solid @border@}
.brand{display:flex;align-items:center;gap:12px;font-size:20px;font-weight:700;
  color:@text@;padding:4px 4px 18px 4px}
.nav-label{color:@muted@;font-size:14px;padding:0 6px 6px 6px}
[data-testid="stSidebar"] [role="radiogroup"]{gap:4px}
[data-testid="stSidebar"] [role="radiogroup"] label{display:flex;align-items:center;
  gap:12px;width:100%;padding:11px 12px;border-radius:10px;cursor:pointer}
[data-testid="stSidebar"] [role="radiogroup"] label>div:first-child{display:none}
[data-testid="stSidebar"] [role="radiogroup"] label p{color:@text@;font-size:15px;margin:0}
[data-testid="stSidebar"] [role="radiogroup"] label::before{content:"";flex:0 0 20px;
  width:20px;height:20px;background-repeat:no-repeat;background-size:contain}
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked){background:@active@}
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) p{color:#2563EB;font-weight:600}
[data-testid="stSidebar"] .stButton button{background:transparent;border:none;
  color:#EF4444;font-weight:600;box-shadow:none}
[data-testid="stSidebar"] .stButton button:hover{background:rgba(239,68,68,.1);color:#EF4444}
.syn{margin-top:12px !important;font-size:14px !important;font-weight:600;
  background:rgba(255,255,255,.18);display:inline-block;padding:5px 12px;border-radius:8px}
.spark{margin-top:6px;display:flex;align-items:center;justify-content:space-between}
.trend{font-size:12px;font-weight:700}
.ring{display:flex;justify-content:center;padding:4px 0}
.ring .num{font-size:28px;font-weight:700;fill:@text@}
.dl{display:flex;align-items:center;gap:10px;padding:8px 0;font-size:13px;
  border-bottom:1px solid @border@;color:@text@}
.pill{margin-left:auto;font-size:12px;font-weight:700;padding:3px 10px;
  border-radius:999px;white-space:nowrap}
.stack{display:flex;height:16px;border-radius:8px;overflow:hidden;margin:14px 0 12px 0}
.leg{display:flex;gap:14px;flex-wrap:wrap;font-size:12px;color:@muted@}
.leg i{display:inline-block;width:9px;height:9px;border-radius:50%;margin-right:5px}
.hbar{height:10px;border-radius:5px;margin:4px 0 12px 0}
.quick-marker{display:none}
.element-container:has(.quick-marker)+div button{border-radius:999px !important;
  border:1px solid #2563EB !important;color:#2563EB !important;background:transparent !important;
  font-weight:600;min-height:38px}
.element-container:has(.quick-marker)+div button:hover{background:#2563EB !important;color:#fff !important}
.kpi{min-height:150px}
@navicons@

/* ----- Bandeau d'accueil ----- */
.hero{position:relative;overflow:hidden;background:@hero@;color:#fff;
  padding:26px 30px;border-radius:14px;margin-bottom:22px}
.hero h1{color:#fff !important;margin:10px 0 4px 0;font-size:30px;padding:0}
.hero p{color:#fff !important;margin:0;font-size:15px;opacity:.95}
.live{display:inline-flex;align-items:center;gap:6px;background:#22C55E;color:#fff;
  font-size:12px;font-weight:600;padding:4px 12px;border-radius:20px}
.live i{width:7px;height:7px;border-radius:50%;background:#fff;display:inline-block}
.hero-shield{position:absolute;right:70px;top:50%;transform:translateY(-50%);
  width:86px;height:86px;border-radius:50%;background:rgba(255,255,255,.18);
  display:flex;align-items:center;justify-content:center}

/* ----- Cartes KPI ----- */
.kpi{border-radius:12px;padding:18px 18px 14px 18px;min-height:128px}
.kpi .ico{width:36px;height:36px;border-radius:10px;display:flex;
  align-items:center;justify-content:center;margin-bottom:10px}
.kpi .lbl{font-size:14px;font-weight:600;color:@text@}
.kpi .val{font-size:28px;font-weight:700;color:@text@;margin-top:2px}

/* ----- Alertes ----- */
.alert{border-radius:10px;padding:16px 20px;margin:18px 0 22px 0}
.alert.bad{background:rgba(239,68,68,.10);border:1px solid rgba(239,68,68,.30)}
.alert.ok{background:rgba(16,185,129,.10);border:1px solid rgba(16,185,129,.30)}
.alert h4{display:flex;align-items:center;gap:10px;margin:0 0 6px 0;font-size:17px;padding:0}
.alert.bad h4{color:#DC2626}
.alert.ok h4{color:#059669}
.alert ul{margin:0;padding-left:46px;font-size:14px;color:@text@}

/* ----- Cartes ----- */
.card{background:@card@;border:1px solid @border@;border-radius:10px;padding:16px 18px;height:100%}
.card h3{font-size:17px;margin:0 0 12px 0;padding:0;color:@text@;display:flex;
  align-items:center;justify-content:space-between;gap:10px}
.sec-title{font-size:19px;font-weight:700;margin:14px 0 10px 0;color:@text@}
table.recap{width:100%;border-collapse:collapse;font-size:14px;
  background:@card@;border:1px solid @border@;border-radius:8px;overflow:hidden}
table.recap th{background:@thead@;text-align:left;padding:10px 16px;font-weight:600;color:@text@}
table.recap td{padding:9px 16px;border-top:1px solid @border@;color:@text@}
.bars{display:flex;align-items:flex-end;justify-content:space-around;height:190px;
  border-bottom:1px solid @border@;padding:0 6px}
.bar-col{display:flex;flex-direction:column;align-items:center;justify-content:flex-end;
  height:100%;width:22%}
.bar{width:100%;border-radius:4px 4px 0 0;min-height:3px}
.bar-n{font-size:12px;color:@muted@;margin-bottom:3px}
.bar-names{display:flex;justify-content:space-around;margin-top:8px;font-size:12px;color:@muted@}
.bar-names span{width:22%;text-align:center}
.act{display:flex;align-items:center;gap:10px;padding:8px 0;font-size:13px;color:@text@}
.act .dot{width:10px;height:10px;border-radius:50%;flex:0 0 10px}
.act .when{margin-left:auto;color:@muted@;font-size:12px;white-space:nowrap}

/* ----- Radios horizontaux (zone principale) ----- */
.block-container [role="radiogroup"]{flex-direction:row !important;flex-wrap:wrap;gap:6px 22px}

/* ----- Tableaux colorés ----- */
.dt-wrap{overflow-x:auto;border:1px solid @border@;border-radius:10px;background:@card@;margin:6px 0 8px 0}
table.dt{width:100%;border-collapse:collapse;font-size:13.5px}
table.dt th{background:@thead@;text-align:left;padding:10px 12px;font-weight:600;color:@text@;white-space:nowrap}
table.dt td{padding:9px 12px;border-top:1px solid @border@;color:@text@;vertical-align:middle}
table.dt tr td:first-child{border-left:5px solid var(--rc,transparent)}
.chip{display:inline-block;padding:3px 12px;border-radius:999px;font-size:12px;font-weight:700;white-space:nowrap}
.card table.dt th,.card table.dt td{padding:8px 10px}
"""

nav_css = ""
for i, key in enumerate(NAV, start=1):
    sel = f'[data-testid="stSidebar"] [role="radiogroup"] label:nth-of-type({i})'
    nav_css += (
        f'{sel}::before{{background-image:url("{svg_url(key, "#64748B")}")}}\n'
        f'{sel}:has(input:checked)::before{{background-image:url("{svg_url(key, "#2563EB")}")}}\n'
    )
CSS = (CSS + TOPBAR_CSS).replace("@navicons@", nav_css)
for k, v in P.items():
    CSS = CSS.replace(f"@{k}@", v)
CSS += bg_css(st.session_state.bg, st.session_state.dark_mode)
if st.session_state.lang == "العربية":
    CSS += (".card,.alert,.hero,table.recap,table.dt,.dl,.act,.leg,.kpi{direction:rtl;text-align:right}"
            "table.recap th,table.dt th{text-align:right}.hero-shield{right:auto;left:70px}"
            ".alert ul{padding-left:0;padding-right:46px}table.dt tr td:first-child{border-left:0;border-right:5px solid var(--rc,transparent)}")
st.markdown(f"<style>{CSS}</style>", unsafe_allow_html=True)

# ---------------------------------------------------------
# BARRE LATÉRALE
# ---------------------------------------------------------
lg = logo_data()
brand_icon = (f'<img src="{lg}" style="width:38px;height:38px;object-fit:contain;border-radius:8px;">' if lg
              else svg("shield", "#2563EB", 34, 1.8, "rgba(37,99,235,.15)"))
st.sidebar.markdown(
    f'<div class="brand">{brand_icon}<span>{t["title"]}</span></div>'
    f'<div class="nav-label">{t["navigation"]}</div>',
    unsafe_allow_html=True,
)
if "goto" not in st.session_state:
    st.session_state["goto"] = ""
if st.session_state["goto"]:
    st.session_state["nav_menu"] = st.session_state["goto"]
    st.session_state["goto"] = ""
st.sidebar.markdown(f'<div class="nav-label">👤 {html.escape(st.session_state["display"])} · {html.escape(tr(st.session_state["role"]) if False else st.session_state["role"])}</div>',
                    unsafe_allow_html=True)
nav_items = [k for k in NAV if k != "history" or st.session_state["role"] == "Administrateur"]
if st.session_state.get("nav_menu") not in nav_items:
    st.session_state["nav_menu"] = "home"
menu = st.sidebar.radio(" ", nav_items, format_func=lambda k: t[k], key="nav_menu")
st.sidebar.markdown("---")
st.sidebar.selectbox("🎨 " + L("Arrière-plan", "Background", "الخلفية"), BG_KEYS,
                     format_func=lambda k: tr(BG_FR[k]), key="bg")
if st.sidebar.button("⎋  " + t["logout"], key="logout_btn"):
    st.session_state.connected = False
    rerun()

# ---------------------------------------------------------
# BARRE DU HAUT (mode + langue)
# ---------------------------------------------------------
st.markdown('<div class="topbar-marker"></div>', unsafe_allow_html=True)
bar = cols([9, 1.4, 0.7, 0.7, 0.7])
with bar[1]:
    dark_on = st.session_state.dark_mode
    label = ("☀️ " + t["mode_light"]) if dark_on else ("🌙 " + t["mode_dark"])
    if st.button(label, key="menu_theme_button"):
        st.session_state.dark_mode = not dark_on
        rerun()
lang_switch(bar[2:])


# ---------------------------------------------------------
# OUTILS DE LA PAGE D'ACCUEIL
# ---------------------------------------------------------
def filter_period(df, days):
    if days is None or df.empty or "Date" not in df.columns:
        return df
    d = pd.to_datetime(df["Date"], errors="coerce")
    return df[d >= pd.Timestamp.now().normalize() - pd.Timedelta(days=days)]


def monthly_counts(df, n=6):
    if df.empty or "Date" not in df.columns:
        return [0] * n
    d = pd.to_datetime(df["Date"], errors="coerce").dropna()
    now, out = pd.Timestamp.now(), []
    for k in range(n - 1, -1, -1):
        y, m = now.year, now.month - k
        while m <= 0:
            m += 12
            y -= 1
        out.append(int(((d.dt.year == y) & (d.dt.month == m)).sum()))
    return out


def sparkline(vals, color):
    w, h = 100, 28
    mx, step = max(vals + [1]), w / (len(vals) - 1)
    pts = " ".join(f"{i * step:.1f},{h - 2 - (v / mx) * (h - 6):.1f}" for i, v in enumerate(vals))
    return (f'<svg width="{w}" height="{h}" viewBox="0 0 {w} {h}"><polyline points="{pts}" '
            f'fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" '
            'stroke-linejoin="round"/></svg>')


def trend_html(vals, color):
    d = vals[-1] - vals[-2]
    if d > 0:
        return f'<span class="trend" style="color:{color};">▲ +{d}</span>'
    if d < 0:
        return f'<span class="trend" style="color:{color};">▼ {d}</span>'
    return '<span class="trend" style="color:#94A3B8;">＝ 0</span>'


def score_ring(score):
    color = "#10B981" if score >= 75 else ("#F59E0B" if score >= 50 else "#EF4444")
    circ = 339.3
    return (
        '<div class="ring"><svg width="150" height="150" viewBox="0 0 140 140">'
        '<circle cx="70" cy="70" r="54" fill="none" stroke="rgba(148,163,184,.25)" stroke-width="12"/>'
        f'<circle cx="70" cy="70" r="54" fill="none" stroke="{color}" stroke-width="12" '
        f'stroke-linecap="round" stroke-dasharray="{score / 100 * circ:.1f} {circ}" '
        'transform="rotate(-90 70 70)"/>'
        f'<text x="70" y="78" text-anchor="middle" class="num">{score:.0f}%</text></svg></div>'
    )


def pill(text, color):
    return f'<span class="pill" style="background:{color}22;color:{color};">{html.escape(str(text))}</span>'


def dot(color):
    return f'<span style="width:10px;height:10px;border-radius:50%;background:{color};flex:0 0 10px;"></span>'


def iso_data(nc, rec, fourn):
    """Indicateurs ISO 9001 : (libellé, valeur, couleur)."""
    out = []
    d = None
    if not nc.empty and {"Date", "Date clôture"} <= set(nc.columns):
        diff = (pd.to_datetime(nc["Date clôture"], errors="coerce") - pd.to_datetime(nc["Date"], errors="coerce")).dt.days.dropna()
        if len(diff):
            d = diff.mean()
    out.append((L("Délai moyen de clôture d'une NC", "Avg. NC closing time", "متوسط مدة إغلاق عدم المطابقة"),
                "—" if d is None else f"{d:.1f} " + L("j", "d", "يوم"), "#3B82F6"))
    rate = None
    if not rec.empty and {"Date", "Date clôture", "Délai"} <= set(rec.columns):
        dd = pd.DataFrame({"d": (pd.to_datetime(rec["Date clôture"], errors="coerce") - pd.to_datetime(rec["Date"], errors="coerce")).dt.days,
                           "lim": pd.to_numeric(rec["Délai"], errors="coerce")}).dropna()
        if len(dd):
            rate = (dd["d"] <= dd["lim"]).mean() * 100
    out.append((L("Réclamations traitées dans les délais", "Complaints handled on time", "الشكاوى المعالجة في الوقت"),
                "—" if rate is None else f"{rate:.0f} %",
                "#94A3B8" if rate is None else ("#10B981" if rate >= 80 else ("#F59E0B" if rate >= 50 else "#EF4444"))))
    sup = None
    if not fourn.empty and "Évaluation" in fourn.columns:
        ev = fourn["Évaluation"].dropna().astype(str).str[:1]
        if len(ev):
            sup = ev.isin(["A", "B"]).mean() * 100
    out.append((L("Fournisseurs conformes (A/B)", "Compliant suppliers (A/B)", "الموردون المطابقون (A/B)"),
                "—" if sup is None else f"{sup:.0f} %",
                "#94A3B8" if sup is None else ("#10B981" if sup >= 80 else ("#F59E0B" if sup >= 50 else "#EF4444"))))
    return out


def iso_html(nc, rec, fourn):
    return "".join(f'<div class="dl"><span>{lab}</span>{pill(val, c)}</div>' for lab, val, c in iso_data(nc, rec, fourn))


def audit_calendar_html(audits):
    empty = f'<div class="dl">{L("Aucun audit planifié.", "No planned audit.", "لا توجد عمليات تدقيق مخططة.")}</div>'
    if audits.empty or "Date prévue" not in audits.columns:
        return empty
    a = audits.copy()
    a["_d"] = pd.to_datetime(a["Date prévue"], errors="coerce")
    if "Réalisation" in a.columns:
        a = a[a["Réalisation"] == "Planifié"]
    a = a.dropna(subset=["_d"]).sort_values("_d").head(5)
    if a.empty:
        return empty
    today, rows = pd.Timestamp.now().normalize(), ""
    for _, r in a.iterrows():
        d = (r["_d"].normalize() - today).days
        day = r["_d"].strftime("%d/%m/%Y")
        if d < 0:
            c, txt = "#EF4444", L(f"retard {-d} j · {day}", f"{-d} d late · {day}", f"متأخر {-d} يوم · {day}")
        elif d <= 7:
            c, txt = "#F59E0B", L(f"rappel : dans {d} j · {day}", f"reminder: in {d} d · {day}", f"تذكير: بعد {d} يوم · {day}")
        else:
            c, txt = "#10B981", day
        rows += f'<div class="dl">{dot(c)}<span>{html.escape(str(r.get("Réf", "")))} {html.escape(str(r.get("Type", ""))[:22])}</span>{pill(txt, c)}</div>'
    return rows


def pareto_svg(nc):
    if nc.empty or "Cause" not in nc.columns or nc["Cause"].dropna().empty:
        return None
    vc = nc["Cause"].dropna().astype(str).value_counts()
    cum = (vc.cumsum() / vc.sum() * 100).tolist()
    W, H, ml, mb, mt = 560, 270, 36, 58, 16
    pw, ph, n = W - ml - 16, H - mb - mt, len(vc)
    bw, mx, mu = pw / n, int(vc.max()), P["muted"]
    parts = [f'<line x1="{ml}" y1="{mt + ph * 0.2:.1f}" x2="{W - 16}" y2="{mt + ph * 0.2:.1f}" stroke="#F59E0B" stroke-dasharray="5 4"/>',
             f'<text x="{W - 16}" y="{mt + ph * 0.2 - 4:.1f}" text-anchor="end" font-size="10" fill="#F59E0B">80 %</text>']
    pts = []
    for i, (name, v) in enumerate(vc.items()):
        x, h = ml + i * bw, v / mx * ph * 0.95
        parts.append(f'<rect x="{x + 8:.1f}" y="{mt + ph - h:.1f}" width="{bw - 16:.1f}" height="{h:.1f}" rx="4" fill="#3B82F6"/>')
        parts.append(f'<text x="{x + bw / 2:.1f}" y="{mt + ph - h - 4:.1f}" text-anchor="middle" font-size="11" fill="{mu}">{v}</text>')
        parts.append(f'<text x="{x + bw / 2:.1f}" y="{H - 34}" text-anchor="middle" font-size="11" fill="{mu}">{html.escape(tr(name)[:12])}</text>')
        pts.append((x + bw / 2, mt + ph - cum[i] / 100 * ph))
    line = " ".join(f"{a:.1f},{b:.1f}" for a, b in pts)
    parts.append(f'<polyline points="{line}" fill="none" stroke="#EF4444" stroke-width="2"/>')
    parts += [f'<circle cx="{a:.1f}" cy="{b:.1f}" r="3.5" fill="#EF4444"/>' for a, b in pts]
    parts.append(f'<line x1="{ml}" y1="{mt + ph}" x2="{W - 16}" y2="{mt + ph}" stroke="{mu}" stroke-opacity=".4"/>')
    return f'<svg width="100%" viewBox="0 0 {W} {H}">{"".join(parts)}</svg>'


SHEETS = ["NonConformites", "Audits", "Reclamations", "Actions", "Fournisseurs", "FluxLogistique", "Historique"]


def excel_bytes():
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for sh in SHEETS:
            df = load_sheet(sh)
            (df if not df.empty else pd.DataFrame({"Info": ["Aucune donnée"]})).to_excel(w, sheet_name=sh, index=False)
    return buf.getvalue()


def localize_df(df):
    """Traduit les en-têtes et les valeurs (pour le rapport exporté)."""
    if st.session_state.lang == "Français" or df.empty:
        return df
    return df.rename(columns=tr).replace({k: tr(k) for k in TR})


def report_html():
    body = ""
    for sh in SHEETS[:-1]:
        df = load_sheet(sh)
        body += f"<h2>{html.escape(mod_name(sh))}</h2>" + (localize_df(df).to_html(index=False, border=0) if not df.empty
                                                          else "<p>" + L("Aucune donnée.", "No data.", "لا توجد بيانات.") + "</p>")
    iso = "".join(f"<li>{lab} : <b>{val}</b></li>" for lab, val, _ in iso_data(load_sheet("NonConformites"), load_sheet("Reclamations"), load_sheet("Fournisseurs")))
    css = ("body{font-family:Arial,sans-serif;margin:30px;color:#0F172A}h1{color:#2563EB}"
           "table{border-collapse:collapse;width:100%;margin-bottom:20px;font-size:13px}"
           "th{background:#F1F5F9;text-align:left;padding:6px}td{padding:6px;border-top:1px solid #E2E8F0}")
    direction = "rtl" if st.session_state.lang == "العربية" else "ltr"
    return (f"<html dir='{direction}'><head><meta charset='utf-8'><title>SMQ</title><style>{css}</style></head><body>"
            f"<h1>{L('Rapport SMQ', 'QMS Report', 'تقرير نظام إدارة الجودة')} — {datetime.now().strftime('%d/%m/%Y %H:%M')}</h1>"
            f"<h2>{L('Indicateurs ISO 9001', 'ISO 9001 indicators', 'مؤشرات ISO 9001')}</h2><ul>{iso}</ul>{body}</body></html>")


def offer_download(label, data, filename, mime):
    if hasattr(st, "download_button"):
        st.download_button(label, data, file_name=filename, mime=mime)
    else:
        b64 = base64.b64encode(data).decode()
        st.markdown(f'<a href="data:{mime};base64,{b64}" download="{filename}" style="display:inline-block;'
                    f'padding:8px 18px;border-radius:999px;border:1px solid #2563EB;color:#2563EB;'
                    f'font-weight:600;text-decoration:none;margin:4px 8px 4px 0;">{label}</a>', unsafe_allow_html=True)


def pending_with_due(actions):
    if actions.empty or not {"Statut", "Action", "Échéance"} <= set(actions.columns):
        return pd.DataFrame()
    a = actions[actions["Statut"] != "Terminée"].copy()
    a["_due"] = pd.to_datetime(a["Échéance"], errors="coerce")
    return a.dropna(subset=["_due"]).sort_values("_due")


def deadlines_html(a):
    if a.empty:
        return f'<div class="dl">{L("Aucune échéance à venir.", "No upcoming deadline.", "لا توجد مواعيد قادمة.")}</div>'
    today, rows = pd.Timestamp.now().normalize(), ""
    for _, r in a.head(5).iterrows():
        d = (r["_due"].normalize() - today).days
        if d < 0:
            txt, c = L(f"en retard de {-d} j", f"{-d} d overdue", f"متأخر {-d} يوم"), "#EF4444"
        elif d == 0:
            txt, c = L("aujourd'hui", "today", "اليوم"), "#F59E0B"
        else:
            txt, c = L(f"dans {d} j", f"in {d} d", f"بعد {d} يوم"), ("#10B981" if d > 7 else "#F59E0B")
        rows += f'<div class="dl">{dot(c)}<span>{html.escape(str(r["Action"])[:30])}</span>{pill(txt, c)}</div>'
    return rows


def audits_html(audits):
    colors = [("Conforme", "#10B981"), ("À améliorer", "#F59E0B"), ("Non conforme", "#EF4444")]
    if audits.empty or "Résultat" not in audits.columns:
        return f'<div class="dl">{L("Aucun audit.", "No audit.", "لا توجد عمليات تدقيق.")}</div>'
    counts = [(n, c, int((audits["Résultat"] == n).sum())) for n, c in colors]
    total = max(sum(c for _, _, c in counts), 1)
    bar_ = "".join(f'<div style="width:{c / total * 100:.1f}%;background:{col};"></div>' for _, col, c in counts)
    leg = "".join(f'<span><i style="background:{col};"></i>{html.escape(tr(n))} ({c})</span>' for n, col, c in counts)
    return f'<div class="stack">{bar_}</div><div class="leg">{leg}</div>'


def suppliers_html(df):
    if df.empty or "Nom" not in df.columns:
        return f'<div class="dl">{L("Aucun fournisseur.", "No supplier.", "لا يوجد موردون.")}</div>'
    blank = pd.Series([""] * len(df), index=df.index)
    ev = df["Évaluation"].astype(str).str[:1] if "Évaluation" in df.columns else blank
    stt = df["Statut"].astype(str) if "Statut" in df.columns else blank
    risk = df[ev.isin(["C", "D"]) | stt.isin(["Sous observation", "Suspendu"])].head(5)
    if risk.empty:
        return f'<div class="dl">{L("Aucun fournisseur à risque ✅", "No supplier at risk ✅", "لا موردون في خطر ✅")}</div>'
    rows = ""
    for idx, r in risk.iterrows():
        letter = str(ev[idx])
        c = "#EF4444" if letter == "D" else ("#F59E0B" if letter == "C" else "#3B82F6")
        rows += (f'<div class="dl">{dot(c)}<span>{html.escape(str(r["Nom"])[:24])}</span>'
                 f'{pill((letter or "-") + " · " + tr(stt[idx]), c)}</div>')
    return rows


def logistics_html(flux):
    if flux.empty or "Type de Flux" not in flux.columns:
        return f'<div class="dl">{L("Aucun flux.", "No flow.", "لا توجد حركات.")}</div>'
    typ = flux["Type de Flux"].astype(str)
    qty = pd.to_numeric(flux["Quantité"], errors="coerce").fillna(0) if "Quantité" in flux.columns else typ.map(lambda x: 0)
    e, s = float(qty[typ.str.startswith("Entrée")].sum()), float(qty[typ.str.startswith("Sortie")].sum())
    mx = max(e, s, 1)
    nonconf = 0.0
    if "Contrôle" in flux.columns:
        nonconf = round(flux["Contrôle"].astype(str).str.startswith("Réservé").mean() * 100, 1)
    c = "#EF4444" if nonconf > 10 else "#10B981"
    return (
        f'<div class="leg">{L("Entrées", "Inbound", "الدخول")} : <b>{e:.0f}</b></div>'
        f'<div class="hbar" style="width:{e / mx * 100:.0f}%;background:#10B981;"></div>'
        f'<div class="leg">{L("Sorties", "Outbound", "الخروج")} : <b>{s:.0f}</b></div>'
        f'<div class="hbar" style="width:{s / mx * 100:.0f}%;background:#3B82F6;"></div>'
        f'<div class="dl">{L("Lots réservés / non conformes", "Held / non-conforming lots", "دفعات غير مطابقة")}{pill(str(nonconf) + " %", c)}</div>'
    )


# ---------------------------------------------------------
# ÉVALUATION : critères automatiques pour chaque table
# Format d'un critère : (libellé FR, valeur, unité, objectif, True si "plus haut = mieux")
# Pour ajouter / modifier un critère : éditez build_criteria() + ajoutez la traduction dans TR.
# ---------------------------------------------------------
def pct(num, den):
    return None if not den else float(num) / den * 100


def _s(df, c):
    return df[c].astype(str) if c in df.columns else pd.Series([""] * len(df), index=df.index)


def _d(df, c):
    if c in df.columns:
        return pd.to_datetime(df[c], errors="coerce")
    return pd.to_datetime(pd.Series([None] * len(df), index=df.index))


def build_criteria():
    today = pd.Timestamp.now().normalize()
    out = {}

    nc = load_sheet("NonConformites"); n = len(nc); s = _s(nc, "Statut")
    dd = (_d(nc, "Date clôture") - _d(nc, "Date")).dt.days.dropna()
    linked = ~_s(nc, "Action liée").isin(["—", "nan", "None", ""])
    out["NonConformites"] = [
        ("Taux de clôture", pct((s == "Clôturée").sum(), n), "%", 80, True),
        ("Délai moyen de clôture", float(dd.mean()) if len(dd) else None, "j", 15, False),
        ("NC critiques non clôturées", int(((_s(nc, "Gravité") == "Critique") & (s != "Clôturée")).sum()), "", 0, False),
        ("NC avec action corrective liée", pct(linked.sum(), n), "%", 80, True),
    ]

    a = load_sheet("Audits"); n = len(a)
    late = int(((_s(a, "Réalisation") == "Planifié") & (_d(a, "Date prévue").dt.normalize() < today)).sum())
    out["Audits"] = [
        ("Taux de réalisation", pct((_s(a, "Réalisation") == "Réalisé").sum(), n), "%", 80, True),
        ("Taux de conformité", pct((_s(a, "Résultat") == "Conforme").sum(), n), "%", 80, True),
        ("Audits planifiés en retard", late, "", 0, False),
    ]

    r = load_sheet("Reclamations"); n = len(r); s = _s(r, "Statut")
    lim = pd.to_numeric(r["Délai"], errors="coerce") if "Délai" in r.columns else pd.Series([float("nan")] * len(r), index=r.index)
    dr = pd.DataFrame({"d": (_d(r, "Date clôture") - _d(r, "Date")).dt.days, "lim": lim}).dropna()
    out["Reclamations"] = [
        ("Taux de traitement", pct(s.isin(["Traitée", "Clôturée"]).sum(), n), "%", 80, True),
        ("Réclamations traitées dans les délais", pct((dr["d"] <= dr["lim"]).sum(), len(dr)), "%", 80, True),
        ("Réclamations non prises en charge", int((s == "Nouvelle").sum()), "", 0, False),
    ]

    ac = load_sheet("Actions"); n = len(ac); s = _s(ac, "Statut")
    out["Actions"] = [
        ("Taux de réalisation", pct((s == "Terminée").sum(), n), "%", 80, True),
        ("Actions en retard", int(((s != "Terminée") & (_d(ac, "Échéance").dt.normalize() < today)).sum()), "", 0, False),
    ]

    f = load_sheet("Fournisseurs"); n = len(f)
    ev = _s(f, "Évaluation").str[:1]
    risk = ev.isin(["C", "D"]) | _s(f, "Statut").isin(["Sous observation", "Suspendu"])
    out["Fournisseurs"] = [
        ("Fournisseurs conformes (A/B)", pct(ev.isin(["A", "B"]).sum(), n), "%", 80, True),
        ("Fournisseurs à risque", int(risk.sum()), "", 0, False),
    ]

    fl = load_sheet("FluxLogistique"); n = len(fl); c = _s(fl, "Contrôle")
    out["FluxLogistique"] = [
        ("Taux de lots conformes", pct(c.str.startswith("Validé").sum(), n), "%", 90, True),
        ("Lots réservés / non conformes", int(c.str.startswith("Réservé").sum()), "", 0, False),
    ]
    return out


def crit_eval(v, target, higher):
    """Retourne (score 0-100, couleur) ; (None, 'gray') si pas de donnée."""
    if v is None:
        return None, "gray"
    if higher:
        score, met = (min(v / target, 1) * 100 if target else 100), v >= target
    else:
        score, met = (100 if v <= target else max(0, 100 - (v - target) / max(target, 1) * 100)), v <= target
    return score, ("green" if met else ("orange" if score >= 50 else "red"))


def fmt_val(v, unit):
    if v is None:
        return "—"
    if unit == "%":
        return f"{v:.0f} %"
    if unit == "j":
        return f"{v:.1f} " + L("j", "d", "يوم")
    return f"{int(v)}"


CRIT_LABEL = {"green": "Atteint", "orange": "Proche de l'objectif", "red": "Non atteint", "gray": "Pas de données"}


def score_key(sc):
    return "gray" if sc is None else ("green" if sc >= 80 else ("orange" if sc >= 50 else "red"))


# ---------------------------------------------------------
# PAGE : VUE D'ENSEMBLE
# ---------------------------------------------------------
if menu == "home":
    nc_all = load_sheet("NonConformites")
    audits_all = load_sheet("Audits")
    rec_all = load_sheet("Reclamations")
    actions_all = load_sheet("Actions")
    fourn_all = load_sheet("Fournisseurs")
    flux_all = load_sheet("FluxLogistique")

    due_df = pending_with_due(actions_all)
    today = pd.Timestamp.now().normalize()
    n_late = int((due_df["_due"].dt.normalize() < today).sum()) if not due_df.empty else 0
    n_open = int((nc_all["Statut"] != "Clôturée").sum()) if (not nc_all.empty and "Statut" in nc_all.columns) else 0
    if n_late == 0 and n_open == 0:
        syn = L("Tout est sous contrôle ✅", "Everything is under control ✅", "كل شيء تحت السيطرة ✅")
    else:
        syn = L(f"{n_late} action(s) en retard · {n_open} non-conformité(s) non clôturée(s)",
                f"{n_late} overdue action(s) · {n_open} open non-conformity(ies)",
                f"{n_late} إجراء متأخر · {n_open} عدم مطابقة مفتوحة")

    hello = t["hello_day"] if datetime.now().hour < 18 else t["hello_night"]
    st.markdown(
        f"""<div class="hero">
        <span class="live"><i></i>Database Sync Live</span>
        <h1>{hello}, {html.escape(st.session_state["display"])} !</h1>
        <p>{t['desc']}</p>
        <p class="syn">{syn}</p>
        <div class="hero-shield">{svg('shield', '#FFFFFF', 46, 1.6)}</div>
        </div>""",
        unsafe_allow_html=True,
    )

    st.markdown('<div class="quick-marker"></div>', unsafe_allow_html=True)
    q = cols([1.2, 1, 1.2, 1.6])
    for col, target, label in [
        (q[0], "nc", L("＋ Non-conformité", "＋ Non-conformity", "＋ عدم مطابقة")),
        (q[1], "audits", L("＋ Audit", "＋ Audit", "＋ تدقيق")),
        (q[2], "reclamations", L("＋ Réclamation", "＋ Complaint", "＋ شكوى")),
    ]:
        with col:
            if st.button(label, key="quick_" + target):
                st.session_state["goto"] = target
                rerun()
    with q[3]:
        period_labels = [L("Toute la période", "All time", "كل الفترة"), L("7 derniers jours", "Last 7 days", "آخر 7 أيام"),
                         L("30 derniers jours", "Last 30 days", "آخر 30 يوم"), L("Cette année", "This year", "هذه السنة")]
        p_idx = st.selectbox(L("Période", "Period", "الفترة"), [0, 1, 2, 3],
                             format_func=lambda i: period_labels[i], key="period")
    days = [None, 7, 30, 365][p_idx]

    nc = filter_period(nc_all, days)
    audits = filter_period(audits_all, days)
    reclamations = filter_period(rec_all, days)
    actions = filter_period(actions_all, days)
    fournisseurs = filter_period(fourn_all, days)
    flux = filter_period(flux_all, days)

    rate_nc = get_resolution_rate(nc, ["Clôturée"])
    rate_rec = get_resolution_rate(reclamations, ["Traitée", "Clôturée"])
    rate_act = get_resolution_rate(actions, ["Terminée"])
    rates = [r for r, d in ((rate_nc, nc), (rate_rec, reclamations), (rate_act, actions)) if not d.empty]
    score = sum(rates) / len(rates) if rates else 0.0

    kpis = [
        ("nc", "nc", nc, nc_all, "#EF4444"), ("audits", "audits", audits, audits_all, "#10B981"),
        ("reclamations", "reclamations", reclamations, rec_all, "#F59E0B"),
        ("actions", "actions", actions, actions_all, "#8B5CF6"),
        ("fournisseurs", "fournisseurs", fournisseurs, fourn_all, "#3B82F6"),
        ("logistique", "logistique", flux, flux_all, "#14B8A6"),
    ]
    for col, (key, icon, df, full, color) in zip(cols(6), kpis):
        mv = monthly_counts(full)
        with col:
            st.markdown(
                f"""<div class="kpi" style="background:{color}14;border:1px solid {color}40;">
                <div class="ico" style="background:{color}26;">{svg(icon, color, 20)}</div>
                <div class="lbl" style="color:{color};">{t[key]}</div>
                <div class="val">{len(df)}</div>
                <div class="spark">{sparkline(mv, color)}{trend_html(mv, color)}</div></div>""",
                unsafe_allow_html=True,
            )

    pending = []
    if not actions_all.empty and {"Statut", "Action"} <= set(actions_all.columns):
        for _, row in actions_all[actions_all["Statut"] != "Terminée"].iterrows():
            pending.append(f"{L('Action en cours', 'Action in progress', 'إجراء جارٍ')} : {html.escape(str(row['Action']))} "
                           f"({L('Responsable', 'Owner', 'المسؤول')}: {html.escape(str(row.get('Responsable', 'N/A')))})")
    if pending:
        items = "".join(f"<li>{p}</li>" for p in pending[:3])
        st.markdown(f'<div class="alert bad"><h4>{svg("nc", "#DC2626", 22, 2)}{t["alerts"]}</h4><ul>{items}</ul></div>',
                    unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="alert ok"><h4>{svg("audits", "#059669", 22, 2)}{t["alerts"]}</h4>'
                    f'<ul><li>{L("Aucune alerte urgente : tous les dossiers sont sous contrôle.", "No urgent alert: everything is under control.", "لا توجد تنبيهات عاجلة: كل الملفات تحت السيطرة.")}</li></ul></div>',
                    unsafe_allow_html=True)

    c_score, c_due, c_aud = cols([1, 1.4, 1.2])
    with c_score:
        st.markdown(f'<div class="card"><h3>{L("Score qualité global", "Overall quality score", "مؤشر الجودة العام")}</h3>'
                    f'{score_ring(score)}</div>', unsafe_allow_html=True)
    with c_due:
        st.markdown(f'<div class="card"><h3>{L("Échéances à venir", "Upcoming deadlines", "المواعيد القادمة")}</h3>'
                    f'{deadlines_html(due_df)}</div>', unsafe_allow_html=True)
    with c_aud:
        st.markdown(f'<div class="card"><h3>{L("Résultats des audits", "Audit results", "نتائج التدقيق")}</h3>'
                    f'{audits_html(audits)}</div>', unsafe_allow_html=True)

    st.write("")
    c_sup, c_log = cols([1, 1])
    with c_sup:
        st.markdown(f'<div class="card"><h3>{L("Fournisseurs à risque", "Suppliers at risk", "موردون في خطر")}</h3>'
                    f'{suppliers_html(fourn_all)}</div>', unsafe_allow_html=True)
    with c_log:
        st.markdown(f'<div class="card"><h3>{L("Logistique : entrées / sorties", "Logistics: inbound / outbound", "اللوجستيك: دخول / خروج")}</h3>'
                    f'{logistics_html(flux)}</div>', unsafe_allow_html=True)

    st.write("")
    c_cal, c_iso = cols([1, 1])
    with c_cal:
        st.markdown(f'<div class="card"><h3>{L("Calendrier des audits", "Audit calendar", "جدول التدقيق")}</h3>'
                    f'{audit_calendar_html(audits_all)}</div>', unsafe_allow_html=True)
    with c_iso:
        st.markdown(f'<div class="card"><h3>{L("Indicateurs ISO 9001", "ISO 9001 indicators", "مؤشرات ISO 9001")}</h3>'
                    f'{iso_html(nc_all, rec_all, fourn_all)}</div>', unsafe_allow_html=True)

    st.write("")
    col_recap, col_chart, col_recent = cols([1.55, 1, 0.95])
    with col_recap:
        recap = [
            (mod_name("NonConformites"), len(nc), L(f"{rate_nc}% résolu", f"{rate_nc}% resolved", f"{rate_nc}% محلولة"), "#10B981"),
            (mod_name("Audits"), len(audits), tr("Actif"), "#2563EB"),
            (mod_name("Reclamations"), len(reclamations), L(f"{rate_rec}% traité", f"{rate_rec}% handled", f"{rate_rec}% معالجة"), "#F59E0B"),
            (mod_name("Actions"), len(actions), L(f"{rate_act}% fait", f"{rate_act}% done", f"{rate_act}% منجزة"), "#8B5CF6"),
            (mod_name("Fournisseurs"), len(fournisseurs), L(f"{len(fournisseurs)} évalués", f"{len(fournisseurs)} rated", f"{len(fournisseurs)} مقيّم"), "#2563EB"),
            (mod_name("FluxLogistique"), len(flux), L(f"{len(flux)} mouvements", f"{len(flux)} movements", f"{len(flux)} حركة"), "#10B981"),
        ]
        rows = "".join(f'<tr><td>{m}</td><td>{n}</td><td style="color:{c};">{s}</td></tr>' for m, n, s, c in recap)
        st.markdown(f'<div class="sec-title">{t["summary"]}</div><table class="recap"><tr><th>{tr("Module")}</th>'
                    f'<th>{L("Total Enregistrements", "Total records", "مجموع التسجيلات")}</th>'
                    f'<th>{L("Taux / Statut Global", "Rate / Global status", "النسبة / الحالة العامة")}</th></tr>{rows}</table>',
                    unsafe_allow_html=True)
    with col_chart:
        counts = count_statuses([nc, reclamations, actions])
        biggest = max([c for _, _, c in counts] + [1])
        bars = "".join(f'<div class="bar-col"><div class="bar-n">{c}</div><div class="bar" '
                       f'style="height:{max(c / biggest * 100, 2)}%;background:{col};"></div></div>'
                       for _, col, c in counts)
        names = "".join(f"<span>{html.escape(tr(label))}</span>" for label, _, _ in counts)
        st.markdown(f'<div class="card"><h3>{t["status_chart"]}</h3><div class="bars">{bars}</div>'
                    f'<div class="bar-names">{names}</div></div>', unsafe_allow_html=True)
    with col_recent:
        recent = get_recent_activities()
        lines = "".join(
            f'<div class="act"><span class="dot" style="background:{color};"></span>'
            f'<span>{html.escape(tr(cat))} · {html.escape(detail[:30])}</span><span class="when">{d_}</span></div>'
            for d_, cat, detail, color in recent
        ) if recent else f'<div class="act">{L("Aucune activité récente.", "No recent activity.", "لا توجد أنشطة حديثة.")}</div>'
        st.markdown(f'<div class="card"><h3>{t["recent"]}</h3>{lines}</div>', unsafe_allow_html=True)

# ---------------------------------------------------------
# PAGES DE GESTION (ajout, recherche, vues par statut, modification, suppression)
# ---------------------------------------------------------
CLOSED = {"NonConformites": {"Clôturée"}, "Reclamations": {"Traitée", "Clôturée"}, "Actions": {"Terminée"}}


def make_widget(kind, label, opts, key, default=None):
    label = tr(label)
    blank = default is None or (not isinstance(default, (list, dict)) and pd.isna(default))
    if kind == "date":
        try:
            d = date.today() if blank else pd.to_datetime(default).date()
        except Exception:
            d = date.today()
        return st.date_input(label, value=d, key=key)
    if kind in ("select", "ref_select"):
        return st.selectbox(label, opts, index=opts.index(default) if (not blank and default in opts) else 0,
                            format_func=tr, key=key)
    if kind == "number":
        try:
            v = int(default) if not blank else (opts if isinstance(opts, int) else 10)
        except Exception:
            v = 10
        return st.number_input(label, min_value=1, value=v, key=key)
    return st.text_input(label, value="" if blank else str(default), key=key)


def stamp_closing(row, sheet, old):
    """Renseigne la date de clôture automatiquement quand le statut devient clôturé."""
    if sheet in CLOSED:
        prev = None if old is None else old.get("Date clôture")
        if row.get("Statut") in CLOSED[sheet]:
            row["Date clôture"] = prev if (prev is not None and not pd.isna(prev)) else date.today()
        else:
            row["Date clôture"] = None


def _norm(s):
    """Minuscules + sans accents."""
    s = unicodedata.normalize("NFD", str(s).lower())
    return "".join(c for c in s if unicodedata.category(c) != "Mn")


def _row_text(row):
    parts = []
    for v in row:
        try:
            if pd.isna(v):
                continue
        except (TypeError, ValueError):
            pass
        parts += [str(v), tr(v)]  # cherche aussi dans le texte traduit
    return _norm(" ".join(parts))


def _word_match(qw, text, tokens):
    if qw in text:                      # contient les lettres : "clot" -> "Clôturée"
        return True
    if len(qw) < 3:
        return False
    if any(sorted(tk) == sorted(qw) for tk in tokens):   # mêmes lettres (ordre différent)
        return True
    cutoff = 0.8 if len(qw) <= 4 else 0.7                # fautes de frappe / mots proches
    return bool(difflib.get_close_matches(qw, list(tokens), n=1, cutoff=cutoff))


def fuzzy_filter(df, q):
    """Recherche tolérante : sans accents, sans majuscules, mots approchants."""
    qs = [w for w in re.split(r"\s+", _norm(q).strip()) if w]
    if not qs:
        return df
    mask = []
    for row in df.itertuples(index=False, name=None):
        text = _row_text(row)
        tokens = set(re.findall(r"\w+", text))
        mask.append(all(_word_match(w, text, tokens) for w in qs))
    return df[pd.Series(mask, index=df.index)]


def crud_page(title_key, sheet, form_title, fields, list_title, empty_msg, ok_msg, submit="Enregistrer"):
    role = st.session_state["role"]
    can_edit, can_delete = role != "Lecteur", role == "Administrateur"
    st.title(t[title_key])
    if st.session_state["flash"]:
        st.success(st.session_state["flash"])
        st.session_state["flash"] = ""
    df = load_sheet(sheet)
    prefix = REFS[sheet]

    if can_edit:
        values = {}
        with st.form(f"form_{sheet}"):
            st.subheader(tr(form_title))
            for col, label, kind, opts in fields:
                values[col] = make_widget(kind, label, opts, f"{sheet}_{col}")
            submitted = st.form_submit_button(tr(submit))
        if submitted:
            ref = f"{prefix}-{next_number(df):03d}"
            row = {"Réf": ref, **values}
            stamp_closing(row, sheet, None)
            new = pd.DataFrame([row])
            save_sheet(new if df.empty else pd.concat([df, new], ignore_index=True), sheet)
            log_event(sheet, ref, "Ajout", str(values.get("Description", values.get("Action", values.get("Nom", values.get("Type", ""))))))
            st.session_state["flash"] = f"{tr(ok_msg)} ({ref})"
            rerun()
    else:
        st.info(L("Mode lecture seule : vous ne pouvez pas modifier les données.", "Read-only mode.", "وضع القراءة فقط."))

    st.subheader(tr(list_title))
    if df.empty:
        st.info(tr(empty_msg))
        return

    sc = STATUS_COLS.get(sheet)
    has_sc = sc in df.columns
    gser = df[sc].map(lambda v: VALUE_COLOR.get(str(v))) if has_sc else None
    fcol = "Statut" if "Statut" in df.columns else ("Résultat" if "Résultat" in df.columns else None)

    f1, f2 = cols([2, 1])
    q = f1.text_input("🔍 " + L("Rechercher", "Search", "بحث"), key=f"q_{sheet}")
    sel = ALL
    if fcol:
        sel = f2.selectbox(L("Filtrer", "Filter", "تصفية") + " : " + tr(fcol),
                           [ALL] + sorted(df[fcol].dropna().astype(str).unique().tolist()),
                           format_func=lambda v: tr("Tous") if v == ALL else tr(v), key=f"f_{sheet}")

    # Table VIRTUELLE : vue filtrée par classification (rouge / jaune / vert)
    vsel = "all"
    if has_sc:
        gcount = {k: int((gser == k).sum()) for k, _, _ in GROUPS}

        def _lab(k):
            if k == "all":
                return f"{tr('Tous')} ({len(df)})"
            e, lab = GROUP_META[k]
            return f"{e} {tr(lab)} ({gcount[k]})"

        vsel = radio_h(L("Vue (classification)", "View (classification)", "عرض (تصنيف)"),
                       ["all"] + [k for k, _, _ in GROUPS], format_func=_lab, key=f"v_{sheet}")

    view = df
    if q:
        view = fuzzy_filter(view, q)
    if fcol and sel != ALL:
        view = view[view[fcol].astype(str) == sel]
    if has_sc and vsel != "all":
        view = view[gser.loc[view.index] == vsel]
    if view.empty:
        st.info(L("Aucun résultat.", "No result.", "لا توجد نتائج."))
    else:
        st.markdown(html_table(view, sc if has_sc else None), unsafe_allow_html=True)
    st.caption(f"{len(view)} / {len(df)}")

    if can_edit:
        with expander(L("✏️ Modifier / clôturer / supprimer une ligne", "✏️ Edit / close / delete a row", "✏️ تعديل / إغلاق / حذف سطر")):
            ref = st.selectbox(L("Référence", "Reference", "المرجع"), df["Réf"].astype(str).tolist(), key=f"sel_{sheet}")
            idx = df.index[df["Réf"].astype(str) == ref][0]
            cur = df.loc[idx]
            with st.form(f"edit_{sheet}"):
                newvals = {col: make_widget(kind, label, opts, f"e_{sheet}_{ref}_{col}", cur.get(col))
                           for col, label, kind, opts in fields}
                upd = st.form_submit_button(L("Enregistrer les modifications", "Save changes", "حفظ التعديلات"))
            if upd:
                new_row = {"Réf": ref, **newvals}
                stamp_closing(new_row, sheet, cur)
                for k, v in new_row.items():
                    df.loc[idx, k] = v
                save_sheet(df, sheet)
                log_event(sheet, ref, "Modification", ", ".join(f"{k}={v}" for k, v in newvals.items() if k in ("Statut", "Résultat", "Réalisation")))
                st.session_state["flash"] = L(f"{ref} modifié.", f"{ref} updated.", f"تم تعديل {ref}.")
                rerun()
            if can_delete:
                ok = st.checkbox(L("Confirmer la suppression", "Confirm deletion", "تأكيد الحذف"), key=f"ok_{sheet}_{ref}")
                if st.button("🗑 " + L("Supprimer", "Delete", "حذف"), key=f"del_{sheet}_{ref}"):
                    if ok:
                        save_sheet(df.drop(index=idx), sheet)
                        log_event(sheet, ref, "Suppression", "")
                        st.session_state["flash"] = L(f"{ref} supprimé.", f"{ref} deleted.", f"تم حذف {ref}.")
                        rerun()
                    else:
                        st.warning(L("Cochez d'abord la confirmation.", "Tick the confirmation first.", "ضع علامة التأكيد أولاً."))


if menu == "nc":
    act_refs = ["—"] + load_sheet("Actions").get("Réf", pd.Series(dtype=str)).astype(str).tolist()
    crud_page("nc", "NonConformites", "＋ Ajouter une Non-Conformité", [
        ("Date", "Date", "date", None),
        ("Description", "Description", "text", None),
        ("Service", "Service concerné", "text", None),
        ("Responsable", "Responsable", "text", None),
        ("Gravité", "Gravité", "select", ["Mineure", "Majeure", "Critique"]),
        ("Cause", "Cause (5M)", "select", ["Matière", "Méthode", "Main-d'œuvre", "Machine", "Milieu"]),
        ("Action liée", "Action corrective liée", "ref_select", act_refs),
        ("Statut", "Statut", "select", ["Ouverte", "En cours", "Clôturée"]),
    ], "Liste des Non-Conformités", "Aucune donnée disponible.", "Non-Conformité enregistrée !")

elif menu == "audits":
    crud_page("audits", "Audits", "＋ Ajouter un Audit", [
        ("Date", "Date", "date", None),
        ("Type", "Type d'audit", "text", None),
        ("Auditeur", "Auditeur", "text", None),
        ("Service", "Service audité", "text", None),
        ("Date prévue", "Date prévue (planning)", "date", None),
        ("Réalisation", "Réalisation", "select", ["Planifié", "Réalisé"]),
        ("Résultat", "Résultat", "select", ["Conforme", "Non conforme", "À améliorer"]),
    ], "Liste des Audits", "Aucun audit enregistré.", "Audit enregistré !")

elif menu == "reclamations":
    crud_page("reclamations", "Reclamations", "＋ Ajouter une Réclamation", [
        ("Date", "Date", "date", None),
        ("Client", "Client", "text", None),
        ("Description", "Description", "text", None),
        ("Délai", "Délai cible (jours)", "number", 7),
        ("Statut", "Statut", "select", ["Nouvelle", "En cours", "Traitée", "Clôturée"]),
    ], "Liste des Réclamations", "Aucune réclamation enregistrée.", "Réclamation ajoutée !")

elif menu == "actions":
    crud_page("actions", "Actions", "＋ Ajouter une Action", [
        ("Date", "Date", "date", None),
        ("Action", "Action", "text", None),
        ("Responsable", "Responsable", "text", None),
        ("Échéance", "Échéance", "date", None),
        ("Statut", "Statut", "select", ["À faire", "En cours", "Terminée"]),
    ], "Plan d'Actions", "Aucune action enregistrée.", "Action enregistrée !")

elif menu == "fournisseurs":
    crud_page("fournisseurs", "Fournisseurs", "＋ Référencer un Fournisseur / Évaluation", [
        ("Date", "Date d'évaluation", "date", None),
        ("Nom", "Nom du Fournisseur", "text", None),
        ("Catégorie", "Matière / Service fourni", "text", None),
        ("Évaluation", "Évaluation Qualité", "select",
         ["A - Excellent", "B - Conforme", "C - À surveiller", "D - Critique"]),
        ("Statut", "Statut Partenariat", "select", ["Actif", "Sous observation", "Suspendu"]),
    ], "Annuaire et Évaluation des Fournisseurs", "Aucun fournisseur enregistré.",
        "Fournisseur enregistré !", "Enregistrer le Fournisseur")

elif menu == "logistique":
    crud_page("logistique", "FluxLogistique", "＋ Enregistrer un Mouvement de Stock", [
        ("Date", "Date du mouvement", "date", None),
        ("Type de Flux", "Type de Flux", "select",
         ["Entrée (Réception Marchandise)", "Sortie (Expédition / Vente)", "Retour Interne"]),
        ("Article", "Référence / Nom de l'Article", "text", None),
        ("Quantité", "Quantité", "number", None),
        ("Tiers", "Fournisseur (si entrée) ou Client (si sortie)", "text", None),
        ("Contrôle", "Contrôle Qualité Réception/Expédition", "select",
         ["Validé / Conforme", "Réservé / Non-conforme"]),
    ], "Historique des Flux Logistiques", "Aucun flux enregistré.",
        "Mouvement enregistré !", "Valider le Mouvement")

# ---------------------------------------------------------
# PAGE : VUES PAR STATUT (table virtuelle tous modules)
# ---------------------------------------------------------
elif menu == "views":
    st.title(t["views"])
    st.caption(L("Table virtuelle regroupant tous les modules, classée par avancement : pas encore commencé, en cours, terminé.",
                 "Virtual table gathering all modules, sorted by progress: not started, in progress, done.",
                 "جدول افتراضي يجمع كل الوحدات، مصنف حسب التقدم: لم يبدأ، قيد التنفيذ، منجز."))
    vv = virtual_view()
    if vv.empty:
        st.info(tr("Aucune donnée disponible."))
    else:
        mods = st.multiselect(L("Modules", "Modules", "الوحدات"), list(REFS), default=list(REFS),
                              format_func=mod_name, key="views_mods")
        vv = vv[vv["Module"].isin(mods)]
        g = vv["Statut"].map(lambda v: VALUE_COLOR.get(str(v)))
        for col, (k, e, lab) in zip(cols(3), GROUPS):
            bg_, _ = COLORS[k]
            col.markdown(f'<div class="kpi" style="background:{bg_}18;border:1px solid {bg_}55;border-left:6px solid {bg_};'
                         f'min-height:100px;"><div class="lbl">{e} {html.escape(tr(lab))}</div>'
                         f'<div class="val">{int((g == k).sum())}</div></div>', unsafe_allow_html=True)
        st.write("")
        vsel = radio_h(L("Vue (classification)", "View (classification)", "عرض (تصنيف)"),
                       ["all"] + [k for k, _, _ in GROUPS],
                       format_func=lambda k: tr("Tous") if k == "all" else f"{GROUP_META[k][0]} {tr(GROUP_META[k][1])}",
                       key="views_group")
        for k, e, lab in GROUPS:
            if vsel not in ("all", k):
                continue
            sub = vv[g == k]
            st.markdown(f'<div class="sec-title">{e} {html.escape(tr(lab))} ({len(sub)})</div>', unsafe_allow_html=True)
            if sub.empty:
                st.caption("—")
            else:
                st.markdown(html_table(sub, "Statut"), unsafe_allow_html=True)

# ---------------------------------------------------------
# PAGE : ÉVALUATION (critères pour chaque table)
# ---------------------------------------------------------
elif menu == "evaluation":
    st.title(t["evaluation"])
    st.caption(L("Chaque module est évalué automatiquement selon des critères et des objectifs. Vert = atteint, orange = proche, rouge = non atteint.",
                 "Each module is automatically evaluated against criteria and targets. Green = met, orange = close, red = not met.",
                 "يتم تقييم كل وحدة تلقائيًا حسب معايير وأهداف. أخضر = محقق، برتقالي = قريب، أحمر = غير محقق."))
    crit = build_criteria()
    cards, summary, mscores = [], "", []
    for sh, items in crit.items():
        rows, scores = "", []
        for lab, v, unit, target, higher in items:
            sc_, key = crit_eval(v, target, higher)
            if sc_ is not None:
                scores.append(sc_)
            tgt = ("≥ " if higher else "≤ ") + fmt_val(target, unit)
            rows += (f"<tr><td>{html.escape(tr(lab))}</td><td><b>{fmt_val(v, unit)}</b></td><td>{tgt}</td>"
                     f"<td>{chip(tr(CRIT_LABEL[key]), key)}</td></tr>")
        msc = sum(scores) / len(scores) if scores else None
        if msc is not None:
            mscores.append(msc)
        mk = score_key(msc)
        badge = chip("—" if msc is None else f"{msc:.0f} %", mk)
        summary += f'<div class="dl"><span>{html.escape(mod_name(sh))}</span><span style="margin-left:auto">{badge}</span></div>'
        cards.append(f'<div class="card"><h3>{html.escape(mod_name(sh))}{badge}</h3><table class="dt"><tr>'
                     f'<th>{tr("Critère")}</th><th>{tr("Valeur")}</th><th>{tr("Objectif")}</th><th>{tr("Résultat")}</th></tr>{rows}</table></div>')
    overall = sum(mscores) / len(mscores) if mscores else 0.0
    c1, c2 = cols([1, 1.6])
    with c1:
        st.markdown(f'<div class="card"><h3>{L("Score global d’évaluation", "Overall evaluation score", "نتيجة التقييم العام")}</h3>'
                    f'{score_ring(overall)}</div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="card"><h3>{L("Score par module", "Score per module", "النتيجة حسب الوحدة")}</h3>{summary}</div>',
                    unsafe_allow_html=True)
    st.write("")
    for i in range(0, len(cards), 2):
        for col, card_html in zip(cols(2), cards[i:i + 2]):
            col.markdown(card_html, unsafe_allow_html=True)
        st.write("")

elif menu == "stats":
    st.title(t["stats"])
    nc_s, rec_s, fou_s = load_sheet("NonConformites"), load_sheet("Reclamations"), load_sheet("Fournisseurs")
    st.markdown(f'<div class="card"><h3>{L("Indicateurs ISO 9001", "ISO 9001 indicators", "مؤشرات ISO 9001")}</h3>'
                f'{iso_html(nc_s, rec_s, fou_s)}</div>', unsafe_allow_html=True)
    st.write("")
    pareto = pareto_svg(nc_s)
    st.markdown(f'<div class="card"><h3>{L("Pareto des causes de non-conformité", "Pareto of non-conformity causes", "باريتو لأسباب عدم المطابقة")}</h3>'
                + (pareto if pareto else f'<div class="dl">{L("Ajoutez des non-conformités avec une cause.", "Add non-conformities with a cause.", "أضف عدم مطابقات مع السبب.")}</div>')
                + "</div>", unsafe_allow_html=True)
    sheets = ["NonConformites", "Audits", "Reclamations", "Actions", "Fournisseurs", "FluxLogistique"]
    stats = pd.DataFrame({"Catégorie": [mod_name(s) for s in sheets], "Nombre": [len(load_sheet(s)) for s in sheets]})
    st.subheader(L("Statistiques Générales & Opérationnelles", "General & Operational Statistics", "إحصائيات عامة وتشغيلية"))
    st.bar_chart(stats.set_index("Catégorie"))
    st.subheader(L("Tableau Récapitulatif", "Summary Table", "جدول ملخص"))
    st.markdown(html_table(stats), unsafe_allow_html=True)
    st.subheader(L("Exporter le rapport", "Export report", "تصدير التقرير"))
    offer_download("📥 Excel (.xlsx)", excel_bytes(), "rapport_smq.xlsx",
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    offer_download("📄 " + L("Rapport imprimable (PDF via Ctrl+P)", "Printable report (PDF via Ctrl+P)", "تقرير للطباعة (PDF عبر Ctrl+P)"),
                   report_html().encode("utf-8"), "rapport_smq.html", "text/html")

elif menu == "history":
    st.title(t["history"])
    h = load_sheet("Historique")
    if h.empty:
        st.info(L("Aucune modification enregistrée.", "No change recorded.", "لا توجد تعديلات مسجلة."))
    else:
        st.markdown(html_table(h.iloc[::-1].head(200)), unsafe_allow_html=True)
