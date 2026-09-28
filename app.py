import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, time
import io
import urllib.request
import re
import collections
import streamlit.components.v1 as components
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from openpyxl.worksheet.pagebreak import Break
from openpyxl.utils import get_column_letter

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(page_title="Générateur Officiel FFLDA", page_icon="🤼", layout="wide")

# --- SYSTÈME D'ACCÈS TEMPORAIRES (GOOGLE SHEETS / CODES D'ACCÈS) ---
if "authentifie" not in st.session_state:
    st.session_state["authentifie"] = False

try:
    URL_GOOGLE_SHEETS_DEFAUT = st.secrets.get("GSHEETS_CODES_URL", "https://docs.google.com/spreadsheets/d/1VQ_L4oy_587wFpbEnuAeehimbV8ax813m3_YQq1Uzuk/edit?usp=sharing")
except Exception:
    URL_GOOGLE_SHEETS_DEFAUT = "https://docs.google.com/spreadsheets/d/1VQ_L4oy_587wFpbEnuAeehimbV8ax813m3_YQq1Uzuk/edit?usp=sharing"

def verifier_code_acces(code_saisi, url_csv):
    if not code_saisi or not str(code_saisi).strip():
        return False, "Veuillez saisir un code d'accès."
    code_clean = str(code_saisi).strip().upper()
    
    if code_clean in ["FFLDA-ADMIN", "FFLDA2026"]:
        return True, "✨ Accès administrateur déverrouillé !"
        
    if not url_csv or "http" not in url_csv:
        return True, "Code validé (mode démo sans tableau Google Sheets)."

    if "docs.google.com/spreadsheets" in url_csv:
        if "/edit" in url_csv:
            url_csv = re.sub(r'/edit.*$', '/export?format=csv', url_csv)
        elif not ("output=csv" in url_csv or "format=csv" in url_csv):
            url_csv = url_csv.rstrip("/") + "/export?format=csv"
        
    try:
        sep = "&" if "?" in url_csv else "?"
        url_fresh = f"{url_csv}{sep}_cb={int(datetime.now().timestamp())}"
        
        df_codes = pd.read_csv(url_fresh)
        df_codes.columns = [str(c).strip().lower() for c in df_codes.columns]
        
        col_code = next((c for c in df_codes.columns if 'code' in c), None)
        col_exp = next((c for c in df_codes.columns if 'expir' in c or 'date' in c), None)
        
        if not col_code or not col_exp:
            return False, "⚠️ Le tableau Google Sheets doit contenir au moins les colonnes 'Code' et 'Expiration'."
            
        df_match = df_codes[df_codes[col_code].astype(str).str.strip().str.upper() == code_clean]
        
        if df_match.empty:
            return False, "❌ Code d'accès invalide. Vérifiez la saisie ou contactez la FFLDA."
            
        date_exp_str = str(df_match.iloc[0][col_exp]).strip()
        date_clean = re.sub(r'(\d+)\s*[hH]\s*(\d*)', lambda m: f"{m.group(1)}:{m.group(2) if m.group(2) else '00'}", date_exp_str)
        
        try:
            date_exp = pd.to_datetime(date_clean, dayfirst=True)
        except Exception:
            date_exp = pd.to_datetime(date_clean)
            
        if hasattr(date_exp, 'to_pydatetime'):
            date_exp = date_exp.to_pydatetime()

        try:
            from zoneinfo import ZoneInfo
            now_fr = datetime.now(ZoneInfo("Europe/Paris")).replace(tzinfo=None)
        except Exception:
            now_fr = datetime.utcnow() + timedelta(hours=2)
            
        if now_fr > date_exp:
            return False, f"❌ Ce code d'accès a expiré le {date_exp.strftime('%d/%m/%Y à %H:%M')} (Heure actuelle en France : {now_fr.strftime('%H:%M')})."
            
        return True, "✨ Accès autorisé !"
    except Exception as e:
        return False, f"⚠️ Erreur lors de la vérification du code : {e}"

URL_BILLING_WEBHOOK_DEFAUT = "https://script.google.com/macros/s/AKfycbxboXVY0FbLYQX6ZeBgDt4lg2fJ6eDIxfW-1BswZRoz5sLqAqBLCUGk7sLHoqpMW_C0/exec"

def enregistrer_log_facturation(code_organisateur, nom_tournoi, nb_inscrits, nb_peses, nb_matchs, details_resultats=None):
    try:
        import json
        url_webhook = None
        try:
            url_webhook = st.secrets.get("BILLING_WEBHOOK_URL", URL_BILLING_WEBHOOK_DEFAUT)
        except Exception:
            pass
        if not url_webhook:
            url_webhook = st.session_state.get("url_billing_webhook", URL_BILLING_WEBHOOK_DEFAUT)
            
        if not url_webhook or "http" not in url_webhook:
            return
            
        try:
            from zoneinfo import ZoneInfo
            now_str = datetime.now(ZoneInfo("Europe/Paris")).strftime("%d/%m/%Y %H:%M:%S")
        except Exception:
            now_str = (datetime.utcnow() + timedelta(hours=2)).strftime("%d/%m/%Y %H:%M:%S")
            
        payload = {
            "code_organisateur": str(code_organisateur or "ANONYME"),
            "nom_tournoi": str(nom_tournoi or "Tournoi sans nom"),
            "date": now_str,
            "total_inscrits": int(nb_inscrits or 0),
            "total_peses": int(nb_peses or 0),
            "total_matchs": int(nb_matchs or 0),
            "details_resultats": json.dumps(details_resultats or [], ensure_ascii=False)
        }
        
        data = urllib.parse.urlencode(payload).encode('utf-8')
        req = urllib.request.Request(url_webhook, data=data, headers={'User-Agent': 'FFLDA-Billing/1.0'})
        with urllib.request.urlopen(req, timeout=3.0):
            pass
    except Exception:
        pass

if not st.session_state["authentifie"]:
    st.markdown("
