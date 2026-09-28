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

    if "[docs.google.com/spreadsheets](https://docs.google.com/spreadsheets)" in url_csv:
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
    st.markdown("<br><br>", unsafe_allow_html=True)
    col_acc1, col_acc2, col_acc3 = st.columns([1, 2, 1])
    with col_acc2:
        import os
        if os.path.exists("logo_fflda.png"):
            st.image("logo_fflda.png", use_container_width=True)
        else:
            st.image("https://www.fflutte.com/content/uploads/2021/10/fflutte-bleu-1024x842.png", use_container_width=True)
        
        st.markdown("<h2 style='text-align: center; color: #0055A4;'>🔒 Espace Sécurisé Organisateur FFLDA</h2>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: #555;'>Veuillez saisir votre code d'accès temporaire pour déverrouiller l'application.</p>", unsafe_allow_html=True)
        st.markdown("---")
        
        code_saisi = st.text_input("🔑 Code d'accès", type="password", placeholder="Ex: PARIS-24H")
        
        if st.button("🚀 Se Connecter", use_container_width=True):
            valide, message = verifier_code_acces(code_saisi, URL_GOOGLE_SHEETS_DEFAUT)
            if valide:
                st.session_state["authentifie"] = True
                st.session_state["code_session"] = code_saisi.strip().upper()
                st.success(message)
                st.rerun()
            else:
                st.error(message)
                
        st.markdown("---")
        st.caption("Fédération Française de Lutte et Disciplines Associées — Plateforme Officielle de Gestion de Tournois")
    st.stop()

# --- VÉRIFICATION CONTINUELLE EN TEMPS RÉEL DE L'EXPIRATION ---
if st.session_state.get("authentifie"):
    code_sess = st.session_state.get("code_session", "")
    if code_sess not in ["FFLDA-ADMIN", "FFLDA2026"]:
        valide_encore, msg_exp = verifier_code_acces(code_sess, URL_GOOGLE_SHEETS_DEFAUT)
        if not valide_encore:
            st.session_state["authentifie"] = False
            st.error(f"⏰ {msg_exp}")
            st.stop()

# --- MENU LATÉRAL ---
with st.sidebar:
    import os
    col_l1, col_l2, col_l3 = st.columns([1, 2, 1])
    with col_l2:
        if os.path.exists("logo_fflda.png"):
            st.image("logo_fflda.png", use_container_width=True)
        else:
            st.image("https://www.fflutte.com/content/uploads/2021/10/fflutte-bleu-1024x842.png", use_container_width=True)
    if st.session_state.get("authentifie"):
        col_s1, col_s2 = st.columns([3, 1])
        with col_s1:
            st.caption("🟢 **Session active**")
        with col_s2:
            if st.button("🔒", help="Se déconnecter de l'application"):
                st.session_state["authentifie"] = False
                st.rerun()

    st.markdown("### Paramètres du tournoi")
    st.markdown("---")
    nom_competition = st.text_input("🏆 Nom de la compétition", value="Tournoi Officiel FFLDA - U7/U9/U11/U13")
    
    with st.expander("⚙️ 1. Logistique & Pesées", expanded=True):
        nb_tapis = st.number_input("Nombre de tapis", min_value=1, max_value=10, value=3)
        type_pesee = st.radio("Format des pesées", ["1 Pesée (Générale)", "2 Pesées (U7/U9 puis U11/U13)"], index=1)
        label_pesee_1 = "1ère pesée" if "1" in type_pesee else "Pesée U7/U9"
        heure_pesee_u9 = st.time_input(label_pesee_1, value=time(9, 0))
        duree_pesee = st.selectbox("Durée allouée à la pesée + échauffement (min)", [30, 45, 60, 90], index=1)
        
    with st.expander("⏱️ 2. Pause de la compétition", expanded=False):
        activer_pause = st.checkbox("Activer la pause", value=True)
        duree_pause = st.selectbox("Durée de la pause (min)", [30, 45, 60, 75, 90], index=2) if activer_pause else 0
    
    with st.expander("🤼 3. Règles Sportives & Temps", expanded=False):
        mixte_active = st.checkbox("Catégories Mixtes (U7, U9, U11 uniquement)", value=True)
        poules_par_niveau = st.checkbox("Créer des poules par niveau (débutants/confirmés)", value=True)
        separer_clubs = st.checkbox("Éviter les lutteurs d'un même club dans la même poule", value=True)
        eviter_arbitre_meme_club = st.checkbox("Éviter les matchs entre arbitres et lutteurs du même club", value=True)
        meme_tapis_poule = st.checkbox("Maintenir chaque poule / lutteur sur un même tapis", value=True)
        tolerance_poids = st.number_input("Tolérance d'écart de poids (%) [U7, U9, U11]", min_value=10, max_value=15, value=10, step=1)
        repos_matchs = st.number_input("Matchs de repos minimum", min_value=1, max_value=10, value=3)
        duree_plateau_u7 = st.number_input("Temps de chaque plateau U7 (min)", min_value=3, max_value=30, value=10, step=1)
        duree_plateau_totale = duree_plateau_u7 * 3
        duree_u9 = st.number_input("Temps total U9 (min)", value=3)
        duree_u11 = st.number_input("Temps total U11 (min)", value=4)
        duree_u13 = st.number_input("Temps total U13 (min)", value=5)

def formater_poids(val):
    if val is None or str(val).strip() in ['', 'None', 'nan']:
        return ''
    val_str = str(val).lower().replace('kg', '').replace(',', '.').strip()
    try:
        val_float = round(float(val_str), 1)
        return f"{int(val_float)} kg" if val_float == int(val_float) else f"{val_float} kg"
    except (ValueError, TypeError):
        s = str(val).strip()
        return f"{s} kg" if (s and not s.lower().endswith('kg')) else s

def formater_poids_court(val):
    if val is None or str(val).strip() in ['', 'None', 'nan']:
        return ''
    val_str = str(val).lower().replace('kg', '').replace(',', '.').strip()
    try:
        val_float = round(float(val_str), 1)
        return f"{int(val_float)}kg" if val_float == int(val_float) else f"{val_float}kg"
    except (ValueError, TypeError):
        s = str(val).strip()
        return f"{s}kg" if (s and not s.lower().endswith('kg')) else s

def abreger_nom_onglet(nom_poule):
    txt = str(nom_poule).replace("Mixte (LL/LF)", "Mxt").replace("LG (Gréco)", "LG").replace("LL (Libre)", "LL").replace("LF (Féminine)", "LF")
    txt = txt.replace(" | ", " ").replace(" (", " ").replace(")", "").replace(" - ", "-")
    txt = txt.replace("/", "-").replace("\\", "-").replace(":", "-").replace("?", "").replace("*", "")
    return re.sub(r'\s+', ' ', txt)[:31].strip()

def nettoyer_nom_tour(val):
    if not val:
        return ""
    val_str = str(val).strip()
    return val_str[5:].strip() if val_str.lower().startswith("tour ") else val_str

def charger_liste_arbitres(fichier_arbitres_in=None):
    arbitres = []
    filepath = fichier_arbitres_in
    if filepath is None:
        import os
        default_name = 'FFLDA - Inscription arbitres.xlsx'
        if os.path.exists(default_name):
            filepath = default_name
            
    if filepath is None:
        return arbitres
        
    try:
        if hasattr(filepath, 'name') and filepath.name.endswith('.csv'):
            df_arb = pd.read_csv(filepath, sep=';', encoding='utf-8')
            if len(df_arb.columns) == 1:
                filepath.seek(0)
                df_arb = pd.read_csv(filepath, sep=',', encoding='utf-8')
        elif isinstance(filepath, str) and filepath.endswith('.csv'):
            df_arb = pd.read_csv(filepath, sep=';', encoding='utf-8')
            if len(df_arb.columns) == 1:
                df_arb = pd.read_csv(filepath, sep=',', encoding='utf-8')
        else:
            df_temp = pd.read_excel(filepath, nrows=5)
            header_row = 0
            for i, row in df_temp.iterrows():
                if 'Licence' in str(row.values) or 'Nom' in str(row.values) or "Inscrit Par" in str(row.values):
                    header_row = i + 1
                    break
            if hasattr(filepath, 'seek'):
                filepath.seek(0)
            df_arb = pd.read_excel(filepath, header=header_row)

        col_nom = next((c for c in df_arb.columns if str(c).strip().lower() == 'nom'), None)
        col_prenom = next((c for c in df_arb.columns if 'prénom' in str(c).strip().lower() or 'prenom' in str(c).strip().lower()), None)
        col_club = next((c for c in df_arb.columns if any(k in str(c).strip().lower() for k in ['sigle du club', 'club'])), None)
        col_comite = next((c for c in df_arb.columns if any(k in str(c).strip().lower() for k in ['comité', 'comite', 'ligue', 'région', 'region'])), None)
        col_licence = next((c for c in df_arb.columns if 'licence' in str(c).strip().lower()), None)

        for _, row in df_arb.iterrows():
            nom_val = str(row[col_nom]).strip() if (col_nom and pd.notna(row[col_nom])) else ''
            prenom_val = str(row[col_prenom]).strip() if (col_prenom and pd.notna(row[col_prenom])) else ''
            if not nom_val or nom_val.lower() in ['nan', 'none', 'photo']:
                continue
            
            club_val = str(row[col_club]).strip() if (col_club and pd.notna(row[col_club])) else 'Indépendant'
            comite_val = str(row[col_comite]).strip() if (col_comite and pd.notna(row[col_comite])) else 'Comité Non Renseigné'
            licence_val = str(row[col_licence]).strip() if (col_licence and pd.notna(row[col_licence])) else ''

            arbitres.append({
                'Nom_Complet': f"{nom_val} {prenom_val}".strip(),
                'Nom': nom_val,
                'Prenom': prenom_val,
                'Licence': licence_val,
                'Club': club_val if club_val not in ['', 'None', 'nan', '-'] else 'Indépendant',
                'Comite': comite_val if comite_val not in ['', 'None', 'nan', '-'] else 'Comité Non Renseigné'
            })
    except Exception:
        pass
        
    return arbitres

st.title(f"🏆 {nom_competition}")
st.markdown("**Plateforme officielle d'optimisation des tournois de jeunes et d'édition des bilans fédéraux.**")
st.markdown("---")

def generer_rondes_fflda(participants_in):
    participants = list(participants_in)
    n = len(participants)
    if n <= 1:
        return []
    elif n == 2:
        return [[(participants[0], participants[1])]]
    elif n == 3:
        return [[(participants[0], participants[1])],
                [(participants[2], participants[0])],
                [(participants[1], participants[2])]]
    elif n == 4:
        return [[(participants[0], participants[1]), (participants[2], participants[3])],
                [(participants[0], participants[2]), (participants[1], participants[3])],
                [(participants[0], participants[3]), (participants[1], participants[2])]]
    elif n == 5:
        return [[(participants[0], participants[1]), (participants[2], participants[3])],
                [(participants[4], participants[0]), (participants[1], participants[2])],
                [(participants[3], participants[4]), (participants[0], participants[2])],
                [(participants[1], participants[3]), (participants[2], participants[4])],
                [(participants[0], participants[3]), (participants[1], participants[4])]]
    else:
        if len(participants) % 2 != 0:
            participants.append({"Nom": "BYE", "Club": "-"})
        num_p = len(participants)
        rondes = []
        for i in range(num_p - 1):
            matchs_ronde = []
            for j in range(num_p // 2):
                p1 = participants[j]
                p2 = participants[num_p - 1 - j]
                if p1["Nom"] != "BYE" and p2["Nom"] != "BYE":
                    matchs_ronde.append((p1, p2))
            rondes.append(matchs_ronde)
            participants.insert(1, participants.pop())
        return rondes

def attribuer_categorie_poids_u13(poids_val):
    try:
        p = float(poids_val) if (poids_val is not None and str(poids_val).strip() != '') else 0.0
    except (ValueError, TypeError):
        p = 0.0
    if p <= 30.0: return "30 kg"
    elif p <= 33.0: return "33 kg"
    elif p <= 36.0: return "36 kg"
    elif p <= 39.0: return "39 kg"
    elif p <= 42.0: return "42 kg"
    elif p <= 46.0: return "46 kg"
    elif p <= 50.0: return "50 kg"
    elif p <= 55.0: return "55 kg"
    elif p <= 60.0: return "60 kg"
    else: return "+60 kg"

ORDRE_POIDS_U13 = ["30 kg", "33 kg", "36 kg", "39 kg", "42 kg", "46 kg", "50 kg", "55 kg", "60 kg", "+60 kg"]

def interleave_bracket_slots(list_a, list_b):
    res = []
    i, j = 0, 0
    while i < len(list_a) or j < len(list_b):
        if i < len(list_a):
            res.append(list_a[i])
            i += 1
        if j < len(list_b):
            res.append(list_b[j])
            j += 1
    return res

def repartir_tableau_protection_clubs(participants, nb_byes, nb_prelim):
    clubs = collections.defaultdict(list)
    for p in participants:
        c = str(p.get('Club', '') or '').strip()
        if not c or c in ['-', 'Comité Non Renseigné', 'Sans club']:
            clubs[f"_indiv_{id(p)}"].append(p)
        else:
            clubs[c].append(p)
            
    sorted_clubs = sorted(clubs.values(), key=len, reverse=True)
    byes_list = []
    prelim_list = []
    
    if nb_byes == 0:
        prelim_list = list(participants)
    else:
        club_byes_count = {i: 0 for i in range(len(sorted_clubs))}
        total_byes_assigned = 0
        
        for i, c_members in enumerate(sorted_clubs):
            surplus = len(c_members) - nb_prelim
            if surplus > 0:
                take = min(surplus, nb_byes - total_byes_assigned)
                club_byes_count[i] += take
                total_byes_assigned += take
                
        while total_byes_assigned < nb_byes:
            progress = False
            for i, c_members in enumerate(sorted_clubs):
                if total_byes_assigned >= nb_byes:
                    break
                if len(c_members) >= 2 and club_byes_count[i] < len(c_members):
                    club_byes_count[i] += 1
                    total_byes_assigned += 1
                    progress = True
            if not progress:
                for i, c_members in enumerate(sorted_clubs):
                    if total_byes_assigned >= nb_byes:
                        break
                    if club_byes_count[i] < len(c_members):
                        club_byes_count[i] += 1
                        total_byes_assigned += 1
                        progress = True
            if not progress:
                break
                
        for i, c_members in enumerate(sorted_clubs):
            n_b = club_byes_count[i]
            byes_list.extend(c_members[:n_b])
            prelim_list.extend(c_members[n_b:])
            
    prelim_matches = []
    if nb_prelim > 0:
        club_groups = collections.defaultdict(list)
        for p in prelim_list:
            c = str(p.get('Club', '') or '').strip()
            club_groups[c].append(p)
        sorted_prelim_clubs = sorted(club_groups.values(), key=len, reverse=True)
        flattened = [p for grp in sorted_prelim_clubs for p in grp]
        
        M = nb_prelim
        half1 = flattened[:M]
        half2 = flattened[M:]
        
        pairs = []
        for i in range(M):
            p1 = half1[i]
            p2 = half2[i]
            if p1.get('Club') and p1.get('Club') == p2.get('Club') and p1.get('Club') not in ['-', '']:
                for j in range(M):
                    if j != i and half2[j].get('Club') != p1.get('Club') and half2[i].get('Club') != half1[j].get('Club'):
                        half2[i], half2[j] = half2[j], half2[i]
                        p2 = half2[i]
                        break
            pairs.append((p1, p2))
        prelim_matches = pairs

    return byes_list, prelim_matches

def generer_competition_u13(age, style_grp, suffixe_niveau, cat_poids, participants, separer_clubs=True):
    n = len(participants)
    if n == 0:
        return None
    elif n == 1:
        nom = f"{age} | {style_grp}{suffixe_niveau} | {cat_poids} (1 seul inscrit)"
        return {'nom': nom, 'participants': list(participants), 'rondes': [], 'type_formule': 'seul', 'cat_poids': cat_poids, 'style_grp': style_grp}
    elif n < 6:
        nom = f"{age} | {style_grp}{suffixe_niveau} | {cat_poids} (Poule unique)"
        return {'nom': nom, 'participants': list(participants), 'rondes': generer_rondes_fflda(participants), 'type_formule': 'poule', 'cat_poids': cat_poids, 'style_grp': style_grp}
    elif n == 6:
        if separer_clubs:
            clubs_vu = {}
            poule_a, poule_b = [], []
            for p in sorted(participants, key=lambda x: x.get('Club', '')):
                c = p.get('Club', '')
                if clubs_vu.get(c, 0) % 2 == 0:
                    if len(poule_a) < 3: poule_a.append(p)
                    else: poule_b.append(p)
                else:
                    if len(poule_b) < 3: poule_b.append(p)
                    else: poule_a.append(p)
                clubs_vu[c] = clubs_vu.get(c, 0) + 1
        else:
            poule_a = [participants[0], participants[2], participants[4]]
            poule_b = [participants[1], participants[3], participants[5]]
        
        while len(poule_a) < 3 and len(poule_b) > 3: poule_a.append(poule_b.pop())
        while len(poule_b) < 3 and len(poule_a) > 3: poule_b.append(poule_a.pop())

        rondes_a = generer_rondes_fflda(poule_a)
        rondes_b = generer_rondes_fflda(poule_b)
        
        r1 = (rondes_a[0] if len(rondes_a) > 0 else []) + (rondes_b[0] if len(rondes_b) > 0 else [])
        r2 = (rondes_a[1] if len(rondes_a) > 1 else []) + (rondes_b[1] if len(rondes_b) > 1 else [])
        r3 = (rondes_a[2] if len(rondes_a) > 2 else []) + (rondes_b[2] if len(rondes_b) > 2 else [])
        
        sf1 = ({"Nom": f"1er Poule A ({cat_poids})", "Club": "Qualifié A", "Comité": "-"}, {"Nom": f"2ème Poule B ({cat_poids})", "Club": "Qualifié B", "Comité": "-"})
        sf2 = ({"Nom": f"1er Poule B ({cat_poids})", "Club": "Qualifié B", "Comité": "-"}, {"Nom": f"2ème Poule A ({cat_poids})", "Club": "Qualifié A", "Comité": "-"})
        r4 = [sf1, sf2]
        
        f_or = ({"Nom": f"Vainqueur 1/2 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}, {"Nom": f"Vainqueur 1/2 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"})
        f_bronze = ({"Nom": f"Perdant 1/2 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}, {"Nom": f"Perdant 1/2 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"})
        r5 = [f_or, f_bronze]
        
        nom = f"{age} | {style_grp}{suffixe_niveau} | {cat_poids} (2 Poules + Finales)"
        return {'nom': nom, 'participants': list(participants), 'rondes': [r1, r2, r3, r4, r5], 'type_formule': 'poules_croisees', 'poule_a': poule_a, 'poule_b': poule_b, 'cat_poids': cat_poids, 'style_grp': style_grp}
    else:
        rondes = []
        if n == 7:
            if separer_clubs:
                clubs = collections.defaultdict(list)
                for p in participants:
                    c = str(p.get('Club', '') or '').strip()
                    if not c or c in ['-', 'Comité Non Renseigné', 'Sans club']:
                        clubs[f"_indiv_{id(p)}"].append(p)
                    else:
                        clubs[c].append(p)
                sorted_clubs = sorted(clubs.values(), key=len, reverse=True)
                p_exempt = sorted_clubs[0][0]
                reste = []
                first_skipped = False
                for grp in sorted_clubs:
                    for p in grp:
                        if not first_skipped and p is p_exempt:
                            first_skipped = True
                        else:
                            reste.append(p)
                club_groups = collections.defaultdict(list)
                for p in reste:
                    c = str(p.get('Club', '') or '').strip()
                    club_groups[c].append(p)
                sorted_reste_clubs = sorted(club_groups.values(), key=len, reverse=True)
                flattened = [p for grp in sorted_reste_clubs for p in grp]
                h1 = flattened[:3]
                h2 = flattened[3:]
                pairs = []
                for i in range(3):
                    p1 = h1[i]
                    p2 = h2[i]
                    if p1.get('Club') and p1.get('Club') == p2.get('Club') and p1.get('Club') not in ['-', '']:
                        for j in range(3):
                            if j != i and half2[j].get('Club') != p1.get('Club') and half2[i].get('Club') != half1[j].get('Club'):
                                half2[i], half2[j] = half2[j], half2[i]
                                p2 = half2[i]
                                break
                    pairs.append((p1, p2))
                q1, q2, q3 = pairs[0], pairs[1], pairs[2]
            else:
                q1, q2, q3 = (participants[0], participants[1]), (participants[2], participants[3]), (participants[4], participants[5])
                p_exempt = participants[6]

            rondes.append([q1, q2, q3])
            sf1 = ({"Nom": f"Vainqueur 1/4 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}, {"Nom": f"Vainqueur 1/4 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"})
            sf2 = ({"Nom": f"Vainqueur 1/4 (3) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}, {"Nom": p_exempt['Nom'], "Club": p_exempt.get('Club', ''), "Comité": p_exempt.get('Comité', '-')})
            rep1 = ({"Nom": f"Perdant 1/4 (1) [{cat_poids}]", "Club": "Repêché", "Comité": "-"}, {"Nom": f"Perdant 1/4 (2) [{cat_poids}]", "Club": "Repêché", "Comité": "-"})
            rondes.append([sf1, sf2, rep1])
            f_or = ({"Nom": f"Vainqueur 1/2 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}, {"Nom": f"Vainqueur 1/2 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"})
            f_b1 = ({"Nom": f"Vainqueur Repêchage 1 [{cat_poids}]", "Club": "Repêché", "Comité": "-"}, {"Nom": f"Perdant 1/2 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"})
            f_b2 = ({"Nom": f"Perdant 1/4 (3) [{cat_poids}]", "Club": "Repêché", "Comité": "-"}, {"Nom": f"Perdant 1/2 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"})
            rondes.append([f_or, f_b1, f_b2])

        elif n <= 16:
            nb_prelim = n - 8
            nb_byes = 16 - n
            if separer_clubs:
                byes, prelim_matches = repartir_tableau_protection_clubs(participants, nb_byes, nb_prelim)
                prelim_winners = [{"Nom": f"Vainqueur Prél. {i+1} [{cat_poids}]", "Club": "Qualifié", "Comité": "-"} for i in range(nb_prelim)]
            else:
                byes = participants[:nb_byes]
                prelim_pts = participants[nb_byes:]
                prelim_matches = [(prelim_pts[i*2], prelim_pts[i*2+1]) for i in range(nb_prelim)]
                prelim_winners = [{"Nom": f"Vainqueur Prél. {i+1} [{cat_poids}]", "Club": "Qualifié", "Comité": "-"} for i in range(nb_prelim)]
                
            if prelim_matches:
                rondes.append(prelim_matches)
                
            slots_qf = [None] * 8
            order_prelim_slots = [7, 3, 5, 1, 6, 2, 4, 0]
            for w_idx, slot_idx in enumerate(order_prelim_slots[:nb_prelim]):
                slots_qf[slot_idx] = prelim_winners[w_idx]
                
            if separer_clubs:
                order_upper = [s for s in [0, 2, 1, 3] if slots_qf[s] is None]
                order_lower = [s for s in [4, 6, 5, 7] if slots_qf[s] is None]
                interleaved_bye_slots = interleave_bracket_slots(order_upper, order_lower)
                for b_idx, s_idx in enumerate(interleaved_bye_slots):
                    if b_idx < len(byes):
                        slots_qf[s_idx] = byes[b_idx]
            else:
                bye_idx = 0
                for s_idx in range(8):
                    if slots_qf[s_idx] is None:
                        slots_qf[s_idx] = byes[bye_idx]
                        bye_idx += 1
                        
            qf_matches = [(slots_qf[0], slots_qf[1]), (slots_qf[2], slots_qf[3]), (slots_qf[4], slots_qf[5]), (slots_qf[6], slots_qf[7])]
            rondes.append(qf_matches)
            
            sf1 = ({"Nom": f"Vainqueur 1/4 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}, {"Nom": f"Vainqueur 1/4 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"})
            sf2 = ({"Nom": f"Vainqueur 1/4 (3) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}, {"Nom": f"Vainqueur 1/4 (4) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"})
            rep1 = ({"Nom": f"Perdant 1/4 (1) [{cat_poids}]", "Club": "Repêché", "Comité": "-"}, {"Nom": f"Perdant 1/4 (2) [{cat_poids}]", "Club": "Repêché", "Comité": "-"})
            rep2 = ({"Nom": f"Perdant 1/4 (3) [{cat_poids}]", "Club": "Repêché", "Comité": "-"}, {"Nom": f"Perdant 1/4 (4) [{cat_poids}]", "Club": "Repêché", "Comité": "-"})
            rondes.append([sf1, sf2, rep1, rep2])
            
            f_or = ({"Nom": f"Vainqueur 1/2 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"}, {"Nom": f"Vainqueur 1/2 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"})
            f_b1 = ({"Nom": f"Vainqueur Repêchage 1 [{cat_poids}]", "Club": "Repêché", "Comité": "-"}, {"Nom": f"Perdant 1/2 (2) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"})
            f_b2 = ({"Nom": f"Vainqueur Repêchage 2 [{cat_poids}]", "Club": "Repêché", "Comité": "-"}, {"Nom": f"Perdant 1/2 (1) [{cat_poids}]", "Club": "Qualifié", "Comité": "-"})
            rondes.append([f_or, f_b1, f_b2])

        nom = f"{age} | {style_grp}{suffixe_niveau} | {cat_poids} (Tableau élimination & repêchages)"
        res = {'nom': nom, 'participants': list(participants), 'rondes': rondes, 'type_formule': 'tableau', 'cat_poids': cat_poids, 'style_grp': style_grp}
        if n == 7: res['p_exempt'] = p_exempt
        return res

# --- LOGIQUE D'EXTRACTION DES RÉSULTATS ET BILANS ---
def nettoyer_rang(val):
    if val is None:
        return "NR"
    s = str(val).strip()
    if not s or s.lower() in ["nr", "none", "nan", "<na>", "nonetype", "-", "0", "en attente"]:
        return "NR"
    try:
        v_num = int(float(s))
        return v_num if v_num > 0 else "NR"
    except (ValueError, TypeError):
        pass
    m = re.search(r'(\d+)', s)
    if m:
        v_num = int(m.group(1))
        return v_num if v_num > 0 else "NR"
    return "NR"

def normaliser_nom_comparaison(txt):
    if not txt:
        return ""
    import unicodedata
    txt = re.sub(r'[🔴🔵🥇🥈🥉🏆🛡️]|\([^\)]*\)', ' ', str(txt))
    txt_nfkd = unicodedata.normalize('NFD', txt)
    txt_sans_accents = "".join([c for c in txt_nfkd if unicodedata.category(c) != 'Mn'])
    return re.sub(r'[^a-z0-9]', '', txt_sans_accents.lower())

def extraire_valeur_score(val):
    if val is None:
        return 0.0
    s = str(val).strip()
    if not s or s.startswith('=') or s.lower() in ['none', 'nan', '-', 'nr']:
        return 0.0
    s_clean = s.replace(',', '.')
    try:
        return float(s_clean)
    except ValueError:
        pass
    m = re.search(r'(-?\d+(?:\.\d+)?)', s_clean)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            pass
    return 0.0

def extraire_resultats_classeur_excel(wb_data, wb_formula=None):
    tous_les_resultats = []
    
    mots_exclus = [
        "résumé", "resume", "grille", "passage", "planning", "programme", 
        "schedule", "recap", "récap", "arbitr"
    ]
    onglets_poules = [
        f for f in wb_data.sheetnames 
        if not any(x in f.lower() for x in mots_exclus) and not any(k in f.lower() for k in ["classement club", "classement comit"])
    ]
    
    onglets_poules.sort(key=lambda x: (
        0 if "u7" in x.lower() else (1 if "u9" in x.lower() else (2 if "u11" in x.lower() else 3)),
        x
    ))
    
    def est_ligne_lutteur(nom_val, club_val):
        if nom_val is None: return False
        nom_s = str(nom_val).strip().lower()
        club_s = str(club_val or "").strip().lower()
        if not nom_s or nom_s in ["none", "nan", "-", "", "0"]: return False
        invalides = ["nom", "nom prénom", "lutteur", "clt", "rang", "n°", "club", "poids", "points", "total", "arbitre"]
        if nom_s in invalides or club_s in invalides: return False
        if any(m in nom_s for m in ["catégorie", "poule", "tournoi", "pause", "pesée", "tapis", "combat n"]): return False
        return True

    for nom_feuille in onglets_poules:
        ws = wb_data[nom_feuille]
        titre_lower = nom_feuille.lower()
        
        # Cas U7 (Plateaux)
        if "u7" in titre_lower or "plateau" in titre_lower:
            col_nom, col_club, col_poids = 2, 3, 4
            h_row = 4
            for r_s in range(1, 6):
                for c_i in range(1, ws.max_column + 1):
                    val = str(ws.cell(row=r_s, column=c_i).value or "").lower()
                    if "nom" in val: col_nom = c_i; h_row = r_s
                    elif "club" in val: col_club = c_i
                    elif "poids" in val: col_poids = c_i
                    
            r = h_row + 1
            while r <= ws.max_row:
                nom = ws.cell(row=r, column=col_nom).value
                club = ws.cell(row=r, column=col_club).value if col_club else "Indépendant"
                poids = ws.cell(row=r, column=col_poids).value if col_poids else ""
                if est_ligne_lutteur(nom, club):
                    tous_les_resultats.append({
                        "Poule": nom_feuille,
                        "Nom": str(nom).strip(),
                        "Club": str(club or "Indépendant").strip(),
                        "Comité": "Comité Non Renseigné",
                        "Poids": str(poids or "").strip(),
                        "Points": 0,
                        "Clt": 1
                    })
                r += 1
            continue

        # Cas Poules / Tableaux
        h_row = None
        col_clt, col_nom, col_club, col_comite, col_tot_pts, col_poids = None, None, None, None, None, None
        
        for r_s in range(1, min(15, ws.max_row + 1)):
            for c_i in range(1, min(25, ws.max_column + 1)):
                val = str(ws.cell(row=r_s, column=c_i).value or "").strip().lower()
                if any(k in val for k in ["nom", "prénom", "prenom", "lutteur", "athlete"]) and "catégorie" not in val:
                    col_nom = c_i
                    h_row = r_s
                elif any(k in val for k in ["clt", "rang", "classt", "classement"]) and c_i != col_nom:
                    col_clt = c_i
                elif any(k in val for k in ["club", "équipe"]):
                    col_club = c_i
                elif any(k in val for k in ["comité", "comite", "ligue"]):
                    col_comite = c_i
                elif any(k in val for k in ["total pts", "total points", "pts total", "points total"]):
                    col_tot_pts = c_i
                elif any(k in val for k in ["poids", "kg"]):
                    col_poids = c_i
            if col_nom and h_row:
                break
                
        if not h_row: h_row = 4
        if not col_nom: col_nom = 3

        combats_joues = []
        for r_m in range(h_row + 1, ws.max_row + 1):
            val_cell = str(ws.cell(row=r_m, column=1).value or "").strip()
            if "COMBAT" in val_cell.upper() or "MATCH" in val_cell.upper():
                c_rouge = str(ws.cell(row=r_m+1, column=1).value or "").strip()
                pt_r = extraire_valeur_score(ws.cell(row=r_m+1, column=ws.max_column).value or ws.cell(row=r_m+1, column=2).value)
                c_bleu = str(ws.cell(row=r_m+2, column=1).value or "").strip()
                pt_b = extraire_valeur_score(ws.cell(row=r_m+2, column=ws.max_column).value or ws.cell(row=r_m+2, column=2).value)
                if c_rouge and c_bleu:
                    combats_joues.append((c_rouge, pt_r, c_bleu, pt_b))

        lutteurs_poule = []
        r = h_row + 1
        while r <= min(ws.max_row, h_row + 35):
            nom_raw = ws.cell(row=r, column=col_nom).value
            club_raw = ws.cell(row=r, column=col_club).value if col_club else "Indépendant"
            
            if est_ligne_lutteur(nom_raw, club_raw):
                nom_s = str(nom_raw).strip()
                club_s = str(club_raw or "Indépendant").strip()
                comite_s = str(ws.cell(row=r, column=col_comite).value or "Comité Non Renseigné").strip() if col_comite else "Comité Non Renseigné"
                poids_s = str(ws.cell(row=r, column=col_poids).value or "").strip() if col_poids else ""
                
                clt_brut = ws.cell(row=r, column=col_clt).value if col_clt else None
                clt_net = nettoyer_rang(clt_brut)
                
                pts_val = 0.0
                if col_tot_pts:
                    pts_val = extraire_valeur_score(ws.cell(row=r, column=col_tot_pts).value)
                
                if pts_val == 0.0:
                    start_t = (col_comite or col_club or col_nom) + 1
                    end_t = col_tot_pts if col_tot_pts else (start_t + 5)
                    for c_t in range(start_t, min(end_t, ws.max_column + 1)):
                        pts_val += extraire_valeur_score(ws.cell(row=r, column=c_t).value)
                        
                if pts_val == 0.0 and combats_joues:
                    norm_n = normaliser_nom_comparaison(nom_s)
                    for (cr, pr, cb, pb) in combats_joues:
                        if norm_n in normaliser_nom_comparaison(cr):
                            pts_val += pr
                        elif norm_n in normaliser_nom_comparaison(cb):
                            pts_val += pb

                lutteurs_poule.append({
                    "Poule": nom_feuille,
                    "Nom": nom_s,
                    "Club": club_s,
                    "Comité": comite_s,
                    "Poids": poids_s,
                    "Points": int(round(pts_val)),
                    "Clt_Cellule": clt_net if isinstance(clt_net, int) else None
                })
            r += 1

        if lutteurs_poule:
            lutteurs_poule.sort(key=lambda x: (
                x["Clt_Cellule"] if x["Clt_Cellule"] is not None else 999,
                -x["Points"]
            ))
            
            rang_courant = 1
            for idx_p, p_item in enumerate(lutteurs_poule):
                if p_item["Clt_Cellule"] is not None:
                    p_item["Clt"] = p_item["Clt_Cellule"]
                    rang_courant = max(rang_courant, p_item["Clt_Cellule"] + 1)
                else:
                    if idx_p > 0 and p_item["Points"] == lutteurs_poule[idx_p-1]["Points"] and p_item["Points"] > 0:
                        p_item["Clt"] = lutteurs_poule[idx_p-1]["Clt"]
                    else:
                        p_item["Clt"] = rang_courant
                        rang_courant += 1
                        
            tous_les_resultats.extend(lutteurs_poule)

    return tous_les_resultats

def trier_dataframe_bilan(df):
    if df is None or df.empty:
        return df
    
    def key_clt(v):
        r_clean = nettoyer_rang(v)
        return (0, r_clean) if isinstance(r_clean, int) else (1, 999)

    df_copy = df.copy()
    df_copy["_sort_clt"] = df_copy["Clt"].apply(key_clt)
    df_copy["_sort_pts"] = pd.to_numeric(df_copy["Points"], errors="coerce").fillna(0)
    
    return df_copy.sort_values(
        by=["Poule", "_sort_clt", "_sort_pts"], 
        ascending=[True, True, False]
    ).drop(columns=["_sort_clt", "_sort_pts"]).reset_index(drop=True)

# --- SÉLECTION DU MODE PRINCIPAL ---
col_mode1, col_mode2 = st.columns([2, 3])
with col_mode1:
    mode_app = st.selectbox("📌 Sélectionnez le mode de travail", ["1. Générer un Tournoi (Planning & Poules)", "2. Importer les scores & Éditer le Bilan (Excel)"])

if mode_app.startswith("2"):
    st.markdown("### 📂 Module de Fin de Tournoi & Bilans Fédéraux")
    st.info(f"Importez votre fichier Excel complété pour la compétition **{nom_competition}** afin de générer automatiquement les classements officiels individuels et par club.")
    
    fichier_resultats = st.file_uploader("Sélectionner le fichier Excel complété (.xlsx)", type=["xlsx"])
    
    if fichier_resultats is not None:
        try:
            wb_res = openpyxl.load_workbook(fichier_resultats, data_only=True)
            fichier_resultats.seek(0)
            try:
                wb_f = openpyxl.load_workbook(fichier_resultats, data_only=False)
            except Exception:
                wb_f = None
            
            nom_comp_officiel = nom_competition
            for s_name in wb_res.sheetnames:
                ws_chk = wb_res[s_name]
                for r_chk in [1, 2]:
                    val_c = str(ws_chk.cell(row=r_chk, column=1).value or "").strip()
                    m_comp = re.search(r"COMPÉTITION\s*:\s*([^—\n]+)", val_c, re.IGNORECASE)
                    if m_comp:
                        nom_comp_officiel = m_comp.group(1).strip()
                        break
                if nom_comp_officiel != nom_competition:
                    break

            tous_les_resultats = extraire_resultats_classeur_excel(wb_res, wb_f)
            df_bilan = pd.DataFrame(tous_les_resultats)
            
            if not df_bilan.empty:
                df_bilan["Clt"] = df_bilan["Clt"].apply(lambda v: str(nettoyer_rang(v)))
                df_bilan["Points"] = pd.to_numeric(df_bilan["Points"], errors="coerce").fillna(0).astype(int)
                for col_str in ["Nom", "Club", "Comité", "Poids", "Poule"]:
                    if col_str in df_bilan.columns:
                        df_bilan[col_str] = df_bilan[col_str].astype(str).replace(["None", "nan"], "")
                df_bilan = trier_dataframe_bilan(df_bilan)
            
            st.success(f"✨ Résultats extraits avec succès pour : **{nom_comp_officiel}** ({len(df_bilan)} lutteurs recensés) !")
            
            st.markdown("---")
            with st.expander("✏️ Vérification & Ajustement des Résultats (Aperçu Interactif)", expanded=True):
                st.info("💡 Vous pouvez corriger directement le rang (`Clt`), le `Nom`, le `Club`, ou les `Points` ci-dessous si nécessaire.")
                df_bilan_edited = st.data_editor(
                    df_bilan,
                    column_config={
                        "Clt": st.column_config.TextColumn("Clt / Rang", help="Rang officiel (1, 2, 3, 4, 5...)"),
                        "Nom": st.column_config.TextColumn("Nom Prénom"),
                        "Club": st.column_config.TextColumn("Club"),
                        "Comité": st.column_config.TextColumn("Comité Régional"),
                        "Poids": st.column_config.TextColumn("Poids"),
                        "Points": st.column_config.NumberColumn("Points Victoire", step=1),
                        "Poule": st.column_config.TextColumn("Poule / Catégorie"),
                    },
                    use_container_width=True,
                    hide_index=True,
                    key="editor_bilan_mode2"
                )

            # CALCUL DES POINTS CLUBS ET COMITÉS
            bareme_points = {1: 4, 2: 3, 3: 2, 4: 1}
            points_clubs = {}
            points_comites = {}

            for _, row in df_bilan_edited.iterrows():
                club = str(row.get("Club", "")).strip()
                comite = str(row.get("Comité", "Comité Non Renseigné")).strip()
                if not comite or comite in ["None", "nan", "-"]:
                    comite = "Comité Non Renseigné"
                
                if not club or club.lower() in ["", "-", "none", "nan", "club", "indépendant"] or "kg" in club.lower():
                    continue
                
                if club not in points_clubs:
                    points_clubs[club] = {"Club": club, "Points Club": 0, "1ers": 0, "2èmes": 0, "3èmes": 0, "4èmes": 0}
                if comite and comite not in points_comites:
                    points_comites[comite] = {"Comité Régional": comite, "Points Comité": 0, "1ers": 0, "2èmes": 0, "3èmes": 0, "4èmes": 0}
                
                clt_val = nettoyer_rang(row.get("Clt"))
                if clt_val == "NR" or not isinstance(clt_val, int):
                    continue
                
                clt_num = clt_val
                is_u7_poule = "u7" in str(row.get("Poule", "")).lower()
                pts_attribués = 1 if is_u7_poule else bareme_points.get(clt_num, 0)
                
                if club in points_clubs:
                    points_clubs[club]["Points Club"] += pts_attribués
                    if not is_u7_poule:
                        if clt_num == 1: points_clubs[club]["1ers"] += 1
                        elif clt_num == 2: points_clubs[club]["2èmes"] += 1
                        elif clt_num == 3: points_clubs[club]["3èmes"] += 1
                        elif clt_num == 4: points_clubs[club]["4èmes"] += 1

                if comite in points_comites:
                    points_comites[comite]["Points Comité"] += pts_attribués
                    if not is_u7_poule:
                        if clt_num == 1: points_comites[comite]["1ers"] += 1
                        elif clt_num == 2: points_comites[comite]["2èmes"] += 1
                        elif clt_num == 3: points_comites[comite]["3èmes"] += 1
                        elif clt_num == 4: points_comites[comite]["4èmes"] += 1

            if points_clubs:
                df_clubs = pd.DataFrame(list(points_clubs.values())).sort_values(
                    by=["Points Club", "1ers", "2èmes", "3èmes", "4èmes"], 
                    ascending=False
                ).reset_index(drop=True)
                df_clubs.index = range(1, len(df_clubs) + 1)
                df_clubs.insert(0, "Clt Club", df_clubs.index)
            else:
                df_clubs = pd.DataFrame(columns=["Clt Club", "Club", "Points Club", "1ers", "2èmes", "3èmes", "4èmes"])

            if points_comites:
                df_comites = pd.DataFrame(list(points_comites.values())).sort_values(
                    by=["Points Comité", "1ers", "2èmes", "3èmes", "4èmes"], 
                    ascending=False
                ).reset_index(drop=True)
                df_comites.index = range(1, len(df_comites) + 1)
                df_comites.insert(0, "Clt Comité", df_comites.index)
            else:
                df_comites = pd.DataFrame(columns=["Clt Comité", "Comité Régional", "Points Comité", "1ers", "2èmes", "3èmes", "4èmes"])

            tab_bilan_1, tab_bilan_2, tab_bilan_3 = st.tabs([
                "🏆 Classements Individuels (U7 / U9 / U11 / U13)", 
                "🛡️ Classement Général des Clubs",
                "🏛️ Classement des Comités Régionaux"
            ])
            
            with tab_bilan_1:
                st.subheader(f"Classements Individuels Officiels — {nom_comp_officiel}")
                for poule in df_bilan_edited['Poule'].unique():
                    st.markdown(f"#### 🤼 {poule}")
                    sous_df = df_bilan_edited[df_bilan_edited['Poule'] == poule][['Clt', 'Nom', 'Club', 'Poids', 'Points']].copy()
                    st.dataframe(sous_df, use_container_width=True, hide_index=True)

            with tab_bilan_2:
                st.subheader(f"🛡️ Podium des Clubs Engagés — {nom_comp_officiel}")
                st.markdown("*Barème officiel FFLDA : 1er = 4 pts | 2ème = 3 pts | 3ème = 2 pts | 4ème = 1 pt*")
                st.dataframe(df_clubs, use_container_width=True, hide_index=True)

            with tab_bilan_3:
                st.subheader(f"🏛️ Classement Officiel des Comités Régionaux — {nom_comp_officiel}")
                st.markdown("*Barème officiel FFLDA : 1er = 4 pts | 2ème = 3 pts | 3ème = 2 pts | 4ème = 1 pt*")
                st.dataframe(df_comites, use_container_width=True, hide_index=True)

            output_excel_bilan = io.BytesIO()
            with pd.ExcelWriter(output_excel_bilan, engine='openpyxl') as writer:
                df_clubs.to_excel(writer, sheet_name="Classement Clubs", index=False)
                df_comites.to_excel(writer, sheet_name="Classement Comités", index=False)
                df_bilan_edited.to_excel(writer, sheet_name="Classements Individuels", index=False)

            st.markdown("---")
            st.markdown("### 📥 Téléchargements du Bilan")
            st.download_button(
                label="📥 Télécharger le Bilan Officiel Excel (.xlsx)",
                data=output_excel_bilan.getvalue(),
                file_name=f"Bilan_Officiel_{nom_comp_officiel.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_dl_bilan"
            )

        except Exception as e:
            st.error(f"Erreur lors de l'analyse du fichier : {e}")

else:
    # --- MODE 1 : GÉNÉRATION DU TOURNOI ---
    fichier_upload = st.file_uploader("📂 Importez votre liste d'inscrits (.csv ou .xlsx)", type=["xlsx", "csv"])
    if fichier_upload is None:
        st.info("👈 Veuillez importer un fichier d'inscrits pour générer la compétition.")
    else:
        try:
            if fichier_upload.name.endswith('.csv'):
                try:
                    df_raw = pd.read_csv(fichier_upload, sep=';', encoding='utf-8')
                    if len(df_raw.columns) == 1:
                        fichier_upload.seek(0)
                        df_raw = pd.read_csv(fichier_upload, sep=',', encoding='utf-8')
                except UnicodeDecodeError:
                    fichier_upload.seek(0)
                    df_raw = pd.read_csv(fichier_upload, sep=';', encoding='latin-1')
            else:
                df_raw = pd.read_excel(fichier_upload)

            st.success(f"Fichier chargé avec succès ! {len(df_raw)} lignes trouvées.")
            st.dataframe(df_raw.head())
            st.info("💡 Vous pouvez paramétrer vos options dans le panneau latéral gauche avant d'éditer ou de passer au Mode 2 pour les bilans finaux.")
        except Exception as e:
            st.error(f"Erreur de lecture du fichier : {e}")
