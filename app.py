import streamlit as st
import re
import pandas as pd
import requests
from datetime import datetime

# ── Configuration de la page ──────────────────────────────────────────────────

st.set_page_config(
    page_title="Airstock → Airbus CSV",
    page_icon="🚁",
    layout="centered"
)

# ── Moteur de conversion ─────────────────────────────────────────────────────
#
# Stratégie : on délègue intégralement la lecture du binaire .xls à pandas/xlrd.
# xlrd est une bibliothèque éprouvée qui gère nativement toutes les variantes
# du format BIFF8 (SST fragmentée, Crystal Reports, enregistrements CONTINUE…).
# Toute tentative de parser le binaire manuellement est fragile par construction :
# le format contient des cas limites impossibles à couvrir exhaustivement.
#
# Structure du fichier Airstock :
#   - Lignes 0–3  : en-têtes (ignorées)
#   - Col 1 (B)   : Part Number
#   - Col 2 (C)   : Quantité commandée
#   - Entre chaque article : 1 ligne "Livraison avant le…" + 1 ligne vide
#   - Fin de fichier : lignes de pied de page (Total HT, mentions légales…)

import io as _io
import subprocess as _subprocess
import sys as _sys
import importlib as _importlib

def _ensure(pkg, import_as=None):
    """Installe pkg si absent (silencieux)."""
    try:
        _importlib.import_module(import_as or pkg)
    except ImportError:
        _subprocess.check_call([_sys.executable, "-m", "pip", "install", pkg, "-q"],
                               stdout=_subprocess.DEVNULL, stderr=_subprocess.DEVNULL)

_ensure("xlrd")

_FOOTER_KW = (
    "Total", "Toutes les", "HELILAGON", "Page", "COMMANDE", "EUR",
    "Livraison avant", "PN - description", "Quantité", "Payment terms",
    "Cpte fourn", "Destinataire",
)

def _is_footer(val) -> bool:
    """
    Retourne True pour toute valeur qui ne peut pas être un Part Number :
    lignes vides, en-têtes, pieds de page, dates, mentions légales…
    N'applique AUCUN filtre sur le format du PN lui-même — tout ce qui
    reste est conservé, peu importe sa forme (alphanumérique, court,
    avec tirets, points, lettres seules…).
    """
    if val is None:
        return True
    s = str(val).strip()
    if not s or s.lower() == "nan":
        return True
    # Cellules multi-paragraphes (en-tête Crystal Reports fusionné)
    if "\n" in s:
        return True
    # Mots-clés de pied de page / en-tête
    if any(kw in s for kw in _FOOTER_KW):
        return True
    # Date dd/mm/yyyy
    if re.match(r"\d{2}/\d{2}/\d{4}", s):
        return True
    return False

def convert_xls_bytes_to_csv(xls_bytes: bytes):
    """
    Lit le fichier .xls Airstock et retourne (csv_content, results).
    results est une liste de tuples (pn_str, qty_str).
    Lève ValueError si le fichier n'est pas reconnu par xlrd.
    """
    try:
        df = pd.read_excel(_io.BytesIO(xls_bytes), engine="xlrd", header=None)
    except Exception as exc:
        raise ValueError(
            f"Impossible de lire le fichier : {exc}\n"
            "Vérifiez que le fichier est bien exporté depuis Airstock "
            "au format \"Microsoft Excel (97-2003) Données uniquement (*.xls)\"."
        )

    results = []
    for i, row in df.iterrows():
        if i < 4:          # lignes 0-3 = en-têtes Airstock
            continue

        pn_val  = row[1]   # colonne B
        qty_val = row[2]   # colonne C

        # Ignorer si PN absent ou ligne de structure (livraison, pied de page…)
        if pd.isna(pn_val) or _is_footer(pn_val):
            continue

        pn_str = str(pn_val).strip()
        if not pn_str:
            continue

        # Ignorer si quantité absente ou non numérique
        if pd.isna(qty_val):
            continue
        try:
            qty_float = round(float(qty_val), 3)
        except (ValueError, TypeError):
            continue

        if qty_float <= 0:
            continue

        # Formatage : entier si .0, sinon décimal avec virgule (norme FR)
        qty_str = (str(int(qty_float))
                   if qty_float == int(qty_float)
                   else str(qty_float).replace(".", ","))

        results.append((pn_str, qty_str))

    csv_lines = ["Ordered Reference;Quantity"]
    for pn, qty_str in results:
        csv_lines.append(f"{pn};{qty_str}")

    return "\n".join(csv_lines), results


# ── Monitoring Supabase ───────────────────────────────────────────────────────

def _supa_headers():
    return {
        "apikey":        st.secrets["SUPABASE_KEY"],
        "Authorization": f"Bearer {st.secrets['SUPABASE_KEY']}",
        "Content-Type":  "application/json",
        "Prefer":        "return=minimal",
    }

def log_conversion(nom_fichier: str, nb_pn: int, pn_list: list):
    """
    Enregistre une conversion dans Supabase.
    Silencieux en cas d'erreur — ne bloque jamais l'app.
    """
    try:
        url  = st.secrets["SUPABASE_URL"] + "/rest/v1/conversions"
        payload = {
            "nom_fichier": nom_fichier,
            "nb_pn":       nb_pn,
            "pn_list":     ", ".join(pn_list),
        }
        requests.post(url, json=payload, headers=_supa_headers(), timeout=5)
    except Exception:
        pass  # Le monitoring ne doit jamais faire planter l'app

def fetch_conversions():
    """Récupère tout l'historique depuis Supabase."""
    try:
        url = (
            st.secrets["SUPABASE_URL"]
            + "/rest/v1/conversions"
            + "?select=*&order=created_at.desc"
        )
        r = requests.get(url, headers=_supa_headers(), timeout=5)
        r.raise_for_status()
        return r.json()
    except Exception as e:
        return None

CONVERSION_LIMIT = 10

def count_conversions():
    """Retourne le nombre de conversions enregistrées, ou None si erreur Supabase."""
    try:
        headers = _supa_headers()
        headers["Prefer"] = "count=exact"
        url = (
            st.secrets["SUPABASE_URL"]
            + "/rest/v1/conversions"
            + "?select=id"
        )
        r = requests.get(url, headers=headers, timeout=5)
        r.raise_for_status()
        # Supabase retourne le total dans le header Content-Range : "0-N/TOTAL"
        content_range = r.headers.get("Content-Range", "")
        if "/" in content_range:
            return int(content_range.split("/")[1])
        return len(r.json())
    except Exception:
        return None


# ── Interface Streamlit ───────────────────────────────────────────────────────

# ── Bandeau de progression paiement ─────────────────────────────────────────

recu    = 239
total   = 478
restant = total - recu
pct     = int(recu / total * 100)

# Récupération du compteur de conversions pour le bandeau
_nb_conv   = count_conversions()
_conv_used = _nb_conv if _nb_conv is not None else 0
_conv_left = max(0, CONVERSION_LIMIT - _conv_used)

if _nb_conv is None:
    _conv_line = "<em style='color:#aaa;'>Compteur indisponible</em>"
elif _conv_left == 0:
    _conv_line = "🔴 <b>Limite atteinte — outil bloqué</b> (0 conversion restante)"
elif _conv_left <= 10:
    _conv_line = f"🟠 <b>{_conv_left} conversion(s) restante(s)</b> avant blocage de l'outil"
else:
    _conv_line = f"🟢 <b>{_conv_left} conversion(s) restante(s)</b> sur {CONVERSION_LIMIT}"

st.markdown(f"""
<div style="
    background: #f0f0f0;
    border-radius: 12px;
    padding: 16px 20px;
    margin-bottom: 20px;
    border: 2px solid #cc0000;
">
    <div style="display: flex; justify-content: space-between; margin-bottom: 8px;">
        <span style="font-weight: bold; color: #cc0000; font-size: 15px;">
            💸 Paiement outil de conversion Airstock → Airbus
        </span>
        <span style="font-weight: bold; color: #cc0000; font-size: 15px;">
            {recu}€ reçus / {total}€ — il reste <b>{restant}€</b> 👀
        </span>
    </div>
    <div style="background:#e0e0e0; border-radius:8px; height:28px; width:100%; overflow:hidden;">
        <div style="
            background: linear-gradient(90deg, #cc0000, #ff1a1a);
            width: {pct}%;
            height: 100%;
            border-radius: 8px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-weight: bold;
            font-size: 14px;
            transition: width 0.5s;
        ">{pct}% payé</div>
    </div>
    <div style="text-align:center; margin-top:8px; color:#888; font-size:12px;">
        ⏳ Départ lundi — la montre tourne 🕐
    </div>
    <div style="text-align:center; margin-top:4px; font-size:13px;">
        {_conv_line}
    </div>
</div>
""", unsafe_allow_html=True)

st.title("🚁 Airstock → Airbus CSV")
st.markdown(
    "Convertit automatiquement un fichier `.xls` exporté depuis **Airstock** "
    "en fichier `.csv` prêt à être déposé sur le portail **Airbus Helicopters** "
    "pour l'import multi-PN."
)

st.divider()

# ── Instructions d'export Airstock ───────────────────────────────────────────

with st.expander("📖 Comment exporter le fichier depuis Airstock ?", expanded=False):
    st.markdown("Suis ces étapes dans Airstock **avant** d'utiliser cet outil :")

    col1, col2 = st.columns([1, 20])
    with col1: st.markdown("**1**")
    with col2: st.markdown("Dans ta commande Airstock, fais un **clic droit** sur la commande → sélectionne **« Imprimer »**")

    col1, col2 = st.columns([1, 20])
    with col1: st.markdown("**2**")
    with col2: st.markdown("Le menu **« Impression du détail des commandes »** s'ouvre → clique sur **« Print (F5) »** (le bouton imprimante)")

    col1, col2 = st.columns([1, 20])
    with col1: st.markdown("**3**")
    with col2: st.markdown("Le menu **« Edition des commandes »** s'ouvre → clique sur **« Exporter le rapport »**")

    col1, col2 = st.columns([1, 20])
    with col1: st.markdown("**4**")
    with col2: st.markdown("Donne un **nom au fichier** (ex : `PO-26-3097`)")

    col1, col2 = st.columns([1, 20])
    with col1: st.markdown("**5**")
    with col2:
        st.markdown("⚠️ **Très important** : dans le menu déroulant du format, sélectionne obligatoirement :")
        st.code("Microsoft Excel (97-2003) Données uniquement (*.xls)", language=None)

    col1, col2 = st.columns([1, 20])
    with col1: st.markdown("**6**")
    with col2: st.markdown("Clique **Enregistrer** — le fichier `.xls` est maintenant enregistré dans le dossier **Documents du bureau à distance**")

    st.markdown("---")
    st.markdown("**📋 Transfert du fichier vers ton ordinateur local**")

    col1, col2 = st.columns([1, 20])
    with col1: st.markdown("**7**")
    with col2: st.markdown("Ouvre l'**Explorateur de fichiers** du bureau à distance et navigue jusqu'au dossier **Documents**")

    col1, col2 = st.columns([1, 20])
    with col1: st.markdown("**8**")
    with col2: st.markdown("Repère ton fichier `.xls` (ex : `PO-26-3097.xls`) → **clic droit** → **Copier**")

    col1, col2 = st.columns([1, 20])
    with col1: st.markdown("**9**")
    with col2: st.markdown("Dans l'Explorateur de fichiers, clique sur **Ce PC** dans le panneau de gauche → tu verras apparaître ton ordinateur local sous la forme **`Lecteur (\\\\tsclient\\...)`** ou similaire → ouvre-le")

    col1, col2 = st.columns([1, 20])
    with col1: st.markdown("**10**")
    with col2: st.markdown("Navigue jusqu'au dossier de ton choix sur ton **PC local** (ex : Documents) → **clic droit** → **Coller**")

    col1, col2 = st.columns([1, 20])
    with col1: st.markdown("**11**")
    with col2: st.markdown("Le fichier est maintenant sur ton PC local ✅ Tu peux le déposer dans l'outil ci-dessous !")

    st.info(
        "💡 **Attention au format !** Si tu choisis un autre format que "
        "*\"Données uniquement (*.xls)\"*, le fichier ne sera pas reconnu par cet outil.",
        icon="⚠️"
    )
    st.info(
        "💡 **Bureau à distance non visible dans Ce PC ?** Vérifie que le partage de lecteurs locaux "
        "est bien activé dans les **options de connexion Bureau à distance** (onglet *Ressources locales* "
        "→ *Plus...* → coche *Lecteurs*).",
        icon="🖥️"
    )

with st.expander("🎬 [TUTO] Comment exporter une commande dans un fichier .xls ?"):
    st.video("https://raw.githubusercontent.com/HugoDrs/Helilagon/main/videos/%5BTUTO%5D%20Export%20fichier%20xls%20depuis%20Airstock.mp4")

with st.expander("🎬 [TUTO] Comment importer un fichier .csv dans Airbus ?"):
    st.video("https://raw.githubusercontent.com/HugoDrs/Helilagon/main/videos/%5BTUTO%5D%20Import%20fichier%20csv%20dans%20Airbus.mp4")

st.divider()

# ── Zone de conversion ────────────────────────────────────────────────────────

st.subheader("📂 Dépose ton fichier .xls")

uploaded_file = st.file_uploader(
    "Sélectionne le fichier exporté depuis Airstock",
    type=["xls"],
    help="Fichier au format Microsoft Excel (97-2003) Données uniquement (*.xls)"
)

if uploaded_file is not None:
    st.info(f"📄 Fichier reçu : **{uploaded_file.name}** ({uploaded_file.size // 1024} Ko)")

    # ── Vérification de la limite avant conversion ────────────────────────────
    _current_count = count_conversions()
    if _current_count is not None and _current_count >= CONVERSION_LIMIT:
        st.error(
            f"🔴 **Limite de {CONVERSION_LIMIT} conversions atteinte.** "
            "L'outil est temporairement bloqué. "
            "Contacte l'administrateur pour débloquer l'accès."
        )
        st.stop()

    with st.spinner("⏳ Conversion en cours..."):
        try:
            xls_bytes            = uploaded_file.read()
            csv_content, results = convert_xls_bytes_to_csv(xls_bytes)

            csv_filename = uploaded_file.name.replace('.xls', '_airbus.csv')
            csv_bytes    = csv_content.encode('utf-8-sig')

            st.success(f"✅ **{len(results)} Part Number(s)** extraits avec succès !")

            with st.expander("📋 Voir le détail des lignes extraites"):
                df = pd.DataFrame({
                    "N°":          list(range(1, len(results) + 1)),
                    "Part Number": [r[0] for r in results],
                    "Quantité":    [r[1] for r in results],
                })
                st.dataframe(
                    df,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "N°":          st.column_config.NumberColumn(width="10"),
                        "Part Number": st.column_config.TextColumn(width="medium"),
                        "Quantité":    st.column_config.TextColumn(width="small"),
                    }
                )

            # ── Enregistrement monitoring au clic sur Télécharger ───────────
            # on_click est appelé UNE seule fois au moment du clic,
            # pas à chaque re-exécution du script Streamlit.
            def _log():
                log_conversion(
                    nom_fichier = uploaded_file.name,
                    nb_pn       = len(results),
                    pn_list     = [r[0] for r in results],
                )

            st.download_button(
                label="📥 Télécharger le CSV Airbus",
                data=csv_bytes,
                file_name=csv_filename,
                mime="text/csv",
                on_click=_log,
            )

            st.markdown("---")
            st.markdown("### 🔜 Étape suivante")
            st.markdown(
                "Une fois le CSV téléchargé, rends-toi sur le portail Airbus "
                "pour déposer le fichier et importer tous tes PN en une seule fois :"
            )
            st.link_button(
                label="🌐 Accéder au portail Airbus Helicopters (Mass Upload)",
                url="https://keycopter.airbushelicopters.com/sparesstorefront/AHFWebsite/mass-upload",
                use_container_width=True,
            )

        except Exception as e:
            st.error(f"❌ Erreur lors de la conversion : {e}")
            st.warning(
                "Vérifiez que le fichier est bien exporté depuis Airstock "
                "au format **\"Microsoft Excel (97-2003) Données uniquement (*.xls)\"**. "
                "Tout autre format ne sera pas reconnu."
            )

else:
    st.markdown(
        "👆 Clique sur **Browse files** (ou glisse ton fichier) pour lancer la conversion automatique."
    )

st.divider()

# ── Tableau de bord monitoring (protégé par mot de passe) ────────────────────

with st.expander("🔒 Tableau de bord — Accès administrateur"):
    pwd = st.text_input("Mot de passe", type="password", key="admin_pwd")

    ADMIN_PWD = st.secrets.get("ADMIN_PASSWORD", "helilagon2024")

    if pwd == ADMIN_PWD:

        data = fetch_conversions()

        if data is None:
            st.error("❌ Impossible de contacter Supabase.")

        elif len(data) == 0:
            st.info("Aucune conversion enregistrée pour le moment.")

        else:
            # ── Stats globales ──────────────────────────────────────────────
            nb_conversions = len(data)
            nb_pn_total    = sum(row["nb_pn"] for row in data)
            derniere       = data[0]["created_at"][:10]  # YYYY-MM-DD

            col1, col2, col3 = st.columns(3)
            col1.metric("📁 Fichiers convertis", nb_conversions)
            col2.metric("🔩 PN traités au total", nb_pn_total)
            col3.metric("📅 Dernière conversion", derniere)

            st.markdown("---")

            # ── Graphique conversions par jour ──────────────────────────────
            df_hist = pd.DataFrame(data)
            df_hist["date"] = pd.to_datetime(df_hist["created_at"]).dt.date
            df_par_jour = (
                df_hist.groupby("date")
                .agg(nb_fichiers=("id", "count"), nb_pn=("nb_pn", "sum"))
                .reset_index()
                .sort_values("date")
            )

            st.markdown("#### 📈 Activité par jour")
            col_g1, col_g2 = st.columns(2)
            with col_g1:
                st.markdown("**Fichiers convertis**")
                st.bar_chart(df_par_jour.set_index("date")["nb_fichiers"])
            with col_g2:
                st.markdown("**PN traités**")
                st.bar_chart(df_par_jour.set_index("date")["nb_pn"])

            st.markdown("---")

            # ── Historique détaillé ─────────────────────────────────────────
            st.markdown("#### 🗂️ Historique des conversions")
            df_display = pd.DataFrame([{
                "Date"        : row["created_at"][:19].replace("T", " "),
                "Fichier"     : row["nom_fichier"],
                "Nb PN"       : row["nb_pn"],
                "Liste PN"    : row["pn_list"],
            } for row in data])

            st.dataframe(
                df_display,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Liste PN": st.column_config.TextColumn(width="large"),
                }
            )

            # ── Export de l'historique ──────────────────────────────────────
            csv_monitoring = df_display.to_csv(index=False, sep=";").encode("utf-8-sig")
            st.download_button(
                label="📥 Exporter l'historique en CSV",
                data=csv_monitoring,
                file_name=f"monitoring_airstock_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
            )

    elif pwd != "":
        st.error("❌ Mot de passe incorrect.")

st.caption("Outil interne — Helilagon Logistique")
