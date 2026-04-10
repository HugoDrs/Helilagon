import streamlit as st
import struct
import re
import pandas as pd
import requests
from datetime import datetime, timezone

# ── Configuration de la page ──────────────────────────────────────────────────

st.set_page_config(
    page_title="Airstock → Airbus CSV",
    page_icon="🚁",
    layout="centered"
)

# ── Moteur de conversion BIFF8 ────────────────────────────────────────────────

COL_PN     = 1
COL_QTY    = 2
HEADER_ROW = 3

def u16(data, pos): return struct.unpack_from('<H', data, pos)[0]
def u32(data, pos): return struct.unpack_from('<I', data, pos)[0]
def f64(data, pos): return struct.unpack_from('<d', data, pos)[0]

def find_sst_offset(data):
    pos = 0
    while True:
        idx = data.find(b'\xfc\x00', pos)
        if idx == -1: return None
        rec_len = u16(data, idx + 2)
        if rec_len >= 8 and idx + 12 <= len(data):
            total  = u32(data, idx + 4)
            unique = u32(data, idx + 8)
            if 0 < unique <= 10000 and total >= unique:
                return idx
        pos = idx + 1

def parse_sst(data, sst_offset):
    unique_count = u32(data, sst_offset + 8)
    pos = sst_offset + 12
    strings = []
    for _ in range(unique_count):
        if pos + 3 > len(data): strings.append(None); continue
        str_len    = u16(data, pos)
        flags      = data[pos + 2]
        compressed = not (flags & 0x01)
        has_rich   = bool(flags & 0x08)
        has_phonet = bool(flags & 0x04)
        pos += 3
        rich_count = phonetic_size = 0
        if has_rich:   rich_count    = u16(data, pos); pos += 2
        if has_phonet: phonetic_size = u32(data, pos); pos += 4
        byte_count = str_len if compressed else str_len * 2
        if pos + byte_count > len(data):
            strings.append(None)
            pos += byte_count + rich_count * 4 + phonetic_size
            continue
        try:
            enc = 'latin-1' if compressed else 'utf-16-le'
            s   = data[pos:pos + byte_count].decode(enc, errors='replace')
            if s.count('\ufffd') > 2 or any(ord(c) > 0x2000 for c in s):
                s = None
        except:
            s = None
        pos += byte_count + rich_count * 4 + phonetic_size
        strings.append(s)
    return strings

def parse_early_pool(data, sst_offset):
    pn_re     = re.compile(r'^[A-Z0-9][A-Z0-9\-\.\,]{2,49}$')
    pn_sep_re = re.compile(r'[0-9\-\.\,]')
    def is_pn(s): return bool(pn_re.match(s) and pn_sep_re.search(s))
    pns = []
    pos = 0
    while pos < sst_offset - 3:
        if pos + 3 > len(data): break
        str_len = u16(data, pos)
        flags   = data[pos + 2]
        if flags == 0x01 and 3 <= str_len <= 50:
            bc  = str_len * 2
            end = pos + 3 + bc
            if end <= sst_offset:
                try:
                    s = data[pos + 3:end].decode('utf-16-le', errors='strict')
                    if s.isprintable() and is_pn(s): pns.append(s)
                except: pass
        pos += 1
    return pns

def build_fallback_pn_list(early_pns):
    pn_re     = re.compile(r'^[A-Z0-9][A-Z0-9\-\.\,]{2,49}$')
    pn_sep_re = re.compile(r'[0-9\-\.\,]')
    result = []
    for pn in early_pns:
        if pn.endswith('LE') and len(pn) > 4:
            t = pn[:-2]
            if pn_re.match(t) and pn_sep_re.search(t) and t not in result:
                result.append(t)
        if pn not in result:
            result.append(pn)
    return result

def parse_cells(data, strings):
    cells = {}
    pos = 0
    while True:
        idx = data.find(b'\xfd\x00', pos)
        if idx == -1: break
        if idx + 14 <= len(data) and u16(data, idx + 2) == 10:
            row     = u16(data, idx + 4)
            col     = u16(data, idx + 6)
            sst_idx = u32(data, idx + 10)
            cells[(row, col)] = strings[sst_idx] if sst_idx < len(strings) else None
        pos = idx + 1
    pos = 0
    while True:
        idx = data.find(b'\x03\x02', pos)
        if idx == -1: break
        if idx + 18 <= len(data) and u16(data, idx + 2) == 14:
            row = u16(data, idx + 4)
            col = u16(data, idx + 6)
            cells[(row, col)] = f64(data, idx + 10)
        pos = idx + 1
    return cells

def fill_corrupted_pns(cells, fallback_pns):
    if not fallback_pns: return cells
    broken_rows = sorted(
        row for (row, col), val in cells.items()
        if col == COL_PN and val is None
    )
    for row, pn in zip(broken_rows, fallback_pns):
        cells[(row, COL_PN)] = pn
    return cells

def is_header_or_footer(val):
    if val is None: return True
    s = str(val).strip()
    if not s: return True
    if '\n' in s: return True
    if re.match(r'^\d{2}/\d{2}/\d{4}', s): return True
    return False

def convert_xls_bytes_to_csv(xls_bytes):
    data = xls_bytes
    sst_offset = find_sst_offset(data)
    if sst_offset is None:
        raise ValueError(
            "Table SST introuvable. "
            "Vérifiez que le fichier est bien un .xls Airstock (Excel 97-2003)."
        )
    strings      = parse_sst(data, sst_offset)
    early_pns    = parse_early_pool(data, sst_offset)
    fallback_pns = build_fallback_pn_list(early_pns)
    cells        = parse_cells(data, strings)
    cells        = fill_corrupted_pns(cells, fallback_pns)
    all_rows     = sorted(set(row for (row, _) in cells))
    results      = []
    for row in all_rows:
        if row <= HEADER_ROW: continue
        pn_val  = cells.get((row, COL_PN))
        qty_val = cells.get((row, COL_QTY))
        if qty_val is None: continue
        if is_header_or_footer(pn_val): continue
        pn_str = str(pn_val).strip()
        if not pn_str: continue
        try:
            qty_float = round(float(qty_val), 3)
        except (ValueError, TypeError):
            continue
        if qty_float <= 0: continue
        qty_str = str(int(qty_float)) if qty_float == int(qty_float) else str(qty_float).replace('.', ',')
        results.append((pn_str, qty_str))
    csv_lines = ["PN;Quantité"]
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


# ── Interface Streamlit ───────────────────────────────────────────────────────

# ── Bandeau de progression paiement ─────────────────────────────────────────

recu    = 239
total   = 500
restant = total - recu
pct     = int(recu / total * 100)

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
        ⏳ Départ lundi — la montre tourne Nadège 🕐
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

    with st.spinner("⏳ Conversion en cours..."):
        try:
            xls_bytes            = uploaded_file.read()
            csv_content, results = convert_xls_bytes_to_csv(xls_bytes)

            csv_filename = uploaded_file.name.replace('.xls', '_airbus.csv')
            csv_bytes    = csv_content.encode('utf-8-sig')

            st.success(f"✅ **{len(results)} Part Number(s)** extraits avec succès !")

            with st.expander("📋 Voir le détail des lignes extraites"):
                df = pd.DataFrame(
                    {"Part Number": [r[0] for r in results],
                     "Quantité":    [r[1] for r in results]},
                    index=range(1, len(results) + 1)
                )
                st.table(df)

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
