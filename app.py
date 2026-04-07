import streamlit as st
import struct
import re
import pandas as pd
from datetime import datetime
import json

# ── Configuration de la page ──────────────────────────────────────────────────

st.set_page_config(
    page_title="Airstock → Airbus CSV",
    page_icon="🚁",
    layout="centered"
)

# ── Fonctions de conversion (moteur) ─────────────────────────────────────────

COL_PN     = 1
COL_QTY    = 2
HEADER_ROW = 3

def u16(data, pos): return struct.unpack_from('<H', data, pos)[0]
def u32(data, pos): return struct.unpack_from('<I', data, pos)[0]
def f64(data, pos): return struct.unpack_from('<d', data, pos)[0]

def find_sst_offset(data):
    """
    Cherche la vraie SST dans le fichier BIFF8.
    Certains fichiers Airstock contiennent de faux enregistrements 0xFC00
    avant la vraie SST (artefacts Crystal Reports). On valide donc que :
    - total_strings et unique_strings sont cohérents (égaux et > 0)
    - unique_strings est une valeur raisonnable (< 10 000)
    """
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
    sst_data_end = sst_offset + 4 + u16(data, sst_offset + 2)  # fin de l'enregistrement SST
    pos          = sst_offset + 12
    strings      = []
    corrupted    = False  # flag : dès qu'une corruption est détectée, tout le reste est None

    for _ in range(unique_count):
        if corrupted or pos + 3 > len(data) or pos >= sst_data_end:
            strings.append(None)
            continue

        str_len    = u16(data, pos)
        flags      = data[pos + 2]
        compressed = not (flags & 0x01)
        has_rich   = bool(flags & 0x08)
        has_phonet = bool(flags & 0x04)
        header_end = pos + 3
        if has_rich   and header_end + 2 <= len(data): header_end += 2
        if has_phonet and header_end + 4 <= len(data): header_end += 4

        byte_count = str_len if compressed else str_len * 2
        str_end    = header_end + byte_count

        # Détection de corruption Crystal Reports :
        # si la string dépasse largement la fin de la SST, le str_len est aberrant.
        if str_end > sst_data_end + 100:
            strings.append(None)
            corrupted = True  # toutes les strings suivantes seront None -> pool précoce prendra le relais
            continue

        if str_end > len(data):
            strings.append(None)
            corrupted = True
            continue

        try:
            enc = 'latin-1' if compressed else 'utf-16-le'
            s   = data[header_end:str_end].decode(enc, errors='replace')
            if s.count('\ufffd') > 2 or any(ord(c) > 0x2000 for c in s): s = None
        except: s = None

        pos = str_end
        strings.append(s)
    return strings

def parse_early_pool(data, sst_offset):
    pn_re      = re.compile(r'^[A-Z0-9][A-Z0-9\-\.\,]{2,49}$')
    pn_digit   = re.compile(r'\d')  # un PN Airbus contient toujours au moins un chiffre
    def is_pn(s): return bool(pn_re.match(s) and pn_digit.search(s))
    pns=[]; pos=0
    while pos < sst_offset - 3:
        if pos + 3 > len(data): break
        str_len = u16(data, pos); flags = data[pos + 2]
        if flags == 0x01 and 3 <= str_len <= 50:
            bc = str_len * 2; end = pos + 3 + bc
            if end <= sst_offset:
                try:
                    s = data[pos + 3:end].decode('utf-16-le', errors='strict')
                    if s.isprintable() and is_pn(s): pns.append(s)
                except: pass
        pos += 1
    return pns

def build_fallback_pn_list(early_pns):
    pn_re=re.compile(r'^[A-Z0-9][A-Z0-9\-\.\,]{2,49}$'); pn_sep_re=re.compile(r'[0-9\-\.\,]')
    result=[]
    for pn in early_pns:
        if pn.endswith('LE') and len(pn) > 4:
            t = pn[:-2]
            if pn_re.match(t) and pn_sep_re.search(t) and t not in result: result.append(t)
        if pn not in result: result.append(pn)
    return result

def parse_cells(data, strings):
    cells = {}
    pos = 0
    while True:
        idx = data.find(b'\xfd\x00', pos)
        if idx == -1: break
        if idx + 14 <= len(data) and u16(data, idx + 2) == 10:
            row = u16(data, idx + 4); col = u16(data, idx + 6)
            sst_idx = u32(data, idx + 10)
            cells[(row, col)] = strings[sst_idx] if sst_idx < len(strings) else None
        pos = idx + 1
    pos = 0
    while True:
        idx = data.find(b'\x03\x02', pos)
        if idx == -1: break
        if idx + 18 <= len(data) and u16(data, idx + 2) == 14:
            row = u16(data, idx + 4); col = u16(data, idx + 6)
            cells[(row, col)] = f64(data, idx + 10)
        pos = idx + 1
    return cells

def fill_corrupted_pns(cells, fallback_pns):
    if not fallback_pns: return cells
    broken_rows = sorted(row for (row, col), val in cells.items() if col == COL_PN and val is None)
    for row, pn in zip(broken_rows, fallback_pns): cells[(row, COL_PN)] = pn
    return cells

def is_footer_value(val):
    if val is None: return True
    s = str(val).strip()
    if not s: return True
    # Si la valeur est déjà un float (cellule numérique Excel) -> montant footer
    # Les PNs tout-chiffres sont stockés comme STR (cellule texte) -> ne pas filtrer
    if isinstance(val, float):
        return True
    # Mots-clés de pied de page
    if any(kw in s for kw in ('Page', 'COMMANDE', 'sur 1', 'EUR', 'Toutes les pi')): return True
    # Date/heure (ex: "27/03/2026  12:01")
    if re.match(r'\d{2}/\d{2}/\d{4}', s): return True
    return False

def convert_xls_bytes_to_csv(xls_bytes):
    data = xls_bytes
    sst_offset = find_sst_offset(data)
    if sst_offset is None:
        raise ValueError("Table SST introuvable — vérifiez que le fichier est bien un .xls Airstock.")
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
        if pn_val is None: continue
        pn_str = str(pn_val).strip()
        if not pn_str or is_footer_value(pn_str) or '\n' in pn_str: continue
        if qty_val is None: continue
        try: qty_float = round(float(qty_val), 3)
        except: continue
        if qty_float <= 0: continue
        qty_str = str(int(qty_float)) if qty_float == int(qty_float) else str(qty_float).replace('.', ',')
        results.append((pn_str, qty_str))
    csv_lines = ["PN;Quantité"] + [f"{pn};{qty}" for pn, qty in results]
    return "\n".join(csv_lines), results

# ── Interface Streamlit ───────────────────────────────────────────────────────

tab_app, tab_dashboard = st.tabs(["🚁 Convertisseur", "📊 Dashboard"])

# ════════════════════════════════════════════════════════════════════════════
# ONGLET 1 — CONVERTISSEUR
# ════════════════════════════════════════════════════════════════════════════
with tab_app:

    st.title("🚁 Airstock → Airbus CSV")
    st.markdown(
        "Convertit automatiquement un fichier `.xls` exporté depuis **Airstock** "
        "en fichier `.csv` prêt à être déposé sur le portail **Airbus Helicopters** "
        "pour l'import multi-PN."
    )

    st.divider()

    # ── Instructions d'export Airstock ───────────────────────────────────────
    with st.expander("📖 Comment exporter le fichier depuis Airstock ?", expanded=True):
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

    # ── Zone de conversion ────────────────────────────────────────────────────
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
                csv_filename         = uploaded_file.name.replace('.xls', '_airbus.csv')
                csv_bytes            = csv_content.encode('utf-8-sig')

                st.success(f"✅ **{len(results)} Part Number(s)** extraits avec succès !")

                with st.expander("📋 Voir le détail des lignes extraites"):
                    df = pd.DataFrame(
                        {"Part Number": [r[0] for r in results],
                         "Quantité":    [r[1] for r in results]},
                        index=range(1, len(results) + 1)
                    )
                    st.table(df)

                st.download_button(
                    label="📥 Télécharger le CSV Airbus",
                    data=csv_bytes,
                    file_name=csv_filename,
                    mime="text/csv",
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
                # ── Log échec ─────────────────────────────────────────────────
                log_conversion(
                    filename  = uploaded_file.name,
                    nb_pn     = 0,
                    status    = "❌ Échec",
                    error_msg = str(e)
                )
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
    st.caption("Outil interne — Helilagon Logistique")

