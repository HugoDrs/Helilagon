import streamlit as st
import tempfile
import os
import struct
import re

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

_PN_RE     = re.compile(r'^[A-Z0-9][A-Z0-9\-\.]{2,49}$')
_PN_SEP_RE = re.compile(r'[0-9\-\.]')

def is_valid_pn(s):
    if not isinstance(s, str):
        return False
    s = s.strip()
    return bool(_PN_RE.match(s) and _PN_SEP_RE.search(s))

_IGNORED_PREFIXES = (
    'Livraison avant le', 'Livraison\n', 'Toutes les pi', 'Total',
    'HELILAGON', 'Destinataire', 'Facturation', 'Cpte fourn',
    'Payment', 'Check', 'Expédiée',
)

def is_ignored_text(s):
    if not isinstance(s, str):
        return False
    return any(s.startswith(p) for p in _IGNORED_PREFIXES)

def u16(data, pos): return struct.unpack_from('<H', data, pos)[0]
def u32(data, pos): return struct.unpack_from('<I', data, pos)[0]
def f64(data, pos): return struct.unpack_from('<d', data, pos)[0]

def find_sst_offset(data):
    pos = 0
    while True:
        idx = data.find(b'\xfc\x00', pos)
        if idx == -1:
            return None
        if u16(data, idx + 2) >= 8:
            return idx
        pos = idx + 1

def parse_sst(data, sst_offset):
    unique_count = u32(data, sst_offset + 8)
    pos          = sst_offset + 12
    strings      = []
    for _ in range(unique_count):
        if pos + 3 > len(data):
            strings.append(None)
            continue
        str_len    = u16(data, pos)
        flags      = data[pos + 2]
        compressed = not (flags & 0x01)
        has_rich   = bool(flags & 0x08)
        has_phonet = bool(flags & 0x04)
        pos += 3
        rich_count = phonetic_size = 0
        if has_rich:
            rich_count = u16(data, pos);    pos += 2
        if has_phonet:
            phonetic_size = u32(data, pos); pos += 4
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
        except Exception:
            s = None
        pos += byte_count + rich_count * 4 + phonetic_size
        strings.append(s)
    return strings

def parse_early_pool(data, sst_offset):
    pns = []
    pos = 0
    while pos < sst_offset - 3:
        if pos + 3 > len(data):
            break
        str_len = u16(data, pos)
        flags   = data[pos + 2]
        if flags == 0x01 and 3 <= str_len <= 50:
            byte_count = str_len * 2
            end        = pos + 3 + byte_count
            if end <= sst_offset:
                try:
                    s = data[pos + 3:end].decode('utf-16-le', errors='strict')
                    if s.isprintable() and is_valid_pn(s):
                        pns.append(s)
                except Exception:
                    pass
        pos += 1
    return pns

def build_fallback_pn_list(early_pns):
    result = []
    for pn in early_pns:
        if pn.endswith('LE') and len(pn) > 4:
            truncated = pn[:-2]
            if is_valid_pn(truncated) and truncated not in result:
                result.append(truncated)
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
            val     = strings[sst_idx] if sst_idx < len(strings) else None
            cells[(row, col)] = val
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
    if not fallback_pns:
        return cells
    broken_rows = sorted(
        row for (row, col), val in cells.items()
        if col == COL_PN and val is None
    )
    for row, pn in zip(broken_rows, fallback_pns):
        cells[(row, COL_PN)] = pn
    return cells

def convert_xls_bytes_to_csv(xls_bytes):
    """Convertit les bytes d'un .xls en contenu CSV (str)."""
    data = xls_bytes

    sst_offset = find_sst_offset(data)
    if sst_offset is None:
        raise ValueError("Table SST introuvable — vérifiez que le fichier est bien un .xls Airstock.")

    strings      = parse_sst(data, sst_offset)
    early_pns    = parse_early_pool(data, sst_offset)
    fallback_pns = build_fallback_pn_list(early_pns)
    cells        = parse_cells(data, strings)
    cells        = fill_corrupted_pns(cells, fallback_pns)

    all_rows = sorted(set(row for (row, _) in cells))
    results  = []

    for row in all_rows:
        if row <= HEADER_ROW:
            continue
        pn_val  = cells.get((row, COL_PN))
        qty_val = cells.get((row, COL_QTY))
        if qty_val is None:
            continue
        pn_str = str(pn_val).strip() if pn_val is not None else ''
        if not is_valid_pn(pn_str) or is_ignored_text(pn_str):
            continue
        try:
            qty = int(round(float(qty_val)))
        except (ValueError, TypeError):
            continue
        if qty <= 0:
            continue
        results.append((pn_str, qty))

    csv_lines = ["PN;Quantité"]
    for pn, qty in results:
        csv_lines.append(f"{pn};{qty}")

    return "\n".join(csv_lines), results

# ── Interface Streamlit ───────────────────────────────────────────────────────

st.title("🚁 Airstock → Airbus CSV")
st.markdown(
    "Convertit automatiquement un fichier `.xls` exporté depuis **Airstock** "
    "en fichier `.csv` prêt à être déposé sur le portail **Airbus Helicopters** "
    "pour l'import multi-PN."
)

st.divider()

# ── Instructions d'export Airstock ───────────────────────────────────────────

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
    with col2: st.markdown("Clique **Enregistrer** — ton fichier `.xls` est prêt à être déposé ci-dessous ✅")

    st.info(
        "💡 **Attention au format !** Si tu choisis un autre format que "
        "*\"Données uniquement (*.xls)\"*, le fichier ne sera pas reconnu par cet outil.",
        icon="⚠️"
    )

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
            csv_bytes    = csv_content.encode('utf-8-sig')  # BOM pour Excel

            st.success(f"✅ **{len(results)} Part Number(s)** extraits avec succès !")

            # Tableau récapitulatif
            with st.expander("📋 Voir le détail des lignes extraites"):
                import pandas as pd
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

            st.markdown(
                "➡️ Dépose ensuite ce fichier `.csv` sur le portail Airbus Helicopters "
                "pour importer tous tes PN en une seule fois."
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
st.caption("Outil interne — Helilagon Logistique")
