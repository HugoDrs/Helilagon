import streamlit as st
import io
import re
import pandas as pd

# ── Configuration de la page ──────────────────────────────────────────────────

st.set_page_config(
    page_title="Airstock → Airbus CSV",
    page_icon="🚁",
    layout="centered"
)

# ── Fonctions de conversion (moteur) ─────────────────────────────────────────

# Mots-clés présents dans les lignes de pied de page / en-tête à ignorer
FOOTER_KEYWORDS = (
    'Total', 'Toutes les', 'HELILAGON', 'Page', 'COMMANDE',
    'EUR', 'Livraison avant', 'PN - description', 'Quantité',
    'Payment terms', 'Cpte fourn', 'Destinataire',
)

def is_footer(val):
    """Retourne True si la valeur correspond à une ligne à ignorer."""
    if val is None:
        return True
    s = str(val).strip()
    if not s or s == 'nan':
        return True
    # Lignes multi-paragraphes (cellules fusionnées d'en-tête)
    if '\n' in s:
        return True
    # Mots-clés de pied de page / en-tête
    if any(kw in s for kw in FOOTER_KEYWORDS):
        return True
    # Date (ex: "18/04/2026")
    if re.match(r'\d{2}/\d{2}/\d{4}', s):
        return True
    return False

def convert_xls_bytes_to_csv(xls_bytes):
    """
    Lit le fichier .xls Airstock avec pandas/xlrd (moteur fiable).
    Structure du fichier :
      - Lignes 0–3 : en-têtes à ignorer (header=None)
      - Col 1 (B) : Part Number
      - Col 2 (C) : Quantité
    Entre chaque article, Airstock insère une ligne "Livraison avant le…"
    et une ligne vide — pandas les lit normalement, elles sont filtrées
    par is_footer() et par le test qty > 0.
    """
    df = pd.read_excel(io.BytesIO(xls_bytes), engine='xlrd', header=None)

    results = []
    for i, row in df.iterrows():
        # Ignorer les 4 premières lignes (en-têtes Airstock)
        if i < 4:
            continue

        pn_val  = row[1]
        qty_val = row[2]

        # Ignorer si PN absent ou ligne de pied de page
        if pd.isna(pn_val) or is_footer(pn_val):
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

        # Formater : entier si .0, sinon décimal avec virgule (norme française)
        if qty_float == int(qty_float):
            qty_str = str(int(qty_float))
        else:
            qty_str = str(qty_float).replace('.', ',')

        results.append((pn_str, qty_str))

    csv_lines = ["Ordered Reference;Quantity"]
    for pn, qty_str in results:
        csv_lines.append(f"{pn};{qty_str}")

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
    with col2: st.markdown("Dans l'Explorateur de fichiers, clique sur **Ce PC** dans le panneau de gauche → tu verras apparaître ton ordinateur local sous la forme **`Lecteur (\\tsclient\...)`** ou similaire → ouvre-le")

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
                     "Quantité":    [r[1] for r in results]},  # r[1] est déjà une str formatée
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
