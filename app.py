import streamlit as st
import pandas as pd
import io
import time

import geo_utils as gu

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(
    page_title="Lab_Math Smart Deduplicator",
    page_icon="🧬",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- STYLES CSS PERSONNALISÉS (Branding) ---
# C'est ici que j'ai modifié la couleur du pied de page
st.markdown("""
<style>
    /* Style du pied de page */
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: #0E1117; /* Fond noir */
        color: #FAFAFA; /* Texte principal blanc */
        text-align: center;
        padding: 10px;
        font-size: 12px;
        border-top: 1px solid #333;
        z-index: 1000;
    }
    /* C'EST ICI QUE LA COULEUR A CHANGÉ */
    .footer a, .footer-highlight {
        color: #4DA8DA !important; /* <-- BLEU PROFESSIONNEL (au lieu du rouge #FF4B4B) */
        text-decoration: none;
        font-weight: bold;
    }
    /* Style des onglets */
    .stTabs [data-baseweb="tab-list"] {
        gap: 24px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #f0f2f6;
        border-radius: 4px 4px 0px 0px;
        padding-left: 10px;
        padding-right: 10px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #FFFFFF;
        border-top: 2px solid #4DA8DA; /* Changé aussi le haut de l'onglet actif en bleu */
    }
    /* Masquer le pied de page par défaut de Streamlit */
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# --- FONCTIONS MÉTIERS ---

@st.cache_data(show_spinner=False)
def charger_donnees(fichier):
    """Charge les données et retourne un DataFrame."""
    extension = fichier.name.split('.')[-1].lower()
    try:
        if extension in ['xlsx', 'xls']:
            return pd.read_excel(fichier)
        elif extension == 'csv':
            try:
                return pd.read_csv(fichier)
            except:
                return pd.read_csv(fichier, sep=';')
        elif extension == 'txt':
            return pd.read_csv(fichier, sep='\t')
        elif extension == 'sav':
            # Nécessite pyreadstat
            return pd.read_spss(fichier)
        elif extension == 'dta':
            # Nécessite pyreadstat
            return pd.read_stata(fichier)
    except Exception as e:
        st.error(f"Erreur technique lors du chargement : {e}")
        return None
    return None

@st.cache_data(show_spinner=False)
def charger_geo(kind, fichiers):
    """Charge un jeu de données géospatial (shapefile, KML, KMZ, GeoJSON, GPKG, ZIP).
    `fichiers` est un tuple de (nom, octets). Retourne (GeoDataFrame, notes)."""
    return gu.lire_geospatial(kind, list(fichiers))

def algo_resolution_intelligente(df, cols_cles):
    """
    Algorithme 'Smart Resolve' :
    1. Calcule un score de qualité (nombre de cellules remplies) pour chaque ligne.
    2. Trie les doublons pour mettre la ligne la plus complète en premier.
    3. Garde la première ligne (la meilleure) et rejette les autres.
    """
    # 1. Calcul du Score de Qualité (Densité de l'information)
    df = df.copy()
    # On compte les valeurs non nulles par ligne
    df['__Score_Qualite__'] = df.notna().sum(axis=1)
    
    # 2. Tri intelligent : D'abord par les clés (pour grouper), puis par Score décroissant
    # ascending=[True... pour les clés, False pour le score]
    sort_order = [True] * len(cols_cles) + [False]
    df_sorted = df.sort_values(by=cols_cles + ['__Score_Qualite__'], ascending=sort_order)
    
    # 3. Séparation
    # Base propre : on garde la première occurrence (celle avec le meilleur score grâce au tri)
    df_clean = df_sorted.drop_duplicates(subset=cols_cles, keep='first').copy()
    
    # Base rejetée (les doublons moins bons)
    # On prend les index qui ne sont PAS dans la base propre
    df_rejets = df_sorted[~df_sorted.index.isin(df_clean.index)].copy()
    
    # Nettoyage de la colonne temporaire
    df_clean = df_clean.drop(columns=['__Score_Qualite__'])
    df_rejets = df_rejets.drop(columns=['__Score_Qualite__'])
    
    return df_clean, df_rejets

def to_excel(df):
    """Convertit un DataFrame en fichier Excel binaire pour téléchargement."""
    output = io.BytesIO()
    # Nécessite xlsxwriter
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Donnees')
    return output.getvalue()

# --- INTERFACE PRINCIPALE ---

TYPES_ACCEPTES = ['xlsx', 'xls', 'csv', 'txt', 'sav', 'dta',
                  'shp', 'shx', 'dbf', 'prj', 'cpg', 'kml', 'kmz', 'geojson', 'json', 'gpkg', 'zip']


def vue_apercu(df_ou_gdf, masque, ordre=None):
    """Table d'aperçu des lignes concernées ; ajoute la géométrie en WKT court pour les données SIG.
    `ordre` : index (déjà trié) servant à regrouper les doublons côte à côte."""
    sous = df_ou_gdf[masque]
    if isinstance(sous, gu.gpd.GeoDataFrame):
        vue = pd.DataFrame(sous.drop(columns=sous.geometry.name))
        vue["géométrie"] = sous.geometry.to_wkt(rounding_precision=6).str.slice(0, 80)
    else:
        vue = sous
    if ordre is not None:
        vue = vue.loc[ordre.loc[vue.index].sort_values(kind="stable").index]
    return vue


def reinitialiser_diagnostic():
    st.session_state.search_done = False
    st.session_state.resolved = False
    st.session_state.cols_cles = []
    st.session_state.nb_doublons = 0


def main():
    # En-tête
    st.title("🧬 Smart Deduplicator Pro")
    st.write("### Vérification et Traitement Intelligent des Doublons (données tabulaires et géospatiales)")
    st.markdown("---")

    # Initialisation des états de session
    if 'data' not in st.session_state:
        st.session_state.data = None
        st.session_state.gdf = None
        st.session_state.is_geo = False
        st.session_state.notes = []
        st.session_state.signature = None
        st.session_state.search_done = False
        st.session_state.resolved = False
        st.info("👋 Bienvenue. Formats supportés : **.xlsx, .csv, .txt, .sav (SPSS), .dta (Stata)** et, pour les données "
                "géospatiales, **.shp (avec .shx, .dbf, .prj), .kml, .kmz, .geojson, .gpkg ou une archive .zip**.")

    # 1. IMPORTATION
    uploaded_files = st.file_uploader(
        "📂 Téléchargez votre base de données ici (pour un shapefile, ajoutez tous ses fichiers : .shp, .shx, .dbf, .prj)",
        type=TYPES_ACCEPTES, accept_multiple_files=True)

    if uploaded_files:
        jeux = gu.regrouper_fichiers(uploaded_files)
        libelles = list(jeux.keys())
        if len(libelles) > 1:
            choix = st.selectbox("🗂️ Plusieurs jeux de données détectés : lequel traiter ?", libelles)
        else:
            choix = libelles[0]
        jeu = jeux[choix]
        signature = (choix, tuple(sorted((f.name, f.size) for f in jeu["objs"])))

        # Chargement initial ou si le fichier change
        if st.session_state.data is None or st.session_state.signature != signature:
            with st.spinner("🔄 Analyse et chargement du fichier en cours..."):
                try:
                    if jeu["kind"] == "table":
                        df_charge = charger_donnees(jeu["objs"][0])
                        gdf_charge, notes = None, []
                    else:
                        gdf_charge, notes = charger_geo(jeu["kind"], tuple((f.name, f.getvalue()) for f in jeu["objs"]))
                        df_charge = pd.DataFrame(gdf_charge.drop(columns=gdf_charge.geometry.name))
                except Exception as e:
                    st.error(f"Erreur technique lors du chargement : {e}")
                    df_charge, gdf_charge, notes = None, None, []
                if df_charge is not None:
                    st.session_state.data = df_charge
                    st.session_state.gdf = gdf_charge
                    st.session_state.is_geo = gdf_charge is not None
                    st.session_state.notes = notes
                    st.session_state.signature = signature
                    st.session_state.filename = choix
                    reinitialiser_diagnostic()
                    st.toast("Fichier chargé avec succès !", icon="✅")

        # Si les données sont chargées, on affiche l'interface
        if st.session_state.data is not None:
            df = st.session_state.data
            est_geo = st.session_state.is_geo
            gdf = st.session_state.gdf
            nom_base = st.session_state.filename
            st.success(f"✅ **Base chargée :** {nom_base} | **{len(df):,}** lignes | **{len(df.columns)}** colonnes".replace(",", " ")
                       + (" (attributs) + géométrie" if est_geo else ""))
            for note in st.session_state.notes:
                st.warning("ℹ️ " + note)

            # --- ONGLETS DE TRAVAIL ---
            noms_onglets = ["🔍 1. Recherche & Diagnostic", "✨ 2. Résolution Intelligente"]
            if est_geo:
                noms_onglets.append("🗺️ 3. Contrôle géométrique")
            onglets = st.tabs(noms_onglets)
            tab1, tab2 = onglets[0], onglets[1]

            # === ONGLET 1 : RECHERCHE ===
            with tab1:
                st.subheader("Configuration de la recherche")
                st.write("Sélectionnez les variables qui définissent l'unicité d'un individu (ex: ID unique, ou combinaison Nom + Date de naissance)."
                         + (" Pour les données géospatiales, vous pouvez aussi choisir la **position** (deux objets au même endroit sont en doublon)." if est_geo else ""))

                options = list(df.columns) + ([gu.GEOM_LABEL] if est_geo else [])
                col_sel1, col_sel2 = st.columns([3, 1])
                with col_sel1:
                    cols_cles_sel = st.multiselect(
                        "Variables identifiantes (Clé primaire) :",
                        options=options,
                        placeholder="Choisissez une ou plusieurs colonnes..."
                    )
                    decimales = 6
                    if est_geo and gu.GEOM_LABEL in cols_cles_sel:
                        decimales = st.slider(
                            "Précision de la comparaison des positions (décimales des coordonnées)",
                            min_value=0, max_value=9, value=gu.suggerer_decimales(gdf),
                            help="6 décimales en degrés ≈ 0,1 m. Baissez la valeur pour regrouper des positions voisines "
                                 "(5 ≈ 1 m, 4 ≈ 10 m). Pour un système projeté en mètres, 1 décimale = 10 cm.")

                with col_sel2:
                    st.write("")  # Espace pour aligner
                    st.write("")
                    bouton_check = st.button("Lancer la recherche 🔎", type="primary", width="stretch", disabled=not cols_cles_sel)

                if bouton_check:
                    with st.spinner("Algorithme de détection en cours..."):
                        time.sleep(0.5)
                        utiliser_geom = est_geo and gu.GEOM_LABEL in cols_cles_sel
                        cols_attr = [c for c in cols_cles_sel if c != gu.GEOM_LABEL]
                        if est_geo:
                            table, cles = gu.table_cles(gdf, cols_attr, utiliser_geom, decimales)
                            nb_doublons = int(table.duplicated(subset=cles).sum())
                        else:
                            nb_doublons = int(df.duplicated(subset=cols_attr).sum())
                        st.session_state.cols_cles = cols_cles_sel
                        st.session_state.cols_attr = cols_attr
                        st.session_state.utiliser_geom = utiliser_geom
                        st.session_state.decimales = decimales
                        st.session_state.nb_doublons = nb_doublons
                        st.session_state.search_done = True
                        st.session_state.resolved = False  # Reset résolution si nouvelle recherche

                # Affichage des résultats de recherche (si une recherche a été faite)
                if st.session_state.search_done:
                    nb = st.session_state.nb_doublons
                    keys = st.session_state.cols_cles

                    st.divider()
                    if nb > 0:
                        st.warning(f"⚠️ **Diagnostic terminé : {nb:,} lignes en doublon trouvées** sur la base des variables : *{', '.join(keys)}*.".replace(",", " "))
                        if est_geo and st.session_state.utiliser_geom and not st.session_state.cols_attr:
                            st.info("ℹ️ La recherche porte sur la position seule. Deux enregistrements différents peuvent légitimement partager "
                                    "le même point : vérifiez avec les attributs (ajoutez des colonnes comme clé) avant de supprimer.")
                        st.info("👉 Veuillez passer à l'onglet **'Résolution Intelligente'** pour corriger automatiquement ce problème.")

                        with st.expander("👁️ Voir un aperçu des lignes concernées (Groupées)"):
                            if est_geo:
                                table, cles = gu.table_cles(gdf, st.session_state.cols_attr, st.session_state.utiliser_geom, st.session_state.decimales)
                                masque = table.duplicated(subset=cles, keep=False)
                                # regroupe les doublons : tri par position, sinon par les colonnes clés
                                ordre = table["__geom__"] if "__geom__" in table else table[cles].astype(str).agg("|".join, axis=1)
                                doublons_view = vue_apercu(gdf, masque, ordre)
                            else:
                                doublons_view = df[df.duplicated(subset=st.session_state.cols_attr, keep=False)].sort_values(by=st.session_state.cols_attr)
                            st.dataframe(doublons_view.head(1000), width="stretch")
                            if len(doublons_view) > 1000:
                                st.caption("Aperçu limité aux 1000 premières lignes de doublons.")
                    else:
                        st.success(f"✅ **Diagnostic terminé : Aucun doublon détecté.** Votre base est propre selon les critères : *{', '.join(keys)}*.")
                        st.balloons()

            # === ONGLET 2 : RÉSOLUTION ===
            with tab2:
                st.subheader("Résolution & Exportation")

                # Blocage si la recherche n'est pas faite ou s'il n'y a pas de doublons
                if not st.session_state.search_done:
                    st.warning("Veuillez d'abord lancer une recherche dans l'onglet 1.")
                elif st.session_state.nb_doublons == 0:
                    st.success("Aucun doublon à résoudre. Votre base est déjà propre !")
                else:
                    st.markdown("""
                    <div style="background-color: #f0f2f6; padding: 15px; border-radius: 5px; margin-bottom: 20px;">
                        <strong>🧠 Logique de l'Intelligence Artificielle (Algorithme de Complétude) :</strong><br>
                        L'application analyse chaque groupe de doublons détecté et <b>conserve automatiquement la ligne qui contient le plus d'informations</b> (celle avec le moins de cellules vides). C'est la solution optimale sans équivoque.
                    </div>
                    """, unsafe_allow_html=True)

                    col_res1, col_res2 = st.columns([2, 1])
                    with col_res1:
                        if st.button("🚀 Appliquer la Résolution Intelligente", type="primary", width="stretch"):
                            with st.spinner("Traitement intelligent et séparation des données..."):
                                if est_geo:
                                    clean, rejets = gu.resolution_intelligente_geo(
                                        gdf, st.session_state.cols_attr, st.session_state.utiliser_geom, st.session_state.decimales)
                                else:
                                    clean, rejets = algo_resolution_intelligente(df, st.session_state.cols_attr)
                                st.session_state.df_clean = clean
                                st.session_state.df_rejets = rejets
                                st.session_state.resolved = True
                                st.toast("Nettoyage effectué avec succès !", icon="✨")

                    # Affichage des résultats après résolution
                    if st.session_state.resolved:
                        st.divider()
                        st.success("✅ **Problème résolu.** La base propre a été générée en conservant les enregistrements les plus complets.")

                        # Métriques Avant / Après
                        kpi1, kpi2, kpi3 = st.columns(3)
                        kpi1.metric("Lignes Totales (Avant)", f"{len(df):,}".replace(",", " "))
                        kpi2.metric("Base Propre (Après)", f"{len(st.session_state.df_clean):,}".replace(",", " "), delta=f"- {len(st.session_state.df_rejets)} doublons retirés", delta_color="inverse")
                        kpi3.metric("Rejets (Doublons moins bons)", f"{len(st.session_state.df_rejets):,}".replace(",", " "))

                        st.divider()
                        st.write("### 💾 Zone de Téléchargement")

                        if est_geo:
                            fmt = st.selectbox("Format d'export", list(gu.FORMATS_EXPORT.keys()))
                            c_dl1, c_dl2 = st.columns(2)
                            for colonne, cle_df, etiquette, fichier, legende, principal in (
                                (c_dl1, "df_clean", "📥 Télécharger la BASE PROPRE", "Base_Nettoyee_Smart", "C'est votre fichier final, prêt à l'emploi.", True),
                                (c_dl2, "df_rejets", "🗑️ Télécharger les REJETS (Audit)", "Doublons_Rejetes_Audit", "Pour archivage ou vérification manuelle si nécessaire.", False),
                            ):
                                with colonne:
                                    try:
                                        octets, ext, mime = gu.exporter_geo(st.session_state[cle_df], fmt)
                                        st.download_button(label=f"{etiquette} (.{ext})", data=octets, file_name=f"{fichier}.{ext}",
                                                           mime=mime, width="stretch", type="primary" if principal else "secondary")
                                        st.caption(legende)
                                    except Exception as e:
                                        st.error(f"Export impossible dans ce format : {e}")
                        else:
                            c_dl1, c_dl2 = st.columns(2)
                            with c_dl1:
                                st.download_button(
                                    label="📥 Télécharger la BASE PROPRE (.xlsx)",
                                    data=to_excel(st.session_state.df_clean),
                                    file_name="Base_Nettoyee_Smart.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    width="stretch",
                                    type="primary"
                                )
                                st.caption("C'est votre fichier final, prêt à l'emploi.")
                            with c_dl2:
                                st.download_button(
                                    label="🗑️ Télécharger les REJETS (Audit .xlsx)",
                                    data=to_excel(st.session_state.df_rejets),
                                    file_name="Doublons_Rejetes_Audit.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    width="stretch"
                                )
                                st.caption("Pour archivage ou vérification manuelle si nécessaire.")

            # === ONGLET 3 : CONTRÔLE GÉOMÉTRIQUE (données SIG uniquement) ===
            if est_geo:
                with onglets[2]:
                    st.subheader("Contrôle et apurement des géométries")
                    d = gu.diagnostic_geometrique(gdf)
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("Objets", f"{d['lignes']:,}".replace(",", " "))
                    m2.metric("Géométries absentes ou vides", d["nulles"] + d["vides"])
                    m3.metric("Géométries invalides", d["invalides"])
                    m4.metric("Hors Cameroun (WGS 84)", d["hors_cameroun"] if d["hors_cameroun"] is not None else "n/d")
                    st.write(f"**Types de géométrie :** {', '.join(f'{k} ({v})' for k, v in d['types'].items()) or 'aucun'}")
                    st.write(f"**Système de coordonnées :** {d['crs'] or 'inconnu'}")
                    if d["bornes"]:
                        st.write(f"**Étendue (xmin, ymin, xmax, ymax) :** {d['bornes']}")
                    st.caption("Les valeurs « hors Cameroun » indiquent des objets dont les coordonnées sortent de l'emprise du pays : "
                               "souvent une inversion latitude/longitude ou une erreur de saisie.")
                    st.divider()
                    st.write("#### Actions d'apurement")
                    a1, a2 = st.columns(2)
                    with a1:
                        if st.button("🛠️ Réparer les géométries invalides", disabled=d["invalides"] == 0, width="stretch"):
                            st.session_state.gdf, n = gu.reparer_geometries(gdf)
                            reinitialiser_diagnostic()
                            st.toast(f"{n} géométrie(s) réparée(s).", icon="🛠️")
                            st.rerun()
                    with a2:
                        if st.button("🧹 Supprimer les géométries absentes ou vides", disabled=(d["nulles"] + d["vides"]) == 0, width="stretch"):
                            st.session_state.gdf, n = gu.supprimer_vides(gdf)
                            st.session_state.data = pd.DataFrame(st.session_state.gdf.drop(columns=st.session_state.gdf.geometry.name))
                            reinitialiser_diagnostic()
                            st.toast(f"{n} objet(s) supprimé(s).", icon="🧹")
                            st.rerun()
                    st.write("#### Système de coordonnées")
                    c1, c2, c3 = st.columns([2, 2, 1])
                    with c1:
                        code_epsg = st.number_input("Code EPSG", min_value=1000, max_value=999999, value=4326, step=1)
                    with c2:
                        st.write("")
                        st.write("")
                        if gdf.crs is None:
                            if st.button("📌 Définir ce système (données sans CRS)", width="stretch"):
                                st.session_state.gdf = gdf.set_crs(int(code_epsg))
                                reinitialiser_diagnostic()
                                st.rerun()
                        else:
                            if st.button("🔄 Reprojeter vers ce système", width="stretch"):
                                st.session_state.gdf = gdf.to_crs(int(code_epsg))
                                reinitialiser_diagnostic()
                                st.rerun()
                    with c3:
                        st.caption("4326 : WGS 84 (degrés). 32632 / 32633 : UTM 32N / 33N (mètres).")
                    st.info("Après chaque action, relancez la recherche de doublons (onglet 1) : elle s'appuie sur la table à jour.")

    # --- PIED DE PAGE (Footer Fixe) ---
    # J'ai ajouté des spans avec la classe 'footer-highlight' pour cibler précisément les textes à colorer
    st.markdown("""
    <div class="footer">
        <p>Power by <b>Lab_Math</b>, du groupe <b>SCSM et CIE</b> (<a href="http://www.scsmaubmar.org" target="_blank">www.scsmaubmar.org</a>) et copyright 2026.</p>
        <p><span class="footer-highlight">📞 Contact:</span> +237 620 30 74 39 / +237 674 65 18 56 &nbsp; | &nbsp; <span class="footer-highlight">✉️ Email:</span> <a href="mailto:scsm@scsmaubmar.org">scsm@scsmaubmar.org</a> ou <a href="mailto:marcialteda@scsmaubmar.org">marcialteda@scsmaubmar.org</a></p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()