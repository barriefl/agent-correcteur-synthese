import base64
import io
import json
import os
import re
import shutil
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st

# --- INITIALISATION DU STATE. ---
if "session_active" not in st.session_state:
    st.session_state.session_active = False


# --- NETTOYAGE DES SESSIONS VIDES. ---
def nettoyer_sessions_vides(dossier_racine="./data"):
    if not os.path.exists(dossier_racine):
        return

    for nom_dossier in os.listdir(dossier_racine):
        chemin_dossier = os.path.join(dossier_racine, nom_dossier)

        if os.path.isdir(chemin_dossier):
            if not os.listdir(chemin_dossier):
                try:
                    shutil.rmtree(chemin_dossier)
                except any:
                    pass


nettoyer_sessions_vides()

svg_icon = """
<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#FF4B4B" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
    <path d="M22 10v6M2 10l10-5 10 5-10 5z"/><path d="M6 12v5c3 3 9 3 12 0v-5"/>
</svg>
"""


def svg_to_data_uri(svg_str):
    b64 = base64.b64encode(svg_str.encode("utf-8")).decode("utf-8")
    return f"data:image/svg+xml;base64,{b64}"


st.set_page_config(
    page_title="Assistant de Correction IA",
    page_icon=svg_to_data_uri(svg_icon),
    layout="wide",
)

# --- DÉFINITION DES ICÔNES SVG. ---
ICON_UPLOAD = '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>'
ICON_GEAR = '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>'
ICON_CHART = '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>'
ICON_SEARCH = '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>'
ICON_TRASH = '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>'
ICON_SAVE = '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>'

SVG_CHECK = "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9IiMyM2M1NWUiIHN0cm9rZS13aWR0aD0iMyIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW51PSJyb3VuZCI+PHBvbHlsaW5lIHBvaW50cz0iMjAgNiA5IDE3IDQgMTIiPjwvcG9seWxpbmU+PC9zdmc+"
SVG_ERROR = "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9IiNmODcxNzEiIHN0cm9rZS13aWR0aD0iMyIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW51PSJyb3VuZCI+PGxpbmUgeDE9IjE4IiB5MT0iNiIgeDI9IjYiIHkyPSIxOCI+PC9saW5lPjxsaW5lIHgxPSI2IiB5MT0iNiIgeDI9IjE4IiB5Mj0iMTgiPjwvbGluZT48L3N2Zz4="
SVG_WAIT = "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9IiM5NGExYjIiIHN0cm9rZS13aWR0aD0iMyIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW51PSJyb3VuZCI+PGNpcmNsZSBjeD0iMTIiIGN5PSIxMiIgcj0iMTAiPjwvY2lyY2xlPjxwb2x5bGluZSBwb2ludHM9IjEyIDYgMTIgMTIgMTYgMTQiPjwvcG9seWxpbmU+PC9zdmc+"

st.markdown("<h1>Assistant de Correction IA</h1>", unsafe_allow_html=True)
st.markdown(
    "Uploadez une archive `.zip` contenant les dossiers des étudiants, puis lancez l'analyse."
)

# --- BARRE LATÉRALE : PARAMÈTRES ET SESSION. ---
with st.sidebar:
    st.markdown(f"<h3>{ICON_GEAR} Configuration</h3>", unsafe_allow_html=True)
    api_key_input = st.text_input(
        "Clé API Gemini",
        type="password",
        help="Clé sur Google AI Studio (https://aistudio.google.com).",
    )

    st.divider()

    st.markdown(f"<h3>{ICON_SEARCH} Session</h3>", unsafe_allow_html=True)
    session_name_input = st.text_input(
        "Identifiant de session",
        value="MaSession_01",
        help="Identifiant pour retrouver vos copies.",
    )

    col_btn_1, col_btn_2 = st.columns(2)
    with col_btn_1:
        if st.button("Valider"):
            if session_name_input:
                st.session_state.session_active = True
                st.rerun()
            else:
                st.error("ID Session requis")
    with col_btn_2:
        if st.button("Quitter"):
            st.session_state.session_active = False
            st.rerun()


# --- LOGIQUE D'AFFICHAGE. ---
if not st.session_state.session_active:
    st.info("Bienvenue ! Veuillez configurer votre session dans le menu latéral pour accéder à l'outil.")
    st.stop()

safe_session_name = re.sub(r"[^a-zA-Z0-9_-]", "", session_name_input)
USER_DATA_PATH = os.path.join("./data", safe_session_name)

if not os.path.exists(USER_DATA_PATH):
    pass


def preparer_dossier_data():
    """Nettoie le dossier data pour éviter de mélanger les anciennes copies avec les nouvelles."""
    if os.path.exists(USER_DATA_PATH):
        shutil.rmtree(USER_DATA_PATH)
    os.makedirs(USER_DATA_PATH)


def obtenir_statuts(dossiers_list):
    """Génère un tableau de suivi basé sur les fichiers présents dans chaque dossier."""
    donnees = []
    for dossier in dossiers_list:
        chemin = os.path.join(USER_DATA_PATH, dossier)

        if os.path.exists(os.path.join(chemin, "rapport_ia_brut.json")):
            icon, label = SVG_CHECK, "Analysé"
        elif os.path.exists(os.path.join(chemin, "erreur_json_brut.txt")):
            icon, label = SVG_ERROR, "Erreur"
        else:
            icon, label = SVG_WAIT, "En attente"

        donnees.append({"Statut": icon, "Dossier du Groupe": dossier, "État": label})

    return pd.DataFrame(donnees)


# --- DÉPÔT DES COPIES. ---
st.markdown(f"<h3>{ICON_UPLOAD} Dépôt des copies</h3>", unsafe_allow_html=True)

fichier_zip = st.file_uploader(
    "Glissez le fichier ZIP contenant les dossiers de groupes", type=["zip"]
)

if fichier_zip is not None:
    if st.button("Charger / Décompresser les dossiers", width="stretch"):
        preparer_dossier_data()

        chemin_zip_temp = os.path.join(USER_DATA_PATH, "temp.zip")
        with open(chemin_zip_temp, "wb") as f:
            f.write(fichier_zip.getbuffer())

        with zipfile.ZipFile(chemin_zip_temp, "r") as zip_ref:
            zip_ref.extractall(USER_DATA_PATH)
        os.remove(chemin_zip_temp)

        st.success("Chargement réussi !")
        st.rerun()

# --- VÉRIFICATION ET STATUTS. ---
dossiers = []
if os.path.exists(USER_DATA_PATH):
    dossiers = [
        d
        for d in os.listdir(USER_DATA_PATH)
        if os.path.isdir(os.path.join(USER_DATA_PATH, d)) and not d.startswith("__")
    ]

if len(dossiers) > 0:
    st.info(f"**{len(dossiers)} groupe(s)** détecté(s) :")
    df_statuts = obtenir_statuts(dossiers)
    st.dataframe(
        df_statuts,
        column_config={
            "Statut": st.column_config.ImageColumn(" ", width="small"),
            "Dossier du Groupe": st.column_config.TextColumn("Dossier du Groupe"),
            "État": st.column_config.TextColumn("État"),
        },
        width="stretch",
        hide_index=True,
    )

    st.divider()

    # --- ACTIONS (LANCEMENT ET SUPPRESSION). ---
    col_run, col_export, col_delete = st.columns([2, 1, 1])

    with col_run:
        st.markdown(f"<h3>{ICON_GEAR} Analyse IA</h3>", unsafe_allow_html=True)
        import correcteur_ia

        if st.button("Lancer l'Analyse IA", type="primary", width="stretch"):
            if not api_key_input:
                st.error("Veuillez entrer votre clé API Gemini dans le menu de gauche.")
            elif len(dossiers) == 0:
                st.error("Ajoutez et décompressez d'abord des dossiers.")
            else:
                progress_bar = st.progress(0)
                status_text = st.empty()

                def update_progress(current, total, nom_dossier):
                    if total > 0:
                        percent = int((current / total) * 100)
                        progress_bar.progress(percent)
                        if current < total:
                            status_text.info(
                                f"Analyse en cours ({current + 1}/{total}) : **{nom_dossier}**..."
                            )
                        else:
                            status_text.success("Terminée !")

                try:
                    correcteur_ia.lancer_analyse_globale(
                        update_progress, USER_DATA_PATH, api_key_input
                    )
                    st.success("Correction des copies terminée !")
                    st.rerun()
                except Exception as e:
                    if str(e) == "QUOTA_429":
                        status_text.error(
                            "STOP : Quota API épuisé pour aujourd'hui (Erreur 429)."
                        )
                    else:
                        status_text.error(f"Une erreur inattendue est survenue : {e}")

    with col_export:
        st.markdown(f"<h3>{ICON_SAVE} Sauvegarde</h3>", unsafe_allow_html=True)
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "x", zipfile.ZIP_DEFLATED) as zf:
            for root, dirs, files in os.walk(USER_DATA_PATH):
                for file in files:
                    if file != "Resultats_Corrections.xlsx":
                        file_path = os.path.join(root, file)
                        zf.write(file_path, os.path.relpath(file_path, USER_DATA_PATH))

        st.download_button(
            label="Exporter le ZIP",
            data=buf.getvalue(),
            file_name=f"Sauvegarde_Correction_{datetime.now().strftime('%d_%m')}.zip",
            mime="application/zip",
            width="stretch",
            help="Téléchargez ce ZIP pour reprendre votre travail plus tard sur n'importe quel PC.",
        )

    with col_delete:
        st.markdown(f"<h3>{ICON_TRASH} Réinitialiser</h3>", unsafe_allow_html=True)

        if "confirm_delete" not in st.session_state:
            st.session_state.confirm_delete = False

        if not st.session_state.confirm_delete:
            if st.button("Tout supprimer", width="stretch"):
                st.session_state.confirm_delete = True
                st.rerun()
        else:
            st.warning("Confirmez la suppression (irréversible) :")
            if st.button("Oui, tout effacer", type="primary", width="stretch"):
                if os.path.exists(USER_DATA_PATH):
                    shutil.rmtree(USER_DATA_PATH)
                st.session_state.confirm_delete = False
                st.success("Toutes les données ont été supprimées.")
                st.rerun()
            if st.button("Annuler", width="stretch"):
                st.session_state.confirm_delete = False
                st.rerun()

st.divider()

# --- RÉSULTATS. ---
st.markdown(f"<h3>{ICON_CHART} Résultats des corrections</h3>", unsafe_allow_html=True)

chemin_excel = os.path.join(USER_DATA_PATH, "Resultats_Corrections.xlsx")

if os.path.exists(chemin_excel):
    df = pd.read_excel(chemin_excel)
    st.dataframe(df, width="stretch", hide_index=True)

    with open(chemin_excel, "rb") as file:
        st.download_button(
            label="Télécharger le fichier Excel",
            data=file,
            file_name="Resultats_Corrections.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.divider()

    st.markdown(
        f"<h3>{ICON_SEARCH} Détail par groupe (Rapports JSON)</h3>",
        unsafe_allow_html=True,
    )

    groupes_analyses = [
        d
        for d in dossiers
        if os.path.exists(os.path.join(USER_DATA_PATH, d, "rapport_ia_brut.json"))
    ]

    if groupes_analyses:
        groupe_selectionne = st.selectbox(
            "Choisissez un groupe pour lire son rapport JSON :",
            groupes_analyses,
        )

        if groupe_selectionne:
            chemin_json = os.path.join(
                USER_DATA_PATH, groupe_selectionne, "rapport_ia_brut.json"
            )
            with open(chemin_json, "r", encoding="utf-8") as f:
                donnees_json = json.load(f)

            feedback = donnees_json.get("3_feedback", {})
            st.success(f"**Points forts :** {feedback.get('points_forts', '')}")
            st.warning(
                f"**Axes d'amélioration :** {feedback.get('axes_amelioration', '')}"
            )

            with st.expander("Voir l'analyse des sources"):
                st.json(donnees_json.get("1_analyse", {}))

            with st.expander("Voir la grille d'évaluation"):
                st.json(donnees_json.get("2_grille", {}))

            with st.expander("Voir le feeback"):
                st.json(donnees_json.get("3_feedback", {}))
    else:
        st.info("Aucun rapport individuel n'est disponible pour le moment.")
else:
    st.info("Le grand tableau des notes apparaîtra ici une fois l'analyse terminée.")
