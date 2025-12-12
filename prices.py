import streamlit as st
import pandas as pd
import io

# --- Configuration de la Page ---
st.set_page_config(page_title="Comparateur de Prix", page_icon="📈")

st.title("📈 Comparateur de Listes de Prix")
st.markdown("""
**Instructions :**
1. Téléchargez votre **Ancien Fichier** (ex: 2025) à gauche.
2. Téléchargez votre **Nouveau Fichier** (ex: 2026) à droite.
3. L'application détectera automatiquement les codes et les prix pour générer un rapport complet.
""")

# --- Fonction de Nettoyage ---
def clean_code(series):
    """Standardise les codes (supprime les espaces, met en majuscules)"""
    return series.astype(str).str.strip().str.upper()

def find_column(df, keywords):
    """Cherche une colonne qui contient un des mots-clés"""
    cols = df.columns.astype(str).tolist()
    for col in cols:
        for kw in keywords:
            if kw.lower() in col.lower():
                return col
    return None

def process_generic_files(file_old, file_new):
    # -----------------------------------------------------
    # 1. Traitement du Fichier ANCIEN (Reference)
    # -----------------------------------------------------
    # On lit la première feuille (sheet_name=0) pour être générique
    df_old = pd.read_excel(file_old, sheet_name=0)
    df_old.columns = df_old.columns.astype(str).str.replace('\n', ' ').str.strip()
    
    # Détection des colonnes
    col_code_old = find_column(df_old, ['Code', 'Nomenclature', 'Ref', 'Article Code'])
    col_price_old = find_column(df_old, ['PCI', 'Prix', 'Price', 'Montant', 'Cost'])
    col_name_old = find_column(df_old, ['Article', 'Description', 'Designation', 'Libellé'])

    if not col_code_old or not col_price_old:
        st.error(f"❌ Impossible de trouver les colonnes 'Code' ou 'Prix' dans le fichier {file_old.name}. Vérifiez les entêtes.")
        return None

    # Renommage et Nettoyage
    df_old = df_old.rename(columns={col_code_old: 'Code', col_name_old: 'Article_Old', col_price_old: 'Prix_Ancien'})
    df_old['Code'] = clean_code(df_old['Code'])
    
    # Si la colonne Article n'existe pas, on met vide
    if not col_name_old:
        df_old['Article_Old'] = ""
        
    df_old = df_old[['Code', 'Article_Old', 'Prix_Ancien']]

    # -----------------------------------------------------
    # 2. Traitement du Fichier NOUVEAU (Cible)
    # -----------------------------------------------------
    # Parfois les fichiers commencent à la ligne 2 (header=1). On teste.
    df_new = pd.read_excel(file_new, sheet_name=0)
    
    # Petite astuce : si la colonne "Code" n'est pas dans la ligne 0, on essaye la ligne 1
    if not find_column(df_new, ['Code', 'Nomenclature']):
        df_new = pd.read_excel(file_new, sheet_name=0, header=1)

    df_new.columns = df_new.columns.astype(str).str.strip()

    col_code_new = find_column(df_new, ['Code', 'Nomenclature', 'Ref'])
    col_price_new = find_column(df_new, ['PCI', 'Prix', 'Price', 'Montant', 'USD'])
    col_name_new = find_column(df_new, ['Article', 'Description', 'Designation', 'Libellé'])

    if not col_code_new or not col_price_new:
        st.error(f"❌ Impossible de trouver les colonnes 'Code' ou 'Prix' dans le fichier {file_new.name}.")
        return None

    # Renommage
    df_new = df_new.rename(columns={col_code_new: 'Code', col_name_new: 'Article_New', col_price_new: 'Prix_Nouveau'})
    df_new['Code'] = clean_code(df_new['Code'])
    
    if not col_name_new:
        df_new['Article_New'] = ""

    df_new = df_new[['Code', 'Article_New', 'Prix_Nouveau']]

    # -----------------------------------------------------
    # 3. Fusion (Outer Join pour tout garder)
    # -----------------------------------------------------
    df_merged = pd.merge(df_old, df_new, on='Code', how='outer')

    # -----------------------------------------------------
    # 4. Nettoyage Final
    # -----------------------------------------------------
    # Consolider le nom de l'article (Priorité au Nouveau, puis Ancien)
    df_merged['Article'] = df_merged['Article_New'].fillna(df_merged['Article_Old'])
    df_merged = df_merged.drop(columns=['Article_New', 'Article_Old'])

    # Formatage des prix
    for col in ['Prix_Ancien', 'Prix_Nouveau']:
        df_merged[col] = pd.to_numeric(df_merged[col], errors='coerce').round(2)

    # Réorganisation
    df_merged = df_merged[['Code', 'Article', 'Prix_Ancien', 'Prix_Nouveau']]
    df_merged = df_merged.sort_values(by='Code')

    return df_merged

# --- Interface Utilisateur ---
col1, col2 = st.columns(2)

with col1:
    uploaded_old = st.file_uploader("📂 Fichier Ancien / Référence", type=['xlsx'])

with col2:
    uploaded_new = st.file_uploader("📂 Fichier Nouveau / Cible", type=['xlsx'])

if uploaded_old and uploaded_new:
    with st.spinner('Traitement en cours...'):
        try:
            result_df = process_generic_files(uploaded_old, uploaded_new)
            
            if result_df is not None:
                st.success(f"✅ Succès ! {len(result_df)} produits traités.")
                
                # Aperçu
                st.subheader("Aperçu des données")
                st.dataframe(result_df.head(100))
                
                # Export Excel
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    result_df.to_excel(writer, index=False)
                
                st.download_button(
                    label="📥 Télécharger le Rapport Excel",
                    data=buffer,
                    file_name="Comparatif_Prix.xlsx",
                    mime="application/vnd.ms-excel"
                )
                
        except Exception as e:
            st.error(f"Une erreur s'est produite : {e}")