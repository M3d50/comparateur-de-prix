import streamlit as st
import pandas as pd
import io

# --- Configuration ---
st.set_page_config(page_title="Consolidateur de Prix", page_icon="⚡")

st.title("⚡ Consolidateur de Listes de Prix")
st.markdown("""
**Logique Appliquée :**
1. **Référence** = Ancien Fichier.
2. **Mise à jour** = Nouveau Fichier.
3. Si un produit **existe dans les deux** : On affiche l'ancien et le nouveau prix.
4. Si un produit **manque dans le nouveau** : Le nouveau prix prend la valeur de l'ancien.
5. Si un produit est **nouveau** (ajouté) : L'ancien prix est vide, le nouveau est affiché.
""")

# --- Fonctions ---

def clean_code(series):
    """Nettoyage des codes pour garantir la correspondance"""
    return series.astype(str).str.strip().str.upper()

def find_column(df, keywords):
    """Trouve la colonne correspondant aux mots-clés"""
    for col in df.columns:
        for kw in keywords:
            if kw.upper() in str(col).upper():
                return col
    return None

def process_logic(file_ref, file_new):
    # 1. Chargement des Fichiers
    # On essaie de lire header=0, si pas de code, header=1
    try:
        df_ref = pd.read_excel(file_ref, header=0)
        if not find_column(df_ref, ['Code', 'Nomenclature']):
             df_ref = pd.read_excel(file_ref, header=1)
    except:
        return None, "Erreur lecture Fichier Référence"

    try:
        df_new = pd.read_excel(file_new, header=0)
        if not find_column(df_new, ['Code', 'Nomenclature']):
             df_new = pd.read_excel(file_new, header=1)
    except:
        return None, "Erreur lecture Fichier Nouveau"

    # 2. Identification des Colonnes
    # Référence (Old)
    ref_code = find_column(df_ref, ['Code', 'Nomenclature', 'Ref'])
    ref_price = find_column(df_ref, ['PCI', 'Prix', 'Price', 'Montant'])
    ref_art = find_column(df_ref, ['Article', 'Designation', 'Description'])

    # Nouveau (Update)
    new_code = find_column(df_new, ['Code', 'Nomenclature', 'Ref'])
    new_price = find_column(df_new, ['PCI', 'Prix', 'Price', 'Montant'])
    new_art = find_column(df_new, ['Article', 'Designation', 'Description'])

    if not ref_code or not ref_price or not new_code or not new_price:
        return None, "Colonnes 'Code' ou 'Prix' introuvables. Vérifiez les fichiers."

    # 3. Préparation des DataFrames
    # On ne garde que l'essentiel pour éviter les conflits
    df_ref = df_ref[[ref_code, ref_art, ref_price]].copy()
    df_ref.columns = ['Code', 'Article_Ref', 'Prix_OLD']
    df_ref['Code'] = clean_code(df_ref['Code'])

    df_new = df_new[[new_code, new_art, new_price]].copy()
    df_new.columns = ['Code', 'Article_New', 'Prix_NEW_Raw']
    df_new['Code'] = clean_code(df_new['Code'])

    # 4. FUSION (Outer Join)
    # Cela inclut : Les communs, ceux uniquement dans Old, et ceux uniquement dans New
    df_final = pd.merge(df_ref, df_new, on='Code', how='outer')

    # 5. APPLICATION DE LA LOGIQUE (Le cœur du problème)
    
    # Gestion des noms d'articles (Prendre le nouveau s'il existe, sinon l'ancien)
    df_final['Article'] = df_final['Article_New'].fillna(df_final['Article_Ref'])

    # Nettoyage des prix (convertir en nombre)
    df_final['Prix_OLD'] = pd.to_numeric(df_final['Prix_OLD'], errors='coerce')
    df_final['Prix_NEW_Raw'] = pd.to_numeric(df_final['Prix_NEW_Raw'], errors='coerce')

    # LOGIQUE PRINCIPALE : Colonne "Nouveau Prix Final"
    # Si Prix_NEW_Raw existe -> On le garde
    # Si Prix_NEW_Raw est vide (produit manquant dans le nouveau fichier) -> On prend Prix_OLD
    df_final['Prix_NEW_Final'] = df_final['Prix_NEW_Raw'].fillna(df_final['Prix_OLD'])

    # Note sur "Nouveau produit" :
    # Si c'est un nouveau produit, 'Prix_OLD' sera déjà NaN (None) grâce au merge Outer.
    # Donc pas besoin de forcer à None manuellement.

    # 6. Formatage Final
    df_final = df_final[['Code', 'Article', 'Prix_OLD', 'Prix_NEW_Final']]
    df_final = df_final.sort_values(by='Code')
    
    # Arrondir
    df_final['Prix_OLD'] = df_final['Prix_OLD'].round(2)
    df_final['Prix_NEW_Final'] = df_final['Prix_NEW_Final'].round(2)

    return df_final, None

# --- Interface ---
col1, col2 = st.columns(2)

with col1:
    f_ref = st.file_uploader("📂 Fichier Référence (Ancien)", type=['xlsx', 'csv'])
with col2:
    f_upd = st.file_uploader("📂 Fichier Mise à jour (Nouveau)", type=['xlsx', 'csv'])

if f_ref and f_upd:
    if st.button("Consolider les Prix"):
        with st.spinner("Application de la logique..."):
            
            df_result, error = process_logic(f_ref, f_upd)
            
            if error:
                st.error(error)
            else:
                st.success(f"Traitement terminé ! {len(df_result)} produits générés.")
                
                # Aperçu
                st.dataframe(df_result.head(50))
                
                # Export
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_result.to_excel(writer, index=False)
                    
                    # Formatage Excel simple (Largeur colonnes)
                    worksheet = writer.sheets['Sheet1']
                    worksheet.set_column('A:A', 20) # Code
                    worksheet.set_column('B:B', 50) # Article
                    worksheet.set_column('C:D', 15) # Prix
                
                st.download_button(
                    label="📥 Télécharger le Fichier Final",
                    data=buffer,
                    file_name="Prix_Consolides.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
