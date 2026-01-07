import streamlit as st
import pandas as pd
import io

# --- Configuration ---
st.set_page_config(page_title="Consolidateur de Prix", page_icon="⚡")

st.title("⚡ Consolidateur de Listes de Prix")
st.markdown("""
**Correctif Colonnes :**
Cette version donne la priorité absolue aux colonnes contenant le mot **"CAISSE"** pour éviter de prendre le "Prix Pièce".

**Logique :**
1. **Référence (Ancien)** : Prix de base.
2. **Mise à jour (Nouveau)** : Nouveaux prix.
3. Si le produit existe dans le nouveau -> On affiche le Nouveau Prix.
4. Si le produit n'existe PAS dans le nouveau -> On garde l'Ancien Prix.
5. Si c'est un nouveau produit -> On affiche le Nouveau Prix (Ancien vide).
""")

# --- Fonctions ---

def clean_code(series):
    """Nettoyage des codes pour garantir la correspondance"""
    return series.astype(str).str.strip().str.upper()

def find_column_name(df, possible_names):
    """Cherche une colonne spécifique (Article, Code)"""
    for col in df.columns:
        c_str = str(col).upper()
        for name in possible_names:
            if name.upper() in c_str:
                return col
    return None

def find_price_column(df):
    """
    Logique intelligente pour trouver le VRAI prix.
    Priorité : 'CAISSE' > 'PCI' > 'PRIX'
    """
    cols = df.columns.tolist()
    
    # 1. Priorité ABSOLUE : Chercher "PCI" ou "PRIX" AVEC "CAISSE"
    for col in cols:
        c_str = str(col).upper()
        if ("PCI" in c_str or "PRIX" in c_str or "PRICE" in c_str) and "CAISSE" in c_str:
            return col
            
    # 2. Priorité Moyenne : Chercher "PCI" spécifique (ex: PCI 2026, PCI PCI)
    for col in cols:
        c_str = str(col).upper()
        # Éviter "Piece" si possible
        if ("PCI" in c_str) and "PIECE" not in c_str:
            return col

    # 3. Dernier Recours : N'importe quoi avec "Prix" ou "Price" ou "Montant"
    for col in cols:
        c_str = str(col).upper()
        if "PRIX" in c_str or "PRICE" in c_str or "MONTANT" in c_str:
            return col
            
    return None

def process_logic(file_ref, file_new):
    # 1. Chargement des Fichiers
    try:
        df_ref = pd.read_excel(file_ref, header=0)
        # Si pas de colonne Code, essayer header=1
        if not find_column_name(df_ref, ['Code', 'Nomenclature']):
             df_ref = pd.read_excel(file_ref, header=1)
    except:
        return None, "Erreur lecture Fichier Référence", None, None

    try:
        df_new = pd.read_excel(file_new, header=0)
        if not find_column_name(df_new, ['Code', 'Nomenclature']):
             df_new = pd.read_excel(file_new, header=1)
    except:
        return None, "Erreur lecture Fichier Nouveau", None, None

    # 2. Identification des Colonnes
    # Référence (Old)
    ref_code = find_column_name(df_ref, ['Code', 'Nomenclature', 'Ref'])
    ref_price = find_price_column(df_ref) # Utilise la nouvelle logique stricte
    ref_art = find_column_name(df_ref, ['Article', 'Designation', 'Description', 'Libellé'])

    # Nouveau (Update)
    new_code = find_column_name(df_new, ['Code', 'Nomenclature', 'Ref'])
    new_price = find_price_column(df_new) # Utilise la nouvelle logique stricte
    new_art = find_column_name(df_new, ['Article', 'Designation', 'Description', 'Libellé'])

    # Debug info pour l'utilisateur
    debug_msg = {
        "Ref_File": file_ref.name,
        "Ref_Col_Prix_Trouvee": ref_price,
        "New_File": file_new.name,
        "New_Col_Prix_Trouvee": new_price
    }

    if not ref_code or not ref_price or not new_code or not new_price:
        return None, "Colonnes introuvables. Voir détails ci-dessous.", debug_msg, None

    # 3. Préparation des DataFrames
    df_ref = df_ref[[ref_code, ref_art, ref_price]].copy()
    df_ref.columns = ['Code', 'Article_Ref', 'Prix_OLD']
    df_ref['Code'] = clean_code(df_ref['Code'])

    df_new = df_new[[new_code, new_art, new_price]].copy()
    df_new.columns = ['Code', 'Article_New', 'Prix_NEW_Source']
    df_new['Code'] = clean_code(df_new['Code'])

    # 4. FUSION (Outer Join)
    df_final = pd.merge(df_ref, df_new, on='Code', how='outer')

    # 5. APPLICATION DE LA LOGIQUE
    
    # Consolider Article Name
    df_final['Article'] = df_final['Article_New'].fillna(df_final['Article_Ref'])

    # Nettoyage prix
    df_final['Prix_OLD'] = pd.to_numeric(df_final['Prix_OLD'], errors='coerce')
    df_final['Prix_NEW_Source'] = pd.to_numeric(df_final['Prix_NEW_Source'], errors='coerce')

    # LOGIQUE :
    # Si le produit existe dans le fichier Nouveau -> On prend le Prix Nouveau
    # Si le produit est MANQUANT dans le fichier Nouveau -> On garde le Prix Ancien
    
    # On crée une colonne finale "Prix 2026 (Consolidé)"
    # fillna() remplit les trous du nouveau fichier avec les valeurs de l'ancien
    df_final['Prix_Final'] = df_final['Prix_NEW_Source'].fillna(df_final['Prix_OLD'])

    # Si le produit est NOUVEAU (n'existait pas avant), Prix_OLD est déjà NaN, ce qui est correct.

    # 6. Formatage
    df_final = df_final[['Code', 'Article', 'Prix_OLD', 'Prix_NEW_Source', 'Prix_Final']]
    df_final = df_final.sort_values(by='Code')
    
    for c in ['Prix_OLD', 'Prix_NEW_Source', 'Prix_Final']:
        df_final[c] = df_final[c].round(2)

    return df_final, None, debug_msg, (ref_price, new_price)

# --- Interface ---
col1, col2 = st.columns(2)

with col1:
    f_ref = st.file_uploader("📂 Fichier Référence (Ancien)", type=['xlsx', 'csv'])
with col2:
    f_upd = st.file_uploader("📂 Fichier Mise à jour (Nouveau)", type=['xlsx', 'csv'])

if f_ref and f_upd:
    if st.button("Consolider les Prix"):
        with st.spinner("Analyse des colonnes..."):
            
            df_result, error, debug, cols_used = process_logic(f_ref, f_upd)
            
            if error:
                st.error(error)
                st.json(debug) # Affiche quelle colonne a posé problème
            else:
                st.success("Traitement terminé !")
                
                # Afficher les colonnes utilisées pour rassurer l'utilisateur
                st.info(f"""
                ℹ️ **Colonnes détectées et utilisées :**
                * Dans l'Ancien Fichier : `{cols_used[0]}`
                * Dans le Nouveau Fichier : `{cols_used[1]}`
                *(Vérifiez que ce sont bien les colonnes 'Caisse' et non 'Pièce')*
                """)
                
                st.dataframe(df_result.head(50))
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_result.to_excel(writer, index=False)
                    worksheet = writer.sheets['Sheet1']
                    worksheet.set_column('A:A', 20)
                    worksheet.set_column('B:B', 50)
                    worksheet.set_column('C:E', 15)
                
                st.download_button(
                    label="📥 Télécharger le Fichier Final",
                    data=buffer,
                    file_name="Prix_Consolides.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
