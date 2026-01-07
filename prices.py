import streamlit as st
import pandas as pd
import io

# --- Configuration ---
st.set_page_config(page_title="Comparateur de Prix & Variation", page_icon="📊")

st.title("📊 Comparateur de Prix avec Variation")
st.markdown("""
**Résultat généré :**
1. **Prix Ancien** (Fichier Référence)
2. **Prix Nouveau** (Fichier Mise à jour)
3. **Variation %** : 
   - <span style='color:green'>**Vert**</span> si le prix a augmenté.
   - <span style='color:red'>**Rouge**</span> si le prix a baissé.
""", unsafe_allow_html=True)

# --- Fonctions ---

def clean_code(series):
    return series.astype(str).str.strip().str.upper()

def find_column_name(df, possible_names):
    for col in df.columns:
        c_str = str(col).upper()
        for name in possible_names:
            if name.upper() in c_str:
                return col
    return None

def find_price_column(df):
    """Logique stricte : Priorité à CAISSE > PCI > PRIX"""
    cols = df.columns.tolist()
    # 1. Priorité ABSOLUE : CAISSE
    for col in cols:
        c_str = str(col).upper()
        if ("PCI" in c_str or "PRIX" in c_str) and "CAISSE" in c_str:
            return col
    # 2. Priorité Moyenne : PCI (sans Piece)
    for col in cols:
        c_str = str(col).upper()
        if ("PCI" in c_str) and "PIECE" not in c_str:
            return col
    # 3. Dernier Recours
    for col in cols:
        c_str = str(col).upper()
        if "PRIX" in c_str or "PRICE" in c_str:
            return col
    return None

def process_data(file_ref, file_new):
    # 1. Chargement
    try:
        df_ref = pd.read_excel(file_ref, header=0)
        if not find_column_name(df_ref, ['Code', 'Nomenclature']):
             df_ref = pd.read_excel(file_ref, header=1)
        
        df_new = pd.read_excel(file_new, header=0)
        if not find_column_name(df_new, ['Code', 'Nomenclature']):
             df_new = pd.read_excel(file_new, header=1)
    except Exception as e:
        return None, f"Erreur de lecture : {e}"

    # 2. Identification Colonnes
    ref_code = find_column_name(df_ref, ['Code', 'Nomenclature', 'Ref'])
    ref_price = find_price_column(df_ref)
    ref_art = find_column_name(df_ref, ['Article', 'Designation', 'Description'])

    new_code = find_column_name(df_new, ['Code', 'Nomenclature', 'Ref'])
    new_price = find_price_column(df_new)
    new_art = find_column_name(df_new, ['Article', 'Designation', 'Description'])

    if not ref_code or not ref_price or not new_code or not new_price:
        return None, "Colonnes Code ou Prix introuvables."

    # 3. Nettoyage
    df_ref = df_ref[[ref_code, ref_art, ref_price]].copy()
    df_ref.columns = ['Code', 'Article_Ref', 'Prix_Ancien']
    df_ref['Code'] = clean_code(df_ref['Code'])
    df_ref['Prix_Ancien'] = pd.to_numeric(df_ref['Prix_Ancien'], errors='coerce')

    df_new = df_new[[new_code, new_art, new_price]].copy()
    df_new.columns = ['Code', 'Article_New', 'Prix_Nouveau']
    df_new['Code'] = clean_code(df_new['Code'])
    df_new['Prix_Nouveau'] = pd.to_numeric(df_new['Prix_Nouveau'], errors='coerce')

    # 4. Fusion
    df_final = pd.merge(df_ref, df_new, on='Code', how='outer')
    
    # 5. Calculs
    df_final['Article'] = df_final['Article_New'].fillna(df_final['Article_Ref'])
    
    # Calcul de la variation % : (Nouveau - Ancien) / Ancien
    df_final['Variation %'] = (df_final['Prix_Nouveau'] - df_final['Prix_Ancien']) / df_final['Prix_Ancien']

    # 6. Mise en forme finale
    df_final = df_final[['Code', 'Article', 'Prix_Ancien', 'Prix_Nouveau', 'Variation %']]
    df_final = df_final.sort_values(by='Code')
    
    # Arrondir pour l'affichage (pas pour le calcul Excel)
    # On laisse la Variation en décimal (ex: 0.12) pour que Excel la formate en % (12%)
    
    return df_final, None

# --- Interface ---
col1, col2 = st.columns(2)
with col1:
    f_ref = st.file_uploader("📂 Fichier Référence (Ancien)", type=['xlsx'])
with col2:
    f_new = st.file_uploader("📂 Fichier Mise à jour (Nouveau)", type=['xlsx'])

if f_ref and f_new:
    if st.button("Générer le Comparatif 🚀"):
        with st.spinner("Calcul des variations..."):
            df_res, err = process_data(f_ref, f_new)
            
            if err:
                st.error(err)
            else:
                st.success("Fichier généré avec succès !")
                
                # Aperçu (Streamlit ne montre pas les couleurs Excel, mais on formate le %)
                st.write("Aperçu (Les couleurs apparaîtront dans le fichier Excel téléchargé) :")
                st.dataframe(df_res.head(50).style.format({
                    'Prix_Ancien': '{:.2f}', 
                    'Prix_Nouveau': '{:.2f}', 
                    'Variation %': '{:.2%}'
                }))

                # --- Moteur d'exportation Excel avec Couleurs ---
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_res.to_excel(writer, index=False, sheet_name='Comparatif')
                    
                    workbook = writer.book
                    worksheet = writer.sheets['Comparatif']
                    
                    # Formats
                    fmt_currency = workbook.add_format({'num_format': '#,##0.00'})
                    fmt_percent = workbook.add_format({'num_format': '0.00%'})
                    
                    # Couleurs conditionnelles
                    # Vert (Augmentation)
                    fmt_green = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
                    # Rouge (Diminution)
                    fmt_red = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

                    # Appliquer formats colonnes
                    worksheet.set_column('A:A', 20) # Code
                    worksheet.set_column('B:B', 40) # Article
                    worksheet.set_column('C:D', 15, fmt_currency) # Prix
                    worksheet.set_column('E:E', 15, fmt_percent)  # Variation

                    # Appliquer règles conditionnelles sur la colonne E (Variation)
                    # Note: Row 1 is header, so data starts at row 2 (Excel index 1)
                    last_row = len(df_res) + 1
                    
                    # Règle 1: Supérieur à 0 (Augmentation) -> Vert
                    worksheet.conditional_format(1, 4, last_row, 4, {
                        'type': 'cell',
                        'criteria': '>',
                        'value': 0,
                        'format': fmt_green
                    })
                    
                    # Règle 2: Inférieur à 0 (Baisse) -> Rouge
                    worksheet.conditional_format(1, 4, last_row, 4, {
                        'type': 'cell',
                        'criteria': '<',
                        'value': 0,
                        'format': fmt_red
                    })

                st.download_button(
                    label="📥 Télécharger Excel (Avec Couleurs)",
                    data=buffer,
                    file_name="Comparatif_Prix_Couleurs.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
