import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import io

# --- Configuration de la Page ---
st.set_page_config(page_title="Mise à jour Prix (Format Préservé)", page_icon="🔄")

st.title("🔄 Mise à Jour des Prix Excel")
st.markdown("""
**Instructions :**
Cette application met à jour les prix d'un fichier Excel **sans casser le formatage** (couleurs, formules, etc.).

1. **Fichier Cible** : Le fichier Excel que vous voulez modifier (ex: `PCI 2026`).
2. **Fichier Source** : Le fichier qui contient les bons prix (ex: `PCI 2024`).
3. L'appli va chercher les Codes produits et remplacer les Prix correspondants.
""")

# --- Fonctions Utilitaires ---

def clean_code(val):
    """Nettoie le code produit pour la comparaison"""
    if val is None: return ""
    return str(val).strip().upper()

def get_source_prices(file_source):
    """Lit le fichier source avec Pandas pour extraire {Code: Prix} rapidement"""
    # 1. Essayer de lire avec l'en-tête standard
    try:
        df = pd.read_excel(file_source, header=0)
        # Vérifier si on trouve une colonne "Code"
        cols = [str(c).upper() for c in df.columns]
        if not any("CODE" in c for c in cols):
            # Si pas trouvé, essayer la ligne suivante (header=1)
            df = pd.read_excel(file_source, header=1)
    except:
        return None, "Erreur lecture fichier source"

    # 2. Trouver les colonnes dynamiquement
    col_code = None
    col_price = None

    for col in df.columns:
        c_str = str(col).upper()
        if "CODE" in c_str or "NOMENCLATURE" in c_str:
            col_code = col
        if "PCI" in c_str or "PRIX" in c_str or "PRICE" in c_str:
            # On préfère "PCI CAISSE" si dispo, sinon n'importe quel PCI
            if col_price is None or "CAISSE" in c_str:
                col_price = col
    
    if not col_code or not col_price:
        return None, f"Colonnes introuvables. Colonnes détectées : {list(df.columns)}"

    # 3. Créer le dictionnaire {Code: Prix}
    price_dict = {}
    for _, row in df.iterrows():
        code = clean_code(row[col_code])
        price = row[col_price]
        if pd.notna(price) and isinstance(price, (int, float)):
            price_dict[code] = round(price, 2)
            
    return price_dict, None

def update_excel_file(file_target, price_dict):
    """Ouvre le fichier cible avec OpenPyXL et met à jour les cellules"""
    # Charger le classeur (Workbook)
    wb = load_workbook(file_target)
    ws = wb.active # Feuille active

    # 1. Scanner les 5 premières lignes pour trouver l'en-tête
    header_row_idx = None
    col_map = {}

    for r in range(1, 6): # On teste les lignes 1 à 5
        row_values = [cell.value for cell in ws[r]]
        # On cherche "Code" et "PCI" ou "Price" dans cette ligne
        str_values = [str(v).upper() if v else "" for v in row_values]
        
        if any("CODE" in s for s in str_values) and (any("PCI" in s for s in str_values) or any("PRIX" in s for s in str_values)):
            header_row_idx = r
            # Créer la map {NomColonne: IndexColonne} (1-based)
            for i, val in enumerate(str_values):
                col_map[val] = i + 1 
            break
    
    if header_row_idx is None:
        return None, "Impossible de trouver la ligne d'en-tête (Code/Prix) dans les 5 premières lignes."

    # 2. Identifier les index précis des colonnes
    idx_code = None
    idx_price = None

    # Recherche un peu floue pour trouver la bonne colonne
    for name, idx in col_map.items():
        if "CODE" in name or "NOMENCLATURE" in name:
            idx_code = idx
        if "PCI" in name or "PRIX" in name:
             # Priorité à "PCI CAISSE" ou "PCI PCI"
             if idx_price is None or "CAISSE" in name or "PCI PCI" in name:
                idx_price = idx

    if not idx_code or not idx_price:
        return None, f"Colonnes cibles non identifiées. En-têtes trouvés : {list(col_map.keys())}"

    # 3. Mettre à jour les lignes
    count = 0
    # On itère à partir de la ligne suivant l'en-tête
    for row in ws.iter_rows(min_row=header_row_idx + 1):
        cell_code = row[idx_code - 1]
        cell_price = row[idx_price - 1]
        
        code_val = clean_code(cell_code.value)
        
        if code_val in price_dict:
            # On met à jour SEULEMENT si on a un nouveau prix
            cell_price.value = price_dict[code_val]
            count += 1

    # 4. Sauvegarder dans un buffer mémoire (pour le téléchargement)
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    
    return buffer, count


# --- Interface Utilisateur ---
col1, col2 = st.columns(2)

with col1:
    f_target = st.file_uploader("📝 Fichier Cible (À modifier)", type=['xlsx'])

with col2:
    f_source = st.file_uploader("💰 Fichier Source (Prix corrects)", type=['xlsx'])

if f_target and f_source:
    if st.button("Lancer la Mise à Jour 🚀"):
        with st.spinner("Analyse et mise à jour en cours..."):
            
            # 1. Lire les prix sources
            prices, error = get_source_prices(f_source)
            
            if error:
                st.error(f"Erreur Source : {error}")
            else:
                st.info(f"{len(prices)} prix trouvés dans le fichier source.")

                # 2. Mettre à jour le fichier cible
                # Important : on passe f_target directement à openpyxl
                result_buffer, count_or_err = update_excel_file(f_target, prices)

                if isinstance(count_or_err, str): # C'est une erreur
                    st.error(f"Erreur Cible : {count_or_err}")
                else:
                    st.success(f"✅ Succès ! {count_or_err} lignes mises à jour.")
                    
                    # Bouton de téléchargement
                    st.download_button(
                        label="📥 Télécharger le Fichier Mis à Jour",
                        data=result_buffer,
                        file_name="Fichier_Mis_a_Jour.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            st.error(f"Une erreur s'est produite : {e}")
