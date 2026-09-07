import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import datetime

# ------------------------------------------------------------------------------
# CONSTANTES & SUFFIXES DES MATÉRIAUX
# ------------------------------------------------------------------------------
STANDARD_SIEVES = [0.063, 0.08, 0.1, 0.125, 0.16, 0.2, 0.25, 0.315, 0.4, 0.5, 0.63, 0.8, 1.0, 1.25, 1.6, 2.0, 2.5, 3.15, 4.0, 5.0, 5.6, 6.3, 8.0, 10.0, 12.5, 14.0, 16.0, 20.0, 25.0, 28.0, 31.5, 40.0]

SUFFIX_MAP = {
    'GII': '/1',
    'GI': '/2',
    'SC': '/3',
    'SD': '/4'
}

def clean_html(html_str):
    """Supprime les espaces en début de ligne pour éviter le rendu en bloc de code Markdown dans Streamlit."""
    return "\n".join([line.strip() for line in html_str.splitlines() if line.strip()])

# -----------------------------------------------------------------------------
# Fonction de calcul du Tamis D (tamis dont le % passant est le plus proche de 95%)
# -----------------------------------------------------------------------------
def get_tamis_D(df: pd.DataFrame) -> float:
    if df is None or df.empty or "% Passants" not in df.columns:
        return None
    diffs = (df["% Passants"] - 95.0).abs()
    idx_closest = diffs.idxmin()
    val = df.loc[idx_closest, "Tamis (mm)"]
    return int(val) if float(val).is_integer() else float(val)

def update_passants(mat_data):
    """Calcule les passants à partir des masses enregistrées (Méthode NF EN 933-1)"""
    M1 = float(mat_data.get('M1', 1000.0))
    refus = [float(r) for r in mat_data.get('refus', [0.0]*len(mat_data['sieves']))]
    
    if M1 > 0:
        pct_refus = [(r / M1) * 100 for r in refus]
        pct_refus_cum = np.cumsum(pct_refus)
        passants = [100.0 - c for c in pct_refus_cum]
        
        passants_fmt = []
        for s, p in zip(mat_data['sieves'], passants):
            passants_fmt.append(max(0.0, round(p, 1)))
        mat_data['passants'] = passants_fmt
    else:
        mat_data['passants'] = [100.0] * len(mat_data['sieves'])

def compute_sieve_at_passant(sieves, passings, target_passant):
    """ Sélectionne le plus petit tamis normatif (mm) dont le passant atteint au moins le % cible """
    s_arr = np.array(sieves, dtype=float)
    p_arr = np.array(passings, dtype=float)
    
    idx_sort = np.argsort(s_arr)
    s_arr = s_arr[idx_sort]
    p_arr = p_arr[idx_sort]
    
    valid_sieves = s_arr[p_arr >= target_passant]
    if len(valid_sieves) > 0:
        val = valid_sieves[0]
        return int(val) if float(val).is_integer() else float(val)
        
    val = s_arr[-1]
    return int(val) if float(val).is_integer() else float(val)

def get_d_D_from_material(mat_data):
    """
    Calcule dynamiquement la dimension du tamis D (tamis dont le % passant est le plus proche de 95%)
    et déduit d (depuis la classe granulaire d/D).
    """
    sieves = mat_data.get('sieves', [])
    passants = mat_data.get('passants', [])

    d_val = None
    classe_str = mat_data.get('classe', '')
    if '/' in classe_str:
        try:
            parts = classe_str.replace(',', '.').split('/')
            d_val = float(parts[0])
        except ValueError:
            pass

    if sieves and passants:
        diffs = [abs(p - 95.0) for p in passants]
        min_idx = diffs.index(min(diffs))
        D_val = sieves[min_idx]
        
        if d_val is None:
            d_val = compute_sieve_at_passant(sieves, passants, 5.0)
        return float(d_val), float(D_val)

    return d_val or 0.0, 0.0

def get_passant_at_sieve(sieves, passings, target_sieve):
    """ Calcule ou interpole linéairement le passant au tamis cible """
    if target_sieve is None or target_sieve <= 0:
        return 0.0
        
    s_arr = np.array(sieves, dtype=float)
    p_arr = np.array(passings, dtype=float)
    
    idx_sort = np.argsort(s_arr)
    s_arr = s_arr[idx_sort]
    p_arr = p_arr[idx_sort]
    
    if target_sieve >= s_arr[-1]:
        return 100.0
    if target_sieve <= s_arr[0]:
        return float(p_arr[0])
        
    return float(np.interp(target_sieve, s_arr, p_arr))

def calculate_characteristic_data(mat_data, tamis_D=None):
    """
    Déduit automatiquement la liste des tamis (2D, 1.4D, D, d, d/2) 
    en utilisant la valeur exacte de D (ex: D=25 -> 2D=50, 1.4D=35, D=25, d=10, d/2=5)
    et calcule leur pourcentage de passant correspondant.
    """
    d, D_calc = get_d_D_from_material(mat_data)
    D = float(tamis_D) if (tamis_D is not None and tamis_D > 0) else D_calc
    
    def fmt_sieve(val):
        if val is None or val == 0:
            return 0
        val_r = round(val, 2)
        return int(val_r) if float(val_r).is_integer() else val_r

    sieves_dict = {
        '2D': fmt_sieve(2 * D),
        '1.4D': fmt_sieve(1.4 * D),
        'D': fmt_sieve(D),
        'd': fmt_sieve(d),
        'd/2': fmt_sieve(d / 2)
    }
    
    passants_dict = {}
    for key, sieve_size in sieves_dict.items():
        val = get_passant_at_sieve(mat_data['sieves'], mat_data['passants'], sieve_size)
        passants_dict[key] = val

    return sieves_dict, passants_dict

def compute_MF(sieves, passings):
    """
    Calcule automatiquement le Module de Finesse (FM) selon la norme NF EN 12620 :
    FM = Sum( Refus cumulés % sur [4, 2, 1, 0.5, 0.25, 0.125] mm ) / 100
    """
    target_sieves = [4.0, 2.0, 1.0, 0.5, 0.25, 0.125]
    sum_refus_cum = 0.0
    for ts in target_sieves:
        passant = get_passant_at_sieve(sieves, passings, ts)
        if not np.isnan(passant):
            sum_refus_cum += (100.0 - passant)
    return round(sum_refus_cum / 100.0, 2)

# ------------------------------------------------------------------------------
# FONCTION PRINCIPALE
# ------------------------------------------------------------------------------
def show(supabase_client=None, can_edit=True, is_admin=False, **kwargs):
    if 'info_prelevement' not in st.session_state:
        st.session_state['info_prelevement'] = {}

    info_defaults = {
        'chantier': "TRAVAUX D'EXECUTION DE TERRASSEMENT, OUVRAGES D'ART ET RETABLISSEMENTS DE COMMUNICATION ENTRE PK 5+450 et PK 10+000-GARE CASA SUD",
        'client': 'TGCC',
        'dossier_no': '2025-260-05985-2025 0247',
        'date_prelevement': '23/07/2026',
        'lieu_prelevement': 'Stock sur centrale à béton',
        'provenance': 'TG PREFA OULAD SALEH',
        'ref_base': '260/26/100',
        'num_rapport': '26/260/LGV/CS/1237'
    }
    for k, v in info_defaults.items():
        st.session_state['info_prelevement'].setdefault(k, v)

    if 'data_granulats' not in st.session_state:
        st.session_state['data_granulats'] = {
            'GII': {
                'nom': 'Gravillons GII - 10/20',
                'classe': '10/20',
                'ref_client': '', 'date_prelevement': '', 'lieu_prelevement': '',
                'sieves': [40, 31.5, 25, 20, 16, 14, 12.5, 10, 8, 6.3, 5, 4, 3.15, 2.5, 2, 1.6, 1.25, 1, 0.8, 0.63, 0.5, 0.4, 0.315, 0.25, 0.2, 0.16, 0.125, 0.1, 0.08, 0.063],
                'refus': [0.0, 0.0, 0.0, 199.7, 2200.3, 732.7, 308.7, 424.0, 163.2, 45.0, 6.9, 2.1, 0.2, 0.1, 0.2, 0.2, 0.1, 0.0, 0.1, 0.2, 0.1, 0.1, 0.1, 0.1, 0.0, 0.2, 0.1, 0.1, 0.1, 0.1],
                'M1': 4110.5, 'M2': 4095.2, 'P': 1.3,
                'passants': [], 
                'fi': 16.0, 'la': 26.0, 'mb': None, 'mf': None, 'se': None
            },
            'GI': {
                'nom': 'Gravillons GI - 4/10',
                'classe': '4/10',
                'ref_client': '', 'date_prelevement': '', 'lieu_prelevement': '',
                'sieves': [20, 16, 14, 12.5, 10, 8, 6.3, 5, 4, 3.15, 2.5, 2, 1.6, 1.25, 1, 0.8, 0.63, 0.5, 0.4, 0.315, 0.25, 0.2, 0.16, 0.125, 0.1, 0.08, 0.063],
                'refus': [199.7, 2200.3, 732.7, 308.7, 424.0, 163.2, 45.0, 6.9, 2.1, 0.2, 0.1, 0.2, 0.2, 0.1, 0.0, 0.1, 0.2, 0.1, 0.1, 0.1, 0.1, 0.0, 0.2, 0.1, 0.1, 0.1, 0.1],
                'M1': 2000.0, 'M2': 1990.0, 'P': 0.0,
                'passants': [], 
                'fi': 14.0, 'la': 26.0, 'mb': None, 'mf': None, 'se': None
            },
            'SD': {
                'nom': 'Sable fin 0/0,630 (Dune)',
                'classe': '0/0,63',
                'ref_client': '', 'date_prelevement': '', 'lieu_prelevement': '',
                'sieves': [6.3, 5, 4, 3.15, 2.5, 2, 1.6, 1.25, 1, 0.8, 0.63, 0.5, 0.4, 0.315, 0.25, 0.2, 0.16, 0.125, 0.1, 0.08, 0.063],
                'refus': [45.0, 6.9, 2.1, 0.2, 0.1, 0.2, 0.2, 0.1, 0.0, 0.1, 0.2, 0.1, 0.1, 0.1, 0.1, 0.0, 0.2, 0.1, 0.1, 0.1, 0.1],
                'M1': 1000.0, 'M2': 900.0, 'P': 2.0,
                'passants': [],
                'fi': None, 'la': None, 'mb': 0.7, 'mf': None, 'se': None
            },
            'SC': {
                'nom': 'Sable grossier 0/4 (Concassé)',
                'classe': '0/4',
                'ref_client': '', 'date_prelevement': '', 'lieu_prelevement': '',
                'sieves': [6.3, 5, 4, 3.15, 2.5, 2, 1.6, 1.25, 1, 0.8, 0.63, 0.5, 0.4, 0.315, 0.25, 0.2, 0.16, 0.125, 0.1, 0.08, 0.063],
                'refus': [45.0, 6.9, 2.1, 0.2, 0.1, 0.2, 0.2, 0.1, 0.0, 0.1, 0.2, 0.1, 0.1, 0.1, 0.1, 0.0, 0.2, 0.1, 0.1, 0.1, 0.1],
                'M1': 1000.0, 'M2': 910.0, 'P': 3.0,
                'passants': [],
                'fi': None, 'la': None, 'mb': None, 'mf': 3.50, 'se': 65.0
            }
        }

    for k in st.session_state['data_granulats'].keys():
        if not st.session_state['data_granulats'][k]['passants']:
            update_passants(st.session_state['data_granulats'][k])
        if k in ['SC', 'SD']:
            st.session_state['data_granulats'][k]['mf'] = compute_MF(
                st.session_state['data_granulats'][k]['sieves'],
                st.session_state['data_granulats'][k]['passants']
            )

    if 'historique_pv' not in st.session_state:
        st.session_state['historique_pv'] = []

    if 'pv_info' not in st.session_state:
        st.session_state['pv_info'] = {}

    pv_defaults = {
        'projet': st.session_state['info_prelevement'].get('chantier', ''),
        'client': st.session_state['info_prelevement'].get('client', ''),
        'ref_pv': st.session_state['info_prelevement'].get('num_rapport', ''),
        'date': st.session_state['info_prelevement'].get('date_prelevement', ''),
        'commentaires': "Les essais d'identifications des granulats pour béton sont conformes aux exigences de la norme NF EN 12620 et NF P 18-545",
        'coord_essais': 'O.IKEN',
        'chef_labo': 'H.BAALLAL'
    }
    for k, v in pv_defaults.items():
        st.session_state['pv_info'].setdefault(k, v)

    if 'success_msg' in st.session_state:
        st.success(st.session_state['success_msg'])
        del st.session_state['success_msg']

    st.title("🏗️ Module Identification des Granulats pour Béton")
    
    if not can_edit:
        st.info("👁️ **Mode Consultation** : Vous êtes en lecture seule.")

    tabs = st.tabs([
        "1️⃣ Feuilles d'Essais Complets",
        "2️⃣ PV d'Identification / Synthèse",
        "3️⃣ Historique & Téléchargement de PV"
    ])

    # ------------------------------------------------------------------------------
    # FENÊTRE 1 : FEUILLES D'ESSAIS COMPLETS
    # ------------------------------------------------------------------------------
    with tabs[0]:
        st.header("Feuilles d'Analyse Granulométrique et Caractéristiques")
        
        st.markdown("##### 📍 Informations de prélèvement (Communes à tous les matériaux)")
        c1, c2, c3 = st.columns(3)
        c4, c5, c6 = st.columns(3)
        
        info_p = st.session_state['info_prelevement']
        
        new_client       = c1.text_input("Client", value=info_p.get('client', 'TGCC'), disabled=not can_edit, key="common_client")
        new_chantier     = c2.text_input("Chantier", value=info_p.get('chantier', ''), disabled=not can_edit, key="common_chantier")
        new_dossier      = c3.text_input("N° Dossier", value=info_p.get('dossier_no', ''), disabled=not can_edit, key="common_dossier")
        new_date_prelev  = c4.text_input("Date de prélèvement", value=info_p.get('date_prelevement', '23/07/2026'), disabled=not can_edit, key="common_date_prelev")
        new_lieu_prelev  = c5.text_input("Lieu de prélèvement", value=info_p.get('lieu_prelevement', 'Stock sur centrale à béton'), disabled=not can_edit, key="common_lieu_prelev")
        new_provenance   = c6.text_input("Provenance échantillon", value=info_p.get('provenance', 'TG PREFA OULAD SALEH'), disabled=not can_edit, key="common_provenance")
        
        new_ref_base     = st.text_input("Référence labo (Base)", value=info_p.get('ref_base', '260/26/100'), disabled=not can_edit, key="common_ref_base")
        new_num_rapport  = st.text_input("N° Rapport d'essai", value=info_p.get('num_rapport', '26/260/LGV/CS/1237'), disabled=not can_edit, key="common_num_rapport")

        if can_edit:
            st.session_state['info_prelevement']['chantier'] = new_chantier
            st.session_state['info_prelevement']['client'] = new_client
            st.session_state['info_prelevement']['dossier_no'] = new_dossier
            st.session_state['info_prelevement']['date_prelevement'] = new_date_prelev
            st.session_state['info_prelevement']['lieu_prelevement'] = new_lieu_prelev
            st.session_state['info_prelevement']['provenance'] = new_provenance
            st.session_state['info_prelevement']['ref_base'] = new_ref_base
            st.session_state['info_prelevement']['num_rapport'] = new_num_rapport
            
            st.session_state['pv_info']['projet'] = new_chantier
            st.session_state['pv_info']['client'] = new_client
            st.session_state['pv_info']['ref_pv'] = new_num_rapport

        st.markdown("---")

        selected_mat = st.radio(
            "Sélectionner la fraction d'échantillon à modifier :",
            ["GII (10/20)", "GI (4/10)", "SC (0/4)", "SD (0/0,63)"],
            horizontal=True
        )
        
        mat_key_map = {"GII (10/20)": "GII", "GI (4/10)": "GI", "SC (0/4)": "SC", "SD (0/0,63)": "SD"}
        key = mat_key_map[selected_mat]
        mat_data = st.session_state['data_granulats'][key]
        
        sub_ref = f"{new_ref_base}{SUFFIX_MAP[key]}" if new_ref_base else SUFFIX_MAP[key]
        
        col1, col2 = st.columns([1, 1])
        
        with col1:
            st.subheader(f"Saisie des données : {mat_data['nom']}")
            st.caption(f"Sous-référence Labo générée : {sub_ref}")

            st.markdown("##### ⚖️ Pesées (Procédé : Lavage et tamisage)")
            c_m1, c_m2, c_p = st.columns(3)
            new_M1 = c_m1.number_input("Masse totale M1 (g)", value=float(mat_data.get('M1', 1000.0)), step=0.1, format="%.1f", disabled=not can_edit, key=f"m1_{key}")
            new_M2 = c_m2.number_input("Masse après lavage M2 (g)", value=float(mat_data.get('M2', 1000.0)), step=0.1, format="%.1f", disabled=not can_edit, key=f"m2_{key}")
            new_P  = c_p.number_input("Matériau au fond P (g)", value=float(mat_data.get('P', 0.0)), step=0.1, format="%.1f", disabled=not can_edit, key=f"p_{key}")
            
            st.subheader("Analyse par tamisage (Saisie des refus en g)")
            
            df_display = pd.DataFrame({
                "Tamis (mm)": mat_data['sieves'],
                "Masse de refus Ri (g)": [round(float(r), 1) for r in mat_data.get('refus', [0.0]*len(mat_data['sieves']))]
            })
            
            saved_M1 = float(new_M1) if new_M1 > 0 else 1.0
            pct_r = (df_display["Masse de refus Ri (g)"] / saved_M1) * 100
            pct_r_cum = pct_r.cumsum()
            df_display["% Refus"] = pct_r.round(1)
            df_display["% Refus Cumulés"] = pct_r_cum.round(1)
            df_display["% Passants"] = (100.0 - pct_r_cum).clip(lower=0.0).round(1)

            edited_df = st.data_editor(
                df_display,
                key=f"editor_{key}",
                column_config={
                    "Tamis (mm)": st.column_config.NumberColumn(disabled=True),
                    "Masse de refus Ri (g)": st.column_config.NumberColumn(
                        disabled=not can_edit, 
                        min_value=0.0, 
                        step=0.1, 
                        format="%.1f"
                    ),
                    "% Refus": st.column_config.NumberColumn(disabled=True, format="%.1f %%"),
                    "% Refus Cumulés": st.column_config.NumberColumn(disabled=True, format="%.1f %%"),
                    "% Passants": st.column_config.NumberColumn(disabled=True, format="%.1f %%")
                },
                hide_index=True,
                use_container_width=True,
                height=350
            )

            # Recalcul dynamique immédiat sur edited_df pour la prise en compte en temps réel
            if new_M1 > 0:
                pct_r_edit = (edited_df["Masse de refus Ri (g)"] / new_M1) * 100
                pct_cum_edit = pct_r_edit.cumsum()
                edited_df["% Refus"] = pct_r_edit.round(1)
                edited_df["% Refus Cumulés"] = pct_cum_edit.round(1)
                edited_df["% Passants"] = (100.0 - pct_cum_edit).clip(lower=0.0).round(1)

                mat_data['M1'] = new_M1
                mat_data['refus'] = [round(float(r), 1) for r in edited_df["Masse de refus Ri (g)"].tolist()]
                update_passants(mat_data)

            st.markdown("##### 🔍 Vérifications et Validations (NF EN 933-1)")
            
            refus_array = edited_df["Masse de refus Ri (g)"].values
            somme_Ri = sum(refus_array)
            masse_calc = somme_Ri + new_P
            perte_fraction = 100 * (new_M2 - masse_calc) / new_M2 if new_M2 > 0 else 0
            fines_f = 100 * ((new_M1 - new_M2) + new_P) / new_M1 if new_M1 > 0 else 0
            
            c_v1, c_v2 = st.columns(2)
            c_v1.info(f"**ΣRi + P :** {masse_calc:.1f} g\n\n**% Tamisat fines (f) :** {fines_f:.2f} %")
            
            if perte_fraction < 1.0:
                c_v2.success(f"**Pertes de tamisage :** {perte_fraction:.2f} %\n\n✅ Essai Valide (< 1%)")
            else:
                c_v2.error(f"**Pertes de tamisage :** {perte_fraction:.2f} %\n\n❌ Rejeter l'essai (> 1%)")

            calculated_mf = compute_MF(mat_data['sieves'], mat_data['passants'])

            st.markdown("---")
            st.subheader(f"Caractéristiques de {mat_data['classe']}")
            
            col_a, col_b = st.columns(2)
            
            if key in ["GII", "GI"]:
                with col_a:
                    fi_val = st.number_input("Coeff. Aplatissement (FI)", value=float(mat_data.get('fi') or 0.0), step=0.1, format="%.1f", disabled=not can_edit, key=f"fi_{key}")
                with col_b:
                    la_val = st.number_input("Los Angeles (LA)", value=float(mat_data.get('la') or 0.0), step=0.1, format="%.1f", disabled=not can_edit, key=f"la_{key}")
                mb_val, mf_val, se_val = 0.0, 0.0, 0.0
            else:
                with col_a:
                    mb_val = st.number_input("Valeur de Bleu (MB)", value=float(mat_data.get('mb') or 0.0), step=0.1, format="%.1f", disabled=not can_edit, key=f"mb_{key}")
                with col_b:
                    mf_val = st.number_input(
                        "Module de Finesse (MF - Calculé Auto)", 
                        value=float(calculated_mf), 
                        disabled=True, 
                        help="FM = Σ(Refus cumulés sur 4, 2, 1, 0.5, 0.25, 0.125 mm) / 100",
                        key=f"mf_{key}"
                    )
                    se_val = st.number_input("Équivalent de Sable (SE 10)", value=float(mat_data.get('se') or 0.0), step=0.1, format="%.1f", disabled=not can_edit, key=f"se_{key}")
                fi_val, la_val = 0.0, 0.0
                
            st.markdown("---")
            
            if can_edit:
                if st.button(f"💾 Enregistrer la feuille {mat_data['nom']}", use_container_width=True, type="primary"):
                    st.session_state['data_granulats'][key]['M1'] = new_M1
                    st.session_state['data_granulats'][key]['M2'] = new_M2
                    st.session_state['data_granulats'][key]['P'] = new_P
                    st.session_state['data_granulats'][key]['refus'] = [round(float(r), 1) for r in edited_df["Masse de refus Ri (g)"].tolist()]
                    
                    st.session_state['data_granulats'][key]['fi'] = fi_val if fi_val > 0 else None
                    st.session_state['data_granulats'][key]['la'] = la_val if la_val > 0 else None
                    st.session_state['data_granulats'][key]['mb'] = mb_val if mb_val > 0 else None
                    st.session_state['data_granulats'][key]['se'] = se_val if se_val > 0 else None

                    update_passants(st.session_state['data_granulats'][key])
                    
                    if key in ["SC", "SD"]:
                        st.session_state['data_granulats'][key]['mf'] = compute_MF(
                            st.session_state['data_granulats'][key]['sieves'],
                            st.session_state['data_granulats'][key]['passants']
                        )

                    st.session_state['success_msg'] = f"✅ Calculs enregistrés pour {mat_data['nom']} (Sous-réf: {sub_ref})"
                    st.rerun()

        with col2:
            st.subheader("Synthèse des Tamis Caractéristiques")

            # Calcul dynamique du Tamis D à partir du tableau d'analyse
            tamis_D = get_tamis_D(edited_df)
            char_sieves, char_passants = calculate_characteristic_data(mat_data, tamis_D=tamis_D)

            col_m1, col_m2 = st.columns([1, 1.5])
            with col_m1:
                st.markdown(f"**Tamis D (~95%)**")
                st.markdown(f"# {tamis_D} mm")

            with col_m2:
                df_char_mat = pd.DataFrame({
                    "Grandeur": ["2D", "1.4D", "D", "d", "d/2"],
                    "Tamis (mm)": [char_sieves['2D'], char_sieves['1.4D'], char_sieves['D'], char_sieves['d'], char_sieves['d/2']],
                    "% Passant": [
                        f"{char_passants['2D']:.1f} %",
                        f"{char_passants['1.4D']:.1f} %",
                        f"{char_passants['D']:.1f} %",
                        f"{char_passants['d']:.1f} %",
                        f"{char_passants['d/2']:.1f} %"
                    ]
                })
                st.markdown(f"**Tamis normatifs : {mat_data['nom']}**")
                st.dataframe(df_char_mat, hide_index=True, use_container_width=True)

            st.markdown("---")
            st.subheader("Courbe Granulométrique Globale")

            # Tracé de la courbe granulométrique
            fig = go.Figure()
            colors = {'GII': 'navy', 'GI': '#0284c7', 'SC': '#16a34a', 'SD': '#ea580c'}
            
            for k, d in st.session_state['data_granulats'].items():
                s_s, p_s = zip(*sorted(zip(d['sieves'], d['passants'])))
                
                line_width = 2.5 if k == key else 1.5
                opacity = 1.0 if k == key else 0.4
                
                fig.add_trace(go.Scatter(
                    x=s_s, y=p_s,
                    mode="lines+markers",
                    name=f"{d['nom']} ({new_ref_base}{SUFFIX_MAP[k]})" if new_ref_base else d['nom'],
                    line=dict(color=colors[k], width=line_width),
                    marker=dict(size=6),
                    opacity=opacity
                ))

            fig.update_layout(
                xaxis=dict(
                    title="Tamis (mm)",
                    type="log",
                    autorange="reversed"
                ),
                yaxis=dict(
                    title="% Passants Cumulés",
                    range=[0, 105]
                ),
                margin=dict(l=20, r=20, t=30, b=20),
                height=380,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )

            st.plotly_chart(fig, use_container_width=True)

    # ------------------------------------------------------------------------------
    # FENÊTRE 2 : PV D'IDENTIFICATION / SYNTHÈSE
    # ------------------------------------------------------------------------------
    with tabs[1]:
        st.header("PV d'Identification des Granulats pour Béton")
        
        pv_info_dict = st.session_state['pv_info']
        
        with st.expander("⚙️ Modifier les entêtes et signataires du PV", expanded=False):
            c1, c2, c3 = st.columns(3)
            c4, c5, c6 = st.columns(3)
            
            projet_val   = c1.text_input("Chantier / Projet", pv_info_dict.get('projet', ''), disabled=not can_edit, key="pv_proj")
            client_val   = c2.text_input("Client", pv_info_dict.get('client', ''), disabled=not can_edit, key="pv_cli")
            ref_val      = c3.text_input("N° Rapport d'Essai", pv_info_dict.get('ref_pv', ''), disabled=not can_edit, key="pv_ref")
            date_val     = c4.text_input("Date du prélèvement", pv_info_dict.get('date', ''), disabled=not can_edit, key="pv_dt")
            coord_val    = c5.text_input("Coordinateur des essais", pv_info_dict.get('coord_essais', 'O.IKEN'), disabled=not can_edit, key="pv_coo")
            chef_val     = c6.text_input("Chef du laboratoire", pv_info_dict.get('chef_labo', 'H.BAALLAL'), disabled=not can_edit, key="pv_che")

            if can_edit:
                st.session_state['pv_info']['projet'] = projet_val
                st.session_state['pv_info']['client'] = client_val
                st.session_state['pv_info']['ref_pv'] = ref_val
                st.session_state['pv_info']['date'] = date_val
                st.session_state['pv_info']['coord_essais'] = coord_val
                st.session_state['pv_info']['chef_labo'] = chef_val

        gii_data = st.session_state['data_granulats']['GII']
        gi_data  = st.session_state['data_granulats']['GI']
        sc_data  = st.session_state['data_granulats']['SC']
        sd_data  = st.session_state['data_granulats']['SD']

        # Calcul dynamique harmonisé pour chaque matériau
        gii_sieves, gii_passants = calculate_characteristic_data(gii_data)
        gi_sieves, gi_passants   = calculate_characteristic_data(gi_data)
        sc_sieves, sc_passants   = calculate_characteristic_data(sc_data)
        sd_sieves, sd_passants   = calculate_characteristic_data(sd_data)

        info_p = st.session_state['info_prelevement']
        ref_b = info_p.get('ref_base', '260/26/100')

        style_css = """
        <style>
        .lpee-pv-card {
            background-color: #ffffff;
            border: 2px solid #1e3a8a;
            border-radius: 8px;
            padding: 15px;
            font-family: Arial, sans-serif;
            color: #1e293b;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        }
        .lpee-header-title {
            background-color: #1e3a8a;
            color: #ffffff;
            text-align: center;
            font-weight: bold;
            font-size: 16px;
            padding: 10px;
            border-radius: 4px;
            letter-spacing: 0.5px;
            margin-bottom: 12px;
        }
        .lpee-info-grid {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 15px;
            font-size: 12px;
        }
        .lpee-info-grid td {
            border: 1px solid #cbd5e1;
            padding: 6px 10px;
            vertical-align: top;
        }
        .lpee-info-label {
            font-weight: bold;
            color: #0f172a;
            width: 18%;
            background-color: #f8fafc;
        }
        .lpee-norm-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 11px;
            margin-bottom: 15px;
        }
        .lpee-norm-table td {
            border: 1px solid #cbd5e1;
            padding: 4px 8px;
        }
        .lpee-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 8px;
            margin-bottom: 15px;
            font-size: 12px;
        }
        .lpee-table th {
            background-color: #2563eb;
            color: #ffffff;
            border: 1px solid #1d4ed8;
            padding: 6px;
            text-align: center;
            font-weight: bold;
        }
        .lpee-table td {
            border: 1px solid #cbd5e1;
            padding: 5px;
            text-align: center;
        }
        .row-designation {
            background-color: #f1f5f9;
            font-weight: bold;
            text-align: left !important;
        }
        .row-limite {
            background-color: #fafafa;
            font-size: 11px;
            color: #475569;
        }
        .row-fuseau {
            background-color: #f8fafc;
            font-size: 11px;
            color: #334155;
        }
        .signature-box {
            margin-top: 20px;
            width: 100%;
            border-collapse: collapse;
            font-size: 12px;
        }
        .signature-box td {
            width: 33.33%;
            border: 1px solid #cbd5e1;
            padding: 8px;
            text-align: center;
            height: 80px;
            vertical-align: top;
        }
        </style>
        """

        pv_html = f"""
        <div class="lpee-pv-card">
            <div class="lpee-header-title">
                RAPPORT D'ESSAI N° : {st.session_state['pv_info'].get('ref_pv', '')}<br>
                <span style="font-size:13px; font-weight:normal;">OBJET : IDENTIFICATION DES GRANULATS POUR BETON</span>
            </div>

            <table class="lpee-info-grid">
                <tr>
                    <td class="lpee-info-label">Client :</td>
                    <td><b>{st.session_state['pv_info'].get('client', '')}</b></td>
                    <td class="lpee-info-label">N° Dossier :</td>
                    <td>{info_p.get('dossier_no', '-')}</td>
                </tr>
                <tr>
                    <td class="lpee-info-label">Chantier :</td>
                    <td colspan="3">{st.session_state['pv_info'].get('projet', '')}</td>
                </tr>
                <tr>
                    <td class="lpee-info-label">Date du prélèvement :</td>
                    <td>{st.session_state['pv_info'].get('date', '')}</td>
                    <td class="lpee-info-label">Provenance :</td>
                    <td>{info_p.get('provenance', '-')}</td>
                </tr>
                <tr>
                    <td class="lpee-info-label">Lieu de prélèvement :</td>
                    <td>{info_p.get('lieu_prelevement', '-')}</td>
                    <td class="lpee-info-label">Réf. Échantillon :</td>
                    <td><b>{ref_b}</b></td>
                </tr>
            </table>

            <table class="lpee-norm-table">
                <tr style="background-color:#f1f5f9; font-weight:bold; text-align:center;">
                    <td colspan="5">Référence Normative</td>
                </tr>
                <tr>
                    <td><b>A.G :</b> NF EN 933-1</td>
                    <td><b>Équivalent de sable :</b> NF EN 933-8</td>
                    <td><b>VB :</b> NF EN 933-9</td>
                    <td><b>LOS ANGELES :</b> NF EN 1097-2</td>
                    <td><b>CA :</b> NF EN 933-3</td>
                </tr>
            </table>

            <!-- TABLEAU 1 : GRAVILLONS GII 10/20 -->
            <table class="lpee-table">
                <thead>
                    <tr>
                        <th rowspan="2" style="vertical-align:middle;">Désignations</th>
                        <th>2D</th><th>1,4D</th><th>D</th><th>d</th><th>d/2</th><th>f</th><th>FI</th><th>LA</th>
                    </tr>
                    <tr>
                        <th>{gii_sieves['2D']}</th>
                        <th>{gii_sieves['1.4D']}</th>
                        <th>{gii_sieves['D']}</th>
                        <th>{gii_sieves['d']}</th>
                        <th>{gii_sieves['d/2']}</th>
                        <th>% &lt; 63µm</th><th>-</th><th>-</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td class="row-designation">{gii_data['nom']} ({ref_b}/1)</td>
                        <td>{gii_passants['2D']:.0f}</td>
                        <td>{gii_passants['1.4D']:.0f}</td>
                        <td>{gii_passants['D']:.0f}</td>
                        <td>{gii_passants['d']:.0f}</td>
                        <td>{gii_passants['d/2']:.0f}</td>
                        <td>{get_passant_at_sieve(gii_data['sieves'], gii_data['passants'], 0.063):.1f}</td>
                        <td>{gii_data['fi'] if gii_data['fi'] is not None else '-'}</td>
                        <td>{gii_data['la'] if gii_data['la'] is not None else '-'}</td>
                    </tr>
                    <tr class="row-limite">
                        <td class="row-designation">Caractéristique générale de granularité</td>
                        <td>100</td><td>98 - 100</td><td>85 - 99</td><td>0 - 20</td><td>0 - 5</td><td>&lt; 1,5</td><td>FI 20</td><td>&lt; 30</td>
                    </tr>
                    <tr class="row-fuseau">
                        <td class="row-designation">Fuseau de production (Min)</td>
                        <td>100</td><td>100</td><td>94</td><td>11,5</td><td>1.4</td><td>0.6</td><td>4</td><td>23</td>
                    </tr>
                    <tr class="row-fuseau">
                        <td class="row-designation">Fuseau de production (Max)</td>
                        <td>100</td><td>100</td><td>99</td><td>1.4</td><td>0.7</td><td>1.3</td><td>16</td><td>28</td>
                    </tr>
                </tbody>
            </table>

            <!-- TABLEAU 2 : GRAVILLONS GI 4/10 -->
            <table class="lpee-table">
                <thead>
                    <tr>
                        <th rowspan="2" style="vertical-align:middle;">Désignations</th>
                        <th>2D</th><th>1,4D</th><th>D</th><th>d</th><th>d/2</th><th>f</th><th>FI</th><th>LA</th>
                    </tr>
                    <tr>
                        <th>{gi_sieves['2D']}</th>
                        <th>{gi_sieves['1.4D']}</th>
                        <th>{gi_sieves['D']}</th>
                        <th>{gi_sieves['d']}</th>
                        <th>{gi_sieves['d/2']}</th>
                        <th>% &lt; 63µm</th><th>-</th><th>-</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td class="row-designation">{gi_data['nom']} ({ref_b}/2)</td>
                        <td>{gi_passants['2D']:.0f}</td>
                        <td>{gi_passants['1.4D']:.0f}</td>
                        <td>{gi_passants['D']:.0f}</td>
                        <td>{gi_passants['d']:.0f}</td>
                        <td>{gi_passants['d/2']:.0f}</td>
                        <td>{get_passant_at_sieve(gi_data['sieves'], gi_data['passants'], 0.063):.1f}</td>
                        <td>{gi_data['fi'] if gi_data['fi'] is not None else '-'}</td>
                        <td>{gi_data['la'] if gi_data['la'] is not None else '-'}</td>
                    </tr>
                    <tr class="row-limite">
                        <td class="row-designation">Caractéristique générale de granularité</td>
                        <td>100</td><td>98 - 100</td><td>80 - 99</td><td>0 - 20</td><td>0 - 5</td><td>&lt; 1,5</td><td>FI 20</td><td>&lt; 30</td>
                    </tr>
                </tbody>
            </table>

            <!-- TABLEAU 3 : SABLE GROSSIER 0/4 -->
            <table class="lpee-table">
                <thead>
                    <tr>
                        <th rowspan="2" style="vertical-align:middle;">Désignations</th>
                        <th>2D</th><th>1,4D</th><th>D</th><th>% &lt; 1mm</th><th>% &lt; 250µm</th><th>% &lt; 63µm</th><th>MF</th><th>SE (10)</th>
                    </tr>
                    <tr>
                        <th>{sc_sieves['2D']}</th>
                        <th>{sc_sieves['1.4D']}</th>
                        <th>{sc_sieves['D']}</th>
                        <th>-</th><th>-</th><th>-</th><th>CF</th><th>-</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td class="row-designation">{sc_data['nom']} ({ref_b}/3)</td>
                        <td>{sc_passants['2D']:.0f}</td>
                        <td>{sc_passants['1.4D']:.0f}</td>
                        <td>{sc_passants['D']:.0f}</td>
                        <td>{get_passant_at_sieve(sc_data['sieves'], sc_data['passants'], 1.0):.0f}</td>
                        <td>{get_passant_at_sieve(sc_data['sieves'], sc_data['passants'], 0.25):.0f}</td>
                        <td>{get_passant_at_sieve(sc_data['sieves'], sc_data['passants'], 0.063):.1f}</td>
                        <td>{sc_data['mf'] if sc_data['mf'] is not None else '-'}</td>
                        <td>{sc_data['se'] if sc_data['se'] is not None else '-'}</td>
                    </tr>
                    <tr class="row-limite">
                        <td class="row-designation">Caractéristique générale de granularité</td>
                        <td>100</td><td>95 - 100</td><td>85 - 99</td><td>40 (±20)</td><td>50 (±20)</td><td>&le; 16</td><td>2,4 - 4,0</td><td>&ge; 60</td>
                    </tr>
                </tbody>
            </table>

            <!-- TABLEAU 4 : SABLE FIN 0/0,630 -->
            <table class="lpee-table">
                <thead>
                    <tr>
                        <th rowspan="2" style="vertical-align:middle;">Désignations</th>
                        <th>2D</th><th>1,4D</th><th>D</th><th>% &lt; 1mm</th><th>% &lt; 250µm</th><th>% &lt; 63µm</th><th>MB</th>
                    </tr>
                    <tr>
                        <th>{sd_sieves['2D']}</th>
                        <th>{sd_sieves['1.4D']}</th>
                        <th>{sd_sieves['D']}</th>
                        <th>-</th><th>-</th><th>-</th><th>-</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td class="row-designation">{sd_data['nom']} ({ref_b}/4)</td>
                        <td>{sd_passants['2D']:.0f}</td>
                        <td>{sd_passants['1.4D']:.0f}</td>
                        <td>{sd_passants['D']:.0f}</td>
                        <td>{get_passant_at_sieve(sd_data['sieves'], sd_data['passants'], 1.0):.0f}</td>
                        <td>{get_passant_at_sieve(sd_data['sieves'], sd_data['passants'], 0.25):.0f}</td>
                        <td>{get_passant_at_sieve(sd_data['sieves'], sd_data['passants'], 0.063):.1f}</td>
                        <td>{sd_data['mb'] if sd_data['mb'] is not None else '-'}</td>
                    </tr>
                    <tr class="row-limite">
                        <td class="row-designation">Caractéristique générale de granularité</td>
                        <td>100</td><td>95 - 100</td><td>85 - 99</td><td>40 (±20)</td><td>50 (±25)</td><td>&le; 10</td><td>VSS 2</td>
                    </tr>
                </tbody>
            </table>

            <!-- SIGNATURES -->
            <table class="signature-box">
                <tr>
                    <td><b>LE COORDINATEUR DES ESSAIS</b><br><br><span style="color:#64748b;">Nom: {st.session_state['pv_info'].get('coord_essais', 'O.IKEN')}</span><br>Visa:</td>
                    <td><b>LE CHEF DU LABORATOIRE</b><br><br><span style="color:#64748b;">Nom: {st.session_state['pv_info'].get('chef_labo', 'H.BAALLAL')}</span><br>Visa:</td>
                    <td><b>REÇU PAR LE CLIENT</b><br><br><span style="color:#64748b;">Nom:</span><br>Visa:</td>
                </tr>
            </table>
        </div>
        """
        
        st.markdown(style_css, unsafe_allow_html=True)
        st.markdown(clean_html(pv_html), unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        st.subheader("COURBE GRANULOMETRIQUE GLOBALE")
        fig_global_tab2 = go.Figure()
        
        colors = {'GII': '#1e40af', 'GI': '#0284c7', 'SC': '#16a34a', 'SD': '#ea580c'}
        for k, d in st.session_state['data_granulats'].items():
            s_s, p_s = zip(*sorted(zip(d['sieves'], d['passants'])))
            fig_global_tab2.add_trace(go.Scatter(
                x=s_s, y=p_s,
                mode='lines+markers',
                name=f"{d['nom']} ({ref_b}{SUFFIX_MAP[k]})" if ref_b else d['nom'],
                line=dict(color=colors[k], width=2)
            ))
            
        fig_global_tab2.update_layout(
            xaxis=dict(type="log", title="Tamis (mm)", tickvals=[0.063, 0.125, 0.25, 0.5, 1, 2, 4, 8, 16, 31.5, 63]),
            yaxis=dict(title="% Passants Cumulés", range=[0, 105]),
            height=480,
            margin=dict(l=30, r=30, t=30, b=30),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig_global_tab2, use_container_width=True)

        st.text_area("COMMENTAIRES :", value=st.session_state['pv_info'].get('commentaires', ''), disabled=not can_edit, height=80)

    # ------------------------------------------------------------------------------
    # FENÊTRE 3 : HISTORIQUE ET GESTION
    # ------------------------------------------------------------------------------
    with tabs[2]:
        st.header("Historique et Sauvegarde des PV")
        
        if can_edit:
            if st.button("📥 Sauvegarder le PV actuel dans l'historique", type="primary", use_container_width=True):
                st.session_state['historique_pv'].append({
                    'date_creation': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'ref_pv': st.session_state['pv_info'].get('ref_pv', ''),
                    'projet': st.session_state['pv_info'].get('projet', ''),
                    'client': st.session_state['pv_info'].get('client', ''),
                    'data': st.session_state['data_granulats']
                })
                st.success("✅ Le PV a été ajouté à l'historique de la session !")
                
        if st.session_state['historique_pv']:
            st.write("### 📜 Liste des PV sauvegardés (Session Actuelle)")
            for i, pv in enumerate(reversed(st.session_state['historique_pv'])):
                with st.expander(f"📁 {pv['date_creation']} - {pv['ref_pv']} - Chantier: {pv['projet']}", expanded=(i==0)):
                    st.json(pv)
        else:
            st.info("Aucun PV n'a été sauvegardé dans cette session.")
