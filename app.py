import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

st.set_page_config(page_title="Monitor de Precios", layout="wide")

st.markdown("""
    <style>
    .block-container { padding-top: 2rem; padding-bottom: 1rem; padding-left: 1rem; padding-right: 1rem; max-width: 100%; }
    [data-testid="stDataFrame"] { font-size: 13px; }
    h1, h2, h3, h4 { color: #800020 !important; }
    </style>
""", unsafe_allow_html=True)

st.sidebar.header("Estado del Sistema")

fecha_act_str = "Desconocida"
try:
    if os.path.exists("Base_Actualizada.xlsx"):
        timestamp = os.path.getmtime("Base_Actualizada.xlsx")
        fecha_ve = datetime.fromtimestamp(timestamp) - timedelta(hours=4)
        fecha_act_str = fecha_ve.strftime('%d/%m/%Y %I:%M %p')
except: pass

try:
    df_base = pd.read_excel("Base_Actualizada.xlsx")
    
    df_base.columns = [c[:-2] if c.endswith('_x') else c for c in df_base.columns]
    df_base = df_base.loc[:, ~df_base.columns.str.endswith('_y')]
    
    st.sidebar.success("✅ Base conectada.")
    st.sidebar.caption(f"📅 **Última Extracción:**\n{fecha_act_str}")
    
    tasa_usada = df_base["Tasa_BCV"].iloc[0] if "Tasa_BCV" in df_base.columns else 0.0
    if tasa_usada > 0: st.sidebar.info(f"💵 **Tasa BCV Aplicada:** {tasa_usada:.2f} Bs/USD")
    else: st.sidebar.warning("💵 **Tasa BCV Aplicada:** No registrada")
    base_lista = True
except:
    st.sidebar.error("❌ Archivo 'Base_Actualizada.xlsx' no encontrado.")
    base_lista = False

def calc_dcto(full, desc):
    if full > 0 and desc > 0 and full > desc: return (full - desc) / full
    return 0.0

def precio_efectivo(full, desc):
    try:
        d = float(desc) if not pd.isna(desc) else 0.0
        f = float(full) if not pd.isna(full) else 0.0
        return d if d > 0 else f
    except:
        return 0.0

def calc_vs_calox(precio_calox, precio_comp, es_calox):
    if es_calox or precio_calox <= 0 or precio_comp <= 0:
        return None
    return (precio_calox - precio_comp) / precio_comp

def fmt_precio(val):
    try: 
        v = float(val)
        if pd.isna(v) or v == 0.0: return "" 
        return "${:.2f}".format(v)
    except: return str(val) if str(val).lower() not in ["nan", "none", ""] else ""

def fmt_pct(val):
    try: 
        v = float(val)
        if pd.isna(v) or v == 0.0: return ""
        return "{:.1%}".format(v)
    except: return str(val) if str(val).lower() not in ["nan", "none", ""] else ""

def fmt_diff(val):
    try:
        if val is None or pd.isna(val): return ""
        v = float(val)
        return "{:+.1%}".format(v)
    except: return ""

def fmt_pct_ms(val):
    try: return "{:.0%}".format(float(val))
    except: return str(val) if str(val).lower() not in ["nan", "none", ""] else ""

def limpiar_pres(val):
    try: return str(int(float(val)))
    except: return str(val).strip() if str(val).lower() not in ["nan", "none"] else ""

if base_lista:
    st.title("Monitor de Precios Competitivos")
    
    lista_grupos = df_base["Grupo_Calox"].dropna().unique().tolist()
    grupo_seleccionado = st.selectbox("Seleccione el Producto:", lista_grupos, label_visibility="collapsed")
    
    df_grupo = df_base[df_base["Grupo_Calox"] == grupo_seleccionado].copy()
    
    if "Codigo_IMS" in df_grupo.columns:
        df_grupo = df_grupo.drop_duplicates(subset=["Codigo_IMS"])
    
    st.markdown("#### Matriz de Revisión de Precios")
    
    canales = [
        ("Farmatodo", "PFull_Farmatodo", "PDesc_Farmatodo"),
        ("Cobeca", "PFull_Cobeca", "PDesc_Cobeca"),
        ("Nena", "PFull_Nena", "PDesc_Nena"),
        ("SanAnselmo", "PFull_SanAnselmo", "PDesc_SanAnselmo"),
        ("Farmadon", "PFull_Farmadon", "PDesc_Farmadon"),
        ("NuevoSiglo", "PFull_NuevoSiglo", "PDesc_NuevoSiglo"),
        ("Locatel", "PFull_Locatel", "PDesc_Locatel"),
        ("Saas", "PFull_Saas", "PDesc_Saas"),
    ]

    # Extraemos los precios efectivos de Calox (por presentación y general del grupo)
    df_calox = df_grupo[df_grupo["Laboratorio"].astype(str).str.upper().str.contains("CALOX", na=False)]
    calox_por_pres = {}
    calox_general = {c[0]: 0.0 for c in canales}

    for _, fc in df_calox.iterrows():
        p_key = limpiar_pres(fc.get("Presentación", fc.get("Presentacion", "")))
        if p_key not in calox_por_pres:
            calox_por_pres[p_key] = {}
        for nombre_c, col_f, col_d in canales:
            pe = precio_efectivo(fc.get(col_f, 0.0), fc.get(col_d, 0.0))
            if pe > 0:
                if calox_por_pres[p_key].get(nombre_c, 0.0) == 0.0:
                    calox_por_pres[p_key][nombre_c] = pe
                if calox_general[nombre_c] == 0.0:
                    calox_general[nombre_c] = pe

    datos_finales = []
    for _, f in df_grupo.iterrows():
        prod_name = str(f.get("Producto", f.get("PACK", "")))
        if prod_name.strip() == "" or prod_name.lower() in ["nan", "none"]:
            continue
            
        lab_str = str(f.get("Laboratorio", ""))
        es_calox = "CALOX" in lab_str.upper()

        ms_q_val = f.get("%MS Q", 0) 
        if pd.isna(ms_q_val) or str(ms_q_val).strip().lower() in ["nan", "none", ""]: ms_q_val = 0
        ms_q_str = fmt_pct_ms(ms_q_val)

        pres_val = limpiar_pres(f.get("Presentación", f.get("Presentacion", "")))
        
        obj_val = str(f.get("Objetivo de Mercado", ""))
        if obj_val.lower() in ["nan", "none", "null"]: obj_val = ""

        fila_datos = [ms_q_str, prod_name, lab_str, pres_val, obj_val]

        for nombre_c, col_f, col_d in canales:
            p_full = float(f.get(col_f, 0.0) or 0.0)
            p_desc = float(f.get(col_d, 0.0) or 0.0)
            p_comp_ef = precio_efectivo(p_full, p_desc)
            
            # Buscamos el precio de Calox de la misma presentación o el general del grupo
            p_calox_ref = calox_por_pres.get(pres_val, {}).get(nombre_c, 0.0)
            if p_calox_ref == 0.0:
                p_calox_ref = calox_general.get(nombre_c, 0.0)
                
            dif_pct = calc_vs_calox(p_calox_ref, p_comp_ef, es_calox)
            fila_datos.extend([p_full, p_desc, calc_dcto(p_full, p_desc), dif_pct])

        datos_finales.append(fila_datos)

    if not datos_finales:
        st.warning("No hay productos registrados para mostrar en este grupo.")
    else:
        grupos_ui = [
            "🔵 Farmatodo", "🟢 Cobeca", "🟠 Nena", "🟣 San Anselmo",
            "🟡 Farmadon", "⚪ Nuevo Siglo", "🔴 Locatel", "🔘 Farm. Saas"
        ]
        
        tuplas_cols = [("", "%MS Q"), ("", "Producto"), ("", "Laboratorio"), ("", "Pres."), ("", "Obj. Merc.")]
        for g in grupos_ui:
            tuplas_cols.extend([(g, "P. Full"), (g, "P. Oferta"), (g, "% Desc."), (g, "Dif. %")])

        columnas_agrupadas = pd.MultiIndex.from_tuples(tuplas_cols)
        df_final = pd.DataFrame(datos_finales, columns=columnas_agrupadas)
        
        df_final = df_final.drop_duplicates(subset=[("", "%MS Q"), ("", "Producto"), ("", "Laboratorio"), ("", "Pres."), ("", "Obj. Merc.")])
        df_final.set_index([("", "%MS Q"), ("", "Producto"), ("", "Laboratorio"), ("", "Pres."), ("", "Obj. Merc.")], inplace=True)

        def aplicar_estilos_matriz(df):
            styles = pd.DataFrame('', index=df.index, columns=df.columns)
            for i in range(len(df.index)):
                is_calox = "CALOX" in str(df.index[i][2]).upper()
                for col in df.columns:
                    farmacia, campo = col[0], col[1]
                    
                    if "Farmatodo" in farmacia: bg_color = "#eaf2f8"      
                    elif "Cobeca" in farmacia: bg_color = "#e9f7ef"       
                    elif "Nena" in farmacia: bg_color = "#fdf2e9"         
                    elif "San Anselmo" in farmacia: bg_color = "#f5eef8"  
                    elif "Farmadon" in farmacia: bg_color = "#fcf3cf"     
                    elif "Nuevo Siglo" in farmacia: bg_color = "#f2f3f4"  
                    elif "Locatel" in farmacia: bg_color = "#fdedec"      
                    elif "Saas" in farmacia: bg_color = "#f2f4f4"         
                    else: bg_color = "#ffffff"                            
                    
                    font_weight = "bold" if is_calox else "normal"
                    text_color = ""
                    if is_calox: 
                        bg_color = "#e0e0e0" 
                    elif campo == "Dif. %":
                        val = df.iat[i, df.columns.get_loc(col)]
                        if val is not None and not pd.isna(val):
                            if val < 0:
                                text_color = "color: #1e8449; font-weight: bold;" # Verde: Calox más barato
                            elif val > 0:
                                text_color = "color: #c0392b; font-weight: bold;" # Rojo: Calox más caro
                    
                    borde_izq = "border-left: 2px solid #bdc3c7;" if campo == "P. Full" else ""
                    styles.iat[i, df.columns.get_loc(col)] = f"background-color: {bg_color}; font-weight: {font_weight}; {text_color} {borde_izq}"
            
            return styles

        formato_columnas = {}
        for g in grupos_ui:
            formato_columnas[(g, "P. Full")] = fmt_precio
            formato_columnas[(g, "P. Oferta")] = fmt_precio
            formato_columnas[(g, "% Desc.")] = fmt_pct
            formato_columnas[(g, "Dif. %")] = fmt_diff

        df_estilizado = df_final.style.apply(aplicar_estilos_matriz, axis=None).format(formato_columnas)
        
        def estilo_indice(val):
            if 'CALOX' in str(val).upper():
                return 'font-weight: bold; color: black;'
            return ''
        
        df_estilizado = df_estilizado.map_index(estilo_indice, axis=0)
        
        st.dataframe(df_estilizado, use_container_width=True)
