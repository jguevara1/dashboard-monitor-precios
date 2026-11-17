import streamlit as st
import pandas as pd
import os
from datetime import datetime
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
        fecha_act_str = datetime.fromtimestamp(timestamp).strftime('%d/%m/%Y %I:%M %p')
except: pass

try:
    df_base = pd.read_excel("Base_Actualizada.xlsx")
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

def fmt_precio(val):
    try: 
        v = float(val)
        if v == 0.0: return "" 
        return "${:.2f}".format(v)
    except: return str(val) if str(val).lower() not in ["nan", "none", ""] else ""

def fmt_pct(val):
    try: 
        v = float(val)
        if v == 0.0: return ""
        return "{:.1%}".format(v)
    except: return str(val) if str(val).lower() not in ["nan", "none", ""] else ""

def fmt_pct_ms(val):
    try: return "{:.0%}".format(float(val))
    except: return str(val) if str(val).lower() not in ["nan", "none", ""] else ""

if base_lista:
    st.title("Monitor de Precios Competitivos")
    
    lista_grupos = df_base["Grupo_Calox"].dropna().unique().tolist()
    grupo_seleccionado = st.selectbox("Seleccione el Producto:", lista_grupos, label_visibility="collapsed")
    
    df_grupo = df_base[df_base["Grupo_Calox"] == grupo_seleccionado].copy()
    
    st.markdown("#### Matriz de Revisión de Precios")
    
    datos_finales = []
    for _, f in df_grupo.iterrows():
        prod_name = str(f.get("Producto", ""))
        if prod_name.strip() == "" or prod_name.lower() in ["nan", "none"]:
            continue
            
        f_ftodo, d_ftodo = f.get("PFull_Farmatodo", 0.0), f.get("PDesc_Farmatodo", 0.0)
        f_cob, d_cob = f.get("PFull_Cobeca", 0.0), f.get("PDesc_Cobeca", 0.0)
        f_nen, d_nen = f.get("PFull_Nena", 0.0), f.get("PDesc_Nena", 0.0)
        f_sans, d_sans = f.get("PFull_SanAnselmo", 0.0), f.get("PDesc_SanAnselmo", 0.0)
        f_fdon, d_fdon = f.get("PFull_Farmadon", 0.0), f.get("PDesc_Farmadon", 0.0)
        f_ns, d_ns = f.get("PFull_NuevoSiglo", 0.0), f.get("PDesc_NuevoSiglo", 0.0)
        f_loc, d_loc = f.get("PFull_Locatel", 0.0), f.get("PDesc_Locatel", 0.0)
        f_saas, d_saas = f.get("PFull_Saas", 0.0), f.get("PDesc_Saas", 0.0)

        ms_q_val = f.get("%MS Q", 0) 
        if pd.isna(ms_q_val) or str(ms_q_val).strip().lower() in ["nan", "none", ""]: ms_q_val = 0
        ms_q_str = fmt_pct_ms(ms_q_val) # Formateamos el MS directamente aquí

        pres_val = f.get("Presentación", f.get("Presentacion", ""))
        try: pres_val = str(int(float(pres_val))) 
        except: pres_val = str(pres_val) if str(pres_val).lower() not in ["nan", "none"] else ""
        
        obj_val = str(f.get("Objetivo de Mercado", ""))
        if obj_val.lower() in ["nan", "none", "null"]: obj_val = ""

        datos_finales.append([
            ms_q_str, prod_name, f.get("Laboratorio", ""), pres_val, obj_val,                 
            f_ftodo, d_ftodo, calc_dcto(f_ftodo, d_ftodo),
            f_cob, d_cob, calc_dcto(f_cob, d_cob),
            f_nen, d_nen, calc_dcto(f_nen, d_nen),
            f_sans, d_sans, calc_dcto(f_sans, d_sans),
            f_fdon, d_fdon, calc_dcto(f_fdon, d_fdon),
            f_ns, d_ns, calc_dcto(f_ns, d_ns),
            f_loc, d_loc, calc_dcto(f_loc, d_loc),
            f_saas, d_saas, calc_dcto(f_saas, d_saas)
        ])

    columnas_agrupadas = pd.MultiIndex.from_tuples([
        ("", "%MS Q"), ("", "Producto"), ("", "Laboratorio"), ("", "Pres."), ("", "Obj. Merc."),
        ("🔵 Farmatodo", "P. Full"), ("🔵 Farmatodo", "P. Oferta"), ("🔵 Farmatodo", "% Desc."),
        ("🟢 Cobeca", "P. Full"), ("🟢 Cobeca", "P. Oferta"), ("🟢 Cobeca", "% Desc."),
        ("🟠 Nena", "P. Full"), ("🟠 Nena", "P. Oferta"), ("🟠 Nena", "% Desc."),
        ("🟣 San Anselmo", "P. Full"), ("🟣 San Anselmo", "P. Oferta"), ("🟣 San Anselmo", "% Desc."),
        ("🟡 Farmadon", "P. Full"), ("🟡 Farmadon", "P. Oferta"), ("🟡 Farmadon", "% Desc."),
        ("⚪ Nuevo Siglo", "P. Full"), ("⚪ Nuevo Siglo", "P. Oferta"), ("⚪ Nuevo Siglo", "% Desc."),
        ("🔴 Locatel", "P. Full"), ("🔴 Locatel", "P. Oferta"), ("🔴 Locatel", "% Desc."),
        ("🔘 Farm. Saas", "P. Full"), ("🔘 Farm. Saas", "P. Oferta"), ("🔘 Farm. Saas", "% Desc.")
    ])

    df_final = pd.DataFrame(datos_finales, columns=columnas_agrupadas)
    
    # Convertimos las 5 primeras columnas en el Índice para que queden Fijas al hacer scroll
    df_final.set_index([("", "%MS Q"), ("", "Producto"), ("", "Laboratorio"), ("", "Pres."), ("", "Obj. Merc.")], inplace=True)

    def aplicar_estilos_matriz(df):
        styles = pd.DataFrame('', index=df.index, columns=df.columns)
        for i in range(len(df.index)):
            # El laboratorio ahora vive en el nivel 2 del índice
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
                if is_calox: bg_color = "#e0e0e0" 
                
                borde_izq = "border-left: 2px solid #bdc3c7;" if campo == "P. Full" else ""
                styles.iat[i, df.columns.get_loc(col)] = f"background-color: {bg_color}; font-weight: {font_weight}; {borde_izq}"
        
        return styles

    formato_columnas = {
        ("🔵 Farmatodo", "P. Full"): fmt_precio, ("🔵 Farmatodo", "P. Oferta"): fmt_precio, ("🔵 Farmatodo", "% Desc."): fmt_pct,
        ("🟢 Cobeca", "P. Full"): fmt_precio, ("🟢 Cobeca", "P. Oferta"): fmt_precio, ("🟢 Cobeca", "% Desc."): fmt_pct,
        ("🟠 Nena", "P. Full"): fmt_precio, ("🟠 Nena", "P. Oferta"): fmt_precio, ("🟠 Nena", "% Desc."): fmt_pct,
        ("🟣 San Anselmo", "P. Full"): fmt_precio, ("🟣 San Anselmo", "P. Oferta"): fmt_precio, ("🟣 San Anselmo", "% Desc."): fmt_pct,
        ("🟡 Farmadon", "P. Full"): fmt_precio, ("🟡 Farmadon", "P. Oferta"): fmt_precio, ("🟡 Farmadon", "% Desc."): fmt_pct,
        ("⚪ Nuevo Siglo", "P. Full"): fmt_precio, ("⚪ Nuevo Siglo", "P. Oferta"): fmt_precio, ("⚪ Nuevo Siglo", "% Desc."): fmt_pct,
        ("🔴 Locatel", "P. Full"): fmt_precio, ("🔴 Locatel", "P. Oferta"): fmt_precio, ("🔴 Locatel", "% Desc."): fmt_pct,
        ("🔘 Farm. Saas", "P. Full"): fmt_precio, ("🔘 Farm. Saas", "P. Oferta"): fmt_precio, ("🔘 Farm. Saas", "% Desc."): fmt_pct
    }

    df_estilizado = df_final.style.apply(aplicar_estilos_matriz, axis=None).format(formato_columnas)
    
    # Función exclusiva para poner en negrita a Calox en las columnas congeladas
    def estilo_indice(val):
        if 'CALOX' in str(val).upper():
            return 'font-weight: bold; color: black;'
        return ''
    
    df_estilizado = df_estilizado.map_index(estilo_indice, axis=0)
    
    # Renderizamos sin ocultar el índice, ya que ahora contiene la información importante
    st.dataframe(df_estilizado, use_container_width=True)