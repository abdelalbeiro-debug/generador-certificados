# =================================================================
# SOFTWARE: Generador de Certificados Auditoría PRO
# DESARROLLADOR: Abdel Areiza
# VERSIÓN: 1.0
# DESCRIPCIÓN: Automatización de certificados contables en Word
# =================================================================

import pandas as pd
import streamlit as st
from docx import Document
from io import BytesIO
import re
from docx.shared import Pt

def limpiar_monto(valor):
    """Convierte textos contables con puntos y comas a números reales."""
    if pd.isna(valor) or valor == "": return 0.0
    s = str(valor).strip()
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    s = re.sub(r'[^\d\.-]', '', s)
    try:
        return float(s)
    except:
        return 0.0

def crear_word(fila_datos, nombre_t, nit_t, ano_t, rubros_disponibles, mapa_tipos):
    doc = Document()
    # --- PROPIEDADES DE AUTOR (ABDEL AREIZA) ---
    doc.core_properties.author = "Abdel Areiza"
    doc.core_properties.title = f"Certificado de Auditoría - {nombre_t}"
    doc.core_properties.comments = "Documento generado por el aplicativo de Abdel Areiza"
    
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(12)

    # 1. Encabezado
    p1 = doc.add_paragraph()
    p1.alignment = 1
    p1.add_run(f"{nombre_t}").bold = True
    
    p2 = doc.add_paragraph()
    p2.alignment = 1
    p2.add_run(f"NIT {nit_t}").bold = True
    
    doc.add_paragraph("\nCERTIFICA QUE\n").alignment = 1
    
    # 2. Cuerpo
    p3 = doc.add_paragraph()
    p3.add_run("La sociedad ").bold = False
    p3.add_run("AGROPECUARIA SANTAMARIA S.A.").bold = True
    p3.add_run(", identificada según ").bold = False
    p3.add_run("NIT.830.075.074-8").bold = True
    p3.add_run(", teniendo en cuenta los saldos registrados en los libros de contabilidad con esta entidad; al corte 31 de diciembre de ")
    p3.add_run(f"{ano_t}").bold = True
    p3.add_run(", presenta los siguientes saldos contables:").bold = False
    
    doc.add_paragraph()

    # 3. Listado Agrupado
    categorias = {"De Activos": [], "De Pasivos": []}
    for rubro in rubros_disponibles:
        valor = float(fila_datos.get(rubro, 0))
        if abs(valor) > 0.01:
            tipo = mapa_tipos.get(rubro, "Otros")
            if tipo not in categorias: categorias[tipo] = []
            # Trae el rubro tal cual del Excel como pidió
            categorias[tipo].append(f"{rubro}: $ {abs(valor):,.2f}")

    for nombre_cat in ["De Activos", "De Pasivos"]:
        lineas = categorias.get(nombre_cat, [])
        if lineas:
            p_cat = doc.add_paragraph()
            p_cat.add_run(f"{nombre_cat}:").bold = True
            # Quitamos espacio después del título de la categoría
            p_cat.paragraph_format.space_after = Pt(0) 
            
            for linea in lineas:
                p_item = doc.add_paragraph(linea)
                # --- AQUÍ QUITAMOS EL INTERLINEADO ---
                p_item.paragraph_format.space_after = Pt(0) # Espacio cero después
                p_item.paragraph_format.line_spacing = 1.0  # Interlineado sencillo
            
            # Espacio pequeño al final de cada grupo para que no se pegue con el siguiente título
            doc.add_paragraph().paragraph_format.space_after = Pt(0)

    # 4. Firma
    doc.add_paragraph(f"\nEsta certificación se expide el día 7 de mayo de 2026.")
    p_firma = doc.add_paragraph("\nAtentamente\n\nNombre: ______________________\n")
    p_firma.add_run("Firma: _______________\n")
    p_firma.add_run("Contador Público\n")
    p_firma.add_run("CC: _________________\n")
    p_firma.add_run("TP: __________")
    
    target = BytesIO()
    doc.save(target)
    return target

# --- INTERFAZ ---
st.set_page_config(page_title="Generador Auditoría", layout="wide")

# --- INTERFAZ ---
st.set_page_config(page_title="Generador Auditoría", layout="wide")
st.title("📑 Generador de certificados PRO")

# 1. El cargador de archivos
archivo = st.file_uploader("Sube tu Excel", type=["xlsx"])

# 2. Si NO han subido nada, mostramos el mensaje azul de bienvenida
if not archivo:
    st.info("👋 ¡Hola Abdel! Por favor, sube tu archivo de Excel arriba para empezar a generar los certificados.")

# 3. Si YA subieron el archivo, procesamos todo una sola vez
else:
    df = pd.read_excel(archivo)
    df.columns = df.columns.str.upper().str.strip()
    
    # Lista de columnas obligatorias
    req = ['AÑO', 'NOMBRE CUENTA', 'TERCERO', 'NIT', 'VALOR EN LIBROS', 'TIPO']
    
    # Verificamos si el Excel es el correcto
    if all(c in df.columns for c in req):
        st.success("✅ ¡Base de datos cargada con éxito!")
        
        # Limpieza de datos (Asegúrate de que estas líneas tengan 8 espacios de sangría)
        df['NIT'] = df['NIT'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        df['VALOR EN LIBROS'] = df['VALOR EN LIBROS'].apply(limpiar_monto)
        df['NOMBRE CUENTA'] = df['NOMBRE CUENTA'].astype(str).str.strip()
        df['TIPO'] = df['TIPO'].astype(str).str.strip()
        
        # Lógica del certificado
        mapa_tipos = pd.Series(df.TIPO.values, index=df['NOMBRE CUENTA']).to_dict()
        resumen = df.groupby(['AÑO', 'NIT', 'TERCERO', 'NOMBRE CUENTA'])['VALOR EN LIBROS'].sum().unstack(fill_value=0).reset_index()
        rubros = [c for c in resumen.columns if c not in ['AÑO', 'NIT', 'TERCERO']]
        
        resumen['ETIQUETA'] = resumen['TERCERO'] + " (" + resumen['NIT'] + ") - Año: " + resumen['AÑO'].astype(str)
        
        # Buscador de personas
        seleccionados = st.multiselect("Selecciona los certificados:", resumen['ETIQUETA'].unique())
        
        if seleccionados:
            if st.button("🚀 Preparar Descargas"):
                for i, sel in enumerate(seleccionados):
                    fila = resumen[resumen['ETIQUETA'] == sel].iloc[0]
                    doc_final = crear_word(fila, fila['TERCERO'], fila['NIT'], fila['AÑO'], rubros, mapa_tipos)
                    st.download_button(
                        label=f"📥 Descargar: {sel}",
                        data=doc_final.getvalue(),
                        file_name=f"Certificado_{fila['NIT']}.docx",
                        key=f"dl_{i}_{fila['NIT']}"
                    )
    else:
        # Este error solo sale si el archivo está mal
        st.error(f"⚠️ El Excel no tiene las columnas correctas. Necesito: {req}")

# 4. Los créditos (Pega esto al puro final, sin espacios a la izquierda)
st.markdown("---")
st.caption("🛠️ Desarrollado por **Abdel Areiza**")
