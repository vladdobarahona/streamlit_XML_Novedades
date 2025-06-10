# -*- coding: utf-8 -*-
"""
Created on Tue May 20 14:34:03 2025

@author: vbarahona
"""

# importar librerías
import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
import tempfile
import openpyxl
from datetime import date  # Faltaba importar date
from io import BytesIO

# Fondo personalizado y fuente
st.markdown("""
<style>
    body {
        background-color:rgb(171 , 190 , 76);
        font-family: 'Handel Gothic', 'Frutiger light - Roman';
    }
    .stApp {
        background-color: rgb(255, 255, 255);
        font-family: 'Frutiger Bold', sans-serif;
    }
</style>
    """, unsafe_allow_html=True)
# Logo a la izquierda y título a la derecha
col1, col2 = st.columns([1, 2])
with col1:
    st.image('https://www.finagro.com.co/sites/default/files/logo-front-finagro.png', width=200)
with col2:
    st.markdown(
        '<h1 style="color: rgb(120,154,61); font-size: 2.25rem; font-weight: bold;">Convertidor de archivo Excel a XML Novedades</h1>',
        unsafe_allow_html=True
    )

# Cargar el archivo Excel desde el repositorio local o remoto
Plantilla_Excel = pd.read_excel("excel_novedades_xml.xlsx", sheet_name='Novedades', engine="openpyxl", dtype=str)

# Convertir el DataFrame a un archivo Excel en memoria
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Novedades')
    processed_data = output.getvalue()
    return processed_data

excel_bytes = to_excel(Plantilla_Excel)

# Crear el formulario y botón de descarga
with st.form("Plantilla Excel"):
    st.form_submit_button("Preparar descarga")
    st.download_button(
        label="📥 Descargar plantilla Excel",
        data=excel_bytes,
        file_name="excel_novedades_xml.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

#icon=":material/download:",
# Columnas predeterminadas para el archivo Excel
required_columns = [
    'TIPO NOVEDAD',
    'MOTIVO_ABONO',
    'DESTINO_ABONO',
    'TIPO_CARTERA',
    'INTERMEDIARIO',
    'NUMERO_OBLIGACION_AGROS',
    'TIPO_DOCUMENTO',
    'NUMERO_DOCUMENTO',
    'VALOR_CAPITAL_ABONO'
]

st.markdown(
    '<span style="color: rgb(120, 154, 61); font-size: 22px;">Validador de Columnas Requeridas</span>',
    unsafe_allow_html=True
)
xls_file = st.file_uploader("", type=["xlsx", "xls"])

if xls_file:
    df = pd.read_excel(xls_file, engine='openpyxl')
    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        st.error("❌ Faltan las siguientes columnas en el archivo:")
        for col in missing_columns:
            st.markdown(f"- **{col}**")
    else:
        st.success("✅ Todas las columnas requeridas están presentes.")
        df = df.dropna(subset=['NUMERO_OBLIGACION_AGROS'])

        st.dataframe(df)

        Valor_creditos = str(sum(df['VALOR_CAPITAL_ABONO'].astype('float64')))
        Cantidad_registros = str(len(df))

        # Formulario de parámetros
        with st.form("form_parametros"):
            fecha_novedades_str = st.date_input("Fecha de aplicación novedades", value=date.today())
            submitted = st.form_submit_button("Confirmar parámetros")

        if submitted:
            st.subheader("Resumen de datos ingresados:")
            st.write(f"Fecha de aplicación novedades: {fecha_novedades_str.strftime('%Y-%m-%d')}")
            st.write(f"Cantidad de registros: {Cantidad_registros}")
            valor = sum(df['VALOR_CAPITAL_ABONO'].astype('float64'))
            st.markdown(f"<h4 style='color:#789a3d;'>Valor total capital: ${valor:,.2f}</h4>", unsafe_allow_html=True)

            st.header("Generación de XML", divider=True)

            try:
                ET.register_namespace('', "http://www.finagro.com.co/sit")
                abonos = ET.Element("{http://www.finagro.com.co/sit}abonos", cifraDeControl=Cantidad_registros)

                for index, row in df.iterrows():
                    abono = ET.SubElement(abonos, "{http://www.finagro.com.co/sit}abono",
                                          tipoNovedadPago="2",
                                          codigoMotivoAbono=str(row['MOTIVO_ABONO']),
                                          destinoAbono=str(row['DESTINO_ABONO']),
                                          fechaAplicacionPago=str(fecha_novedades_str.strftime('%Y-%m-%d'))
                                          )

                    informacionObligacion = ET.SubElement(abono, "{http://www.finagro.com.co/sit}informacionObligacion",
                                                          tipoCarteraId=str(row['TIPO_CARTERA']),
                                                          codigoIntermediario=str(row['INTERMEDIARIO']),
                                                          numeroObligacion=str(row['NUMERO_OBLIGACION_AGROS']),
                                                          tipoMonedaId="1"
                                                          )

                    informacionBeneficiario = ET.SubElement(informacionObligacion, "{http://www.finagro.com.co/sit}informacionBeneficiario",
                                                            tipoDocumentoId=str(row['TIPO_DOCUMENTO']),
                                                            numeroDocumento=str(row['NUMERO_DOCUMENTO'])
                                                            )

                    valorAbono = ET.SubElement(abono, "{http://www.finagro.com.co/sit}valorAbono")
                    valorAbonoCapital = ET.SubElement(valorAbono, "{http://www.finagro.com.co/sit}valorAbonoCapital", {"xmlns": ""})
                    valorAbonoCapital.text = str(row['VALOR_CAPITAL_ABONO'])

                # Función para sanitizar elementos
                def sanitize_element(element):
                    if element.text is not None and not isinstance(element.text, str):
                        element.text = str(element.text)
                    for key, value in element.attrib.items():
                        if not isinstance(value, str):
                            element.attrib[key] = str(value)
                    for child in element:
                        sanitize_element(child)

                sanitize_element(abonos)

                tree = ET.ElementTree(abonos)
                ET.indent(tree, space="  ", level=0)

                with tempfile.NamedTemporaryFile(delete=False, suffix=".xml") as tmp:
                    tree.write(tmp.name, encoding="UTF-8", xml_declaration=True)
                    st.success("✅ XML de novedades generado exitosamente.")
                    with open(tmp.name, "rb") as f:
                        st.download_button("📥 Descargar XML de Novedades", f, file_name="Novedades.xml", mime="application/xml")

            except Exception as e:
                st.error(f"Ocurrió un error al generar el XML: {e}")
