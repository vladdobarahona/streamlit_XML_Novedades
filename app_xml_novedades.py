# -*- coding: utf-8 -*-
"""
Created on Tue May 20 14:34:03 2025

@author: vbarahona
"""

# importar librerias
import streamlit as st
import xml.etree.ElementTree as ET
from datetime import date
from dateutil.relativedelta import relativedelta
import pandas as pd
from decimal import Decimal
import tempfile
import openpyxl
#background-color: rgb(255, 255, 255);
#background-color: rgb(171 , 190 , 76)
#background-color: rgb(120, 154, 61);
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
        '<h1 style="color: rgb(120,154,61); font-size: 2.25rem; font-weight: bold;">Generador de XML de Obligaciones a partir de un archivo Excel</h1>',
        unsafe_allow_html=True
    )

# Columnas predeterminadas para el archivo Excel
required_columns = [
    'Tipo de Cartera', 'Tipo de Productor', 'Número de Pagare', 'Número de Pagare Anterior',
    'Fecha de Suscripción', 'Ciudad de Inversión', 'Numero de Identificacion',
    'Tipo Identificacion Finagro', 'Nombre Razón Social', 'Email Beneficiario',
    'Teléfono Beneficiario', 'Fecha de Activos', 'Monto Activos', 'Dirección Beneficiario',
    'Plazo', 'Tipo Plan de Pagos', 'Capital Total', 'Porcentaje Fag', 'Indicativo Fag',
    'Tipo Comisión', 'Puntos IBR', 'Ubicación Predio', 'Código Oficina', 'Producto Relacionado',
    'Código Destino 1', 'Unidades Destino 1', 'Costo Inversión 1', 'Valor a Financiar 1',
    'Código Destino 2', 'Unidades Destino 2', 'Costo Inversión 2', 'Valor a Financiar 2',
    'Código Destino 3', 'Unidades Destino 3', 'Costo Inversión 3', 'Valor a Financiar 3',
    'Código Destino 4', 'Unidades Destino 4', 'Costo Inversión 4', 'Valor a Financiar 4',
    'Valor Ingresos', 'Fecha de Ingresos', 'Moneda Ingresos', 'Moneda de Activos'
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
        df = df.dropna(subset=['Número de Pagare'])
        df['Fecha de Ingresos'] = pd.to_datetime(df['Fecha de Ingresos'], format='%Y/%m/%d')
        df['Fecha de Activos'] = pd.to_datetime(df['Fecha de Activos'], format='%Y%m%d')
        df['Fecha de Suscripción'] = pd.to_datetime(df['Fecha de Suscripción'], format='%Y%m%d')

        df['Fecha de Ingresos'] = df['Fecha de Ingresos'].dt.strftime('%Y-%m-%d')
        df['Fecha de Activos'] = df['Fecha de Activos'].dt.strftime('%Y-%m-%d')
        df['Fecha de Suscripción'] = df['Fecha de Suscripción'].dt.strftime('%Y-%m-%d')
	
        st.dataframe(df)
        
        Valor_creditos = str(sum(df['Capital Total'].astype('float64')))
        Cantidad_creditos = str(len(df))
        
        # XML generation logic here
        ET.register_namespace('', "http://www.finagro.com.co/sit")
        obligaciones = ET.Element("{http://www.finagro.com.co/sit}obligaciones",
                                    cifraDeControl=Cantidad_creditos,
                                    cifraDeControlValor=Valor_creditos)
        #fecha_Desembolso = fecha_Desembolso_str.strftime('%Y-%m-%d')
        fecha_Desembolso = date(2025, 5, 9) # indicar fecha desembolso
        cod_programa = '126' # indicar código del programa
        cod_intermediario = '203018' # indicar código del intermediario
        tipo_plan = 0 # solo va 1 o cero | # si tipo_plan = 1 entonces bullet sino cuotas capital simétricas
        
        	#st.write(f"Tipo plan: {tipo_plan}")
        	    
        for index, row in df.iterrows():
            # Crear vencimiento final
            fechaFinal = pd.to_datetime(row['Fecha de Suscripción'],format ='%Y-%m-%d') + relativedelta(months=int(row['Plazo'])) 
            fechaFinal = fechaFinal.strftime('%Y-%m-%d')
            # Crear el elemento 'obligacion'
            obligacion = ET.SubElement(obligaciones, "{http://www.finagro.com.co/sit}obligacion",
                                        tipoCartera= row['Tipo de Cartera'],
                                        programaCredito = str(cod_programa),
                                        tipoOperacion="1",
                                        tipoMoneda="1",
                                        tipoAgrupamiento="1",
                                        numeroPagare= row['Número de Pagare'],
                                        numeroObligacionIntermediario= row['Número de Pagare'],
                                        fechaSuscripcion=str(row['Fecha de Suscripción'] ),
                                        fechaDesembolso=str(fecha_Desembolso))
                
            # Crear el elemento 'intermediario'
            intermediario = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}intermediario",
                                            oficinaPagare=str(row['Código Oficina']),
                                            oficinaObligacion=str(row['Código Oficina']),
                                            codigo=str(cod_intermediario))
                
            # Crear el elemento 'beneficiarios'
            beneficiarios = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}beneficiarios",
                                            cantidad="1")
                
            # Crear el elemento 'beneficiario'
            beneficiario = ET.SubElement(beneficiarios, "{http://www.finagro.com.co/sit}beneficiario",
                                            correoElectronico=str(row['Email Beneficiario']),
                                            tipoAgrupacion="1",
                                            tipoPersona="1",
                                            tipoProductor=str(row['Tipo de Productor']),
                                            actividadEconomica=str(row['Producto Relacionado']),
                                            cumpleCondicionesProductorAgrupacion="true")
                
            # Crear el elemento 'identificacion' dentro de 'beneficiario'
            identificacion_beneficiario = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}identificacion",
                                                        tipo="2",
                                                        numeroIdentificacion=str(row['Numero de Identificacion']))
                
            # Crear el elemento 'nombre' dentro de 'beneficiario'
            #calcular por espacios
            nombre_beneficiario = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}nombre",
                                                primerNombre=row['Nombre Razón Social'],
                                                segundoNombre="",
                                                primerApellido="",
                                                segundoApellido="",
                                                Razonsocial="")
                
            # Crear el elemento 'nombre' dentro de 'beneficiario'
            direccionCorrespondencia = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}direccionCorrespondencia",
                                            direccion="R|"+str(row['Ubicación Predio']),
                                            municipio=str(row['Ciudad de Inversión']))
                
            # Crear el elemento 'nombre' dentro de 'beneficiario'
            numeroTelefono = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}numeroTelefono",
                                            prefijo="6",
                                            numero=str(row['Teléfono Beneficiario']))
                
            # Crear el elemento 'valorActivos' dentro de 'beneficiario'
            valor_activos = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}valorActivos",
                                            valor=str(row['Monto Activos']),
                                            fechaCorte=str(row['Fecha de Activos']),
                                            tipoDato=str(row['Moneda de Activos']))
                
            # Crear el elemento 'valorIngresos' dentro de 'beneficiario'
            valor_ingresos = ET.SubElement(beneficiario, "{http://www.finagro.com.co/sit}valorIngresos",
                                                valor=str(row['Valor Ingresos']),
                                                fechaCorte=str(row['Fecha de Ingresos']),
                                                tipoDato=str(row['Moneda Ingresos']))
                
            # Crear el elemento 'proyecto'
            proyecto = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}proyecto",
                                    fechaInicialEjecucion=str(fecha_Desembolso),
                                    fechaFinalEjecucion=str(fechaFinal))
            # Se podrían agregar 'incentivo' y 'proyectosFinanciados' dentro de 'proyecto' si fuera necesario
                
            # Crear el elemento 'predios'
            predios = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}predios")
                
            # Crear un elemento 'predio' dentro de 'predios'
            predio = ET.SubElement(predios, "{http://www.finagro.com.co/sit}predio",
                                    tipo="1",
                                    municipio=row['Ciudad de Inversión'],
                                    direccion="R|" +str(row['Ubicación Predio']))
                
            #pendiente crear loop a partir de "Indicativo Fag"
            if row['Indicativo Fag'] == "S":
                # Crear el elemento 'garantiaFAG'
                garantiaFAG = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}garantiaFAG",
                                                    tipoComision =str(row['Tipo Comisión']),
                                                    porcentajeCobertura = str(row['Porcentaje Fag'])
                                                    )
                   
                    
            # Crear el elemento 'destinosCredito'
            destinos_credito = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}destinosCredito")
                
            # pendiente loop a partir de la cantida de destinos, solo hay hasta 4 destinos   
            # Crear un elemento 'destinoCredito' dentro de 'destinosCredito'
            destino_credito = ET.SubElement(destinos_credito, "{http://www.finagro.com.co/sit}destinoCredito",
                                            codigo=str(row['Código Destino 1']),
                                            unidadesAFinanciar=str(row['Unidades Destino 1']),
                                            costoInversion=str(row['Costo Inversión 1']))
                
            # Crear el elemento 'destinoCreditoValorAFinanciar' dentro de 'destinoCredito'
            destino_credito_valor = ET.SubElement(destino_credito, "{http://www.finagro.com.co/sit}destinoCreditoValorAFinanciar")
            valor_a_financiar = ET.SubElement(destino_credito_valor, "{http://www.finagro.com.co/sit}valorAFinanciar", {"xmlns": ""})
            valor_a_financiar.text=str(row['Valor a Financiar 1'])
                    
            if not row['Código Destino 2']!= row['Código Destino 2']: 
                # Crear un elemento 'destinoCredito' dentro de 'destinosCredito'
                destino_credito = ET.SubElement(destinos_credito, "{http://www.finagro.com.co/sit}destinoCredito",
                                                codigo=str(row['Código Destino 2']),
                                                unidadesAFinanciar=str(row['Unidades Destino 2']),
                                                costoInversion=str(row['Costo Inversión 2']))
                
                # Crear el elemento 'destinoCreditoValorAFinanciar' dentro de 'destinoCredito'
                destino_credito_valor = ET.SubElement(destino_credito, "{http://www.finagro.com.co/sit}destinoCreditoValorAFinanciar")
                valor_a_financiar = ET.SubElement(destino_credito_valor, "{http://www.finagro.com.co/sit}valorAFinanciar", {"xmlns": ""})
                valor_a_financiar.text=str(row['Valor a Financiar 2'])
                    
            if not row['Código Destino 3']!= row['Código Destino 3']: 
                # Crear un elemento 'destinoCredito' dentro de 'destinosCredito'
                destino_credito = ET.SubElement(destinos_credito, "{http://www.finagro.com.co/sit}destinoCredito",
                                                codigo=str(row['Código Destino 3']),
                                                unidadesAFinanciar=str(row['Unidades Destino 3']),
                                                costoInversion=str(row['Costo Inversión 3']))
                
                # Crear el elemento 'destinoCreditoValorAFinanciar' dentro de 'destinoCredito'
                destino_credito_valor = ET.SubElement(destino_credito, "{http://www.finagro.com.co/sit}destinoCreditoValorAFinanciar")
                valor_a_financiar = ET.SubElement(destino_credito_valor, "{http://www.finagro.com.co/sit}valorAFinanciar", {"xmlns": ""})
                valor_a_financiar.text=str(row['Valor a Financiar 3'])
                    
            if not row['Código Destino 4']!= row['Código Destino 3']: 
                # Crear un elemento 'destinoCredito' dentro de 'destinosCredito'
                destino_credito = ET.SubElement(destinos_credito, "{http://www.finagro.com.co/sit}destinoCredito",
                                                codigo=str(row['Código Destino 4']),
                                                unidadesAFinanciar=str(row['Unidades Destino 4']),
                                                costoInversion=str(row['Costo Inversión 4']))
                
                # Crear el elemento 'destinoCreditoValorAFinanciar' dentro de 'destinoCredito'
                destino_credito_valor = ET.SubElement(destino_credito, "{http://www.finagro.com.co/sit}destinoCreditoValorAFinanciar")
                valor_a_financiar = ET.SubElement(destino_credito_valor, "{http://www.finagro.com.co/sit}valorAFinanciar", {"xmlns": ""})
                valor_a_financiar.text=str(row['Valor a Financiar 4'])
                    
            # Crear el elemento 'financiacion'
            financiacion = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}financiacion",
                                            fechaVencimientoFinal=str(fechaFinal),
                                            plazoCredito=str(row['Plazo']),
                                            valorTotalCredito=str(row['Capital Total']),
                                            porcentaje="100",
                                            valorObligacion=str(row['Capital Total']))
                
            # Datos para las cuotas
            datos_cuotas = []
            cantidad_cuotas= int(int(row['Plazo'])/int(row['Tipo Plan de Pagos']))
            cuota_capital = int(int(row['Capital Total'])/cantidad_cuotas)
                    
            ult_cuota_capital = cuota_capital if int(row['Capital Total']) - (cantidad_cuotas*cuota_capital) == 0 else str(int(row['Capital Total']) - (Decimal(cantidad_cuotas-1)*Decimal(cuota_capital)))
            fHasta = pd.to_datetime(row['Fecha de Suscripción'],format ='%Y-%m-%d')
            for i in range(cantidad_cuotas-1):
                meses = int(row['Tipo Plan de Pagos'])
                fHasta = fHasta + relativedelta(months=meses)
                cuotas = {
                                "registro": str(i+1),
                                "fechaAplicacionHasta":str(date(int(fHasta.strftime('%Y')),int(fHasta.strftime('%m')),10)),
                                "conceptoRegistroCuota": "I" if tipo_plan == 1 else "K",
                                "periodicidadIntereses": "PE",
                                "periodicidadCapital": "" if tipo_plan == 1 else "PE",
                                "tasaBaseBeneficiario": "5",
                                "margenTasaBeneficiario": str(row['Puntos IBR']),
                                "valorCuotaCapital": "0" if tipo_plan == 1 else str(cuota_capital),
                                "porcentajeCapitalizacionIntereses": "0.0",
                                "margenTasaRedescuento": "0"
                                }
                            
                datos_cuotas.append(cuotas)
            cuotas = {
                        "registro": str(cantidad_cuotas),
                        "fechaAplicacionHasta": str(fechaFinal),
                        "conceptoRegistroCuota": "K",
                        "periodicidadIntereses": "PE",
                        "periodicidadCapital": "PE",
                        "tasaBaseBeneficiario": "5",
                        "margenTasaBeneficiario":  str(row['Puntos IBR']),
                        "valorCuotaCapital": str(row['Capital Total']) if tipo_plan == 1 else str(ult_cuota_capital),
                        "porcentajeCapitalizacionIntereses": "0.0",
                        "margenTasaRedescuento": "0"
                    }
                         
            datos_cuotas.append(cuotas)
            # Crear el elemento 'planPagos'
            plan_pagos = ET.SubElement(obligacion, "{http://www.finagro.com.co/sit}planPagos")
                
            # Iterar sobre los datos de las cuotas y crear un elemento 'registroCuota' para cada uno
            for dato_cuota in datos_cuotas:
                registro_cuota = ET.SubElement(plan_pagos, "{http://www.finagro.com.co/sit}registroCuota",
                                                registro=str(dato_cuota["registro"]),
                                                fechaAplicacionHasta=str(dato_cuota["fechaAplicacionHasta"]),
                                                conceptoRegistroCuota=dato_cuota["conceptoRegistroCuota"],
                                                periodicidadIntereses=dato_cuota["periodicidadIntereses"],
                                                periodicidadCapital=dato_cuota["periodicidadCapital"],
                                                tasaBaseBeneficiario=dato_cuota["tasaBaseBeneficiario"],
                                                margenTasaBeneficiario=dato_cuota["margenTasaBeneficiario"],
                                                valorCuotaCapital=dato_cuota.get("valorCuotaCapital"),  # Usamos .get() por si es opcional
                                                porcentajeCapitalizacionIntereses=dato_cuota.get("porcentajeCapitalizacionIntereses"),
                                                margenTasaRedescuento=dato_cuota.get("margenTasaRedescuento"),
                                                )
                    
        # Crear el árbol XML
        
        def sanitize_element(element):
            if element.text is not None and not isinstance(element.text, str):
                element.text = str(element.text)
            for key, value in element.attrib.items():
                if not isinstance(value, str):
                    element.attrib[key] = str(value)
            for child in element:
                sanitize_element(child)
        
        sanitize_element(obligaciones)


        tree = ET.ElementTree(obligaciones)
        ET.indent(tree, space="  ", level=0)
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xml") as tmp:
                    tree.write(tmp.name, encoding="UTF-8", xml_declaration=True)
                    st.success("✅ XML de obligaciones nuevas generado exitosamente.")
                    with open(tmp.name, "rb") as f:
                        st.download_button("📥 Descargar XML de Obligaciones nuevas", f, file_name="Obligaciones.xml", mime="application/xml")
