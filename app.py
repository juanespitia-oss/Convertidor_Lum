import streamlit as st
import pandas as pd
import io
import re

# Configuración estética de la aplicación
st.set_page_config(page_title="LUM Logistic - Convertidor de Gastos", page_icon="📊", layout="wide")

# --- ENCABEZADO ---
st.title("📊 Convertidor Oficial de Legalizaciones de Gastos")
st.subheader("LUM Logistic - Aliados de Última Milla")
st.write("Herramienta corporativa indexada de forma definitiva con el maestro de productos y categorías ODOO.")
st.markdown("---")

# --- CONFIGURACIÓN DINÁMICA DE LA ETIQUETA ---
st.sidebar.header("📋 Parámetros de la Etiqueta ODOO")
param_tipo_leg = st.sidebar.text_input("1. Tipo de legalización:", value="Leg Viaticos alimentacion")
param_nombre = st.sidebar.text_input("2. Nombre de quien legaliza:", value="Yoselin Gomez")
param_ciudad = st.sidebar.text_input("3. Ciudad destino (Abv.):", value="Bog")
param_fechas = st.sidebar.text_input("4. Rango de fechas del viaje:", value="12 may a 16 mayo 26")

def normalizar_texto(texto):
    if not texto:
        return ""
    texto = str(texto).lower().replace('\xa0', ' ').replace('\xad', ' ')
    texto = re.sub(r'[áàäâ]', 'a', texto)
    texto = re.sub(r'[éèëê]', 'e', texto)
    texto = re.sub(r'[íìïî]', 'i', texto)
    texto = re.sub(r'[óòöô]', 'o', texto)
    texto = re.sub(r'[úùüû]', 'u', texto)
    texto = re.sub(r'ñ', 'n', texto)
    texto = re.sub(r'[^a-z0-9]', ' ', texto)
    texto = re.sub(r'\s+', ' ', texto)
    return texto.strip()

# --- MATRICES DE HOMOLOGACIÓN OFICIALES (IDs ODOO) ---
MAPEO_CENTROS = {
    "administracion barranquilla lum": '{"110":100.0}',
    "el poblado": '{"111":100.0}',
    "el prado": '{"112":100.0}',
    "las colinas": '{"113":100.0}',
    "villa de andalucia": '{"114":100.0}',
    "administracion bogota lum": '{"79":100.0}',
    "americas": '{"90":100.0}',
    "barrancas": '{"98":100.0}',
    "cedritos": '{"99":100.0}',
    "chapinero": '{"100":100.0}',
    "chico": '{"101":100.0}',
    "chilacos": '{"102":100.0}',
    "colina": '{"103":100.0}',
    "comuneros": '{"104":100.0}',
    "country": '{"80":100.0}',
    "el lago": '{"81":100.0}',
    "engativa occidental": '{"82":100.0}',
    "esperanza": '{"83":100.0}',
    "granada norte": '{"84":100.0}',
    "iserra 100": '{"85":100.0}',
    "la cabana": '{"86":100.0}',
    "modelia": '{"87":100.0}',
    "niza": '{"88":100.0}',
    "palermo": '{"89":100.0}',
    "parque bavaria": '{"91":100.0}',
    "quinta camacho": '{"92":100.0}',
    "san patricio": '{"93":100.0}',
    "suba turingia": '{"94":100.0}',
    "titan": '{"95":100.0}',
    "usaquen": '{"96":100.0}',
    "verbenal": '{"97":100.0}',
    "prado veraniego": '{"190":100.0}',
    "administracion bucaramanga lum": '{"129":100.0}',
    "cabecera": '{"130":100.0}',
    "canaveral": '{"131":100.0}',
    "administracion cali lum": '{"124":100.0}',
    "c jardin": '{"125":100.0}',
    "gran limonar": '{"126":100.0}',
    "parque del perro": '{"127":100.0}',
    "san vicente": '{"128":100.0}',
    "administracion cartagena lum": '{"115":100.0}',
    "bocagrande": '{"116":100.0}',
    "la castellana": '{"117":100.0}',
    "manga": '{"118":100.0}',
    "mar del norte": '{"119":100.0}',
    "administracion eje cafetero lum": '{"105":100.0}',
    "armenia": '{"106":100.0}',
    "ibague": '{"107":100.0}',
    "pereira": '{"108":100.0}',
    "villavicencio": '{"109":100.0}',
    "administracion medellin lum": '{"68":100.0}',
    "envigado alto": '{"71":100.0}',
    "envigado bajo": '{"72":100.0}',
    "la estrella": '{"73":100.0}',
    "la frontera": '{"74":100.0}',
    "la mota": '{"75":100.0}',
    "la playa boston": '{"76":100.0}',
    "laureles": '{"77":100.0}',
    "manila": '{"78":100.0}',
    "sabaneta": '{"69":100.0}',
    "tesoro": '{"70":100.0}',
    "administracion monteria lum": '{"120":100.0}',
    "alameda": '{"121":100.0}',
    "administracion santa marta lum": '{"122":100.0}',
    "bavaria sm": '{"123":100.0}',
    "rodadero": '{"155":100.0}',
    "valledupar": '{"156":100.0}',
    "cucuta": '{"188":100.0}',
    "manizales": '{"191":100.0}',
    
    # Excluidos
    "compostela": "No cargar", 
    "hipotecho": "No cargar", 
    "santa coloma": "No cargar",
    "engativa services": "No cargar", 
    "kennedy": "No cargar", 
    "suba services": "No cargar",
    "itagui": "No cargar", 
    "la america": "No cargar", 
    "manrique": "No cargar"
}

# --- INTERFAZ DE CARGA ---
archivo_subido = st.file_uploader("📥 Cargar archivo de gastos LUM (Excel)", type=["xlsx", "xlsm"])

if archivo_subido is not None:
    try:
        df_origen = pd.read_excel(archivo_subido, sheet_name=0, skiprows=7, header=None)
        nuevas_filas = []
        lineas_sin_mapeo = []
        ultimo_nro_factura = None
        ultimo_concepto = None
        
        for idx, row in df_origen.iterrows():
            if row.dropna().empty:
                continue
                
            raw_fecha = row[0]
            if isinstance(raw_fecha, pd.Timestamp):
                fecha_factura_str = raw_fecha.strftime('%Y-%m-%d')
            else:
                fecha_factura_str = str(raw_fecha).split(" ")[0].strip() if not pd.isna(raw_fecha) and str(raw_fecha).strip() != "" else None
                
            concepto_gasto = str(row[1]).strip() if not pd.isna(row[1]) else ""
            tipo_compra = str(row[2]).strip().lower() if not pd.isna(row[2]) else ""
            tercero = str(row[3]).strip() if not pd.isna(row[3]) else ""
            centro_costo = str(row[4]).strip() if not pd.isna(row[4]) else ""
            linea_presupuesto = str(row[5]).strip() if not pd.isna(row[5]) else ""
            valor_antes_iva = row[6] if not pd.isna(row[6]) else 0
            
            try:
                iva = float(row[7]) if not pd.isna(row[7]) else 0.0
            except ValueError:
                iva = 0.0
                
            numero_factura = str(row[9]).strip() if not pd.isna(row[9]) else ""
            tipo = str(row[10]).strip().lower() if not pd.isna(row[10]) else ""
            
            etiqueta_final = f"{param_tipo_leg} {concepto_gasto}, {param_nombre} {param_ciudad} {param_fechas}"
            
            cc_normalizado = normalizar_texto(centro_costo)
            dist_analitica = MAPEO_CENTROS.get(cc_normalizado, None)
            
            is_no_cargar = dist_analitica == "No cargar"
            if is_no_cargar:
                dist_analitica = None
                
            if "compras" in tipo_compra and iva > 0:
                impuesto = "IVA DESCONTABLE POR COMPRAS 19%"
            elif "servicios" in tipo_compra and iva > 0:
                impuesto = "IVA DESCONTABLE POR SERVICIOS 19%"
            else:
                impuesto = None

            txt_buscar = normalizar_texto(linea_presupuesto)
            cc_buscar = cc_normalizado
            cuenta = None
            producto = None
            
            # Identificación Jerárquica
            if "proyecto" in txt_buscar or "74" in cc_buscar:
                es_proyecto = True
                es_costo = False
            elif "adm" in txt_buscar:
                es_proyecto = False
                es_costo = False
            elif "cost" in txt_buscar:
                es_proyecto = False
                es_costo = True
            else:
                es_proyecto = False
                es_costo = "73" in cc_buscar or "72" in cc_buscar or any(suc in cc_buscar for suc in ["jardin", "pereira", "armenia", "bucaramanga", "poblado", "medellin", "cali", "cartagena", "bogota"])

            if es_proyecto:
                if "litog" in txt_buscar or "volant" in txt_buscar or "pendon" in txt_buscar or "rappi" in txt_buscar or "insumos" in txt_buscar:
                    producto, cuenta = "74053001 INSUMOS (LITOGRAFIA, VOLANTES, PENDONES Y ACTIVIDADES COMERCIALES RAPPI)", "74053001"
                elif "util" in txt_buscar or "papel" in txt_buscar or "oficn" in txt_buscar or "oficin" in txt_buscar or "fotocop" in txt_buscar:
                    producto, cuenta = "74052001 UTILES, PAPELERIA Y FOTOCOPIAS (OFICINA Y BODEGA)", "74052001"
                elif "licenc" in txt_buscar or "regist" in txt_buscar: producto, cuenta = "(PROYECTOS) REGISTRO LICENCIAS", "74401502"
                elif "camara" in txt_buscar: producto, cuenta = "(PROYECTOS) CAMARA DE COMERCIO", "74401001"
                elif "construc" in txt_buscar or "edific" in txt_buscar: producto, cuenta = "(PROYECTOS) CONSTRUCCIONES Y EDIFICACIONES", "74451003"
                elif "comput" in txt_buscar: producto, cuenta = "(PROYECTOS) EQUIPO DE COMPUTACION Y COMUNICACION MENOR", "74452502"
                elif "oficina" in txt_buscar or "muebles" in txt_buscar: producto, cuenta = "(PROYECTOS) EQUIPO DE OFICINA (MUEBLES Y ENSERES) MENOR", "74453002"
                elif "instalac" in txt_buscar or "adecuac" in txt_buscar: producto, cuenta = "(PROYECTOS) INSTALACIONES, ADECUACIONES, LOCATIVAS", "74451004"
                elif "maquinaria" in txt_buscar: producto, cuenta = "(PROYECTOS) MANTENIMIENTO DE MAQUINARIA (NEVERAS)", "74451501"
                elif "edificio" in txt_buscar: producto, cuenta = "(PROYECTOS) MANTENIMIENTO EDIFICIO (FACHADAS PISOS PAREDES TECHOS JA", "74451001"
                elif "redes" in txt_buscar or "inter" in txt_buscar: producto, cuenta = "(PROYECTOS) MANTENIMIENTO REDES (ACUEDUCTO GAS NATURAL RED INTERNET)", "74451002"
                elif "menor" in txt_buscar: producto, cuenta = "(PROYECTOS) MAQUINARIA Y EQUIPO MENOR", "74451502"
                elif "aseo" in txt_buscar or "vigil" in txt_buscar: producto, cuenta = "(PROYECTOS) SERVICIOS ASEO Y VIGILANCIA", "74350501"
                elif "transp" in txt_buscar or "carga" in txt_buscar: producto, cuenta = "(PROYECTOS) TRANSPORTE CARGA (MERCANCIA ACARREOS)", "74355001"
                elif "viatic" in txt_buscar or "pasaj" in txt_buscar: producto, cuenta = "(PROYECTOS) VIATICOS PASAJES", "74551501"
            else:
                if "contin" in txt_buscar: 
                    producto, cuenta = ("Adm-Servicios Planes de internet contingencias", "51353502") if not es_costo else ("Cost-Servicios Planes de internet contingencias", "73353502")
                elif "inter" in txt_buscar: 
                    producto, cuenta = ("Adm-Servicio de Internet (Claro y Tigo)", "51353503") if not es_costo else ("Cost-Servicio de Internet (Claro y Tigo)", "73353503")
                elif "plan" in txt_buscar or "movil" in txt_buscar or "tigo" in txt_buscar:
                    producto, cuenta = ("Adm-Servicio de Planes móviles (Tigo) asignados", "51353501") if not es_costo else ("Cost-Servicio de Planes móviles (Tigo) asignados", "73353501")
                elif "tecnol" in txt_buscar or "informac" in txt_buscar:
                    producto, cuenta = ("Adm-Servicios , Tecnologia de Ia Información", "51353504") if not es_costo else ("Cost-Servicios , Tecnologia de Ia Información", "73353504")
                elif "acued" in txt_buscar or "agua" in txt_buscar:
                    producto, cuenta = ("Adm-Servicios Publicos Acueducto", "51352501") if not es_costo else ("Cost-Servicios Publicos Acueducto", "73352501")
                elif "energ" in txt_buscar or "electr" in txt_buscar:
                    producto, cuenta = ("Adm-Servicios públicos Energía", "51353001") if not es_costo else ("Cost-Servicios públicos Energía", "73353001")
                elif "gas" in txt_buscar:
                    producto, cuenta = ("Cost-Servicios públicos Gas", "73353101")
                elif "softw" in txt_buscar and "activ" in txt_buscar:
                    producto, cuenta = ("Adm-Servicio de Software Activos fijos", "51204504") if not es_costo else ("Cost-Servicio de Software Activos fijos", "73204504")
                elif "softw" in txt_buscar and "turn" in txt_buscar:
                    producto, cuenta = ("Adm-Servicios Software Turnos", "51204503") if not es_costo else ("Cost-Servicios Software Turnos", "73204503")
                elif "softw" in txt_buscar and "contab" in txt_buscar:
                    producto, cuenta = ("Adm-Servicios Software contable", "51204502") if not es_costo else ("Cost-Servicios Software contable", "73204502")
                elif "softw" in txt_buscar and "nomin" in txt_buscar:
                    producto, cuenta = ("Adm-Servicios Software de nómina", "51204501") if not es_costo else ("Cost-Servicios Software de nómina", "73204501")
                elif "tempor" in txt_buscar or "indirect" in txt_buscar:
                    producto, cuenta = ("Adm-Nomina Indirecta subcontratada (Personal temporal) Nomina Indirecta subcontratada (Personal temporal)", "51351001") if not es_costo else ("Cost-Nomina Indirecta subcontratada (Personal temporal)", "73351001")
                elif "dotac" in txt_buscar or "vestuar" in txt_buscar or "pantalon" in txt_buscar:
                    producto, cuenta = ("Adm-Dotacion Legal 4 meses (Pantalon, camisas, zapatos, chaquetas, overol antifluido tela, otros)", "51057001") if not es_costo else ("Cost-Dotacion Legal 4 meses (Pantalon, camisas, zapatos, chaquetas, overol antifluido tela, otros)", "72057001")
                elif "ocupac" in txt_buscar or "siso" in txt_buscar or "poligr" in txt_buscar:
                    producto, cuenta = ("Adm-Examenes ocupacionales (SISO, poligrafo)", "51056901") if not es_costo else ("Cost-Examenes ocupacionales (SISO, poligrafo)", "72056901")
                elif "exame" in txt_buscar or "ingres" in txt_buscar or "egres" in txt_buscar or "medic" in txt_buscar:
                    producto, cuenta = ("Adm-Examenes medicos (Ingreso /egreso)", "51056902") if not es_costo else ("Cost-Examenes medicos (Ingreso /egreso)", "72056902")
                elif "capacit" in txt_buscar or "higiene" in txt_buscar or "seguridad" in txt_buscar:
                    producto, cuenta = ("Adm-Capacitaciones (Empresa Privada) : Systems de emergencia Higiene, seguridad industrial, manipulacion de alimentos, basico de alturas, quimicos, otros)", "51056301") if not es_costo else ("Cost-Capacitaciones (Empresa Privada) : Sistemas de emergencia Higiene, seguridad industrial, manipulacion de alimentos, basico de alturas, quimicos, otros)", "72056301")
                elif "licenc" in txt_buscar or "tramit" in txt_buscar:
                    if "tramit" in txt_buscar: producto, cuenta = ("Cost-Tramites y licencias", "73401505")
                    else: producto, cuenta = ("Adm-Registro licencias", "51401502") if not es_costo else ("Cost-Registro licencias", "73401502")
                elif "derech" in txt_buscar or "contra" in txt_buscar:
                    producto, cuenta = ("Adm-Derechos de contrato", "51401506") if not es_costo else ("Cost-Derechos de Contrato", "73401506")
                elif "respons" in txt_buscar or "civil" in txt_buscar:
                    producto, cuenta = ("Adm-Responsabilidad civil excontractual", "51301005")
                elif "activ" in txt_buscar and "fijo" in txt_buscar:
                    producto, cuenta = ("Adm-Activos fijos", "51301006")
                elif "contab" in txt_buscar or "merma" in txt_buscar or "sgsst" in txt_buscar:
                    producto, cuenta = ("Adm-Contabilidad/ Mermas/SGSST", "51109501") if not es_costo else ("Cost-Contabilidad/ Mermas/SGSST", "73109501")
                elif "calid" in txt_buscar:
                    producto, cuenta = ("Adm-Calidad", "51109502") if not es_costo else ("Cost-Calidad", "73109502")
                elif "indemn" in txt_buscar or "indemin" in txt_buscar:
                    producto, cuenta = ("Cost-Indemnizaciones", "73957501")
                elif "cumple" in txt_buscar or "bono" in txt_buscar: 
                    producto, cuenta = ("Adm-Bonos Cumpleaños", "51970506")
                elif "cumplim" in txt_buscar:
                    producto, cuenta = ("Adm-Bonos de Cumplimiento", "51970507")
                elif "cine" in txt_buscar:
                    producto, cuenta = ("Adm-Cine", "51970503")
                elif "condol" in txt_buscar:
                    producto, cuenta = ("Adm-Condolencias", "51970505")
                elif "event" in txt_buscar or "recre" in txt_buscar:
                    producto, cuenta = ("Adm-Eventos y Recreacion", "51970501")
                elif "plant" in txt_buscar or "redes" in txt_buscar:
                    producto, cuenta = ("Adm-Plantas y redes Menor", "51451002")
                elif "util" in txt_buscar or "papel" in txt_buscar:
                    producto, cuenta = ("Adm-Insumos: Útiles, papelería y fotocopias (oficina y Bodega) tableros acliricos", "51953001") if not es_costo else ("Cost-Insumos: Útiles, papelería y fotocopias", "73953001")
                elif "aseo" in txt_buscar or "vigil" in txt_buscar:
                    producto, cuenta = ("Adm-Servicios Aseo y Vigilancia", "51350501") if not es_costo else ("Cost-Servicios Aseo y Vigilancia", "73350501")
                elif "element" in txt_buscar or "cafet" in txt_buscar:
                    producto, cuenta = ("Adm-Insumos: Elementos de aseo y cafetería (oficina y Bodega)", "51952501") if not es_costo else ("Cost- Insumos de aseo y cafeteria", "73952501") # AJUSTADO A CUENTA 73952501
                elif "envas" in txt_buscar or "empaq" in txt_buscar:
                    producto, cuenta = ("Adm-Insumos: Envases y empaques (bolsas, cajas,contendores, empaques, canecas)", "51954001")
                elif "estib" in txt_buscar:
                    producto, cuenta = ("Adm-Insumos: Estibas, Butacos, escaleras, bandejas, lockers y compras ocacionales", "51953003")
                elif "litog" in txt_buscar or "volant" in txt_buscar or "pendon" in txt_buscar:
                    producto, cuenta = ("Adm-Insumos: (Litografía, Volantes, Pendones y actividades comerciales Rappi)", "51959502")
                elif "pasaj" in txt_buscar and "terr" in txt_buscar:
                    producto, cuenta = ("Adm-viaticos Pasajes terrestres", "51551501") if not es_costo else ("Cost-viaticos Pasajes terrestres", "73551501")
                elif "pasaj" in txt_buscar and "aer" in txt_buscar:
                    producto, cuenta = ("Adm-viaticos Pasajes aéreos", "51551502") if not es_costo else ("Cost-viaticos Pasajes aéreos", "73551502")
                elif "aloj" in txt_buscar or "manut" in txt_buscar:
                    producto, cuenta = ("Adm-viaticos Alojamiento y manutención", "51550501") if not es_costo else ("Cost-viaticos Alojamiento y manutención", "73550501")
                elif "casin" in txt_buscar or "restaur" in txt_buscar or "refrig" in txt_buscar:
                    producto, cuenta = ("Adm-Casino y restaurante (Refrigerios)", "51956001") if not es_costo else ("Cost-Casino y restaurante (Refrigerios)", "73956001")
                elif "manten" in txt_buscar and "equip" in txt_buscar:
                    producto, cuenta = ("Adm-Mantenimiento de equipos (comunicación, computacion, camaras, extintores, emergencia) Mantenimiento redes (acueducto, Gas Natural, red internet)", "51452501") if not es_costo else ("Cost-Mantenimiento de equipos (comunicación, computacion, camaras, extintores, emergencia) Mantenimiento redes (acueducto, Gas Natural, red internet)", "73451002")
                elif "manten" in txt_buscar and "edific" in txt_buscar:
                    producto, cuenta = ("Adm-Mantenimiento edificio (Fachadas, pisos, paredes, techos, jardineria)", "51451001") if not es_costo else ("Cost- Mantenimiento de edificios", "73451001") # AJUSTADO A CUENTA 73451001
                elif "manten" in txt_buscar and "muebl" in txt_buscar:
                    producto, cuenta = ("Adm-Mantenimiento muebles office", "51453001") if not es_costo else ("Cost-Mantenimiento muebles oficina", "73453001")
                elif "manten" in txt_buscar and "maquin" in txt_buscar:
                    producto, cuenta = ("Adm-Mantenimiento de Maquinaria (Neveras)", "51451501") if not es_costo else ("Cost-Mantenimiento de Maquinaria (Neveras)", "73451501")
                elif "instal" in txt_buscar or "adecu" in txt_buscar or "locat" in txt_buscar:
                    producto, cuenta = ("Adm-Instalaciones, adecuaciones, locativas", "51501502") if not es_costo else ("Cost-Instalaciones, adecuaciones, locativas", "73451004")
                elif "arriend" in txt_buscar and "bodeg" in txt_buscar:
                    producto, cuenta = ("Adm-Arriendo Bodegas /Arrendamiento de equipos (PC, impresoras, Router, plantas eléctricas, otros)", "51201002") if not es_costo else ("Cost-Arriendo Bodegas /Arrendamiento de equipos (PC, impresoras, Router, plantas eléctricas, otros)", "73202501")
                elif "arriend" in txt_buscar and "instal" in txt_buscar:
                    producto, cuenta = ("Cost-Arriendo Instalaciones", "73201001")
                elif "fumig" in txt_buscar or "desinf" in txt_buscar:
                    producto, cuenta = ("Adm-Servicios Fumigaciones, Desinfección, roedores", "51501501") if not es_costo else ("Cost-Servicios Fumigaciones, Desinfección, roedores", "73501501")
                elif "asesor" in txt_buscar:
                    producto, cuenta = ("Adm-Asesoria Jurídica Penal y contractual", "51102504") if not es_costo else ("Cost-Asesoria Jurídica Penal y contractual", "73102504")
                elif "taxi" in txt_buscar or "bus" in txt_buscar:
                    producto, cuenta = ("Adm-Taxis y buses (Del personal)", "51954502") if not es_costo else ("Cost-Taxis y buses (Del personal)", "73954502")
                elif "transp" in txt_buscar or "carga" in txt_buscar:
                    producto, cuenta = ("Adm-Transporte carga (Mercancia acarreos)", "51355001") if not es_costo else ("Cost-Transporte carga (Mercancia acarreos)", "73355001")
                elif "combus" in txt_buscar or "lubric" in txt_buscar:
                    producto, cuenta = ("Adm-Combustibles y lubricantes", "51953501") if not es_costo else ("Cost-Combustibles y lubricantes", "73953501")
                elif "bienest" in txt_buscar:
                    producto, cuenta = ("Adm-Bienestar", "51970508")
                elif "notar" in txt_buscar:
                    producto, cuenta = ("Adm-Notaria", "51401501") if not es_costo else ("Cost-Notaria", "73401501")
                elif "bomb" in txt_buscar:
                    producto, cuenta = ("Adm-Registro bomberos", "51401503") if not es_costo else ("Cost-Registro bomberos", "73401503")
                elif "camar" in txt_buscar:
                    producto, cuenta = ("Adm-Camara de comercio", "51401001")
                elif "comput" in txt_buscar:
                    producto, cuenta = ("Adm-Equipo de Computacion y Comunicación Menor", "51452502") if not es_costo else ("Cost-Equipo de Computacion y Comunicación Menor", "73452502")
                elif "oficin" in txt_buscar:
                    producto, cuenta = ("Adm-Equipo de Oficina (Muebles y Enseres) Menor", "51453002") if not es_costo else ("Cost-Equipo de Oficina (Muebles y Enseres) Menor", "73453002")
                elif "maquin" in txt_buscar:
                    producto, cuenta = ("Adm-Maquinaria y equipo Menor", "51451502") if not es_costo else ("Cost-Maquinaria y equipo Menor", "73451502")

            if not cuenta:
                producto = f"ALERTA: Sin mapear ({linea_presupuesto})"
                cuenta = "REVISAR CUENTA"
                if linea_presupuesto not in lineas_sin_mapeo:
                    lineas_sin_mapeo.append(linea_presupuesto)

            es_misma_factura = False
            if numero_factura != "" and numero_factura != "-" and numero_factura == ultimo_nro_factura:
                es_misma_factura = True
            elif (numero_factura == "" or numero_factura == "-") and concepto_gasto == ultimo_concepto and ultimo_concepto != "":
                es_misma_factura = True

            if es_misma_factura:
                f_factura_val, f_vencimiento_val, contacto_val = None, None, None
                terminos_pago_val, diario_val, cufe_val, referencia_val = None, None, None, None
            else:
                f_factura_val = fecha_factura_str if fecha_factura_str else "Revisar Fecha"
                f_vencimiento_val = fecha_factura_str if fecha_factura_str else "Revisar Fecha"
                contacto_val = tercero if tercero != "" else "TERCERO REQUERIDO"
                terminos_pago_val = "Pago inmediato"
                diario_val = "Documento Soporte Electrónico" if "soporte" in tipo else "Factura Caja Menor - Tarjetas Credito - Legalizaciones"
                cufe_val = "CONTRAPARTIDA MANUAL" if is_no_cargar else "PONER MANUAL"
                referencia_val = numero_factura if numero_factura != "" and numero_factura != "-" else "S/N"

            ultimo_nro_factura = numero_factura
            ultimo_concepto = concepto_gasto

            fila_formateada = {
                "Fecha de factura": f_factura_val,
                "Fecha de vencimiento": f_vencimiento_val,
                "Referencia de pago": None,
                "Contacto": contacto_val,
                "Términos de pago": terminos_pago_val,
                "Diario": diario_val,
                "Referencia": referencia_val,
                "CUFE Factura Proveedor": cufe_val,
                "Líneas de factura/Distribución analítica": dist_analitica,
                "Líneas de factura/Cantidad": 1,
                "Líneas de factura/Etiqueta": etiqueta_final,
                "Líneas de factura/Producto": producto,
                "Líneas de factura/Cuenta": cuenta,
                "Líneas de factura/Precio unitario": valor_antes_iva,
                "Líneas de factura/Impuestos": impuesto
            }
            nuevas_filas.append(fila_formateada)
            
        df_nuevo = pd.DataFrame(nuevas_filas)
        
        if lineas_sin_mapeo:
            st.warning(f"⚠️ ¡Atención! Se detectaron {len(lineas_sin_mapeo)} líneas que requirieron revisión:")
            for item in lineas_sin_mapeo:
                st.markdown(f"* ❌ **Línea sin reconocer**: `{item}`")
        else:
            st.success("🎯 ¡Perfecto! Se barrió por completo el formato de Excel. El 100% de los registros están indexados.")

        st.subheader("📋 Vista previa de la plantilla homologada:")
        st.dataframe(df_nuevo)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_nuevo.to_excel(writer, index=False, sheet_name='Plantilla Importación')
            
        xlsx_data = output.getvalue()
        
        st.download_button(
            label="📥 Descargar Formato de Legalización LUM Oficial",
            data=xlsx_data,
            file_name="Plantilla_LUM_Formato_ODOO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"❌ Ocurrió un inconveniente al procesar las columnas: {e}")
