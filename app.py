import pandas as pd
import streamlit as st
import dropbox
import io
import time
import json
import hashlib
from datetime import datetime

# Configuración inicial de la página
st.set_page_config(page_title="Sistema de Grabaciones Claro", layout="wide", page_icon="📶")

# Clave de API de Dropbox desde las credenciales secretas
DROPBOX_TOKEN = st.secrets.get("DROPBOX_TOKEN", "")

# RUTAS EXACTAS EN TU DROPBOX
RUTA_DROPBOX_TOTAL = "/CLARO/VENTAS - CLARO/SISTEMA_VENTAS_CLARO/total_grabaciones_claro.xlsx"
RUTA_DROPBOX_CONFIG = "/CLARO/VENTAS - CLARO/SISTEMA_VENTAS_CLARO/configuracion de tabulacion v2.xlsx"
RUTA_DROPBOX_USUARIOS = "/CLARO/VENTAS - CLARO/SISTEMA_VENTAS_CLARO/usuarios.json"

dbx = dropbox.Dropbox(DROPBOX_TOKEN)

# ---------------------------------------------------------
# FUNCIONES AUXILIARES DE DROPBOX Y AUDITORÍA
# ---------------------------------------------------------
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def descargar_archivo_dropbox(ruta):
    try:
        _, res = dbx.files_download(ruta)
        return res.content
    except Exception as e:
        return None

def subir_archivo_dropbox(contenido, ruta):
    try:
        dbx.files_upload(contenido, ruta, mode=dropbox.files.WriteMode.overwrite)
        return True
    except Exception as e:
        st.error(f"Error al conectar con Dropbox: {e}")
        return False

def cargar_usuarios():
    data = descargar_archivo_dropbox(RUTA_DROPBOX_USUARIOS)
    if data:
        return json.loads(data.decode('utf-8'))
    return {
        "admin": {
            "password": hash_password("admin123"),
            "nombre": "Administrador Principal",
            "rol": "admin"
        }
    }

def guardar_usuarios(usuarios):
    contenido = json.dumps(usuarios, indent=4).encode('utf-8')
    return subir_archivo_dropbox(contenido, RUTA_DROPBOX_USUARIOS)

def descargar_excel_base():
    content = descargar_archivo_dropbox(RUTA_DROPBOX_TOTAL)
    if content:
        try:
            # Lee la primera hoja disponible sin forzar el nombre 'Base'
            xls = pd.ExcelFile(io.BytesIO(content))
            nombre_hoja = xls.sheet_names[0]
            df = pd.read_excel(xls, sheet_name=nombre_hoja, dtype=str).fillna('')
            
            # Limpiar espacios invisibles en los nombres de las columnas
            df.columns = [str(col).strip() for col in df.columns]
            
            # Columnas requeridas para la aplicación
            columnas_requeridas = [
                'SOT', 'ESTADO_SOT', 'FEC_GEN_SOT', 'FECHA_INSTALACION', 'FECHA_RECHAZO',
                'FECHA_AGENDA', 'PLAN', 'SCORE_CREDITICIO', 'CARGO_FIJO_CON_IGV',
                'USUARIO_VENDEDOR', 'ZONAL_VDD', 'VENDEDOR_REAL', 'SUPERVISOR',
                'MODALIDAD_VDD', 'NOMBRE_CLIENTE', 'DOC_DE_CLIENTE', 'DOC_VDD',
                'VELOCIDAD_PLAN_MBPS', 'NUMERO_DE_TELEFONO', 'CORREO_ELECTRONICO',
                'CORREO_CORRECTO', 'DIRECCION', 'DEPARTAMENTO_INSTALACION',
                'DISTRITO_INSTALACION', 'PROVINCIA_INSTALACION', 'COD_PLANO',
                'CONTRATA', 'TABULACION', 'TIPO_DE_DEV', 'MOTIVO_DE_DEV',
                'DETALLE_DE_DEV', 'VERIFICACION_BO', 'SOT_CORRECTA', 'LLAMADA _BO',
                'RECUPERABLE', 'OBSERVACION', 'MOTIVO_PENALIDAD', 'PUNTO_VENTA',
                'DISTRIBUIDOR', 'COORDINADOR', 'ESTADO_CONTACTO', 'TECNICO',
                'FEC_PROGRAMACION_TOA', 'FRANJA', 'ESTADO_TOA', 'NUMERO_DE_SEC',
                'ESTADO_FINAL', 'ARBITRAJE', 'SOT_INICIAL', 'TECNOLOGIA', 'VENTA',
                'ALTAS', 'USUARIO_MODIFICACION', 'FECHA_MODIFICACION'
            ]
            
            # Si alguna columna falta en el Excel, la crea vacía para evitar fallos
            for col in columnas_requeridas:
                if col not in df.columns:
                    df[col] = ''
                    
            return df
        except Exception as e:
            st.error(f"Error procesando el Excel de base: {e}")
            return pd.DataFrame(columns=['ZONAL_VDD', 'TABULACION', 'SOT', 'VENTA', 'ALTAS'])
            
    # Retorna DataFrame con columnas mínimas si no se pudo descargar
    return pd.DataFrame(columns=['ZONAL_VDD', 'TABULACION', 'SOT', 'VENTA', 'ALTAS'])
    
def guardar_base_con_reintentos(nuevo_registro, modo, usuario_actual, max_reintentos=5):
    fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    nuevo_registro['USUARIO_MODIFICACION'] = usuario_actual
    nuevo_registro['FECHA_MODIFICACION'] = fecha_actual

    for intento in range(max_reintentos):
        try:
            content = descargar_archivo_dropbox(RUTA_DROPBOX_TOTAL)
            df_actual = pd.read_excel(io.BytesIO(content), sheet_name='Base', dtype=str).fillna('')
            sot_id = str(nuevo_registro['SOT']).strip()

            if modo == "Crear Nuevo Registro":
                if sot_id in df_actual['SOT'].astype(str).str.strip().values:
                    return False, f"La SOT {sot_id} ya fue registrada por otro usuario."
                df_actual = pd.concat([df_actual, pd.DataFrame([nuevo_registro])], ignore_index=True)
            else:
                df_actual.loc[df_actual['SOT'].astype(str).str.strip() == sot_id, list(nuevo_registro.keys())] = list(nuevo_registro.values())

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_actual.to_excel(writer, sheet_name='Base', index=False)
            
            if subir_archivo_dropbox(buffer.getvalue(), RUTA_DROPBOX_TOTAL):
                return True, "Registro guardado correctamente."
        except Exception:
            time.sleep(1)
    return False, "El sistema estuvo ocupado por concurrencia. Intente nuevamente."

def cargar_opciones_config():
    content = descargar_archivo_dropbox(RUTA_DROPBOX_CONFIG)
    config = {}
    if content:
        xls = pd.ExcelFile(io.BytesIO(content))
        config['TABULACION'] = pd.read_excel(xls, 'TABULACION')['TABULACION'].dropna().astype(str).str.strip().tolist()
        config['ESTADO_SOT'] = pd.read_excel(xls, 'ESTADO_SOT')['ESTADO_SOT'].dropna().astype(str).str.strip().tolist()
        config['TIPO_DE_DEV'] = pd.read_excel(xls, 'TIPO _DE_ DEV')['TIPO _DE_ DEV'].dropna().astype(str).str.strip().tolist()
        config['MOTIVO_DE_DEV'] = pd.read_excel(xls, 'MOTIVO_DE_DEV')['MOTIVO_DE_DEV'].dropna().astype(str).str.strip().tolist()
        config['VERIFICACION_BO'] = pd.read_excel(xls, 'VERIFICACION_BO')['VERIFICACION_BO'].dropna().astype(str).str.strip().tolist()
        config['LLAMADA_BO'] = pd.read_excel(xls, 'LLAMADO_BO')['LLAMADA BO'].dropna().astype(str).str.strip().tolist()
        config['MOTIVO_PENALIDAD'] = pd.read_excel(xls, 'MOTIVO_PENALIDAD')['MOTIVO_PENALIDAD'].dropna().astype(str).str.strip().tolist()
    return config

# ---------------------------------------------------------
# CONTROL DE SESIÓN Y BIENVENIDA
# ---------------------------------------------------------
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.nombre = ""
    st.session_state.rol = ""
    st.session_state.mostrar_bienvenida = False

if not st.session_state.logged_in:
    st.markdown("<h2 style='text-align: center;'>📶 Sistema de Grabaciones Claro</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: gray;'>Ingrese sus credenciales de acceso</p>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        usuarios_db = cargar_usuarios()
        with st.form("form_login"):
            user_input = st.text_input("DNI / Usuario").strip()
            pass_input = st.text_input("Contraseña", type="password").strip()
            submit_login = st.form_submit_button("Ingresar", use_container_width=True)
            
            if submit_login:
                if user_input in usuarios_db and usuarios_db[user_input]["password"] == hash_password(pass_input):
                    st.session_state.logged_in = True
                    st.session_state.username = user_input
                    st.session_state.nombre = usuarios_db[user_input]["nombre"]
                    st.session_state.rol = usuarios_db[user_input]["rol"]
                    st.session_state.mostrar_bienvenida = True
                    st.rerun()
                else:
                    st.error("❌ DNI o contraseña incorrectos.")
    st.stop()

# ---------------------------------------------------------
# INTERFAZ PRINCIPAL
# ---------------------------------------------------------

# Saludo de Bienvenida Único al Entrar
if st.session_state.mostrar_bienvenida:
    st.success(f"🎉 ¡Bienvenido(a), **{st.session_state.nombre}**! Has iniciado sesión correctamente.")
    st.balloons()
    st.session_state.mostrar_bienvenida = False

st.sidebar.markdown(f"### 👤 Usuario Activo")
st.sidebar.info(f"**Nombre:** {st.session_state.nombre}\n\n**DNI:** {st.session_state.username}\n\n**Rol:** `{st.session_state.rol.upper()}`")

if st.sidebar.button("🚪 Cerrar Sesión", use_container_width=True):
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.nombre = ""
    st.session_state.rol = ""
    st.session_state.mostrar_bienvenida = False
    st.rerun()

# Selección de Vistas según Rol
opciones_menu = ["Gestión de Grabaciones", "📊 Estadísticas y Reportes"]
if st.session_state.rol == "admin":
    opciones_menu.append("⚙️ Administración de Usuarios")

menu_sel = st.sidebar.radio("Menú Principal:", opciones_menu)
df_base = descargar_excel_base()
opciones_combos = cargar_opciones_config()

# ---------------------------------------------------------
# MÓDULO 1: GESTIÓN DE GRABACIONES
# ---------------------------------------------------------
if menu_sel == "Gestión de Grabaciones":
    st.title("📹 Registro y Edición de SOTs")
    
    st.sidebar.subheader("🔍 Filtros de Búsqueda")
    zonales = ["TODAS"] + sorted([z for z in df_base['ZONAL_VDD'].unique().tolist() if str(z).strip()])
    zonal_filtro = st.sidebar.selectbox("Zonal Vendedor (ZONAL_VDD):", zonales)
    
    tabulaciones = ["TODAS"] + opciones_combos.get('TABULACION', [])
    tab_filtro = st.sidebar.selectbox("Tabulación:", tabulaciones)
    
    sot_busqueda = st.sidebar.text_input("Buscar por SOT:")

    df_filtrado = df_base.copy()
    if zonal_filtro != "TODAS":
        df_filtrado = df_filtrado[df_filtrado['ZONAL_VDD'] == zonal_filtro]
    if tab_filtro != "TODAS":
        df_filtrado = df_filtrado[df_filtrado['TABULACION'] == tab_filtro]
    if sot_busqueda:
        df_filtrado = df_filtrado[df_filtrado['SOT'].astype(str).str.contains(sot_busqueda, case=False, na=False)]

    modo = st.radio("Acción:", ["Crear Nuevo Registro", "Actualizar Registro Existente"], horizontal=True)

    def combo_box(label, key_combo, val_actual):
        lista = opciones_combos.get(key_combo, [])
        idx = lista.index(val_actual) if val_actual in lista else 0
        return st.selectbox(label, options=[""] + lista, index=idx + 1 if val_actual in lista else 0)

    registro_actual = {}
    if modo == "Actualizar Registro Existente":
        if not df_filtrado.empty:
            sot_sel = st.selectbox("Seleccione SOT a editar:", df_filtrado['SOT'].astype(str).tolist())
            registro_actual = df_base[df_base['SOT'].astype(str) == sot_sel].iloc[0].to_dict()
            st.info(f"Modificado por última vez por: **{registro_actual.get('USUARIO_MODIFICACION', 'N/A')}** el **{registro_actual.get('FECHA_MODIFICACION', 'N/A')}**")
        else:
            st.warning("No hay SOTs con los filtros seleccionados.")
            st.stop()

    with st.form("form_grabacion"):
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("### 📌 Venta e Identificación")
            sot = st.text_input("SOT", value=str(registro_actual.get('SOT', '')), disabled=(modo == "Actualizar Registro Existente"))
            estado_sot = combo_box("ESTADO_SOT", 'ESTADO_SOT', str(registro_actual.get('ESTADO_SOT', '')))
            fec_gen_sot = st.text_input("FEC_GEN_SOT (dd/mm/yyyy)", value=str(registro_actual.get('FEC_GEN_SOT', '')))
            fecha_instalacion = st.text_input("FECHA_INSTALACION (dd/mm/yyyy)", value=str(registro_actual.get('FECHA_INSTALACION', '')))
            fecha_rechazo = st.text_input("FECHA_RECHAZO (dd/mm/yyyy)", value=str(registro_actual.get('FECHA_RECHAZO', '')))
            fecha_agenda = st.text_input("FECHA_AGENDA (dd/mm/yyyy)", value=str(registro_actual.get('FECHA_AGENDA', '')))
            plan = st.text_input("PLAN", value=str(registro_actual.get('PLAN', '')))
            score_crediticio = st.text_input("SCORE_CREDITICIO", value=str(registro_actual.get('SCORE_CREDITICIO', '')))
            cargo_fijo = st.text_input("CARGO_FIJO_CON_IGV", value=str(registro_actual.get('CARGO_FIJO_CON_IGV', '')))
            usuario_vendedor = st.text_input("USUARIO_VENDEDOR", value=str(registro_actual.get('USUARIO_VENDEDOR', '')))
            zonal_vdd = st.text_input("ZONAL_VDD", value=str(registro_actual.get('ZONAL_VDD', '')))
            vendedor_real = st.text_input("VENDEDOR_REAL", value=str(registro_actual.get('VENDEDOR_REAL', '')))
            supervisor = st.text_input("SUPERVISOR", value=str(registro_actual.get('SUPERVISOR', '')))
            modalidad_vdd = st.text_input("MODALIDAD_VDD", value=str(registro_actual.get('MODALIDAD_VDD', '')))

        with col2:
            st.markdown("### 👤 Datos del Cliente")
            nombre_cliente = st.text_input("NOMBRE_CLIENTE", value=str(registro_actual.get('NOMBRE_CLIENTE', '')))
            doc_cliente = st.text_input("DOC_DE_CLIENTE", value=str(registro_actual.get('DOC_DE_CLIENTE', '')))
            doc_vdd = st.text_input("DOC_VDD", value=str(registro_actual.get('DOC_VDD', '')))
            velocidad_plan = st.text_input("VELOCIDAD_PLAN_MBPS", value=str(registro_actual.get('VELOCIDAD_PLAN_MBPS', '')))
            telefono = st.text_input("NUMERO_DE_TELEFONO", value=str(registro_actual.get('NUMERO_DE_TELEFONO', '')))
            correo = st.text_input("CORREO_ELECTRONICO", value=str(registro_actual.get('CORREO_ELECTRONICO', '')))
            correo_correcto = st.text_input("CORREO_CORRECTO", value=str(registro_actual.get('CORREO_CORRECTO', '')))
            direccion = st.text_input("DIRECCION", value=str(registro_actual.get('DIRECCION', '')))
            dep_inst = st.text_input("DEPARTAMENTO_INSTALACION", value=str(registro_actual.get('DEPARTAMENTO_INSTALACION', '')))
            dist_inst = st.text_input("DISTRITO_INSTALACION", value=str(registro_actual.get('DISTRITO_INSTALACION', '')))
            prov_inst = st.text_input("PROVINCIA_INSTALACION", value=str(registro_actual.get('PROVINCIA_INSTALACION', '')))
            cod_plano = st.text_input("COD_PLANO", value=str(registro_actual.get('COD_PLANO', '')))
            contrata = st.text_input("CONTRATA", value=str(registro_actual.get('CONTRATA', '')))

        with col3:
            st.markdown("### 📋 Tabulación / BO")
            tabulacion = combo_box("TABULACION", 'TABULACION', str(registro_actual.get('TABULACION', '')))
            tipo_dev = combo_box("TIPO_DE_DEV", 'TIPO_DE_DEV', str(registro_actual.get('TIPO_DE_DEV', '')))
            motivo_dev = combo_box("MOTIVO_DE_DEV", 'MOTIVO_DE_DEV', str(registro_actual.get('MOTIVO_DE_DEV', '')))
            detalle_dev = st.text_input("DETALLE_DE_DEV", value=str(registro_actual.get('DETALLE_DE_DEV', '')))
            verificacion_bo = combo_box("VERIFICACION_BO", 'VERIFICACION_BO', str(registro_actual.get('VERIFICACION_BO', '')))
            sot_correcta = st.text_input("SOT_CORRECTA (Libre)", value=str(registro_actual.get('SOT_CORRECTA', '')))
            llamada_bo = combo_box("LLAMADA _BO", 'LLAMADA_BO', str(registro_actual.get('LLAMADA _BO', '')))
            recuperable = st.text_input("RECUPERABLE", value=str(registro_actual.get('RECUPERABLE', '')))
            observacion = st.text_area("OBSERVACION", value=str(registro_actual.get('OBSERVACION', '')))
            motivo_penalidad = combo_box("MOTIVO_PENALIDAD", 'MOTIVO_PENALIDAD', str(registro_actual.get('MOTIVO_PENALIDAD', '')))
            punto_venta = st.text_input("PUNTO_VENTA", value=str(registro_actual.get('PUNTO_VENTA', '')))
            distribuidor = st.text_input("DISTRIBUIDOR", value=str(registro_actual.get('DISTRIBUIDOR', '')))
            coordinador = st.text_input("COORDINADOR", value=str(registro_actual.get('COORDINADOR', '')))
            estado_contacto = st.text_input("ESTADO_CONTACTO", value=str(registro_actual.get('ESTADO_CONTACTO', '')))
            tecnico = st.text_input("TECNICO", value=str(registro_actual.get('TECNICO', '')))
            fec_prog_toa = st.text_input("FEC_PROGRAMACION_TOA", value=str(registro_actual.get('FEC_PROGRAMACION_TOA', '')))
            franja = st.text_input("FRANJA", value=str(registro_actual.get('FRANJA', '')))
            estado_toa = st.text_input("ESTADO_TOA", value=str(registro_actual.get('ESTADO_TOA', '')))
            numero_sec = st.text_input("NUMERO_DE_SEC", value=str(registro_actual.get('NUMERO_DE_SEC', '')))
            estado_final = st.text_input("ESTADO_FINAL", value=str(registro_actual.get('ESTADO_FINAL', '')))
            arbitraje = st.text_input("ARBITRAJE", value=str(registro_actual.get('ARBITRAJE', '')))
            sot_inicial = st.text_input("SOT_INICIAL", value=str(registro_actual.get('SOT_INICIAL', '')))
            tecnologia = st.text_input("TECNOLOGIA", value=str(registro_actual.get('TECNOLOGIA', '')))
            venta = st.text_input("VENTA", value=str(registro_actual.get('VENTA', '')))
            altas = st.text_input("ALTAS", value=str(registro_actual.get('ALTAS', '')))

        btn_guardar = st.form_submit_button("💾 Guardar Registro", use_container_width=True)

    if btn_guardar:
        if not sot:
            st.error("El campo SOT es obligatorio.")
        else:
            nuevo_registro = {
                'SOT': sot, 'ESTADO_SOT': estado_sot, 'FEC_GEN_SOT': fec_gen_sot,
                'FECHA_INSTALACION': fecha_instalacion, 'FECHA_RECHAZO': fecha_rechazo,
                'FECHA_AGENDA': fecha_agenda, 'PLAN': plan, 'SCORE_CREDITICIO': score_crediticio,
                'CARGO_FIJO_CON_IGV': cargo_fijo, 'USUARIO_VENDEDOR': usuario_vendedor,
                'DEPARTAMENTO_INSTALACION': dep_inst, 'DISTRITO_INSTALACION': dist_inst,
                'PROVINCIA_INSTALACION': prov_inst, 'COD_PLANO': cod_plano, 'CONTRATA': contrata,
                'PUNTO_VENTA': punto_venta, 'DISTRIBUIDOR': distribuidor, 'COORDINADOR': coordinador,
                'ESTADO_CONTACTO': estado_contacto, 'TECNICO': tecnico, 'FEC_PROGRAMACION_TOA': fec_prog_toa,
                'FRANJA': franja, 'ESTADO_TOA': estado_toa, 'NUMERO_DE_SEC': numero_sec,
                'ESTADO_FINAL': estado_final, 'ARBITRAJE': arbitraje, 'SOT_INICIAL': sot_inicial,
                'TECNOLOGIA': tecnologia, 'DOC_VDD': doc_vdd, 'NOMBRE_CLIENTE': nombre_cliente,
                'DOC_DE_CLIENTE': doc_cliente, 'VELOCIDAD_PLAN_MBPS': velocidad_plan,
                'NUMERO_DE_TELEFONO': telefono, 'CORREO_ELECTRONICO': correo,
                'CORREO_CORRECTO': correo_correcto, 'DIRECCION': direccion, 'TABULACION': tabulacion,
                'TIPO_DE_DEV': tipo_dev, 'MOTIVO_DE_DEV': motivo_dev, 'DETALLE_DE_DEV': detalle_dev,
                'VERIFICACION_BO': verificacion_bo, 'SOT_CORRECTA': sot_correcta,
                'LLAMADA _BO': llamada_bo, 'RECUPERABLE': recuperable, 'OBSERVACION': observacion,
                'MOTIVO_PENALIDAD': motivo_penalidad, 'VENTA': venta, 'ALTAS': altas,
                'ZONAL_VDD': zonal_vdd, 'VENDEDOR_REAL': vendedor_real, 'SUPERVISOR': supervisor,
                'MODALIDAD_VDD': modalidad_vdd
            }
            éxito, msj = guardar_base_con_reintentos(nuevo_registro, modo, st.session_state.username)
            if éxito:
                st.success(f"¡SOT {sot} procesada exitosamente por {st.session_state.nombre}!")
                time.sleep(1)
                st.rerun()
            else:
                st.error(f"Error: {msj}")

# ---------------------------------------------------------
# MÓDULO 2: ESTADÍSTICAS Y RESÚMENES (DASHBOARD)
# ---------------------------------------------------------
elif menu_sel == "📊 Estadísticas y Reportes":
    st.title("📊 Panel de Métricas y Tabulaciones")
    
    df_stats = df_base.copy()
    df_stats['VENTA_NUM'] = pd.to_numeric(df_stats['VENTA'], errors='coerce').fillna(0)
    df_stats['ALTAS_NUM'] = pd.to_numeric(df_stats['ALTAS'], errors='coerce').fillna(0)
    df_stats['FEC_GEN_SOT_DT'] = pd.to_datetime(df_stats['FEC_GEN_SOT'], format='%d/%m/%Y', errors='coerce')

    col_f1, col_f2 = st.columns(2)
    with col_f1:
        zonales_dash = ["TODAS"] + sorted([z for z in df_stats['ZONAL_VDD'].unique().tolist() if str(z).strip()])
        zonal_sel = st.selectbox("Filtrar por Zonal:", zonales_dash)
    with col_f2:
        anios_disponibles = ["TODOS"] + sorted([str(int(a)) for a in df_stats['FEC_GEN_SOT_DT'].dt.year.dropna().unique()])
        anio_sel = st.selectbox("Filtrar por Año:", anios_disponibles)

    if zonal_sel != "TODAS":
        df_stats = df_stats[df_stats['ZONAL_VDD'] == zonal_sel]
    if anio_sel != "TODOS":
        df_stats = df_stats[df_stats['FEC_GEN_SOT_DT'].dt.year == int(anio_sel)]

    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("Total SOTs", len(df_stats))
    kpi2.metric("Total Ventas", int(df_stats['VENTA_NUM'].sum()))
    kpi3.metric("Total Altas", int(df_stats['ALTAS_NUM'].sum()))
    kpi4.metric("Instalados", len(df_stats[df_stats['TABULACION'].str.upper() == 'INSTALADO']))

    st.markdown("---")
    st.subheader("📅 Tabulaciones Registradas por Mes")
    if not df_stats.empty:
        df_stats['MES_AÑO'] = df_stats['FEC_GEN_SOT_DT'].dt.to_period('M').astype(str)
        pivot_tab = pd.pivot_table(
            df_stats, 
            index='MES_AÑO', 
            columns='TABULACION', 
            values='SOT', 
            aggfunc='count', 
            fill_value=0
        )
        st.dataframe(pivot_tab, use_container_width=True)

        st.subheader("📊 Distribución por Tabulación")
        st.bar_chart(df_stats['TABULACION'].value_counts())

        st.subheader("👥 Auditoría de Modificaciones por Usuario")
        st.dataframe(
            df_stats['USUARIO_MODIFICACION'].value_counts().reset_index().rename(
                columns={'index': 'Usuario (DNI)', 'USUARIO_MODIFICACION': 'Registros / Ediciones'}
            ),
            use_container_width=True
        )

# ---------------------------------------------------------
# MÓDULO 3: ADMINISTRACIÓN DE USUARIOS
# ---------------------------------------------------------
elif menu_sel == "⚙️ Administración de Usuarios":
    st.title("⚙️ Gestión de Usuarios y Accesos")
    usuarios_db = cargar_usuarios()

    st.subheader("👥 Lista de Usuarios Activos")
    tabla_user = [{"DNI / Usuario": u, "Nombre": v["nombre"], "Rol": v["rol"].upper()} for u, v in usuarios_db.items()]
    st.dataframe(pd.DataFrame(tabla_user), use_container_width=True)

    st.markdown("---")
    st.subheader("➕ Registrar Nuevo Operador / Admin")
    with st.form("form_nuevo_usuario"):
        nuevo_user = st.text_input("DNI (Servirá como Usuario de Login)").strip()
        nuevo_nombre = st.text_input("Nombre Completo del Usuario").strip()
        nueva_pass = st.text_input("Contraseña Inicial", type="password").strip()
        nuevo_rol = st.selectbox("Rol de Acceso", ["operador", "admin"])
        btn_crear_user = st.form_submit_button("Crear Usuario", use_container_width=True)

        if btn_crear_user:
            if not nuevo_user or not nueva_pass or not nuevo_nombre:
                st.error("Todos los campos son obligatorios.")
            elif nuevo_user in usuarios_db:
                st.error("El DNI/Usuario ya está registrado.")
            else:
                usuarios_db[nuevo_user] = {
                    "password": hash_password(nueva_pass),
                    "nombre": nuevo_nombre,
                    "rol": nuevo_rol
                }
                if guardar_usuarios(usuarios_db):
                    st.success(f"Usuario '{nuevo_nombre}' creado con éxito.")
                    time.sleep(1)
                    st.rerun()
