# -*- coding: utf-8 -*-
"""
app.py — Visita a Clientes de Pequeña Empresa (CMAC Caja Arequipa)

Flujo: Búsqueda y carga -> Evaluación de crédito (criterios) ->
       Ficha del cliente -> Ingresos y gastos -> Ubicación (visita) -> Reporte

Diseño mobile-first (ver assets/style.css). El procesamiento del Excel
ocurre en el servidor (no en el celular ni la PC del usuario), y se
cachea con @st.cache_data, así que carga igual de rápido en ambos.
"""
from datetime import datetime

import pandas as pd
import streamlit as st

from utils.helpers import (
    load_css, safe_str, safe_float, fmt_money, slug,
    cargar_excel, CRITERIOS_DEF, CLIENTE_VISITADO_OPCIONES,
    hay_borrador, guardar_borrador, cargar_borrador, borrar_borrador,
    registrar_historial, leer_historial, ahora_peru,
    calcular_resultado, criterios_seleccionados_lista,
    generar_word, generar_pdf, guardar_reporte_en_carpeta,
    reporte_consolidado_por_agencia, reporte_consolidado_por_cliente,
    # 📌 NUEVO — para la pantalla de Búsqueda igualando el mockup:
    iniciales, clase_calificacion, clientes_similares, solo_digitos,
)

st.set_page_config(
    page_title="Visita a Clientes - Caja Arequipa",
    page_icon="🏦",
    layout="centered",
    initial_sidebar_state="collapsed",
)
load_css("assets/style.css")

# --------------------------------------------------------------------------
# ESTADO INICIAL
# --------------------------------------------------------------------------
DEFAULTS = {
    "usuario": "",
    "view": "busqueda",
    "df": None,
    "hoja_usada": "",
    "cliente_actual": None,
    "visitas": {},
    "garantias": [],
    "rcc": [],
    "borrador_prompt": False,
    "ultimo_archivo": None,
    "cliente_visitado": "",
    # 📌 NUEVO — estado de la pantalla de Búsqueda/Carga rediseñada:
    "archivo_meta": {},          # nombre, tamaño y fecha del Excel cargado
    "mostrar_preview": False,    # toggle de "Vista previa de datos"
    "mostrar_camara_dni": False,  # toggle de la cámara de "Escanear DNI"
    "uploader_key_version": 0,    # ver comentario en "btn_limpiar_base"
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v


# --------------------------------------------------------------------------
# COMPONENTES COMUNES
# --------------------------------------------------------------------------
def header(icono, titulo, subtitulo="", icono_derecha=None):
    """Encabezado tipo app móvil. `icono_derecha` es puramente visual (p.
    ej. el "❓" de ayuda del mockup); si más adelante quieres que abra
    algo real, lo más simple en Streamlit es agregar un st.popover justo
    debajo de este header() en la pantalla correspondiente."""
    extra = f'<div class="icon-box-derecha">{icono_derecha}</div>' if icono_derecha else ""
    st.markdown(
        f"""<div class="app-header">
                <div class="icon-box">{icono}</div>
                <div class="titles">
                    <h1>{titulo}</h1>
                    <p>{subtitulo}</p>
                </div>
                {extra}
            </div>""",
        unsafe_allow_html=True,
    )


def badge(texto, clase):
    st.markdown(f'<span class="badge {clase}">{texto}</span>', unsafe_allow_html=True)


PASOS = ["busqueda", "evaluacion", "ficha", "ubicacion", "reporte"]
PASOS_LABEL = {
    "busqueda": ("🔍", "Buscar"),
    "evaluacion": ("⚠️", "Criterio"),
    "ficha": ("👤", "Cliente"),
    "ubicacion": ("📍", "Visita"),
    "reporte": ("📄", "Reporte"),
}


def top_menu():
    """Barra de navegación. 📌 MEJORA: antes se dibujaba como una fila de
    botones normales arriba de la pantalla; ahora, con el <div
    class="marker-bottomnav"> de abajo + las reglas CSS en
    assets/style.css, esa misma fila de botones se ve FIJA abajo de la
    pantalla, con apariencia de íconos (no de botones rojos) — igual a
    la barra "Resumen / Clientes / Visitas / Más" del mockup. La lógica
    de a dónde navega cada botón no cambió."""
    st.markdown('<div class="top-menu-spacer"></div>', unsafe_allow_html=True)
    pasos_visibles = list(PASOS) if st.session_state.cliente_actual is not None else ["busqueda"]
    mostrar_consolidado = st.session_state.df is not None
    n = len(pasos_visibles) + (1 if mostrar_consolidado else 0)
    st.markdown('<div class="marker-bottomnav"></div>', unsafe_allow_html=True)
    cols = st.columns(n)
    for i, paso in enumerate(pasos_visibles):
        icono, label = PASOS_LABEL[paso]
        activo = " ✅" if st.session_state.view == paso else ""
        if cols[i].button(f"{icono} {label}", key=f"nav_{paso}", use_container_width=True,
                           help=f"Ir a {label}", type="primary" if st.session_state.view == paso else "secondary"):
            st.session_state.view = paso
            st.rerun()
    if mostrar_consolidado:
        if cols[-1].button("📊 Resumen", key="nav_consolidado", use_container_width=True,
                            help="Reporte consolidado por agencia y cliente",
                            type="primary" if st.session_state.view == "consolidado" else "secondary"):
            st.session_state.view = "consolidado"
            st.rerun()


def ir_a(paso):
    st.session_state.view = paso
    st.rerun()


def cliente():
    return st.session_state.cliente_actual or {}


def guardar_avance():
    c = cliente()
    if c:
        guardar_borrador(st.session_state.usuario, safe_str(c.get("DOCPEN")), c)


# --------------------------------------------------------------------------
# PANTALLA 1 — BÚSQUEDA Y CARGA
# --------------------------------------------------------------------------
# 📌 REDISEÑO COMPLETO de esta pantalla para igualar el mockup
# "busqueda_y_carga.png" (las dos capturas: "Evaluación de Crédito" /
# Carga de Base de Datos, y "Buscar Cliente" / Búsqueda Inteligente).
# Sigue siendo UNA sola pantalla (como en la app original) porque cargar
# y luego buscar es un solo flujo, pero ahora cada bloque imita en CSS
# las tarjetas, colores e iconografía del mockup. Todo el HTML usa CSS
# en assets/style.css (sección "MEJORAS — Pantalla Buscar Cliente").
# --------------------------------------------------------------------------
def pantalla_busqueda():
    header("🔍", "Buscar Cliente", "Encuentra a tu cliente por nombre, DNI o código", icono_derecha="❓")

    # ---- Usuario (necesario para guardar avance/historial; no está en el
    # mockup porque ese diseño no contempló esta validación, pero la app
    # SÍ la necesita para no mezclar el progreso de distintos auditores) --
    cu1, cu2 = st.columns([1, 2.2])
    with cu1:
        st.markdown('<div class="usuario-bar-label">👤 Usuario</div>', unsafe_allow_html=True)
    with cu2:
        usuario = st.text_input(
            "Tu nombre / usuario", value=st.session_state.usuario, key="input_usuario",
            placeholder="Ej: ACEJ", label_visibility="collapsed",
        )
        st.session_state.usuario = usuario.strip()
    if not st.session_state.usuario:
        st.info("Escribe tu nombre de usuario para continuar.")
        return

    # =====================================================================
    # SECCIÓN 1 — CARGA DE BASE DE DATOS
    # =====================================================================
    with st.container(border=True):
        st.markdown("**📂 Carga de Base de Datos**")
        st.caption("Carga tu archivo Excel con la cartera de clientes")

        if st.session_state.df is None:
            # La key incluye "uploader_key_version": al presionar
            # "Limpiar base" más abajo, esa versión se incrementa y
            # Streamlit dibuja un file_uploader NUEVO y vacío — si se
            # usara siempre la misma key, el archivo ya subido "revivía"
            # solo porque Streamlit recuerda el último valor de cada key.
            archivo = st.file_uploader(
                "Seleccionar archivo Excel — Formatos permitidos: .xlsx, .xls",
                type=["xlsx", "xls"],
                key=f"file_uploader_main_{st.session_state.uploader_key_version}",
            )
            if archivo is not None:
                df, hoja_usada, faltantes = cargar_excel(archivo.getvalue())
                st.session_state.df = df
                st.session_state.hoja_usada = hoja_usada
                # 🔧 AQUÍ: metadatos del archivo que se muestran en el panel
                # verde de abajo (nombre, tamaño, fecha de carga), igual al
                # mockup ("cartera_clientes_junio.xlsx · 2.4 MB · fecha").
                st.session_state.archivo_meta = {
                    "nombre": archivo.name,
                    "tamano_mb": len(archivo.getvalue()) / (1024 * 1024),
                    "fecha": ahora_peru().strftime("%d/%m/%Y %I:%M %p"),
                }
                if hoja_usada != "MUESTRA_FINAL":
                    st.warning("No se encontró la hoja 'MUESTRA_FINAL'; se usó la primera hoja del archivo.")
                if faltantes:
                    st.caption(
                        "Columnas no encontradas (quedarán vacías): "
                        + ", ".join(faltantes[:8]) + ("..." if len(faltantes) > 8 else "")
                    )
                st.rerun()
        else:
            df_actual = st.session_state.df
            meta = st.session_state.get("archivo_meta", {}) or {}
            st.markdown(
                f"""<div class="processed-panel">
                        <div class="processed-head">✅ Archivo procesado correctamente</div>
                        <div class="processed-file">
                            <span class="file-icon">📊</span>
                            <div>
                                <div class="file-name">{meta.get('nombre', 'archivo.xlsx')}</div>
                                <div class="file-meta">{meta.get('tamano_mb', 0):.1f} MB · {meta.get('fecha', '')}</div>
                            </div>
                        </div>
                        <div class="processed-count">{len(df_actual):,}</div>
                        <div class="processed-count-label">Registros cargados</div>
                    </div>""",
                unsafe_allow_html=True,
            )

            st.markdown('<div class="marker-vistaprevia"></div>', unsafe_allow_html=True)
            if st.button("📋 Vista previa de datos", use_container_width=True, key="btn_vista_previa"):
                st.session_state.mostrar_preview = not st.session_state.mostrar_preview
            if st.session_state.mostrar_preview:
                st.dataframe(df_actual.head(20), use_container_width=True, hide_index=True)
                st.caption(f"Mostrando 20 de {len(df_actual):,} registros.")

            st.markdown('<div class="marker-limpiar"></div>', unsafe_allow_html=True)
            if st.button("🗑️ Limpiar base — esto eliminará la base actual cargada",
                         use_container_width=True, key="btn_limpiar_base"):
                st.session_state.df = None
                st.session_state.hoja_usada = ""
                st.session_state.archivo_meta = {}
                st.session_state.mostrar_preview = False
                st.session_state.uploader_key_version += 1
                st.rerun()

    # ---- Estado del Sistema --------------------------------------------
    with st.container(border=True):
        st.markdown("**Estado del Sistema**")
        conectado = st.session_state.df is not None
        st.markdown(
            f'<span class="estado-sistema-dot">{"🟢" if conectado else "⚪"}</span> '
            f'**{"Conectado" if conectado else "Sin datos cargados"}**',
            unsafe_allow_html=True,
        )
        meta = st.session_state.get("archivo_meta", {}) or {}
        if meta:
            st.caption(f"Última actualización: {meta.get('fecha', '')}")

    df = st.session_state.df
    if df is None:
        st.info("Sube el archivo Excel para poder buscar clientes.")
        return

    # =====================================================================
    # SECCIÓN 2 — BÚSQUEDA INTELIGENTE
    # =====================================================================
    with st.container(border=True):
        st.markdown("**🔎 Búsqueda Inteligente**")
        st.caption("Encuentra a tu cliente por nombre, DNI o código")

        c_busq, c_scan = st.columns([2.6, 1])
        with c_busq:
            busqueda = st.text_input(
                "Buscar", placeholder="Nombre, DNI o código de cliente",
                label_visibility="collapsed", key="input_busqueda",
            )
        with c_scan:
            st.markdown('<div class="marker-escanear"></div>', unsafe_allow_html=True)
            if st.button("📷 Escanear DNI", use_container_width=True, key="btn_escanear_dni"):
                st.session_state.mostrar_camara_dni = not st.session_state.mostrar_camara_dni

        if st.session_state.mostrar_camara_dni:
            # ----------------------------------------------------------------
            # 🔧 AQUÍ: "Escanear DNI" abre la cámara pero todavía NO lee el
            # número automáticamente (no hay OCR conectado — ver README,
            # sección "Lo que NO incluye"). Para activarlo de verdad:
            #   1) Llama aquí a un servicio de OCR/reconocimiento de DNI
            #      (por ejemplo una API externa) pasándole `foto_dni.getvalue()`.
            #   2) Toma el número que te devuelva ese servicio y ponlo en
            #      st.session_state["input_busqueda"] ANTES de que el
            #      st.text_input de arriba se vuelva a dibujar, y haz
            #      st.rerun() para que la búsqueda se ejecute sola.
            # ----------------------------------------------------------------
            foto_dni = st.camera_input("Foto del DNI", key="camara_dni")
            if foto_dni is not None:
                st.info(
                    "Foto capturada. La lectura automática del número de DNI "
                    "(OCR) todavía no está conectada — escribe el número "
                    "manualmente en el campo de búsqueda."
                )

    if not busqueda:
        return

    st.markdown(
        '<div class="live-indicator"><span class="live-dot"></span>'
        'Buscando coincidencias en tiempo real...</div>',
        unsafe_allow_html=True,
    )

    b = busqueda.strip().lower()
    b_digitos = solo_digitos(busqueda)
    mask = (
        df.get("DOCPEN", pd.Series("", index=df.index)).astype(str).str.contains(b, case=False, na=False)
        | df.get("BCCTA", pd.Series("", index=df.index)).astype(str).str.contains(b, case=False, na=False)
        | df.get("CODCLI", pd.Series("", index=df.index)).astype(str).str.contains(b, case=False, na=False)
        | df.get("CLIENTE", pd.Series("", index=df.index)).astype(str).str.contains(b, case=False, na=False)
    )
    resultados = df[mask].head(10)

    # ---- ¿Hay un match EXACTO? (DNI completo o nombre completo iguales) -
    exacto = None
    for _, row in resultados.iterrows():
        dni_row = solo_digitos(row.get("DOCPEN"))
        nombre_row = safe_str(row.get("CLIENTE")).strip().lower()
        if (b_digitos and dni_row == b_digitos) or (b and nombre_row == b):
            exacto = row
            break
    if exacto is None and len(resultados) == 1:
        exacto = resultados.iloc[0]

    if exacto is not None:
        render_cliente_encontrado(exacto, df)
    elif len(resultados):
        st.markdown("**No encontramos coincidencias exactas**")
        st.caption("Te mostramos clientes similares de tu agencia")
        render_lista_similares(resultados.head(5))
    else:
        st.warning("No se encontraron coincidencias.")

    st.markdown(
        """<div class="tip-box">
                💡 <b>Consejo</b><br>
                Puedes buscar por nombre completo, parcial, DNI o código de
                cliente. O usa el botón de cámara para escanear el DNI.
            </div>""",
        unsafe_allow_html=True,
    )


def render_cliente_encontrado(row, df):
    """Tarjeta de cliente encontrado, igual al mockup: avatar con
    iniciales, badge verde "Cliente encontrado", datos de contacto a la
    izquierda y datos del crédito a la derecha, botón "Confirmar este
    cliente" y, debajo, una lista de clientes con nombre/DNI parecido
    (para detectar homónimos o posibles duplicados)."""
    nombre = safe_str(row.get("CLIENTE"), "Sin nombre")
    dni = safe_str(row.get("DOCPEN"), "-")
    codigo = safe_str(row.get("CODCLI"), "-")
    direccion = safe_str(row.get("DIRECCION_DOM")) or safe_str(row.get("DIRECCION_NEG"), "No registrada")
    dias_atraso_num = safe_float(row.get("DIAS_ATRASO"))
    calificacion = safe_str(row.get("CATEG_RESULTANTE"), "-")

    # ------------------------------------------------------------------
    # 🔧 AQUÍ: teléfono, correo y límite de crédito NO existen todavía en
    # las columnas del Excel (ver EXCEL_COLUMNS en utils/helpers.py). Si
    # tu base sí los tiene, agrega el nombre real de columna en
    # utils/helpers.py → COLUMNAS_OPCIONALES_CONTACTO, y reemplaza aquí
    # "TELEFONO" / "EMAIL" / "LIMITE_CREDITO_MN" por ese mismo nombre.
    # ------------------------------------------------------------------
    telefono = safe_str(row.get("TELEFONO"), "No registrado")
    correo = safe_str(row.get("EMAIL"), "No registrado")
    limite_credito = (
        fmt_money(row.get("LIMITE_CREDITO_MN"))
        if safe_str(row.get("LIMITE_CREDITO_MN")) else "No disponible"
    )

    clase_atraso = "valor-alerta" if dias_atraso_num > 0 else "valor-ok"

    st.markdown(
        f"""<div class="cliente-card">
                <div class="cliente-card-top">
                    <div class="avatar-circle">{iniciales(nombre)}</div>
                    <div class="cliente-nombre-wrap">
                        <div class="cliente-nombre">{nombre}</div>
                    </div>
                    <span class="badge-encontrado">Cliente encontrado</span>
                </div>
                <div class="cliente-grid">
                    <div class="cliente-col">
                        <div class="dato-fila">📞 {telefono}</div>
                        <div class="dato-fila">✉️ {correo}</div>
                        <div class="dato-fila">📍 {direccion}</div>
                        <div class="dato-fila">🆔 DNI: <b>{dni}</b></div>
                    </div>
                    <div class="cliente-col">
                        <div class="dato-etiqueta">Código de Cliente</div>
                        <div class="dato-valor">{codigo}</div>
                        <div class="dato-etiqueta">Límite de Crédito</div>
                        <div class="dato-valor">{limite_credito}</div>
                        <div class="dato-etiqueta">Saldo Actual</div>
                        <div class="dato-valor">{fmt_money(row.get('SALDO_MN'))}</div>
                        <div class="dato-etiqueta">Días de Atraso</div>
                        <div class="dato-valor {clase_atraso}">{dias_atraso_num:.0f} días</div>
                        <div class="dato-etiqueta">Calificación</div>
                        <span class="chip-calif {clase_calificacion(calificacion)}">{calificacion}</span>
                    </div>
                </div>
            </div>""",
        unsafe_allow_html=True,
    )

    st.markdown('<div class="marker-confirmar"></div>', unsafe_allow_html=True)
    if st.button("✅ Confirmar este cliente", use_container_width=True, key="btn_confirmar_cliente", type="primary"):
        seleccionar_cliente(row.to_dict())
    st.caption(f"Continuar con la evaluación de {nombre}")

    similares = clientes_similares(df, row)
    if len(similares):
        st.write("")
        st.markdown("**¿No es este cliente?**")
        st.caption("Estos otros clientes tienen nombre o DNI parecido — verifica que no se trate de la misma persona registrada dos veces.")
        render_lista_similares(similares)


def render_lista_similares(resultados):
    """Lista de clientes parecidos/aproximados, igual al mockup: avatar
    chico, nombre, DNI, saldo, y un botón para abrir directamente la
    evaluación de ese cliente."""
    for idx, row in resultados.iterrows():
        nombre = safe_str(row.get("CLIENTE"), "Sin nombre")
        st.markdown(
            f"""<div class="similar-item">
                    <div class="avatar-circle avatar-chica">{iniciales(nombre)}</div>
                    <div class="similar-info">
                        <div class="similar-nombre">{nombre}</div>
                        <div class="similar-dni">DNI: {safe_str(row.get('DOCPEN'), '-')}</div>
                    </div>
                    <div class="similar-saldo">{fmt_money(row.get('SALDO_MN'))}</div>
                </div>""",
            unsafe_allow_html=True,
        )
        st.markdown('<div class="marker-abrir-similar"></div>', unsafe_allow_html=True)
        if st.button("Abrir →", key=f"abrir_similar_{idx}", use_container_width=True):
            seleccionar_cliente(row.to_dict())


def seleccionar_cliente(fila):
    st.session_state.cliente_actual = fila
    dni = safe_str(fila.get("DOCPEN"))
    if hay_borrador(st.session_state.usuario, dni):
        st.session_state.borrador_prompt = True
    else:
        st.session_state.visitas = {}
        st.session_state.garantias = []
        st.session_state.rcc = []
        st.session_state.cliente_visitado = ""
        ir_a("evaluacion")
    st.rerun()


def prompt_borrador():
    c = cliente()
    with st.container(border=True):
        st.warning(f"Encontramos un avance guardado para **{safe_str(c.get('CLIENTE'))}** con tu usuario.")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔄 Continuar avance", use_container_width=True):
                cargar_borrador(st.session_state.usuario, safe_str(c.get("DOCPEN")))
                st.session_state.borrador_prompt = False
                ir_a("evaluacion")
        with c2:
            if st.button("🆕 Iniciar nuevo", use_container_width=True):
                borrar_borrador(st.session_state.usuario, safe_str(c.get("DOCPEN")))
                st.session_state.visitas = {}
                st.session_state.garantias = []
                st.session_state.rcc = []
                st.session_state.cliente_visitado = ""
                st.session_state.borrador_prompt = False
                ir_a("evaluacion")


# --------------------------------------------------------------------------
# PANTALLA 2 — EVALUACIÓN DE CRÉDITO (CRITERIOS PARA LA VISITA)
# --------------------------------------------------------------------------
def pantalla_evaluacion():
    c = cliente()
    header("⚠️", "Evaluación de Crédito", f"Cliente: {safe_str(c.get('CLIENTE'))}")
    st.caption("Marca los criterios identificados para esta visita.")

    for categoria, items in CRITERIOS_DEF.items():
        keys = [f"chk_{slug(categoria)}_{slug(item)}" for item in items]
        activo = any(st.session_state.get(k, False) for k in keys)
        icono = "🔴" if activo else "⚪"
        with st.container(border=True):
            with st.expander(f"{icono} {categoria}", expanded=activo):
                for item, key in zip(items, keys):
                    st.checkbox(item, key=key)
                    if item == "Calificación diferente a normal" and st.session_state.get(key):
                        st.text_input("Indicar la calificación a la fecha de revisión", key="calif_revision")

    n_marcados = sum(
        1 for cat, items in CRITERIOS_DEF.items() for item in items
        if st.session_state.get(f"chk_{slug(cat)}_{slug(item)}", False)
    )
    if n_marcados:
        badge(f"⚠️ {n_marcados} criterio(s) marcado(s)", "badge-pend")
    else:
        badge("Sin criterios de riesgo marcados", "badge-ok")

    st.write("")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("⬅️ Volver a buscar", use_container_width=True):
            ir_a("busqueda")
    with c2:
        if st.button("Guardar y continuar ➡️", use_container_width=True, type="primary"):
            guardar_avance()
            ir_a("ficha")


# --------------------------------------------------------------------------
# PANTALLA 3 — FICHA DEL CLIENTE
# --------------------------------------------------------------------------
def pantalla_ficha():
    c = cliente()
    header("👤", "Cliente y Crédito", "Ficha de identidad (solo lectura)")

    st.markdown(
        f"""<div class="banner-cliente">
                <div class="nombre">{safe_str(c.get('CLIENTE'))}</div>
                <div class="dni">DNI: {safe_str(c.get('DOCPEN'))} · Cuenta: {safe_str(c.get('BCCTA'))}</div>
            </div>""",
        unsafe_allow_html=True,
    )

    with st.container(border=True):
        st.markdown("**Información del crédito**")
        chips = [
            ("N° de cuenta", safe_str(c.get("BCCTA"), "-")),
            ("Tipo de crédito", safe_str(c.get("PRODUCTO_CAJA"), "-")),
            ("Calificación", safe_str(c.get("CATEG_RESULTANTE"), "-")),
            ("Importe desembolsado", fmt_money(c.get("IMPDESEMB_MN"))),
            ("Saldo actual", fmt_money(c.get("SALDO_MN"))),
            ("Fecha de desembolso", safe_str(c.get("FECDES"), "-")),
            ("Último pago", safe_str(c.get("FECHA_UTLPAGO"), "-")),
        ]
        chips_html = "".join(
            f'<div class="chip"><div class="lbl">{lbl}</div><div class="val">{val}</div></div>'
            for lbl, val in chips
        )
        st.markdown(f'<div class="info-credito">{chips_html}</div>', unsafe_allow_html=True)

        imp = safe_float(c.get("IMPDESEMB_MN"))
        saldo = safe_float(c.get("SALDO_MN"))
        if imp > 0:
            usado_pct = max(0.0, min(1.0, 1 - (saldo / imp)))
            st.progress(usado_pct, text=f"{usado_pct*100:.0f}% pagado del importe original")

    with st.expander("ℹ️ Información adicional"):
        info = [
            ("Agencia", c.get("AGENCIA")), ("Analista vigente", c.get("ANALISTA")),
            ("Analista evaluador", c.get("ANALISTA_EVAL")), ("Tipo SBS", c.get("TIPO_SBS")),
            ("Rubro / Actividad", c.get("ACTIVIDAD_ECON")), ("Segmentación MYPE", c.get("SEGMENTACION_MYPE")),
            ("Cuenta aval", c.get("CUENTA_AVAL")), ("Estado del crédito", c.get("ESTADO_CREDITO")),
        ]
        for label, val in info:
            st.write(f"**{label}:** {safe_str(val, '-')}")

    st.write("")
    st.markdown('<div class="nav-pie">', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        if st.button("⬅️ Criterio", use_container_width=True):
            ir_a("evaluacion")
    with c2:
        if st.button("Ir a la visita ➡️", use_container_width=True, type="primary"):
            guardar_avance()
            ir_a("ubicacion")
    st.markdown('</div>', unsafe_allow_html=True)


# --------------------------------------------------------------------------
# (Se eliminó la pantalla de "Ingresos y Gastos" / evaluación de crédito
#  detallada a pedido — esa vista ya no se muestra en la app. Los cálculos
#  de utilidad neta quedan en 0 por defecto en el reporte si no se usan.)
# --------------------------------------------------------------------------


# --------------------------------------------------------------------------
# PANTALLA 5 — UBICACIÓN (VISITA: DOMICILIO / NEGOCIO / AVAL)
# --------------------------------------------------------------------------
TIPOS_VISITA = {
    "negocio": ("💼", "Negocio", "DIRECCION_NEG", "DISTRITO_NEG", "PROVINCIA_NEG", "DEPARTAMENTO_NEG", True),
    "laboral": ("🏢", "Centro laboral", None, None, None, None, False),
    "aval": ("🧾", "Aval", None, None, None, None, False),
    "domicilio": ("🏠", "Domicilio", "DIRECCION_DOM", "DISTRITO_DOM", "PROVINCIA_DOM", "DEPARTAMENTO_DOM", False),
}


def pantalla_ubicacion():
    c = cliente()
    header("📍", "Nueva Visita", "Verificación: negocio (obligatorio), laboral, aval y domicilio (opcionales)")

    with st.container(border=True):
        st.markdown("**Cliente visitado** — resultado general de esta visita")
        idx_actual = (
            CLIENTE_VISITADO_OPCIONES.index(st.session_state.cliente_visitado)
            if st.session_state.cliente_visitado in CLIENTE_VISITADO_OPCIONES else 0
        )
        seleccion = st.selectbox(
            "Selecciona una opción", CLIENTE_VISITADO_OPCIONES, index=idx_actual,
            key="sel_cliente_visitado", label_visibility="collapsed",
        )
        st.session_state.cliente_visitado = seleccion

    tabs = st.tabs([f"{TIPOS_VISITA[t][0]} {TIPOS_VISITA[t][1]}" for t in TIPOS_VISITA])
    for tab, clave in zip(tabs, TIPOS_VISITA):
        with tab:
            render_visita(clave, c)

    st.write("")
    st.markdown('<div class="nav-pie">', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        if st.button("⬅️ Cliente", use_container_width=True, key="back_ubic"):
            ir_a("ficha")
    with c2:
        if st.button("Ir al reporte ➡️", use_container_width=True, type="primary", key="next_ubic"):
            guardar_avance()
            ir_a("reporte")
    st.markdown('</div>', unsafe_allow_html=True)


def render_visita(clave, c):
    icono, etiqueta, k_dir, k_dist, k_prov, k_depto, obligatoria = TIPOS_VISITA[clave]
    visitas = st.session_state.visitas
    data = visitas.get(clave, {})

    with st.container(border=True):
        etiqueta_foto = "Foto de verificación (obligatoria)" if obligatoria else "Foto de verificación (opcional)"
        st.markdown(f"**Paso 1 · {etiqueta_foto}**")
        foto_camara = st.camera_input("Tomar foto ahora", key=f"camara_{clave}")
        foto_archivo = st.file_uploader("...o subir desde galería", type=["jpg", "jpeg", "png"], key=f"upload_{clave}")
        foto_final = foto_camara if foto_camara is not None else foto_archivo
        if foto_final is None and data.get("foto_bytes"):
            st.image(data["foto_bytes"], caption="Foto guardada previamente", width=200)
        if obligatoria and foto_final is None and not data.get("foto_bytes"):
            st.warning("⚠ Esta sección requiere foto de verificación antes de guardar.")

        st.markdown("**Paso 2 · Ubicación GPS**")
        cgps1, cgps2 = st.columns([1, 2])
        with cgps1:
            capturar = st.button("📡 Capturar GPS", key=f"btn_gps_{clave}")
        lat, lon, precision = data.get("lat"), data.get("lon"), data.get("precision")
        if capturar:
            try:
                from streamlit_js_eval import get_geolocation
                loc = get_geolocation(key=f"geo_{clave}_{datetime.now().timestamp()}")
                if loc and "coords" in loc:
                    lat = loc["coords"]["latitude"]
                    lon = loc["coords"]["longitude"]
                    precision = loc["coords"].get("accuracy")
                else:
                    st.warning("No se pudo obtener la ubicación. Acepta el permiso en el navegador e inténtalo otra vez.")
            except Exception:
                st.warning("Geolocalización no disponible en este entorno. Ingresa la dirección manualmente.")
        with cgps2:
            if lat and lon:
                st.success(f"Lat: {lat:.6f} · Lon: {lon:.6f}" + (f" (±{precision:.0f} m)" if precision else ""))
                st.map(pd.DataFrame({"lat": [lat], "lon": [lon]}), zoom=15, height=160)
            else:
                st.caption("Sin ubicación capturada todavía.")

        st.markdown("**Paso 3 · Datos del lugar**")
        valor_dir = data.get("direccion") or (safe_str(c.get(k_dir)) if k_dir else "")
        direccion = st.text_input("Dirección", value=valor_dir, key=f"dir_{clave}")
        cc1, cc2 = st.columns(2)
        with cc1:
            valor_dist = data.get("distrito") or (safe_str(c.get(k_dist)) if k_dist else "")
            distrito = st.text_input("Distrito", value=valor_dist, key=f"dist_{clave}")
            valor_prov = data.get("provincia") or (safe_str(c.get(k_prov)) if k_prov else "")
            provincia = st.text_input("Provincia", value=valor_prov, key=f"prov_{clave}")
        with cc2:
            valor_depto = data.get("departamento") or (safe_str(c.get(k_depto)) if k_depto else "")
            departamento = st.text_input("Departamento", value=valor_depto, key=f"depto_{clave}")
            referencia = st.text_input("Referencia", value=data.get("referencia", ""), key=f"ref_{clave}")

        st.markdown("**Paso 4 · Observaciones**")
        ahora = ahora_peru()
        fecha_v = st.date_input("Fecha de visita", value=ahora.date(), key=f"fecha_{clave}")
        hora_v = st.time_input("Hora de visita", value=ahora.time(), key=f"hora_{clave}")
        entrevista_con = st.text_input("Entrevista con", value=data.get("entrevista_con", ""), key=f"entrevista_{clave}")
        comentarios = st.text_area("Comentarios", value=data.get("comentarios", ""), key=f"comentarios_{clave}")

        puede_guardar = (not obligatoria) or foto_final is not None or bool(data.get("foto_bytes"))
        if st.button(f"💾 Guardar visita de {etiqueta}", key=f"guardar_{clave}", use_container_width=True,
                     type="primary", disabled=not puede_guardar):
            st.session_state.visitas[clave] = {
                "direccion": direccion, "distrito": distrito, "provincia": provincia,
                "departamento": departamento, "referencia": referencia,
                "fecha": str(fecha_v), "hora": str(hora_v),
                "entrevista_con": entrevista_con, "comentarios": comentarios,
                "lat": lat, "lon": lon, "precision": precision,
                "foto_bytes": foto_final.getvalue() if foto_final is not None else data.get("foto_bytes"),
            }
            guardar_avance()
            st.success(f"✅ Visita de {etiqueta} guardada — {fecha_v} {hora_v} (hora Perú)")
        if not puede_guardar:
            st.caption("Toma o sube la foto obligatoria del negocio para poder guardar esta sección.")

        if clave in visitas:
            badge("Registrada", "badge-ok")
        else:
            badge("Pendiente", "badge-pend")


# --------------------------------------------------------------------------
# PANTALLA 6 — GENERACIÓN DE REPORTE
# --------------------------------------------------------------------------
def pantalla_reporte():
    c = cliente()
    header("📄", "Generación de Reporte", "Revisión final y descarga del documento")

    visitas = st.session_state.visitas
    secciones = [("negocio", "Negocio", True), ("laboral", "Laboral", False),
                 ("aval", "Aval", False), ("domicilio", "Domicilio", False)]
    completas = sum(1 for k, _, _ in secciones if k in visitas)

    with st.container(border=True):
        st.markdown(f"**Resumen de calidad** — {completas} de {len(secciones)} visitas registradas")
        cols = st.columns(4)
        for col, (clave, etiqueta, obligatoria) in zip(cols, secciones):
            ok = clave in visitas
            badge_clase = "badge-ok" if ok else ("badge-pend" if obligatoria else "badge-warn")
            texto = "Foto capturada" if ok else ("Falta (obligatoria)" if obligatoria else "Opcional")
            col.markdown(
                f"""<div style="text-align:center;padding:0.5rem 0.2rem;border-radius:10px;background:{'#F0FDF4' if ok else '#FEF2F2'};">
                        <div style="font-size:1.3rem;">{'✅' if ok else ('⚠️' if obligatoria else '➖')}</div>
                        <div style="font-weight:700;font-size:0.8rem;">{etiqueta}</div>
                        <span class="badge {badge_clase}" style="font-size:0.65rem;">{texto}</span>
                    </div>""",
                unsafe_allow_html=True,
            )

    if "negocio" not in visitas:
        st.warning("Acción requerida — falta la visita obligatoria al **Negocio**. Puedes generar el reporte igual; quedará indicado como pendiente.")

    criterios_dict = {k: v for k, v in st.session_state.items() if k.startswith("chk_")}
    criterios_txt = criterios_seleccionados_lista(criterios_dict, st.session_state.get("calif_revision", ""))
    ing = {k: st.session_state.get(k, 0.0) for k in [
        "ingreso_principal", "otros_ingresos", "op_alquiler", "op_servicios", "op_transporte",
        "op_mercaderia", "op_publicidad", "op_otros", "fam_alimentacion", "fam_vivienda",
        "fam_servicios", "fam_educacion", "fam_salud", "fam_otros",
    ]}
    calc = calcular_resultado(ing)
    cliente_visitado = st.session_state.get("cliente_visitado", "")

    with st.container(border=True):
        st.markdown("**Resumen de la evaluación**")
        st.write(f"**Cuenta cliente:** {safe_str(c.get('BCCTA'))}")
        st.write(f"**N° de operación:** {safe_str(c.get('BCOPER'))}")
        st.write(f"**Nombre del cliente:** {safe_str(c.get('CLIENTE'))}")
        st.write(f"**Módulo:** {safe_str(c.get('MODULO'))}")
        st.write(f"**Analista vigente:** {safe_str(c.get('ANALISTA'))}")
        st.write(f"**Analista evaluador:** {safe_str(c.get('ANALISTA_EVAL'))}")
        st.write(f"**Auditor:** {st.session_state.usuario}")
        st.write(f"**Fecha de visita:** {ahora_peru().strftime('%d/%m/%Y %H:%M')} (hora Perú)")
        st.write(f"**Cliente visitado:** {cliente_visitado or '—'}")
        st.write(f"**Utilidad neta:** {fmt_money(calc['utilidad_neta'])}")
        st.write(f"**Criterio seleccionado:** {len(criterios_txt)}")
        if criterios_txt:
            for ct in criterios_txt:
                st.caption("• " + ct)

    with st.container(border=True):
        st.markdown("**Generar y descargar reporte**")
        st.caption("Disponible en Word (.docx) y PDF. Se guarda automáticamente en la carpeta de reportes configurada.")

        base_nombre = f"Visita_{slug(c.get('CLIENTE'))}_{ahora_peru().strftime('%Y%m%d_%H%M')}"

        c1, c2 = st.columns(2)
        with c1:
            if st.button("📝 Generar Word", use_container_width=True, type="primary"):
                buf = generar_word(c, criterios_txt, calc, ing, visitas, st.session_state.garantias,
                                    st.session_state.rcc, st.session_state.usuario, cliente_visitado)
                nombre = base_nombre + ".docx"
                ruta = guardar_reporte_en_carpeta(nombre, buf.getvalue())
                n_ag, n_gen = registrar_historial(st.session_state.usuario, c, "Word", nombre,
                                                   "; ".join(criterios_txt), cliente_visitado, ruta)
                st.session_state.ultimo_archivo = (nombre, buf.getvalue(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                st.session_state["ultimo_conteo"] = (n_ag, n_gen, ruta)
        with c2:
            if st.button("📕 Generar PDF", use_container_width=True, type="primary"):
                buf = generar_pdf(c, criterios_txt, calc, ing, visitas, st.session_state.garantias,
                                   st.session_state.rcc, st.session_state.usuario, cliente_visitado)
                nombre = base_nombre + ".pdf"
                ruta = guardar_reporte_en_carpeta(nombre, buf.getvalue())
                n_ag, n_gen = registrar_historial(st.session_state.usuario, c, "PDF", nombre,
                                                   "; ".join(criterios_txt), cliente_visitado, ruta)
                st.session_state.ultimo_archivo = (nombre, buf.getvalue(), "application/pdf")
                st.session_state["ultimo_conteo"] = (n_ag, n_gen, ruta)

        if st.session_state.ultimo_archivo:
            nombre, contenido, mime = st.session_state.ultimo_archivo
            st.download_button(f"⬇️ Descargar {nombre}", data=contenido, file_name=nombre, mime=mime, use_container_width=True)
            n_ag, n_gen, ruta = st.session_state.get("ultimo_conteo", (None, None, ""))
            agencia_txt = safe_str(c.get("AGENCIA"), "-")
            if n_ag is not None:
                st.success(
                    f"Reporte generado. Visita N° {n_ag} en la agencia **{agencia_txt}** "
                    f"(N° {n_gen} en general)."
                )
            if ruta:
                st.caption(f"📁 Copia guardada en: `{ruta}`")
            else:
                st.caption("⚠ No se pudo guardar copia automática en la carpeta de reportes; usa el botón de descarga.")

    with st.expander("🗂️ Ver historial de reportes generados"):
        hist = leer_historial()
        if len(hist):
            st.dataframe(hist.tail(20), use_container_width=True, hide_index=True)
        else:
            st.caption("Aún no se ha generado ningún reporte.")

    st.write("")
    st.markdown('<div class="nav-pie">', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        if st.button("⬅️ Visita", use_container_width=True):
            ir_a("ubicacion")
    with c2:
        if st.button("🏁 Terminar y volver a buscar", use_container_width=True):
            c_dni = safe_str(c.get("DOCPEN"))
            borrar_borrador(st.session_state.usuario, c_dni)
            st.session_state.cliente_actual = None
            st.session_state.visitas = {}
            st.session_state.garantias = []
            st.session_state.rcc = []
            st.session_state.ultimo_archivo = None
            st.session_state.cliente_visitado = ""
            ir_a("busqueda")
    st.markdown('</div>', unsafe_allow_html=True)


# --------------------------------------------------------------------------
# PANTALLA 7 — RESUMEN / REPORTE CONSOLIDADO POR AGENCIA Y CLIENTE
# --------------------------------------------------------------------------
def pantalla_consolidado():
    header("📊", "Reporte Consolidado", "Visitas realizadas por agencia y por cliente")
    st.caption(
        "Se genera a partir de todo lo guardado en el historial (carpeta `data/`, "
        "o la carpeta de reportes que configures en `utils/helpers.py` → REPORTES_DIR)."
    )

    resumen_agencia = reporte_consolidado_por_agencia()
    with st.container(border=True):
        st.markdown("**Por agencia** — clientes visitados y reportes generados")
        if len(resumen_agencia):
            st.dataframe(resumen_agencia, use_container_width=True, hide_index=True)
        else:
            st.caption("Aún no hay reportes generados para consolidar.")

    agencias_disponibles = ["(Todas)"] + (
        sorted(resumen_agencia["Agencia"].dropna().astype(str).unique().tolist())
        if len(resumen_agencia) else []
    )
    with st.container(border=True):
        st.markdown("**Detalle por cliente**")
        agencia_sel = st.selectbox("Filtrar por agencia", agencias_disponibles, key="sel_agencia_consolidado")
        filtro = None if agencia_sel == "(Todas)" else agencia_sel
        detalle = reporte_consolidado_por_cliente(filtro)
        if len(detalle):
            st.dataframe(detalle, use_container_width=True, hide_index=True)
        else:
            st.caption("Sin datos para este filtro todavía.")

    st.write("")
    if st.button("⬅️ Volver a buscar", use_container_width=True):
        ir_a("busqueda")


# --------------------------------------------------------------------------
# ROUTER
# --------------------------------------------------------------------------
top_menu()

if st.session_state.borrador_prompt:
    prompt_borrador()
elif st.session_state.view == "consolidado" and st.session_state.df is not None:
    pantalla_consolidado()
elif st.session_state.view == "busqueda" or st.session_state.cliente_actual is None:
    pantalla_busqueda()
elif st.session_state.view == "evaluacion":
    pantalla_evaluacion()
elif st.session_state.view == "ficha":
    pantalla_ficha()
elif st.session_state.view == "ubicacion":
    pantalla_ubicacion()
elif st.session_state.view == "reporte":
    pantalla_reporte()
else:
    pantalla_busqueda()
