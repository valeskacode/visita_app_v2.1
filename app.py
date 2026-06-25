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
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v


# --------------------------------------------------------------------------
# COMPONENTES COMUNES
# --------------------------------------------------------------------------
def header(icono, titulo, subtitulo=""):
    st.markdown(
        f"""<div class="app-header">
                <div class="icon-box">{icono}</div>
                <div class="titles">
                    <h1>{titulo}</h1>
                    <p>{subtitulo}</p>
                </div>
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
    """Submenú superior por íconos, en una sola línea, para moverse entre
    las vistas (reemplaza la barra inferior). Incluye también el acceso
    al Resumen/Consolidado por agencia, disponible desde que hay datos
    cargados (aunque todavía no se haya abierto un cliente)."""
    st.markdown('<div class="top-menu-spacer"></div>', unsafe_allow_html=True)
    pasos_visibles = list(PASOS) if st.session_state.cliente_actual is not None else ["busqueda"]
    mostrar_consolidado = st.session_state.df is not None
    n = len(pasos_visibles) + (1 if mostrar_consolidado else 0)
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
def pantalla_busqueda():
    header("🏦", "Buscar Cliente", "Carga tu base y busca por DNI o cuenta")

    with st.container(border=True):
        usuario = st.text_input(
            "Tu nombre / usuario (para guardar tu progreso e historial)",
            value=st.session_state.usuario, key="input_usuario",
            placeholder="Ej: ACEJ",
        )
        st.session_state.usuario = usuario.strip()

    if not st.session_state.usuario:
        st.info("Escribe tu nombre de usuario para continuar.")
        return

    with st.container(border=True):
        st.markdown("**📂 Carga de Base de Datos**")
        st.caption("Sube el Excel con la hoja 'MUESTRA_FINAL'. Formatos: .xlsx, .xls")
        archivo = st.file_uploader("Seleccionar archivo Excel", type=["xlsx", "xls"], label_visibility="collapsed")
        if archivo is not None:
            df, hoja_usada, faltantes = cargar_excel(archivo.getvalue())
            st.session_state.df = df
            st.session_state.hoja_usada = hoja_usada
            st.success(f"✅ {len(df)} registros cargados desde la hoja **{hoja_usada}**")
            if hoja_usada != "MUESTRA_FINAL":
                st.warning("No se encontró la hoja 'MUESTRA_FINAL'; se usó la primera hoja del archivo.")
            if faltantes:
                st.caption(
                    "Columnas no encontradas (quedarán vacías): "
                    + ", ".join(faltantes[:8]) + ("..." if len(faltantes) > 8 else "")
                )

    df = st.session_state.df
    if df is None:
        st.info("Sube el archivo Excel para poder buscar clientes.")
        return

    with st.container(border=True):
        st.markdown("**🔎 Búsqueda Inteligente**")
        st.caption("Busca por DNI o número de cuenta (también acepta el nombre)")
        busqueda = st.text_input("Buscar", placeholder="DNI, N° de cuenta o nombre", label_visibility="collapsed")

    if busqueda:
        b = busqueda.strip().lower()
        mask = (
            df.get("DOCPEN", pd.Series("", index=df.index)).astype(str).str.contains(b, case=False, na=False)
            | df.get("BCCTA", pd.Series("", index=df.index)).astype(str).str.contains(b, case=False, na=False)
            | df.get("CLIENTE", pd.Series("", index=df.index)).astype(str).str.contains(b, case=False, na=False)
        )
        resultados = df[mask].head(8)

        if len(resultados) == 0:
            st.warning("No se encontraron coincidencias.")
        else:
            for idx, row in resultados.iterrows():
                with st.container(border=True):
                    c1, c2 = st.columns([3, 1])
                    with c1:
                        st.markdown(f"**{safe_str(row.get('CLIENTE'))}**")
                        st.caption(f"DNI: {safe_str(row.get('DOCPEN'))} · Cuenta: {safe_str(row.get('BCCTA'))} · Cód: {safe_str(row.get('CODCLI'))}")
                        st.caption(f"Saldo: {fmt_money(row.get('SALDO_MN'))}")
                    with c2:
                        if st.button("Abrir", key=f"abrir_{idx}", use_container_width=True):
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
