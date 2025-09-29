# app.py
# Streamlit app para generar documentos (Word) por fila de un Excel
# Pagarés individuales por sucursal + Convenio grupal (KGRUPAL)

import io, os, re, zipfile, tempfile
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

# ---------- Utilidades ----------
SAFE_NAME_RE = re.compile(r"[^A-Za-z0-9._\-áéíóúÁÉÍÓÚñÑ ]+")

def safe_name(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    s = SAFE_NAME_RE.sub("_", s)
    s = re.sub(r"\s+", " ", s)
    return s[:150]

@st.cache_data(show_spinner=False)
def read_excel(file) -> pd.DataFrame:
    df = pd.read_excel(file)
    # Si los encabezados están en la primera fila:
    if sum(str(c).startswith("Unnamed") for c in df.columns) > len(df.columns) * 0.6:
        headers = [str(x).strip() for x in df.iloc[0].tolist()]
        df = df.iloc[1:].copy()
        df.columns = headers
    df.columns = [str(c).strip() for c in df.columns]
    return df

def normalize_str(x: str) -> str:
    if x is None:
        return ""
    s = str(x).strip().lower()
    for a, b in (("á","a"),("é","e"),("í","i"),("ó","o"),("ú","u"),("ñ","n")):
        s = s.replace(a,b)
    return s

# Fecha “DD DE MES DEL YYYY”
MESES_MAYUS = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]
def fecha_hoy_es() -> str:
    d = datetime.now()
    return f"{d.day:02d} DE {MESES_MAYUS[d.month-1]} DEL {d.year}"

# Sucursal normalizada para usar como llave de plantillas
def detect_branch(texto_sucursal: str) -> str:
    s = normalize_str(texto_sucursal)
    if "huehuetoca" in s: return "HUEHUETOCA"
    if "tecamac" in s or "tecámac" in s: return "TECAMAC"
    if "zumpango" in s: return "ZUMPANGO"
    return "OTRA"

BRANCH_ADDRESSES = {
    "HUEHUETOCA": "Plaza Comercial El Árbol, Local 07 en Av. Jalapa No. 50. Col. Barrio La Cañada, Huehuetoca C.P. 54685, Edo. De México.",
    "TECAMAC":    "San Rafael No. 22, Tecámac de Felipe Villanueva, Tecámac, C.P. 55740, Edo. De México.",
    "ZUMPANGO":   "Plaza San Juan, Local 9 en Calle Francisco Xavier Mina, No. 37, Col. San Juan, Zumpango de Ocampo, Estado de México. C.P. 55600.",
}

# -------- Números a letras (es-MX, básica) --------
UNIDADES = ["cero","uno","dos","tres","cuatro","cinco","seis","siete","ocho","nueve","diez","once","doce","trece","catorce","quince","dieciséis","diecisiete","dieciocho","diecinueve","veinte"]
DECENAS  = ["","","veinte","treinta","cuarenta","cincuenta","sesenta","setenta","ochenta","noventa"]
CENTENAS = ["","cien","doscientos","trescientos","cuatrocientos","quinientos","seiscientos","setecientos","ochocientos","novecientos"]

def _tens(n:int)->str:
    if n<=20: return UNIDADES[n]
    d,u = divmod(n,10)
    if u==0: return DECENAS[d]
    if d==2: return f"veinti{UNIDADES[u]}".replace("veintiuno","veintiún")
    return f"{DECENAS[d]} y {UNIDADES[u]}"

def _hundreds(n:int)->str:
    if n==0: return ""
    if n==100: return "cien"
    c,r = divmod(n,100)
    pref = "ciento" if c==1 else CENTENAS[c]
    return (pref + (f" {_tens(r)}" if r else "")).strip()

def numero_a_letras(n:int)->str:
    if n==0: return "cero"
    partes=[]
    millones, r = divmod(n, 1_000_000)
    miles, unidades = divmod(r, 1000)
    if millones: partes.append("un millón" if millones==1 else f"{_hundreds(millones)} millones")
    if miles:    partes.append("mil" if miles==1 else f"{_hundreds(miles)} mil")
    if unidades: partes.append(_hundreds(unidades))
    return " ".join(partes).replace("uno mil","un mil").replace("veintiun","veintiún")

def monto_en_letras(mx: float)->str:
    try:
        mx = float(mx)
    except:
        mx = 0.0
    pesos = int(mx)
    cents = int(round((mx - pesos)*100))
    ptxt = numero_a_letras(pesos).upper()
    return f"{ptxt} PESOS {cents:02d}/100 M.N."

# --- util: encuentra un valor por nombres candidatos / substrings ---
def pick_col(row: pd.Series, candidates, contains=None):
    for c in candidates:
        if c in row.index: return row.get(c)
        for col in row.index:
            if str(col).lower() == str(c).lower():
                return row.get(col)
    if contains:
        for col in row.index:
            for frag in contains:
                if frag.lower() in str(col).lower():
                    return row.get(col)
    return None

def parse_money(x) -> float:
    if x is None or x == "": return 0.0
    if isinstance(x, (int,float)): return float(x)
    s = str(x)
    s = s.replace("$","").replace(",","").replace("MXN","").replace("mn","").strip()
    # manejar 12.399,94 (punto miles, coma decimal)
    if s.count(",")==1 and s.count(".")>1:
        s = s.replace(".","").replace(",",".")
    try:
        return float(s or 0)
    except:
        return 0.0

# --------------------------------------------------
def row_to_context(row: pd.Series) -> dict:
    ctx = {col: row[col] for col in row.index}

    nombre  = pick_col(row, ["Nombre Cliente","Nombre"]) or ""
    suc_raw = pick_col(row, ["Sucursal","Municipio"]) or ""
    folio   = pick_col(row, ["Clave Solicitud","Folio"]) or ""

    # MONTO de pagaré (del Excel)
    monto_raw = pick_col(
        row,
        ["Monto Pagaré","Monto Pagare","Monto Pagaré "],
        contains=["monto pagar", "pagare", "pagaré"]
    ) or pick_col(row, ["CUOTA","Monto Autorizado","Monto","Importe","Crédito","Credito"], contains=["cuota","monto","importe","credito"])
    monto_val = parse_money(monto_raw)

    branch = detect_branch(suc_raw)

    ctx["Nombre"]            = str(nombre)
    ctx["Sucursal"]          = branch
    ctx["Municipio"]         = str(suc_raw)
    ctx["Folio"]             = str(folio)
    ctx["CUOTA"]             = float(monto_val)
    ctx["CUOTA_FORMAT"]      = f"{monto_val:,.2f}"
    ctx["CUOTA_LETRAS"]      = monto_en_letras(monto_val)
    ctx["DireccionSucursal"] = BRANCH_ADDRESSES.get(branch, str(suc_raw))
    ctx["FechaHoy"]          = fecha_hoy_es()

    return ctx

# ---------- Detectar bloques KGRUPAL en Excel ----------
def detectar_grupos_kgrupal(df: pd.DataFrame):
    grupos = []
    en_grupo = False
    inicio = None
    for pos, (_, row) in enumerate(df.iterrows()):
        prod = pick_col(row, ["Producto"]) or ""
        es_k = "KGRUPAL" in str(prod).upper()
        if es_k and not en_grupo:
            en_grupo = True
            inicio = pos
        elif (not es_k) and en_grupo:
            grupos.append((inicio, pos - 1))
            en_grupo = False
            inicio = None
    if en_grupo:
        grupos.append((inicio, len(df) - 1))
    return grupos

# ---------- Contexto grupo (para convenio) ----------
def crear_contexto_grupal(grupo_df: pd.DataFrame, datos_grupo: dict, montos_antecedentes=None) -> dict:
    montos_antecedentes = montos_antecedentes or {}
    total_pagare = 0.0
    total_antecedentes = 0.0
    integrantes = []

    for _, row in grupo_df.iterrows():
        nombre = pick_col(row, ["Nombre Cliente","Nombre"]) or ""
        folio  = pick_col(row, ["Clave Solicitud","Folio"]) or ""
        # Pagaré (Excel)
        monto_raw = pick_col(row, ["Monto Pagaré","Monto Pagare","Monto Pagaré "], contains=["monto pagar"]) \
                    or pick_col(row, ["CUOTA","Monto","Importe"], contains=["cuota","monto","importe"])
        monto_pagare = parse_money(monto_raw)
        # Antecedente (UI)
        monto_ant = float(montos_antecedentes.get(str(folio), 0.0))

        total_pagare += monto_pagare
        total_antecedentes += monto_ant

        integrantes.append({
            "Nombre": str(nombre),
            "Folio": str(folio),
            "Monto": monto_pagare,
            "Monto_FORMAT": f"{monto_pagare:,.2f}",
            "MontoAntecedente": monto_ant,
            "MontoAntecedente_FORMAT": f"{monto_ant:,.2f}",
        })

    suc_raw = pick_col(grupo_df.iloc[0], ["Sucursal","Municipio"]) or ""
    branch = detect_branch(suc_raw)
    lista_integrantes = ", ".join([i["Nombre"] for i in integrantes])

    ctx = {
        "GrupoNombre": datos_grupo.get("nombre_grupo", ""),
        "Integrantes": integrantes,
        "lista_integrantes": lista_integrantes,

        # Totales:
        "TotalGrupo": total_pagare,  # suma de pagarés
        "TotalGrupo_FORMAT": f"{total_pagare:,.2f}",
        "TotalGrupo_LETRAS": monto_en_letras(total_pagare),

        # Totales de ANTECEDENTES (para “lo firman por $... pesos”)
        "TotalAntecedentes": total_antecedentes,
        "TotalAntecedentes_FORMAT": f"{total_antecedentes:,.2f}",
        "TotalAntecedentes_LETRAS": monto_en_letras(total_antecedentes),

        # Fechas y comité
        "FechaHoy": fecha_hoy_es(),
        "FechaFirma": datos_grupo.get("fecha_firma", fecha_hoy_es()),
        "Presidenta": datos_grupo.get("presidenta", ""),
        "Secretaria": datos_grupo.get("secretaria", ""),
        "Tesorera": datos_grupo.get("tesorera", ""),
        # Sucursal
        "Sucursal": branch,
        "DireccionSucursal": BRANCH_ADDRESSES.get(branch, str(suc_raw)),
    }
    return ctx

# ---------- Renderers ----------
def letra_abc(idx):
    import string
    i = int(idx) - 1
    if 0 <= i < 26:
        return f"{string.ascii_lowercase[i]})"
    return f"{idx})"

def render_docx(path_tpl: Path, context: dict) -> bytes:
    tpl = DocxTemplate(str(path_tpl))
    try:
        tpl.jinja_env.filters['letra_abc'] = letra_abc
    except Exception:
        pass
    with tempfile.TemporaryDirectory() as td:
        out_path = Path(td) / "out.docx"
        tpl.render(dict(context))
        tpl.save(out_path)
        return out_path.read_bytes()

def render_convenio_con_imagenes(path_tpl: Path, context: dict,
                                 img_pagos_path=None,
                                 img_amort_path=None,
                                 img_control_path=None) -> bytes:
    tpl = DocxTemplate(str(path_tpl))
    try:
        tpl.jinja_env.filters['letra_abc'] = letra_abc
    except Exception:
        pass
    ctx = dict(context)  # no mutar el original
    # Inyectar imágenes si existen (la plantilla debe tener {{ imagen_* }})
    if img_pagos_path:
        ctx["imagen_tabla_pagos"] = InlineImage(tpl, img_pagos_path, width=Mm(160))
    if img_amort_path:
        ctx["imagen_tabla_amort"] = InlineImage(tpl, img_amort_path, width=Mm(160))
    if img_control_path:
        ctx["imagen_control_pagos"] = InlineImage(tpl, img_control_path, width=Mm(160))
    with tempfile.TemporaryDirectory() as td:
        out_path = Path(td) / "out.docx"
        tpl.render(ctx)
        tpl.save(out_path)
        return out_path.read_bytes()

# ---------- Plantillas ----------
TPL_DIR = Path(__file__).parent / "plantillas"
TEMPLATES = {
    # Ajusta los nombres de archivo a tus plantillas reales:
    "HUEHUETOCA": TPL_DIR / "PAGARE 2 HUEHUETOCAej.docx",
    "TECAMAC":    TPL_DIR / "PAGARE 2 TECAMACej.docx",
    "ZUMPANGO":   TPL_DIR / "PAGARE 2 ZUMPANGOej1.docx",
    "CONVENIO":   TPL_DIR / "CONVENIO_GRUPALejefinal10.docx",  # Cambia aquí al nombre exacto de tu plantilla
}
PAGARE_KEYS = ["HUEHUETOCA", "TECAMAC", "ZUMPANGO"]  # claves de plantillas de pagaré

# ---------- UI ----------
st.set_page_config(page_title="Generador de documentos", page_icon="📄", layout="wide")
st.title("📄 Generador de Pagarés y Convenios Grupales")

# Subida de Excel (opcional)
excel_file = st.file_uploader("Excel de entrada (.xlsx) (opcional)", type=["xlsx"], accept_multiple_files=False)

# Cargar DF si hay excel
df = pd.DataFrame()
if excel_file:
    with st.spinner("Leyendo Excel..."):
        df = read_excel(excel_file).fillna("").reset_index(drop=True)

# Pestañas SIEMPRE visibles
tab1, tab2 = st.tabs(["📄 PAGARÉS INDIVIDUALES", "👥 CONVENIO GRUPAL"])

# ============ TAB 1: Pagarés individuales ============
with tab1:
    st.subheader("Generar Pagarés Individuales")
    modo = st.radio("Origen de datos", ["Captura manual", "Desde Excel"], horizontal=True)

    # ---- CAPTURA MANUAL ----
    if "manual_pagares" not in st.session_state:
        st.session_state.manual_pagares = []

    if modo == "Captura manual":
        st.info("Captura uno o más pagarés manualmente. Puedes editar la lista antes de generar.")
        with st.form("form_manual_pagare", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                nombre_m = st.text_input("Nombre del Cliente *")
                folio_m = st.text_input("Folio / Clave Solicitud *")
                suc_sel = st.selectbox("Sucursal/Plantilla *", PAGARE_KEYS, index=0)
            with col2:
                municipio_m = st.text_input("Municipio / Sucursal (texto)")
                monto_m = st.number_input("Monto Pagaré *", min_value=0.0, step=100.0, format="%.2f")
                fecha_m = st.text_input("Fecha (por defecto hoy)", value=fecha_hoy_es())
            direccion_default = BRANCH_ADDRESSES.get(suc_sel, municipio_m or suc_sel)
            direccion_m = st.text_area("Dirección a usar en el pagaré ({{DireccionSucursal}}):", value=direccion_default)

            add = st.form_submit_button("➕ Añadir a la lista")
            if add:
                if not nombre_m or not folio_m:
                    st.warning("Completa al menos: Nombre y Folio.")
                else:
                    ctx = {
                        "Nombre": nombre_m,
                        "Folio": folio_m,
                        "Sucursal": suc_sel,                               # <- plantilla a usar
                        "Municipio": municipio_m or suc_sel,
                        "CUOTA": float(monto_m),
                        "CUOTA_FORMAT": f"{monto_m:,.2f}",
                        "CUOTA_LETRAS": monto_en_letras(float(monto_m)),
                        "DireccionSucursal": (direccion_m or direccion_default).strip(),
                        "FechaHoy": fecha_m or fecha_hoy_es(),
                    }
                    st.session_state.manual_pagares.append(ctx)
                    st.success("Añadido ✅")

        # Lista editable y generación
        if st.session_state.manual_pagares:
            ed_df = pd.DataFrame(st.session_state.manual_pagares)
            ed_df = st.data_editor(ed_df, hide_index=True, use_container_width=True)
            st.session_state.manual_pagares = ed_df.to_dict(orient="records")

            if st.button("🚀 Generar Pagarés (Manual)"):
                with st.spinner("Generando pagarés (manual)..."):
                    try:
                        tmp_root = Path(tempfile.mkdtemp(prefix="pagares_manual_"))
                        total, errors = 0, []

                        for i, ctx in enumerate(st.session_state.manual_pagares):
                            tpl_key = ctx.get("Sucursal", "HUEHUETOCA")
                            tpl_path = TEMPLATES.get(tpl_key)
                            if not tpl_path or not tpl_path.exists():
                                errors.append((i, f"No hay plantilla para '{tpl_key}'"))
                                continue
                            try:
                                docx_bytes = render_docx(tpl_path, ctx)
                                folder = tmp_root / safe_name(f"{ctx.get('Nombre','SIN_NOMBRE')}_{tpl_key}")
                                folder.mkdir(parents=True, exist_ok=True)
                                (folder / safe_name(f"{ctx.get('Folio','')}_{ctx.get('Nombre','')}_{tpl_key}.docx")).write_bytes(docx_bytes)
                                total += 1
                            except Exception as e:
                                errors.append((i, f"Error renderizando: {e}"))

                        # ZIP
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for root, _, files in os.walk(tmp_root):
                                for f in files:
                                    full_path = Path(root)/f
                                    rel = full_path.relative_to(tmp_root)
                                    zf.write(full_path, arcname=str(rel))
                        zip_buffer.seek(0)

                        st.success(f"✅ Generados {total} pagaré(s) manual(es)")
                        st.download_button("⬇️ Descargar ZIP", zip_buffer, "pagares_manual.zip", "application/zip")

                        if errors:
                            st.warning("Avisos:")
                            for idx, msg in errors[:200]:
                                st.text(f"{idx}: {msg}")
                    except Exception as e:
                        st.exception(e)

    # ---- DESDE EXCEL ----
    else:
        if df.empty:
            st.warning("Sube un Excel para generar desde archivo.")
            st.button("🚀 Generar Pagarés (Excel)", disabled=True)
        else:
            # Opciones avanzadas para forzar plantilla / dirección
            with st.expander("⚙️ Opciones avanzadas (plantilla/dirección)", expanded=False):
                plantilla_fallback = st.selectbox(
                    "Si la sucursal detectada no tiene plantilla, usar esta:",
                    PAGARE_KEYS,
                    index=0,
                    key="tpl_fallback"
                )
                forzar_plantilla_todos = st.checkbox(
                    "Usar SIEMPRE la plantilla seleccionada arriba (ignorar sucursal detectada)",
                    value=False,
                    key="force_tpl_all"
                )
                direccion_forzada = st.text_area(
                    "Dirección de sucursal a usar (dejar vacío para no forzar)",
                    value="",
                    key="dir_override"
                )
                direccion_aplicar_todos = st.checkbox(
                    "Aplicar esta dirección a TODOS los pagarés",
                    value=False,
                    key="force_addr_all"
                )

            if st.button("🚀 Generar Pagarés (Excel)"):
                with st.spinner("Generando pagarés desde Excel..."):
                    try:
                        tmp_root = Path(tempfile.mkdtemp(prefix="pagares_excel_"))
                        total, errors = 0, []
                        # Detectar bloques KGRUPAL para excluirlos
                        grupos_kgrupal = detectar_grupos_kgrupal(df)

                        for i, row in df.iterrows():
                            # Saltar filas que pertenecen a un bloque KGRUPAL
                            if any(start <= i <= end for (start, end) in grupos_kgrupal):
                                continue

                            ctx = row_to_context(row)

                            # 1) ¿Qué plantilla usar?
                            branch_detectada = ctx.get("Sucursal", "HUEHUETOCA")
                            tpl_key = branch_detectada
                            if forzar_plantilla_todos:
                                tpl_key = plantilla_fallback
                            elif not (TEMPLATES.get(branch_detectada) and TEMPLATES[branch_detectada].exists()):
                                tpl_key = plantilla_fallback

                            tpl_path = TEMPLATES.get(tpl_key)
                            if not tpl_path or not tpl_path.exists():
                                errors.append((ctx.get('Folio','?'), f"No hay plantilla para '{tpl_key}'"))
                                continue

                            # 2) ¿Qué dirección usar?
                            if direccion_aplicar_todos and direccion_forzada.strip():
                                ctx["DireccionSucursal"] = direccion_forzada.strip()
                            elif (tpl_key != branch_detectada) and direccion_forzada.strip() and not forzar_plantilla_todos:
                                # Solo caímos a fallback por falta de plantilla: usa la dirección escrita
                                ctx["DireccionSucursal"] = direccion_forzada.strip()
                            else:
                                # Asegura un valor por defecto razonable
                                ctx.setdefault(
                                    "DireccionSucursal",
                                    BRANCH_ADDRESSES.get(branch_detectada, ctx.get("Municipio", branch_detectada))
                                )

                            try:
                                docx_bytes = render_docx(tpl_path, ctx)
                                folder = tmp_root / safe_name(f"{ctx.get('Nombre','SIN')}_{tpl_key}")
                                folder.mkdir(parents=True, exist_ok=True)
                                (folder / safe_name(f"{ctx.get('Folio','')}_{ctx.get('Nombre','')}_{tpl_key}.docx")).write_bytes(docx_bytes)
                                total += 1
                            except Exception as e:
                                errors.append((ctx.get('Folio','?'), f"Render error: {e}"))

                        # ZIP
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for root, _, files in os.walk(tmp_root):
                                for f in files:
                                    full_path = Path(root)/f
                                    rel = full_path.relative_to(tmp_root)
                                    zf.write(full_path, arcname=str(rel))
                        zip_buffer.seek(0)

                        st.success(f"✅ Generados {total} pagaré(s) desde Excel")
                        st.download_button("⬇️ Descargar ZIP", zip_buffer, "pagares_excel.zip", "application/zip")

                        if errors:
                            st.warning("Avisos:")
                            for folio, msg in errors[:200]:
                                st.text(f"{folio}: {msg}")
                    except Exception as e:
                        st.exception(e)

# ============ TAB 2: Convenio GRUPAL ============
with tab2:
    st.subheader("Generar Convenio Grupal (KGRUPAL)")
    if df.empty:
        st.info("Sube un Excel para detectar grupos KGRUPAL. Si no tienes Excel, puedes generar un convenio manual mínimo.")
        # MODO MANUAL RÁPIDO (sin Excel): ingresar integrantes a mano
        if "manual_integrantes" not in st.session_state:
            st.session_state.manual_integrantes = []
        with st.form("form_integrantes_manual", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                nombre_i = st.text_input("Nombre integrante")
                folio_i  = st.text_input("Folio/ID")
            with col2:
                monto_i  = st.number_input("Monto Pagaré", min_value=0.0, step=100.0, format="%.2f")
                ant_i    = st.number_input("Monto ANTECEDENTE", min_value=0.0, step=100.0, format="%.2f")
            if st.form_submit_button("➕ Añadir integrante"):
                if nombre_i:
                    st.session_state.manual_integrantes.append({
                        "Nombre": nombre_i, "Folio": folio_i,
                        "Monto": float(monto_i), "MontoAntecedente": float(ant_i),
                        "Monto_FORMAT": f"{monto_i:,.2f}",
                        "MontoAntecedente_FORMAT": f"{ant_i:,.2f}",
                    })
                    st.success("Integrante añadido ✅")

        if st.session_state.manual_integrantes:
            ed_df = pd.DataFrame(st.session_state.manual_integrantes)
            ed_df = st.data_editor(ed_df, hide_index=True, use_container_width=True)
            st.session_state.manual_integrantes = ed_df.to_dict(orient="records")

        st.markdown("### Datos del convenio")
        nombre_grupo_m = st.text_input("Nombre del Grupo", value="Grupo Manual")
        presidenta_m   = st.text_input("Presidenta")
        secretaria_m   = st.text_input("Secretaria")
        tesorera_m     = st.text_input("Tesorera")
        fecha_firma_m  = st.text_input("Fecha de firma", value=fecha_hoy_es())

        st.markdown("### 📎 Imágenes (opcional)")
        up1 = st.file_uploader("Tabla de pagos (imagen)", type=["png","jpg","jpeg"], key="img_pagos_m")
        up2 = st.file_uploader("Tabla de amortización (imagen)", type=["png","jpg","jpeg"], key="img_amort_m")
        up3 = st.file_uploader("Control de pagos (imagen)", type=["png","jpg","jpeg"], key="img_control_m")

        if st.button("🚀 Generar Convenio (manual)"):
            # Contexto manual
            total_pagare = sum(i.get("Monto",0.0) for i in st.session_state.manual_integrantes)
            total_ante   = sum(i.get("MontoAntecedente",0.0) for i in st.session_state.manual_integrantes)
            ctx = {
                "GrupoNombre": nombre_grupo_m,
                "Integrantes": st.session_state.manual_integrantes,
                "lista_integrantes": ", ".join([i["Nombre"] for i in st.session_state.manual_integrantes]),

                "TotalGrupo": total_pagare,
                "TotalGrupo_FORMAT": f"{total_pagare:,.2f}",
                "TotalGrupo_LETRAS": monto_en_letras(total_pagare),

                "TotalAntecedentes": total_ante,
                "TotalAntecedentes_FORMAT": f"{total_ante:,.2f}",
                "TotalAntecedentes_LETRAS": monto_en_letras(total_ante),

                "FechaHoy": fecha_hoy_es(),
                "FechaFirma": fecha_firma_m,
                "Presidenta": presidenta_m,
                "Secretaria": secretaria_m,
                "Tesorera": tesorera_m,
            }
            tpl_path = TEMPLATES.get("CONVENIO")
            if not tpl_path or not tpl_path.exists():
                st.error("❌ No se encontró la plantilla de convenio.")
            else:
                with tempfile.TemporaryDirectory() as td_imgs:
                    p1=p2=p3=None
                    if up1: p1 = str(Path(td_imgs)/("pagos"+Path(up1.name).suffix)); Path(p1).write_bytes(up1.getvalue())
                    if up2: p2 = str(Path(td_imgs)/("amort"+Path(up2.name).suffix)); Path(p2).write_bytes(up2.getvalue())
                    if up3: p3 = str(Path(td_imgs)/("control"+Path(up3.name).suffix)); Path(p3).write_bytes(up3.getvalue())
                    docx_bytes = render_convenio_con_imagenes(tpl_path, ctx, p1, p2, p3)
                st.success("✅ Convenio generado")
                st.download_button(
                    "⬇️ Descargar Convenio",
                    data=docx_bytes,
                    file_name=f"CONVENIO_{safe_name(nombre_grupo_m)}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

    else:
        # Con Excel: detectar grupos KGRUPAL y permitir seleccionar uno
        grupos_kgrupal = detectar_grupos_kgrupal(df)
        if not grupos_kgrupal:
            st.info("ℹ️ No se detectaron bloques KGRUPAL en el Excel.")
        else:
            opciones_grupos = [f"Grupo {i+1} (Filas {start+1}-{end+1})" for i,(start,end) in enumerate(grupos_kgrupal)]
            grupo_sel = st.selectbox("Selecciona un grupo KGRUPAL", opciones_grupos)
            gidx = opciones_grupos.index(grupo_sel)
            start, end = grupos_kgrupal[gidx]
            grupo_df = df.iloc[start:end+1]

            # Mostrar integrantes base
            st.markdown("### 👥 Integrantes del grupo seleccionado")
            integrantes_data = []
            for _, row in grupo_df.iterrows():
                nombre = pick_col(row, ["Nombre Cliente","Nombre"]) or "SIN NOMBRE"
                folio  = pick_col(row, ["Clave Solicitud","Folio"]) or "SIN FOLIO"
                monto_raw = pick_col(row, ["Monto Pagaré","Monto Pagare"], contains=["monto pagar"]) or pick_col(row, ["CUOTA","Monto"])
                monto = parse_money(monto_raw)
                integrantes_data.append({"Nombre": nombre, "Folio": folio, "Monto": monto})
            st.dataframe(pd.DataFrame(integrantes_data))

            # Editor de antecedentes por integrante
            ed_df = pd.DataFrame(integrantes_data)
            if "MontoAntecedente" not in ed_df.columns:
                ed_df["MontoAntecedente"] = 0.0
            ed_df = st.data_editor(
                ed_df,
                key=f"editor_antecedentes_{gidx}",
                num_rows="fixed",
                column_config={
                    "Monto": st.column_config.NumberColumn("Monto Pagaré (Excel)", format="%.2f", step=100.0, disabled=True),
                    "MontoAntecedente": st.column_config.NumberColumn("Monto ANTECEDENTE (captura)", format="%.2f", step=100.0),
                },
                hide_index=True,
            )
            montos_ant_por_folio = {str(r["Folio"]): float(r.get("MontoAntecedente") or 0) for _, r in ed_df.iterrows()}

            # Datos del convenio
            st.subheader("📋 Datos del Convenio")
            col1, col2 = st.columns(2)
            with col1:
                nombre_grupo = st.text_input("Nombre del Grupo", value=f"Grupo_{gidx+1}", key=f"nombre_{gidx}")
                presidenta   = st.text_input("Presidenta del Grupo", key=f"presidenta_{gidx}")
            with col2:
                secretaria   = st.text_input("Secretaria del Grupo", key=f"secretaria_{gidx}")
                tesorera     = st.text_input("Tesorera del Grupo", key=f"tesorera_{gidx}")
            fecha_firma = st.text_input("Fecha de firma del convenio", value=fecha_hoy_es(), key=f"fecha_{gidx}")

            # Imágenes anexas
            st.markdown("### 📎 Archivos adicionales (imágenes)")
            tabla_pagos = st.file_uploader("Tabla de pagos concentrada (imagen)", type=["png","jpg","jpeg"], key=f"pagos_{gidx}")
            tabla_amort = st.file_uploader("Tabla de amortización (imagen)", type=["png","jpg","jpeg"], key=f"amort_{gidx}")
            control_pagos = st.file_uploader("Control de pagos (imagen)", type=["png","jpg","jpeg"], key=f"control_{gidx}")

            # Generar convenio del grupo seleccionado
            if st.button("🚀 Generar Convenio para este Grupo", key=f"btn_grupo_{gidx}"):
                with st.spinner("Generando convenio grupal..."):
                    try:
                        datos_grupo = {
                            'nombre_grupo': nombre_grupo,
                            'presidenta': presidenta,
                            'secretaria': secretaria,
                            'tesorera': tesorera,
                            'fecha_firma': fecha_firma
                        }
                        ctx_grupal = crear_contexto_grupal(grupo_df, datos_grupo, montos_ant_por_folio)

                        tpl_path = TEMPLATES.get("CONVENIO")
                        if not tpl_path or not tpl_path.exists():
                            st.error("❌ No se encontró la plantilla para convenios grupales.")
                        else:
                            with tempfile.TemporaryDirectory() as td_imgs:
                                p1=p2=p3=None
                                if tabla_pagos is not None:
                                    p1 = str(Path(td_imgs)/("pagos"+Path(tabla_pagos.name).suffix)); Path(p1).write_bytes(tabla_pagos.getvalue())
                                if tabla_amort is not None:
                                    p2 = str(Path(td_imgs)/("amort"+Path(tabla_amort.name).suffix)); Path(p2).write_bytes(tabla_amort.getvalue())
                                if control_pagos is not None:
                                    p3 = str(Path(td_imgs)/("control"+Path(control_pagos.name).suffix)); Path(p3).write_bytes(control_pagos.getvalue())

                                docx_bytes = render_convenio_con_imagenes(
                                    tpl_path, ctx_grupal,
                                    img_pagos_path=p1,
                                    img_amort_path=p2,
                                    img_control_path=p3,
                                )

                            # Crear ZIP con el DOCX (y opcionalmente las imágenes sueltas)
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                                zf.writestr(safe_name(f"CONVENIO_{nombre_grupo}.docx"), docx_bytes)
                                if tabla_pagos is not None:
                                    zf.writestr("TABLA_PAGOS"+Path(tabla_pagos.name).suffix, tabla_pagos.getvalue())
                                if tabla_amort is not None:
                                    zf.writestr("TABLA_AMORTIZACION"+Path(tabla_amort.name).suffix, tabla_amort.getvalue())
                                if control_pagos is not None:
                                    zf.writestr("CONTROL_PAGOS"+Path(control_pagos.name).suffix, control_pagos.getvalue())
                            zip_buffer.seek(0)

                            st.success(f"✅ Convenio generado para {grupo_sel}")
                            st.download_button(
                                label=f"⬇️ Descargar Convenio {nombre_grupo}",
                                data=zip_buffer,
                                file_name=f"convenio_{safe_name(nombre_grupo)}.zip",
                                mime="application/zip",
                                key=f"dl_grupo_{gidx}"
                            )

                            # Resumen
                            st.subheader("📊 Resumen del Convenio")
                            colA, colB, colC = st.columns(3)
                            with colA:
                                st.metric("Total ANTECEDENTES", f"${ctx_grupal['TotalAntecedentes_FORMAT']}")
                            with colB:
                                st.metric("Total Pagarés", f"${ctx_grupal['TotalGrupo_FORMAT']}")
                            with colC:
                                st.metric("Integrantes", len(integrantes_data))
                    except Exception as e:
                        st.exception(e)

# Vista previa del Excel (si hay)
if not df.empty:
    st.subheader("📋 Vista previa del Excel")
    st.dataframe(df.head(10).astype(str))


