"""
Generador de informes de evaluación docente — EAFIT
Versión web (Streamlit)
"""

import io, os, re, zipfile, random, tempfile
from collections import defaultdict
from copy import deepcopy
from lxml import etree
import openpyxl
import streamlit as st

# ─── CONFIGURACIÓN EXCEL ───────────────────────────────────────────────────
FILA_ENCABEZADO = 7
FILA_DATOS      = 8

COL_NOMBRE          = "Nombres y apellidos Docente"
COL_CICLO           = "Ciclo"
COL_CURSO           = "Nombre Catalogo"
COL_ID_ESCUELA      = "Id Escuela"
COL_ESCUELA         = "Escuela"
COL_COMPETENCIA     = "Competencia Evaluada"
COL_NOTA_FINAL      = "Nota final por clase"
COL_NOTA_CURSO      = "Nota final por curso"
COL_PREGUNTA        = "Pregunta"
COL_COMENTARIO      = "Comentarios"
COL_TOTAL_GENERADAS = "Total Evaluaciones generadas"
COL_EVALUACIONES    = "Evaluaciones realizadas"

ESCUELAS_EAFIT = {
    "E-ADM": "Escuela de Administración",
    "E-DER": "Escuela de Derecho",
    "E-ECO": "Escuela de Economía y Finanzas",
    "E-HUM": "Escuela de Humanidades",
    "E-ING": "Escuela de Ciencias Aplicadas e Ingeniería",
    "E-MED": "Escuela de Medicina",
    "E-MUS": "Escuela de Música",
    "E-DIS": "Escuela de Arquitectura y Diseño",
    "E-CS":  "Escuela de Ciencias",
    "E-VIS": "Vicerrectoría de Internacionalización",
}

FILTROS_COMENTARIOS = [
    r"^\s*(no|na|n\.a\.?|n/a)\s*$",
    r"^\s*(ninguno?|ninguna?|nada)\s*$",
    r"^\s*(ok|oki|okay)\s*$",
    r"^\s*(todo\s+(bien|esta\s+bien|está\s+bien))\s*$",
    r"^\s*(bien|excelente)\s*$",
    r"^\s*(gracias?)\s*$",
    r"^\s*(no\s+aplica|no\s+apply|n\.?a\.?)\s*$",
    r"^\s*(cumple|cumplido)\s*$",
    r"^\s*s[ií]n?\s+comentarios?\s*$",
    r"^\s*(si|sí)\s*$",
    r"^\s*(nunguno?|nung[uo]no?)\s*$",
    r"^\s*-\s*$",
    r"^\s*$",
]

MAYUSCULAS_FIJAS = {"eafit", "covid", "ia", "ti", "zoom", "teams", "meet"}

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# ─── HELPERS ───────────────────────────────────────────────────────────────

def resolver_escuela(id_escuela: str, escuela_raw: str) -> str:
    key = str(id_escuela or "").strip().upper()
    return ESCUELAS_EAFIT.get(key, str(escuela_raw or "").strip())

def fmt_nota(n) -> str:
    if n is None: return "—"
    try: return f"{float(n):.2f}".replace(".", ",")
    except: return str(n)

def es_valido(texto) -> bool:
    if not texto: return False
    t = str(texto).strip()
    for p in FILTROS_COMENTARIOS:
        if re.match(p, t, re.IGNORECASE): return False
    return len(t) >= 4

def formatear_comentario(texto: str) -> str:
    texto = str(texto).strip()
    if not texto: return texto
    palabras = texto.split()
    resultado = []
    for palabra in palabras:
        if len(palabra) > 1 and palabra.isupper() and palabra.isalpha():
            resultado.append(palabra)
        elif palabra.lower().rstrip(".,;:") in MAYUSCULAS_FIJAS:
            resultado.append(palabra.upper())
        else:
            resultado.append(palabra.lower())
    if resultado:
        primera = resultado[0]
        resultado[0] = primera[0].upper() + primera[1:] if primera else primera
    texto_f = " ".join(resultado)
    if texto_f and texto_f[-1] not in ".!?;:":
        texto_f += "."
    return texto_f

def slugify(texto: str) -> str:
    repl = {'á':'a','é':'e','í':'i','ó':'o','ú':'u','ü':'u','ñ':'n',
            'Á':'A','É':'E','Í':'I','Ó':'O','Ú':'U','Ü':'U','Ñ':'N'}
    for orig, rep in repl.items():
        texto = texto.replace(orig, rep)
    texto = re.sub(r'[^\w\s\-]', '', texto)
    return re.sub(r'\s+', '', texto)

def nombre_archivo_defecto(datos: dict, nombre_prof: str) -> str:
    info       = datos.get("info", {})
    ciclo      = str(info.get("ciclo", "")).strip()
    escuela_id = str(info.get("id_escuela", "")).replace("-", "")
    curso      = slugify(info.get("curso", ""))
    partes     = nombre_prof.strip().split()
    if len(partes) >= 2:
        mitad          = len(partes) // 2
        primer_nombre  = partes[mitad].capitalize()
        primer_apellido= partes[0].capitalize()
        nombre_corto   = primer_nombre + primer_apellido
    else:
        nombre_corto = slugify(nombre_prof)
    return f"{ciclo}_{escuela_id}_{curso}_{nombre_corto}"

def replace_in_subtree(elem, old: str, new: str):
    for t in elem.iter(W+'t'):
        if t.text and old in t.text:
            t.text = t.text.replace(old, new)

def clone_bullet_para(template_para):
    new_p = deepcopy(template_para)
    for attr in list(new_p.attrib):
        if 'paraId' in attr or 'textId' in attr:
            new_p.set(attr, f"{random.randint(0x10000000, 0xFFFFFFFF):08X}")
    return new_p

# ─── LECTURA EXCEL ─────────────────────────────────────────────────────────

def leer_excel(archivo_bytes: bytes) -> dict:
    wb = openpyxl.load_workbook(io.BytesIO(archivo_bytes), read_only=True)
    ws = wb.active

    headers = {}
    for row in ws.iter_rows(min_row=FILA_ENCABEZADO, max_row=FILA_ENCABEZADO, values_only=True):
        for i, val in enumerate(row):
            if val: headers[str(val).strip()] = i
        break

    def col(row_data, nombre):
        idx = headers.get(nombre)
        return row_data[idx] if idx is not None and idx < len(row_data) else None

    profesores = defaultdict(lambda: {
        "info": {}, "nota_final": None, "nota_curso": None,
        "comentarios": defaultdict(list),
        "total_generadas": None, "evaluaciones_realizadas": None,
    })

    for row in ws.iter_rows(min_row=FILA_DATOS, values_only=True):
        nombre = col(row, COL_NOMBRE)
        if not nombre: continue
        nombre = str(nombre).strip()
        p = profesores[nombre]

        if not p["info"]:
            id_esc  = str(col(row, COL_ID_ESCUELA) or "").strip()
            esc_raw = str(col(row, COL_ESCUELA) or "").strip()
            p["info"] = {
                "ciclo":      str(col(row, COL_CICLO) or "").strip(),
                "curso":      str(col(row, COL_CURSO) or "").strip(),
                "id_escuela": id_esc,
                "escuela":    resolver_escuela(id_esc, esc_raw),
            }

        comp      = str(col(row, COL_COMPETENCIA) or "").strip()
        pregunta  = str(col(row, COL_PREGUNTA) or "").strip()
        nota_final= col(row, COL_NOTA_FINAL)
        nota_curso= col(row, COL_NOTA_CURSO)
        comentario= col(row, COL_COMENTARIO)

        total_gen = col(row, COL_TOTAL_GENERADAS)
        eval_real = col(row, COL_EVALUACIONES)
        if total_gen and p["total_generadas"] is None:
            try: p["total_generadas"] = int(float(total_gen))
            except: pass
        if eval_real and p["evaluaciones_realizadas"] is None:
            try: p["evaluaciones_realizadas"] = int(float(eval_real))
            except: pass

        if nota_final and nota_final > 0 and p["nota_final"] is None:
            p["nota_final"] = nota_final
        if nota_curso and nota_curso > 0 and p["nota_curso"] is None:
            p["nota_curso"] = nota_curso

        if comp == "Comentarios" and pregunta and comentario:
            if es_valido(str(comentario)):
                p["comentarios"][pregunta].append(formatear_comentario(str(comentario)))

    return dict(profesores)

# ─── GENERACIÓN WORD EN MEMORIA ────────────────────────────────────────────

def generar_informe_bytes(nombre: str, datos: dict,
                          plantilla_bytes: bytes,
                          nombre_archivo: str = None) -> tuple[bytes, str]:
    """Devuelve (docx_bytes, nombre_archivo)."""
    info                    = datos["info"]
    nota_curso              = datos["nota_curso"]
    nota_final              = datos["nota_final"]
    comentarios             = datos["comentarios"]
    total_generadas         = datos.get("total_generadas")
    evaluaciones_realizadas = datos.get("evaluaciones_realizadas")
    nota_tabla              = nota_curso if nota_curso else nota_final

    if not nombre_archivo:
        nombre_archivo = nombre_archivo_defecto(datos, nombre)
    nombre_archivo = re.sub(r'[<>:"/\\|?*]', '_', nombre_archivo).strip()

    # Trabajar en memoria
    tree = etree.fromstring(
        zipfile.ZipFile(io.BytesIO(plantilla_bytes)).read('word/document.xml')
    )
    body          = tree.find(W+'body')
    body_children = list(body)

    # ── Portada ──
    p0 = body_children[0]
    replace_in_subtree(p0, 'Nombre del curso',     info['curso'])
    replace_in_subtree(p0, 'Escuela del profesor', info['escuela'])
    replace_in_subtree(p0, 'numero',               info['ciclo'])
    replace_in_subtree(p0, ' del semestre',        '')
    replace_in_subtree(p0, 'NOMBRE PROFESOR',      nombre)

    # ── Tabla competencias ──
    tbl_comp = body_children[3]
    for t in tbl_comp.iter(W+'t'):
        if t.text and ('final_course_note' in t.text or t.text.strip() in ('{{', '}}', '{', '}')):
            t.text = ''
    nota_str  = fmt_nota(nota_tabla)
    data_rows = tbl_comp.findall('.//' + W+'tr')[1:]
    for row in data_rows:
        cells = row.findall(W+'tc')
        if len(cells) >= 2:
            nota_cell = cells[1]
            for t in nota_cell.iter(W+'t'):
                t.text = ''
            paras = nota_cell.findall('.//' + W+'p')
            if paras:
                runs = paras[0].findall('.//' + W+'r')
                if runs:
                    t_elems = runs[0].findall(W+'t')
                    if t_elems:
                        t_elems[0].text = nota_str
                    else:
                        etree.SubElement(runs[0], W+'t').text = nota_str

    # ── Tabla estudiantes ──
    tbl_est = body_children[5]
    for t in tbl_est.iter(W+'t'):
        if t.text:
            if 'number_students_answered' in t.text:
                t.text = str(evaluaciones_realizadas) if evaluaciones_realizadas is not None else '—'
            elif 'total_number_students' in t.text:
                t.text = str(total_generadas) if total_generadas is not None else '—'
            elif t.text.strip() in ('{{', '}}'):
                t.text = ''

    # ── Comentarios ──
    def get_comentarios(clave):
        for k, lista in comentarios.items():
            if clave in k.lower(): return lista
        return []

    for clave, start_idx in [("positivo", 12), ("mejorar", 21), ("adicional", 29)]:
        placeholders = [body_children[start_idx + j] for j in range(5)
                        if start_idx + j < len(body_children)]
        if not placeholders: continue
        template_para = placeholders[0]
        insert_pos    = list(body).index(placeholders[0])
        for pp in placeholders:
            body.remove(pp)
        for j, texto in enumerate(get_comentarios(clave)):
            new_p  = clone_bullet_para(template_para)
            all_ts = list(new_p.iter(W+'t'))
            if all_ts:
                all_ts[0].text = texto
                for t in all_ts[1:]: t.text = ''
            body.insert(insert_pos + j, new_p)

    # ── Empaquetar de vuelta en memoria ──
    new_xml = etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
    out_buf  = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(plantilla_bytes), 'r') as zin:
        with zipfile.ZipFile(out_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                zout.writestr(item, new_xml if item.filename == 'word/document.xml'
                              else zin.read(item.filename))
    return out_buf.getvalue(), nombre_archivo

# ─── INTERFAZ STREAMLIT ────────────────────────────────────────────────────

st.set_page_config(
    page_title="Generador de Informes — EAFIT",
    page_icon="📋",
    layout="centered",
)

# ── Estilos ──
st.markdown("""
<style>
    /* Fondo principal */
    .stApp { background-color: #1e1e2e; color: #ececf1; }

    /* Franja roja superior */
    .header-bar {
        background: #c0392b;
        height: 6px;
        border-radius: 2px;
        margin-bottom: 1.2rem;
    }

    /* Título */
    .titulo { font-size: 1.6rem; font-weight: 700; color: #ececf1; margin-bottom: 0; }
    .subtitulo { font-size: 0.9rem; color: #7f8c9a; margin-top: 2px; margin-bottom: 1.5rem; }

    /* Tarjetas de sección */
    .card {
        background: #2a2a3d;
        border-radius: 10px;
        padding: 1.4rem 1.6rem;
        margin-bottom: 1.2rem;
    }
    .card-title {
        font-size: 0.78rem;
        font-weight: 600;
        color: #7f8c9a;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        margin-bottom: 0.8rem;
    }

    /* Inputs */
    .stTextInput > div > div > input {
        background: #12121e !important;
        color: #ececf1 !important;
        border: 1px solid #3a3a55 !important;
        border-radius: 6px !important;
    }

    /* File uploader */
    .stFileUploader > div {
        background: #12121e !important;
        border: 1.5px dashed #3a3a55 !important;
        border-radius: 8px !important;
        color: #7f8c9a !important;
    }

    /* Botón primario */
    .stButton > button {
        background: #c0392b !important;
        color: white !important;
        border: none !important;
        border-radius: 7px !important;
        font-weight: 600 !important;
        padding: 0.55rem 1.6rem !important;
        font-size: 1rem !important;
        transition: background 0.2s;
    }
    .stButton > button:hover {
        background: #e74c3c !important;
    }

    /* Botón de descarga */
    .stDownloadButton > button {
        background: #27ae60 !important;
        color: white !important;
        border: none !important;
        border-radius: 7px !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
        padding: 0.55rem 1.6rem !important;
    }
    .stDownloadButton > button:hover { background: #2ecc71 !important; }

    /* Info / success / warning */
    .stAlert { border-radius: 8px !important; }

    /* Tabla de preview */
    .preview-table { width: 100%; border-collapse: collapse; font-size: 0.85rem; }
    .preview-table th {
        background: #1e1e2e; color: #7f8c9a;
        padding: 6px 10px; text-align: left;
        font-weight: 600; font-size: 0.75rem;
        text-transform: uppercase; letter-spacing: 0.06em;
    }
    .preview-table td { padding: 7px 10px; color: #ececf1; border-bottom: 1px solid #1e1e2e; }
    .preview-table tr:last-child td { border-bottom: none; }

    /* Ocultar menú de Streamlit */
    #MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Encabezado ──
st.markdown('<div class="header-bar"></div>', unsafe_allow_html=True)
st.markdown('<div class="titulo">Generador de informes docentes</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitulo">Universidad EAFIT &nbsp;·&nbsp; Evaluación docente</div>',
            unsafe_allow_html=True)

# ── Sección 1: Archivos ──
st.markdown('<div class="card"><div class="card-title">📂 Archivos</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    archivo_excel = st.file_uploader(
        "Archivo Excel de evaluaciones",
        type=["xlsx", "xls"],
        help="El archivo exportado del sistema de evaluación docente"
    )
with col2:
    archivo_plantilla = st.file_uploader(
        "Plantilla Word (.docx)",
        type=["docx"],
        help="La plantilla base para los informes"
    )

st.markdown('</div>', unsafe_allow_html=True)

# ── Preview automático al subir Excel ──
profesores = {}
if archivo_excel:
    try:
        with st.spinner("Leyendo Excel…"):
            profesores = leer_excel(archivo_excel.getvalue())

        st.markdown('<div class="card"><div class="card-title">👥 Profesores encontrados</div>',
                    unsafe_allow_html=True)

        filas_html = ""
        for nombre, datos in profesores.items():
            info = datos["info"]
            nf   = nombre_archivo_defecto(datos, nombre)
            filas_html += f"""
            <tr>
              <td><strong>{nombre.title()}</strong></td>
              <td>{info.get('curso','—')}</td>
              <td>{info.get('escuela','—')}</td>
              <td>{info.get('ciclo','—')}</td>
              <td style="color:#7f8c9a;font-size:0.78rem">{nf}.docx</td>
            </tr>"""

        st.markdown(f"""
        <table class="preview-table">
          <thead>
            <tr>
              <th>Profesor</th><th>Curso</th><th>Escuela</th>
              <th>Semestre</th><th>Nombre del archivo</th>
            </tr>
          </thead>
          <tbody>{filas_html}</tbody>
        </table>
        """, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    except Exception as e:
        st.error(f"No se pudo leer el Excel: {e}")

# ── Sección 2: Generar ──
if profesores and archivo_plantilla:
    st.markdown('<div class="card"><div class="card-title">⚙️ Generar informes</div>',
                unsafe_allow_html=True)

    total = len(profesores)
    es_uno = total == 1

    nombre_custom = None
    if es_uno:
        nombre_prof, datos_prof = next(iter(profesores.items()))
        defecto = nombre_archivo_defecto(datos_prof, nombre_prof)
        nombre_custom = st.text_input(
            "Nombre del archivo (editable)",
            value=defecto,
            help="Puedes cambiar el nombre antes de generar"
        )

    st.markdown('</div>', unsafe_allow_html=True)

    if st.button(f"Generar {'informe' if es_uno else str(total) + ' informes'}"):
        plantilla_bytes = archivo_plantilla.getvalue()
        errores = []
        archivos_generados = {}  # nombre → bytes

        barra = st.progress(0, text="Generando informes…")
        for i, (nombre, datos) in enumerate(profesores.items()):
            try:
                nf = nombre_custom if (es_uno and nombre_custom) else None
                docx_bytes, nombre_arch = generar_informe_bytes(
                    nombre, datos, plantilla_bytes, nombre_archivo=nf
                )
                archivos_generados[nombre_arch + ".docx"] = docx_bytes
            except Exception as e:
                errores.append(f"{nombre}: {e}")
            barra.progress((i + 1) / total,
                           text=f"Procesando {i+1} de {total}…")

        barra.empty()

        if errores:
            for err in errores:
                st.error(f"❌ {err}")

        if archivos_generados:
            if len(archivos_generados) == 1:
                nombre_arch, docx_bytes = next(iter(archivos_generados.items()))
                st.success(f"✅ Informe generado: **{nombre_arch}**")
                st.download_button(
                    label="⬇️  Descargar informe",
                    data=docx_bytes,
                    file_name=nombre_arch,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            else:
                # Empaquetar todos en un ZIP
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for nombre_arch, docx_bytes in archivos_generados.items():
                        zf.writestr(nombre_arch, docx_bytes)

                st.success(f"✅ {len(archivos_generados)} informes generados correctamente.")
                st.download_button(
                    label="⬇️  Descargar todos los informes (.zip)",
                    data=zip_buf.getvalue(),
                    file_name="Informes_EAFIT.zip",
                    mime="application/zip",
                )

elif archivo_excel and not archivo_plantilla:
    st.info("📄 Ahora sube la **Plantilla Word** para poder generar los informes.")
elif not archivo_excel:
    st.info("📊 Sube el **archivo Excel** para comenzar.")
