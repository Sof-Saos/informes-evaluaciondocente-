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
    # Respuestas vacías o sin contenido
    r"^\s*$",
    r"^\s*-+\s*$",
    r"^\s*\.+\s*$",
    # Afirmaciones/negaciones sin contenido
    r"^\s*(no|na|n\.a\.?|n/a|nada|ninguno?|ninguna?)\s*$",
    r"^\s*(nunguno?|nung[uo]no?)\s*$",
    r"^\s*(si|sí|yes)\s*$",
    r"^\s*(ok|oki|okay|okey)\s*$",
    # Frases cortas irrelevantes
    r"^\s*(todo\s+bien|todo\s+está?\s*bien|todo\s+esta\s*bien)\s*$",
    r"^\s*(bien|muy\s+bien|excelente|perfecto)\s*$",
    r"^\s*(gracias?|thanks?)\s*$",
    r"^\s*(no\s+aplica|no\s+apply|n\.?a\.?)\s*$",
    r"^\s*(cumple|cumplido|cumple\s+con\s+todo)\s*$",
    r"^\s*s[ií]n?\s+comentarios?\s*$",
    r"^\s*s[ií]n?\s+novedad(es)?\s*$",
    r"^\s*(ningún?\s+comentario|no\s+tengo\s+comentarios?)\s*$",
    r"^\s*(no\s+hay\s+comentarios?|sin\s+observaciones?)\s*$",
    r"^\s*(ningun[ao])\s*$",
]

MAYUSCULAS_FIJAS = {"eafit", "covid", "ia", "ti", "zoom", "teams", "meet", "canvas", "moodle"}

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# ─── LANGUAGETOOL ──────────────────────────────────────────────────────────

import urllib.request
import urllib.parse
import json
import time

_LT_CACHE: dict = {}   # cache para no llamar la API dos veces con el mismo texto
_LT_URL = "https://api.languagetool.org/v2/check"
_LT_DISABLED = False   # se pone True si la API falla (fallback silencioso)

def _languagetool_check(texto: str) -> list:
    """Llama a la API pública de LanguageTool y devuelve la lista de matches."""
    global _LT_DISABLED
    if _LT_DISABLED:
        return []
    if texto in _LT_CACHE:
        return _LT_CACHE[texto]
    try:
        data = urllib.parse.urlencode({
            "language": "es",
            "text": texto,
            "enabledOnly": "false",
        }).encode("utf-8")
        req = urllib.request.Request(
            _LT_URL, data=data,
            headers={"Content-Type": "application/x-www-form-urlencoded",
                     "User-Agent": "EXA-InformeDocente/1.0"},
        )
        with urllib.request.urlopen(req, timeout=6) as resp:
            result = json.loads(resp.read().decode("utf-8"))
        matches = result.get("matches", [])
        _LT_CACHE[texto] = matches
        return matches
    except Exception:
        _LT_DISABLED = True   # si falla (sin internet, rate limit, etc.) no reintentar
        return []

def _aplicar_correcciones_lt(texto: str) -> str:
    """Aplica las sugerencias de LanguageTool al texto (primera sugerencia por match)."""
    matches = _languagetool_check(texto)
    if not matches:
        return texto
    # Ordenar de atrás hacia adelante para no desplazar offsets
    matches_sorted = sorted(matches, key=lambda m: m["offset"], reverse=True)
    resultado = texto
    for m in matches_sorted:
        replacements = m.get("replacements", [])
        if not replacements:
            continue
        # Solo aplicar correcciones de ortografía y tipografía, NO de estilo
        rule_id = m.get("rule", {}).get("id", "")
        issue_type = m.get("rule", {}).get("issueType", "")
        if issue_type not in ("misspelling", "typographical"):
            continue
        mejor = replacements[0]["value"]
        offset = m["offset"]
        length = m["length"]
        resultado = resultado[:offset] + mejor + resultado[offset + length:]
    return resultado

# ─── HELPERS ───────────────────────────────────────────────────────────────

def resolver_escuela(id_escuela: str, escuela_raw: str) -> str:
    key = str(id_escuela or "").strip().upper()
    return ESCUELAS_EAFIT.get(key, str(escuela_raw or "").strip())

def fmt_nota(n) -> str:
    if n is None: return "—"
    try: return f"{float(n):.2f}".replace(".", ",")
    except: return str(n)

def es_valido(texto) -> bool:
    """Filtra comentarios sin contenido útil."""
    if not texto: return False
    t = str(texto).strip()
    for p in FILTROS_COMENTARIOS:
        if re.match(p, t, re.IGNORECASE): return False
    # Demasiado corto para ser útil
    if len(t) < 5:
        return False
    return True

def _sentence_case(texto: str) -> str:
    """
    Convierte el texto a sentence case inteligente:
    - Si todo (o casi todo) está en mayúsculas → convierte a minúsculas primero
    - Primera letra de la oración en mayúscula
    - Respeta siglas conocidas (EAFIT, COVID, etc.)
    - Respeta palabras con mayúscula interna (iPad, etc.) — no las toca
    """
    if not texto:
        return texto

    # Detectar si el texto está "gritado" (>55% letras en mayúscula)
    letras = [c for c in texto if c.isalpha()]
    if letras and sum(1 for c in letras if c.isupper()) / len(letras) > 0.55:
        texto = texto.lower()

    # Procesar palabra por palabra
    palabras = texto.split()
    resultado = []
    for i, palabra in enumerate(palabras):
        base = palabra.rstrip(".,;:!?()")
        sufijo = palabra[len(base):]
        base_lower = base.lower()

        if base_lower in MAYUSCULAS_FIJAS:
            resultado.append(base.upper() + sufijo)
        elif len(base) > 1 and base.isupper():
            # Sigla desconocida → dejar en mayúsculas
            resultado.append(palabra)
        else:
            resultado.append(base_lower + sufijo)

    if resultado:
        p = resultado[0]
        resultado[0] = p[0].upper() + p[1:] if p else p

    texto_f = " ".join(resultado)

    # Asegurar punto al final
    if texto_f and texto_f[-1] not in ".!?;:":
        texto_f += "."

    return texto_f

def formatear_comentario(texto: str) -> str:
    """Pipeline completo: sentence case → corrección ortográfica (LanguageTool)."""
    texto = str(texto).strip()
    if not texto:
        return texto
    # 1. Sentence case
    texto = _sentence_case(texto)
    # 2. Corrección ortográfica con LanguageTool (gratuito, sin API key)
    texto = _aplicar_correcciones_lt(texto)
    # 3. Asegurar que después de LT la primera letra sigue en mayúscula
    if texto:
        texto = texto[0].upper() + texto[1:]
    return texto

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

# ── Colores oficiales EAFIT ──
# Azul:     #004B85   Amarillo: #FFB903
# Negro:    #000000   Superficie: #0C0C0E  Borde: #1C1C22

st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

  /* ── BASE ── */
  html, body, .stApp {
    background-color: #000000 !important;
    color: #E8EAF0 !important;
    font-family: 'Inter', sans-serif !important;
  }
  .block-container {
    padding-top: 0 !important;
    padding-bottom: 2rem !important;
    max-width: 720px !important;
  }

  /* ── HEADER ── */
  .exa-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 1.6rem 0 1.4rem 0;
    border-bottom: 1px solid #1C1C22;
    margin-bottom: 2rem;
  }
  .exa-header-left {
    display: flex;
    align-items: center;
    gap: 1rem;
  }
  .exa-logo img {
    height: 38px;
    width: auto;
    display: block;
  }
  .exa-divider-v {
    width: 1px;
    height: 36px;
    background: #1C1C22;
  }
  .exa-title-block {}
  .exa-title {
    font-size: 1.15rem;
    font-weight: 700;
    color: #FFFFFF;
    line-height: 1.25;
    letter-spacing: -0.01em;
    margin: 0;
  }
  .exa-subtitle {
    font-size: 0.76rem;
    color: #4A5068;
    margin: 3px 0 0 0;
    letter-spacing: 0.04em;
    text-transform: uppercase;
  }
  .exa-badge {
    font-size: 0.7rem;
    font-weight: 600;
    color: #000000;
    background: #FFB903;
    padding: 3px 10px;
    border-radius: 20px;
    letter-spacing: 0.05em;
    text-transform: uppercase;
  }

  /* ── ACCENT LINE ── */
  .exa-accent-line {
    height: 2px;
    background: linear-gradient(90deg, #004B85 0%, #FFB903 60%, transparent 100%);
    margin-bottom: 2rem;
    border-radius: 2px;
  }

  /* ── CARDS ── */
  .card {
    background: #0C0C0E;
    border: 1px solid #1C1C22;
    border-radius: 12px;
    padding: 1.4rem 1.6rem;
    margin-bottom: 1.2rem;
  }
  .card-label {
    font-size: 0.68rem;
    font-weight: 700;
    color: #FFB903;
    text-transform: uppercase;
    letter-spacing: 0.12em;
    margin-bottom: 1rem;
    display: flex;
    align-items: center;
    gap: 6px;
  }
  .card-label::after {
    content: '';
    flex: 1;
    height: 1px;
    background: #1C1C22;
  }

  /* ── INPUTS ── */
  .stTextInput > div > div > input {
    background: #070709 !important;
    color: #E8EAF0 !important;
    border: 1px solid #1C1C22 !important;
    border-radius: 8px !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 0.88rem !important;
  }
  .stTextInput > div > div > input:focus {
    border-color: #004B85 !important;
    box-shadow: 0 0 0 3px rgba(0,75,133,0.2) !important;
  }
  .stTextInput label {
    color: #9399A8 !important;
    font-size: 0.82rem !important;
    font-weight: 500 !important;
  }

  /* ── FILE UPLOADER ── */
  .stFileUploader > div {
    background: #070709 !important;
    border: 1.5px dashed #1C1C22 !important;
    border-radius: 10px !important;
    transition: border-color 0.2s;
  }
  .stFileUploader > div:hover {
    border-color: #004B85 !important;
  }
  .stFileUploader label {
    color: #9399A8 !important;
    font-size: 0.82rem !important;
    font-weight: 500 !important;
  }
  /* Upload icon/text color */
  .stFileUploader [data-testid="stFileUploaderDropzone"] p,
  .stFileUploader [data-testid="stFileUploaderDropzone"] span {
    color: #4A5068 !important;
  }

  /* ── BUTTON PRIMARY ── */
  .stButton > button {
    background: #004B85 !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: 'Inter', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.9rem !important;
    padding: 0.55rem 1.6rem !important;
    letter-spacing: 0.01em;
    transition: background 0.15s, box-shadow 0.15s !important;
  }
  .stButton > button:hover {
    background: #005FA8 !important;
    box-shadow: 0 0 0 3px rgba(0,75,133,0.25) !important;
  }
  .stButton > button:active { background: #003D6E !important; }

  /* ── BUTTON DOWNLOAD ── */
  .stDownloadButton > button {
    background: #FFB903 !important;
    color: #000000 !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: 'Inter', sans-serif !important;
    font-weight: 700 !important;
    font-size: 0.9rem !important;
    padding: 0.55rem 1.6rem !important;
    transition: background 0.15s, box-shadow 0.15s !important;
  }
  .stDownloadButton > button:hover {
    background: #FFC72C !important;
    box-shadow: 0 0 0 3px rgba(255,185,3,0.25) !important;
  }

  /* ── PROGRESS BAR ── */
  .stProgress > div > div {
    background: #1C1C22 !important;
    border-radius: 4px !important;
  }
  .stProgress > div > div > div {
    background: linear-gradient(90deg, #004B85, #FFB903) !important;
    border-radius: 4px !important;
  }

  /* ── ALERTS ── */
  .stAlert {
    background: #0C0C0E !important;
    border-radius: 8px !important;
    border-left-width: 3px !important;
  }
  [data-testid="stAlert"][kind="info"] {
    border-color: #004B85 !important;
  }
  [data-testid="stAlert"][kind="success"] {
    border-color: #16A34A !important;
  }
  [data-testid="stAlert"][kind="error"] {
    border-color: #DC2626 !important;
  }

  /* ── SPINNER ── */
  .stSpinner > div { border-top-color: #FFB903 !important; }

  /* ── PREVIEW TABLE ── */
  .preview-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.82rem;
    margin-top: 0.4rem;
  }
  .preview-table thead tr {
    border-bottom: 1px solid #004B85;
  }
  .preview-table th {
    background: transparent;
    color: #4A5068;
    padding: 6px 10px;
    text-align: left;
    font-weight: 600;
    font-size: 0.68rem;
    text-transform: uppercase;
    letter-spacing: 0.1em;
  }
  .preview-table td {
    padding: 9px 10px;
    color: #C8CAD4;
    border-bottom: 1px solid #111116;
    vertical-align: middle;
  }
  .preview-table tr:last-child td { border-bottom: none; }
  .preview-table tr:hover td { background: #0F0F13; }
  .preview-table .name-cell { color: #FFFFFF; font-weight: 600; }
  .preview-table .file-cell { color: #4A5068; font-family: 'Courier New', monospace; font-size: 0.74rem; }
  .badge-ciclo {
    display: inline-block;
    background: #111116;
    color: #FFB903;
    border: 1px solid #1C1C22;
    border-radius: 4px;
    padding: 1px 7px;
    font-size: 0.72rem;
    font-weight: 600;
  }

  /* ── HIDE STREAMLIT CHROME ── */
  #MainMenu, footer, header { visibility: hidden; }
  [data-testid="stDecoration"] { display: none; }
</style>
""", unsafe_allow_html=True)

# ── Header ──
st.markdown("""
<div class="exa-header">
  <div class="exa-header-left">
    <div class="exa-logo">
      <img src="https://www.eafit.edu.co/sites/default/files/2024-07/logo_EAFIT_blanco.svg"
           alt="EAFIT" />
    </div>
    <div class="exa-divider-v"></div>
    <div class="exa-title-block">
      <div class="exa-title">Informes de evaluación docente</div>
      <div class="exa-subtitle">Centro para la Excelencia en el Aprendizaje · EXA</div>
    </div>
  </div>
  <div class="exa-badge">EXA</div>
</div>
<div class="exa-accent-line"></div>
""", unsafe_allow_html=True)

# ── Cargar plantilla desde el repositorio ──
_PLANTILLA_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Plantilla.docx")

@st.cache_data
def _cargar_plantilla():
    if os.path.isfile(_PLANTILLA_PATH):
        with open(_PLANTILLA_PATH, "rb") as f:
            return f.read()
    return None

plantilla_bytes = _cargar_plantilla()

if plantilla_bytes is None:
    st.error("⚠️ No se encontró **Plantilla.docx** en el repositorio. "
             "Asegúrate de subir ese archivo a GitHub junto con `app.py`.")
    st.stop()

# ── Estado LanguageTool (se verifica una vez) ──
@st.cache_data(ttl=300)
def _verificar_lt() -> bool:
    """Verifica si LanguageTool API está disponible. Cachea por 5 min."""
    try:
        data = urllib.parse.urlencode({"language": "es", "text": "prueba"}).encode("utf-8")
        req = urllib.request.Request(
            _LT_URL, data=data,
            headers={"Content-Type": "application/x-www-form-urlencoded",
                     "User-Agent": "EXA-InformeDocente/1.0"},
        )
        with urllib.request.urlopen(req, timeout=5):
            pass
        return True
    except Exception:
        return False

_lt_disponible = _verificar_lt()
_estado_lt = (
    '<span style="color:#16A34A;font-size:0.75rem">● Corrección ortográfica activa</span>'
    if _lt_disponible else
    '<span style="color:#4A5068;font-size:0.75rem">○ Corrección ortográfica no disponible</span>'
)
st.markdown(f'<div style="text-align:right;margin-top:-1.2rem;margin-bottom:1rem">{_estado_lt}</div>',
            unsafe_allow_html=True)


st.markdown('<div class="card"><div class="card-label">📂 Archivos</div>', unsafe_allow_html=True)

archivo_excel = st.file_uploader(
    "Archivo Excel de evaluaciones",
    type=["xlsx", "xls"],
    help="El archivo exportado del sistema de evaluación docente"
)

st.markdown('</div>', unsafe_allow_html=True)

# ── Preview automático al subir Excel ──
profesores = {}
if archivo_excel:
    try:
        with st.spinner("Leyendo Excel y procesando comentarios…"):
            profesores = leer_excel(archivo_excel.getvalue())

        st.markdown('<div class="card"><div class="card-label">👥 Profesores encontrados</div>',
                    unsafe_allow_html=True)

        filas_html = ""
        for nombre, datos in profesores.items():
            info = datos["info"]
            nf   = nombre_archivo_defecto(datos, nombre)
            filas_html += f"""
            <tr>
              <td class="name-cell">{nombre.title()}</td>
              <td>{info.get('curso','—')}</td>
              <td>{info.get('escuela','—')}</td>
              <td><span class="badge-ciclo">{info.get('ciclo','—')}</span></td>
              <td class="file-cell">{nf}.docx</td>
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
if profesores:
    st.markdown('<div class="card"><div class="card-label">⚙️ Generar informes</div>',
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

elif not archivo_excel:
    st.info("📊 Sube el **archivo Excel** para comenzar.")
