"""
Generador de informes de evaluación docente — EAFIT
Versión web (Streamlit)
"""

import io, os, re, zipfile, random, tempfile
import math
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
COL_NOTA_FINAL      = "Nota competencia por clase"
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

# ─── (LanguageTool eliminado por rendimiento) ──────────────────────────────

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
    """Sentence case únicamente (LanguageTool eliminado por rendimiento)."""
    texto = str(texto).strip()
    if not texto:
        return texto
    texto = _sentence_case(texto)
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

COL_CATALOGO = "Catálogo"
COL_NCLASE   = "Nº Clase"

def leer_excel(archivo_bytes: bytes,
               filtro_catalogo: str = None,
               filtro_clase: str = None) -> dict:
    """
    Lee el Excel de evaluaciones docentes.
    Si se pasan filtro_catalogo y filtro_clase, solo procesa las filas
    que coincidan con esa combinación (mucho más rápido).
    """
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

    # Normalizar filtros
    f_cat   = str(filtro_catalogo).strip().upper() if filtro_catalogo else None
    f_clase = str(filtro_clase).strip() if filtro_clase else None

    profesores = defaultdict(lambda: {
        "info": {}, "nota_final": None, "nota_curso": None,
        "comentarios": defaultdict(list),
        "total_generadas": None, "evaluaciones_realizadas": None,
        "notas_competencias": {},
    })

    for row in ws.iter_rows(min_row=FILA_DATOS, values_only=True):
        nombre = col(row, COL_NOMBRE)
        if not nombre: continue

        # Aplicar filtro si se indicó
        if f_cat or f_clase:
            cat_row   = str(col(row, COL_CATALOGO) or "").strip().upper()
            clase_row = str(col(row, COL_NCLASE) or "").strip()
            if f_cat and cat_row != f_cat:
                continue
            if f_clase and clase_row != f_clase:
                continue

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

        if comp and comp != "Comentarios" and nota_final and nota_final > 0:
            if comp not in p["notas_competencias"]:
                p["notas_competencias"][comp] = nota_final

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

    # ── Tabla / gráfico competencias ──
    tbl_comp = body_children[3]
    notas_comp = datos.get("notas_competencias", {})

    if notas_comp and len(notas_comp) >= 3:
        # Generar PNG del diagrama de araña e insertarlo como imagen en el docx
        png_bytes  = _spider_chart_png(notas_comp)
        img_rid    = "rIdSpider"
        img_name   = "spider_chart.png"
        img_target = f"media/{img_name}"

        # Construir elemento <w:p> con la imagen inline
        # Tamaño: ~10 cm de ancho (EMU: 1 cm = 914400 / 100 * 10 = 3600000... → 1 cm = 360000 EMU, 10 cm = 3600000)
        EMU_W = 4200000   # ~11.7 cm
        EMU_H = 4200000

        draw_xml = (
            f'<w:p xmlns:w="{W[1:-1]}"'
            f' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
            f' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            f' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
            f' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<w:pPr><w:jc w:val="center"/></w:pPr>'
            f'<w:r><w:rPr/><w:drawing>'
            f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent cx="{EMU_W}" cy="{EMU_H}"/>'
            f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:docPr id="1" name="SpiderChart"/>'
            f'<wp:cNvGraphicFramePr>'
            f'<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>'
            f'</wp:cNvGraphicFramePr>'
            f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:nvPicPr>'
            f'<pic:cNvPr id="0" name="{img_name}"/>'
            f'<pic:cNvPicPr/>'
            f'</pic:nvPicPr>'
            f'<pic:blipFill>'
            f'<a:blip r:embed="{img_rid}"/>'
            f'<a:stretch><a:fillRect/></a:stretch>'
            f'</pic:blipFill>'
            f'<pic:spPr>'
            f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{EMU_W}" cy="{EMU_H}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            f'</pic:spPr>'
            f'</pic:pic>'
            f'</a:graphicData>'
            f'</a:graphic>'
            f'</wp:inline>'
            f'</w:drawing></w:r></w:p>'
        )
        img_para = etree.fromstring(draw_xml)

        # Insertar el párrafo de imagen donde estaba la tabla, luego eliminar la tabla
        tbl_pos = list(body).index(tbl_comp)
        body.remove(tbl_comp)
        body.insert(tbl_pos, img_para)

        # Guardar la imagen y actualizar relationships en el zip de salida
        _spider_png_data   = png_bytes
        _spider_img_rid    = img_rid
        _spider_img_target = img_target
    else:
        # Sin datos de competencias: limpiar placeholders de la tabla original
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
        _spider_png_data   = None
        _spider_img_rid    = None
        _spider_img_target = None

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
                if item.filename == 'word/document.xml':
                    zout.writestr(item, new_xml)
                elif _spider_png_data and item.filename == 'word/_rels/document.xml.rels':
                    # Inyectar la relationship de la imagen
                    rels_xml = zin.read(item.filename)
                    rels_tree = etree.fromstring(rels_xml)
                    REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
                    new_rel = etree.SubElement(rels_tree, f"{{{REL_NS}}}Relationship")
                    new_rel.set("Id", _spider_img_rid)
                    new_rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
                    new_rel.set("Target", _spider_img_target)
                    zout.writestr(item, etree.tostring(rels_tree, xml_declaration=True,
                                                       encoding='UTF-8', standalone=True))
                else:
                    zout.writestr(item, zin.read(item.filename))
            # Escribir el PNG de la araña dentro del zip
            if _spider_png_data:
                zout.writestr(f"word/{_spider_img_target}", _spider_png_data)
    return out_buf.getvalue(), nombre_archivo

# ─── SPIDER CHART PNG (para insertar en Word) ───────────────────────────────

def _spider_chart_png(notas: dict) -> bytes:
    """
    Genera el diagrama de araña como PNG en memoria usando matplotlib.
    Retorna los bytes del PNG.
    """
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import numpy as np

    labels = list(notas.keys())
    values = [min(max(float(v), 0), 5) for v in notas.values()]
    n = len(labels)

    # Ángulos para cada eje (cerrar el polígono repitiendo el primero)
    angles = np.linspace(0, 2 * np.pi, n, endpoint=False).tolist()
    angles += angles[:1]
    values_plot = values + values[:1]

    fig, ax = plt.subplots(figsize=(5.5, 5.5), subplot_kw=dict(polar=True))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    # Colores EAFIT
    C_BLUE   = "#004B85"
    C_YELLOW = "#FFB903"
    C_GRID   = "#CCCCCC"

    # Rejilla y niveles
    ax.set_ylim(0, 5)
    ax.set_yticks([1, 2, 3, 4, 5])
    ax.set_yticklabels(["1", "2", "3", "4", "5"], color="#888888", fontsize=7)
    ax.yaxis.set_tick_params(labelsize=7)
    ax.grid(color=C_GRID, linewidth=0.8, linestyle="--", alpha=0.7)
    ax.spines["polar"].set_color(C_GRID)

    # Ejes con etiquetas
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(labels, fontsize=8.5, color="#222222",
                       fontfamily="DejaVu Sans", wrap=True)
    # Ajuste de padding para etiquetas largas
    ax.tick_params(axis='x', pad=10)

    # Polígono de datos
    ax.plot(angles, values_plot, color=C_BLUE, linewidth=2, linestyle="solid")
    ax.fill(angles, values_plot, color=C_BLUE, alpha=0.25)

    # Puntos y etiquetas de puntaje
    for angle, val in zip(angles[:-1], values):
        ax.plot(angle, val, "o", color=C_YELLOW, markersize=7,
                markeredgecolor=C_BLUE, markeredgewidth=1.2)
        ax.annotate(
            f"{val:.2f}",
            xy=(angle, val),
            xytext=(angle, val + 0.35),
            ha="center", va="center",
            fontsize=7.5, color=C_BLUE, fontweight="bold",
        )

    plt.tight_layout(pad=1.5)

    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=150, bbox_inches="tight",
                facecolor="white", edgecolor="none")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


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



# ── Cargar base de datos de evaluaciones (embebida en el repo) ──
_DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "evaluaciones.xlsx")

@st.cache_data(show_spinner=False)
def _cargar_db() -> bytes:
    if os.path.isfile(_DB_PATH):
        with open(_DB_PATH, "rb") as f:
            return f.read()
    return None

_db_bytes = _cargar_db()

if _db_bytes is None:
    st.error("⚠️ No se encontró **evaluaciones.xlsx** en el repositorio.")
    st.stop()

# ── Entrada: Catálogo y Nº Clase ──
st.markdown('<div class="card"><div class="card-label">🔍 Buscar clase</div>', unsafe_allow_html=True)

input_codigo = st.text_input(
    "Catálogo – Nº de clase",
    placeholder="Ej: OG2117-5890",
    help="Ingresa el código del catálogo seguido de un guion y el número de clase (ej: OG2117-5890)"
)

buscar = st.button("🔎 Buscar y previsualizar")
st.markdown('</div>', unsafe_allow_html=True)

# ── Estado de sesión para profesores encontrados ──
if "profesores" not in st.session_state:
    st.session_state.profesores = {}

profesores = st.session_state.profesores

if buscar:
    codigo = input_codigo.strip()
    if not codigo or "-" not in codigo:
        st.warning("Ingresa el código en el formato **CATÁLOGO-CLASE**, por ejemplo: `OG2117-5890`.")
    else:
        # Separar por el último guion para admitir catálogos con letras+números
        partes = codigo.rsplit("-", 1)
        if len(partes) != 2 or not partes[0].strip() or not partes[1].strip():
            st.warning("Formato inválido. Usa **CATÁLOGO-CLASE**, por ejemplo: `OG2117-5890`.")
        else:
            input_catalogo = partes[0].strip()
            input_clase    = partes[1].strip()
            try:
                with st.spinner("Buscando en la base de datos…"):
                    resultado = leer_excel(
                        _db_bytes,
                        filtro_catalogo=input_catalogo,
                        filtro_clase=input_clase
                    )
                if not resultado:
                    st.error(f"No se encontraron registros para **{input_catalogo.upper()}-{input_clase}**. "
                             "Verifica el catálogo y número de clase.")
                    st.session_state.profesores = {}
                else:
                    st.session_state.profesores = resultado
                    profesores = resultado
            except Exception as e:
                st.error(f"Error al leer la base de datos: {e}")

# ── Preview de resultados ──
if profesores:
    try:
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
        st.error(f"Error al mostrar resultados: {e}")

# ── Sección: Generar informes ──
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
        archivos_generados = {}

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

elif buscar and not profesores:
    pass  # El error ya se mostró arriba
else:
    st.info("🔍 Ingresa el código en formato **CATÁLOGO-CLASE** (ej: `OG2117-5890`) para comenzar.")

# ── Sección: Actualizar base de datos (al final) ──
import base64
import urllib.request
import urllib.error
import json as _json

_GH_REPO  = "Sof-Saos/informes-evaluaciondocente-"   # usuario/repo
_GH_FILE  = "evaluaciones.xlsx"                        # ruta dentro del repo
_GH_TOKEN = st.secrets.get("GITHUB_TOKEN", "")        # secreto en Streamlit Cloud

def _gh_get_sha(token: str) -> str | None:
    """Obtiene el SHA actual del archivo en GitHub (necesario para actualizarlo)."""
    url = f"https://api.github.com/repos/{_GH_REPO}/contents/{_GH_FILE}"
    req = urllib.request.Request(url, headers={
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    })
    try:
        with urllib.request.urlopen(req, timeout=10) as r:
            return _json.loads(r.read())["sha"]
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None   # archivo no existe todavía
        raise

def _gh_push_file(token: str, content_bytes: bytes, sha: str | None, mensaje: str) -> bool:
    """Sube (o reemplaza) el archivo en GitHub mediante un commit."""
    url = f"https://api.github.com/repos/{_GH_REPO}/contents/{_GH_FILE}"
    payload = {
        "message": mensaje,
        "content": base64.b64encode(content_bytes).decode(),
        "branch": "main",
    }
    if sha:
        payload["sha"] = sha
    data = _json.dumps(payload).encode()
    req = urllib.request.Request(url, data=data, method="PUT", headers={
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "Content-Type": "application/json",
        "X-GitHub-Api-Version": "2022-11-28",
    })
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.status in (200, 201)

st.markdown("---")
with st.expander("🔄 Actualizar base de datos de evaluación docente"):
    if not _GH_TOKEN:
        st.warning(
            "⚠️ No se encontró el secreto **GITHUB_TOKEN** en Streamlit Cloud. "
            "Agrégalo en *Settings → Secrets* para habilitar la actualización permanente."
        )
    else:
        st.markdown(
            "<small style='color:#4A5068'>Sube el Excel del nuevo semestre. "
            "El archivo se guardará directamente en GitHub — el cambio es <b>permanente</b> "
            "y la app se actualizará automáticamente en unos segundos.</small>",
            unsafe_allow_html=True
        )
        nuevo_excel = st.file_uploader(
            "Nuevo archivo de evaluaciones (.xlsx)",
            type=["xlsx"],
            key="uploader_db"
        )
        if nuevo_excel:
            if st.button("✅ Confirmar actualización de base de datos"):
                try:
                    nuevos_bytes = nuevo_excel.getvalue()

                    # 1. Validar estructura del Excel
                    wb_test = openpyxl.load_workbook(io.BytesIO(nuevos_bytes), read_only=True)
                    ws_test = wb_test.active
                    headers_test = {}
                    for row in ws_test.iter_rows(min_row=FILA_ENCABEZADO, max_row=FILA_ENCABEZADO, values_only=True):
                        for i, val in enumerate(row):
                            if val: headers_test[str(val).strip()] = i
                        break
                    cols_requeridas = [COL_NOMBRE, COL_CATALOGO, COL_NCLASE, COL_COMPETENCIA, COL_NOTA_FINAL]
                    faltantes = [c for c in cols_requeridas if c not in headers_test]
                    if faltantes:
                        st.error(f"El archivo no tiene las columnas requeridas: {', '.join(faltantes)}")
                    else:
                        # 2. Subir a GitHub
                        with st.spinner("Subiendo a GitHub…"):
                            sha_actual = _gh_get_sha(_GH_TOKEN)
                            ciclo_val  = ws_test.cell(row=1, column=2).value or ""
                            commit_msg = f"Actualizar evaluaciones.xlsx — ciclo {ciclo_val} ({len(nuevos_bytes)//1024} KB)"
                            ok = _gh_push_file(_GH_TOKEN, nuevos_bytes, sha_actual, commit_msg)
                        if ok:
                            st.cache_data.clear()
                            st.success(
                                f"✅ Base de datos actualizada en GitHub ({len(nuevos_bytes)//1024} KB). "
                                "Streamlit Cloud redeployará la app en unos segundos con los datos nuevos."
                            )
                        else:
                            st.error("No se pudo subir el archivo a GitHub. Verifica el token y los permisos.")
                except urllib.error.HTTPError as e:
                    body = e.read().decode(errors="replace")
                    st.error(f"Error de GitHub ({e.code}): {body}")
                except Exception as e:
                    st.error(f"Error al actualizar: {e}")
