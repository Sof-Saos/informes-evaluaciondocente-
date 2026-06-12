# 📋 Generador de Informes de Evaluación Docente

Herramienta web para generar automáticamente informes individuales de evaluación docente a partir de los resultados del sistema institucional.

---

## ¿Qué hace?

El sistema busca los resultados de un curso por su código de catálogo y número de clase, y genera un informe Word (`.docx`) personalizado por cada profesor, con:

- **Portada** con nombre del profesor, curso, escuela y semestre
- **Diagrama de araña** con las notas por competencia (escala 0–5)
- **Tabla de estudiantes** con el número de evaluaciones generadas y realizadas
- **Comentarios** organizados por sección (aspectos positivos, aspectos a mejorar y comentarios adicionales), filtrados y formateados automáticamente

---

## Cómo usar la aplicación web

1. Entra a la URL de la app
2. Ingresa el **código de catálogo y número de clase** en el formato `CATÁLOGO-CLASE` (ej. `OG2117-5890`)
3. La app muestra una vista previa con los profesores encontrados
4. Sube la **Plantilla Word** (`Plantilla.docx`) si no está cargada aún
5. Haz clic en **Generar informes**
6. Descarga el `.docx` (un profesor) o el `.zip` (varios profesores)

> La base de datos de evaluaciones ya viene incluida en la app — no es necesario subir ningún Excel para generar informes.

---

## Estructura del repositorio

```
├── app.py               # Aplicación web (Streamlit)
├── requirements.txt     # Dependencias de Python
├── evaluaciones.xlsx    # Base de datos de evaluaciones docentes
├── Plantilla.docx       # Plantilla base para los informes
└── README.md
```

---

## Actualizar la base de datos

Cuando haya datos de un nuevo semestre, ve al expander **"Actualizar base de datos de evaluación docente"** al final de la app, sube el nuevo Excel y confirma. El archivo se sube directamente al repositorio en GitHub mediante un commit automático — el cambio es permanente y la app se actualiza en segundos.

Para que esto funcione, el secreto `GITHUB_TOKEN` debe estar configurado en Streamlit Cloud (*Settings → Secrets*):

```toml
GITHUB_TOKEN = "ghp_tuTokenAquí"
```

El token debe tener permisos de **Contents: Read and write** sobre el repositorio.

---

## Formato del Excel de evaluaciones

El archivo debe tener los encabezados en la **fila 7** y los datos desde la **fila 8**. Las columnas que usa el sistema son:

| Columna | Descripción |
|---|---|
| `Nombres y apellidos Docente` | Nombre completo del profesor |
| `Ciclo` | Semestre (ej. `2661`) |
| `Catálogo` | Código del catálogo (ej. `OG2117`) |
| `Nº Clase` | Número de clase (ej. `5890`) |
| `Nombre Catalogo` | Nombre del curso |
| `Id Escuela` | Código de escuela (ej. `E-ADM`) |
| `Escuela` | Nombre abreviado (fallback si el código no está mapeado) |
| `Competencia Evaluada` | Nombre de la competencia |
| `Nota competencia por clase` | Nota de la competencia por sesión |
| `Nota final por curso` | Nota consolidada del curso |
| `Pregunta` | Pregunta de la encuesta |
| `Comentarios` | Respuesta abierta del estudiante |
| `Total Evaluaciones generadas` | Total de estudiantes del curso |
| `Evaluaciones realizadas` | Estudiantes que respondieron |

---

## Escuelas mapeadas

El sistema traduce automáticamente los códigos del sistema a nombres completos:

| Código | Nombre completo |
|---|---|
| `E-ADM` | Escuela de Administración |
| `E-DER` | Escuela de Derecho |
| `E-ECO` | Escuela de Economía y Finanzas |
| `E-HUM` | Escuela de Humanidades |
| `E-ING` | Escuela de Ciencias Aplicadas e Ingeniería |
| `E-MED` | Escuela de Medicina |
| `E-MUS` | Escuela de Música |
| `E-DIS` | Escuela de Arquitectura y Diseño |
| `E-CS` | Escuela de Ciencias |
| `E-VIS` | Vicerrectoría de Internacionalización |

Si aparece un código nuevo que no esté en esta lista, se usa el texto de la columna `Escuela` tal cual viene del Excel.

---

## Formato de los comentarios

El sistema aplica automáticamente estas reglas a cada comentario antes de incluirlo en el informe:

- Primera letra en mayúscula, resto en minúscula
- Siglas en mayúsculas preservadas (ej. `TIC`, `COVID`, `EAFIT`)
- Comentarios vacíos, muy cortos o genéricos (`no`, `ok`, `ninguno`, `n/a`, etc.) son descartados
- Se agrega punto final si el comentario no termina en signo de puntuación

---

## Nombre de los archivos generados

Los informes se nombran automáticamente con el formato:

```
SEMESTRE_ESCUELA_NombreCurso_NombreProfesor.docx
```

**Ejemplo:** `2661_EING_GerenciadeProyectos_LauraOlarte.docx`

Cuando se genera un solo informe, el nombre es editable desde la interfaz antes de descargar.

---

## Dependencias

- [Streamlit](https://streamlit.io/) — interfaz web
- [openpyxl](https://openpyxl.readthedocs.io/) — lectura de archivos Excel
- [lxml](https://lxml.de/) — manipulación del XML interno de los archivos Word
- [matplotlib](https://matplotlib.org/) — generación del diagrama de araña en el informe
- [numpy](https://numpy.org/) — cálculos para el diagrama de araña
