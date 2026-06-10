# 📋 Generador de Informes de Evaluación Docente 

Herramienta web para generar automáticamente informes individuales de evaluación docente a partir de los archivos Excel exportados del sistema institucional.

---

## ¿Qué hace?

El sistema toma el archivo Excel con los resultados de la evaluación docente y genera un informe Word (`.docx`) personalizado por cada profesor, con:

- **Portada** con nombre del profesor, curso, escuela y semestre
- **Tabla de competencias** con las notas del curso
- **Tabla de estudiantes** con el número de evaluaciones generadas y realizadas
- **Comentarios** organizados por sección (aspectos positivos, aspectos a mejorar y comentarios adicionales), filtrados y formateados automáticamente

---

## Cómo usar la aplicación web

1. Entra a la URL de la app
2. Sube el **archivo Excel** de evaluación docente
3. Sube la **Plantilla Word** (archivo `Plantilla.docx`)
4. La app muestra una vista previa con los profesores encontrados
5. Haz clic en **Generar informes**
6. Descarga el `.docx` (un profesor) o el `.zip` (varios profesores)

---

## Estructura del repositorio

```
├── app.py              # Aplicación web (Streamlit)
├── requirements.txt    # Dependencias de Python
├── Plantilla.docx      # Plantilla base para los informes
└── README.md
```

---

## Formato del Excel esperado

El archivo Excel debe tener los encabezados en la **fila 7** y los datos desde la **fila 8**. Las columnas que usa el sistema son:

| Columna | Descripción |
|---|---|
| `Nombres y apellidos Docente` | Nombre completo del profesor |
| `Ciclo` | Semestre (ej. `2661`) |
| `Nombre Catalogo` | Nombre del curso |
| `Id Escuela` | Código de escuela (ej. `E-ADM`) |
| `Escuela` | Nombre abreviado (fallback si el código no está mapeado) |
| `Competencia Evaluada` | Nombre de la competencia |
| `Nota final por clase` | Nota por sesión |
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
