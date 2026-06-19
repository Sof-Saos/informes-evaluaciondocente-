# 📋 Generador de Informes de Evaluación Docente

Herramienta web para generar automáticamente informes individuales de evaluación docente a partir de los resultados del sistema institucional.

---

## ¿Qué hace?

El sistema busca los resultados de un curso por su código de catálogo y número de clase, y genera un informe Word (`.docx`) personalizado por cada profesor, con:

- **Portada** con nombre del profesor, curso, escuela y semestre
- **Diagrama de araña** con las notas por competencia (escala 0–5)
- **Tabla de estudiantes** con el número de evaluaciones generadas y realizadas
- **Comentarios** organizados por sección (positivos, a mejorar, adicionales), filtrados y formateados automáticamente
- **Consideraciones** redactadas con IA, alineadas con documentos institucionales de referencia

---

## Navegación

La app tiene dos módulos accesibles desde el menú lateral:

| Módulo | Descripción |
|---|---|
| 📋 Generar informe | Busca por catálogo-clase y genera el `.docx` base |
| ✨ Alistamiento de Consideraciones | Usa IA para redactar la sección Consideraciones del informe |

---

## Cómo usar — Generar informe

1. Entra a la URL de la app
2. Ingresa el código en el formato `CATÁLOGO-CLASE` (ej. `OG2117-5890`)
3. La app muestra una vista previa con los profesores encontrados
4. Haz clic en **Generar informes**
5. Descarga el `.docx` (un profesor) o el `.zip` (varios profesores)

> La base de datos de evaluaciones ya viene incluida — no es necesario subir ningún Excel.

---

## Cómo usar — Alistamiento de Consideraciones

1. Genera primero el informe base en el módulo anterior
2. Ve al módulo **✨ Alistamiento de Consideraciones**
3. El informe recién generado se detecta automáticamente (sin necesidad de subirlo de nuevo). Si necesitas trabajar sobre otro informe, marca la casilla **"Añadir consideraciones para otro informe"**
4. Selecciona los documentos guía que quieres usar (quedan guardados permanentemente en el repo)
5. Opcionalmente escribe una indicación: qué enfoque quieres, qué rescatar, qué resultado esperas
6. Haz clic en **Generar Consideraciones**
7. Revisa y edita el texto propuesto
8. Haz clic en **Incorporar al informe** y descarga el `.docx` final

---

## Estructura del repositorio

```
├── app.py               # Aplicación web (Streamlit)
├── requirements.txt     # Dependencias de Python
├── evaluaciones.xlsx    # Base de datos de evaluaciones docentes
├── Plantilla.docx       # Plantilla base para los informes
├── docs_guia/           # Documentos de referencia para IA (se crean automáticamente)
│   ├── formacion_exa.pdf
│   └── lineamientos.docx
└── README.md
```

---

## Configuración de secretos en Streamlit Cloud

La app usa **dos tokens de GitHub**, cada uno con un propósito distinto. Ambos se configuran en Streamlit Cloud en **Settings → Secrets**:

```toml
GITHUB_TOKEN = "ghp_tokenParaContenido"
GITHUB_MODELS_TOKEN = "github_pat_tokenParaIA"
```

### GITHUB_TOKEN — para subir el Excel y los documentos guía

**Permisos necesarios:** `Contents: Read and write`  
**Scope:** Solo el repositorio `informes-evaluaciondocente-`

Se usa para:
- Actualizar `evaluaciones.xlsx` cuando hay datos de un nuevo semestre
- Guardar y eliminar documentos en `docs_guia/`

### GITHUB_MODELS_TOKEN — para la IA (Consideraciones)

**Permisos necesarios:** `Models: Read`  
**Scope:** No requiere acceso a ningún repositorio específico — es un permiso de cuenta, no de repo

Se usa exclusivamente para llamar a la API de GitHub Models (GPT-4o mini) y generar las Consideraciones.

---

## Cómo crear el GITHUB_MODELS_TOKEN (paso a paso)

1. Ve a [github.com](https://github.com) e inicia sesión con tu cuenta
2. Haz clic en tu foto de perfil (esquina superior derecha) → **Settings**
3. En el menú lateral izquierdo, baja hasta el final → **Developer settings**
4. Haz clic en **Personal access tokens** → **Fine-grained tokens**
5. Haz clic en **Generate new token**
6. Completa los campos:
   - **Token name:** `EXA-Models-Consideraciones` (o el nombre que prefieras)
   - **Expiration:** elige la duración que quieras (recomendado: 1 año)
   - **Resource owner:** tu usuario personal (no una organización)
7. En **Repository access:** selecciona **"No repositories"** — este token no necesita acceso a ningún repo
8. En **Permissions → Account permissions:** busca **Models** y cambia el acceso a **Read-only**
9. Haz clic en **Generate token**
10. **Copia el token inmediatamente** (empieza con `github_pat_...`) — no se vuelve a mostrar
11. Ve a tu app en Streamlit Cloud → **Settings → Secrets** y agrega:

```toml
GITHUB_MODELS_TOKEN = "github_pat_elTokenQueCopiasté"
```

12. Guarda. La app detecta el secreto automáticamente en el próximo rerun.

> **¿Cómo verificar que funciona?** Ve al módulo de Consideraciones, genera un informe base, y haz clic en "Generar Consideraciones". Si el token es válido y tiene el permiso correcto, el texto aparecerá en unos segundos. Si el token está mal o le falta el permiso, verás un mensaje de error claro indicando exactamente qué falló.

---

## Límites del tier gratuito de GitHub Models

El plan gratuito de GitHub Models (tier "Low") permite:

| Límite | Valor |
|---|---|
| Solicitudes por minuto | 15 |
| Solicitudes por día | 150 |
| Tokens de entrada por solicitud | 8 000 |

Esto es más que suficiente para el flujo de generación de informes. Si se alcanza el límite diario, la app muestra un mensaje claro pidiendo esperar hasta el día siguiente.

---

## Actualizar la base de datos

Cuando haya datos de un nuevo semestre, ve al expander **"Actualizar base de datos de evaluación docente"** al final del módulo Generar informe. El archivo se sube directamente al repositorio mediante un commit automático — el cambio es permanente y la app se actualiza en segundos.

---

## Formato del Excel de evaluaciones

El archivo debe tener los encabezados en la **fila 7** y los datos desde la **fila 8**:

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
| `Nota final por clase` | Nota alternativa (usada para Índice de recomendación y Pacto Pedagógico) |
| `Nota final por curso` | Nota consolidada del curso |
| `Pregunta` | Pregunta de la encuesta |
| `Comentarios` | Respuesta abierta del estudiante |
| `Total Evaluaciones generadas` | Total de estudiantes del curso |
| `Evaluaciones realizadas` | Estudiantes que respondieron |

---

## Escuelas mapeadas

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

---

## Dependencias

- [Streamlit](https://streamlit.io/) — interfaz web
- [openpyxl](https://openpyxl.readthedocs.io/) — lectura de archivos Excel
- [lxml](https://lxml.de/) — manipulación del XML interno de los archivos Word
- [matplotlib](https://matplotlib.org/) — generación del diagrama de araña
- [numpy](https://numpy.org/) — cálculos para el diagrama de araña
- [pypdf](https://pypdf.readthedocs.io/) — lectura de documentos PDF de referencia
- [python-docx](https://python-docx.readthedocs.io/) — lectura de documentos DOCX de referencia
