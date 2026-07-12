# CanónicaDL - Toolkit — Resumen técnico

## Qué es

Aplicación de escritorio para Windows que organiza bibliotecas digitales:
analiza una carpeta de archivos (PDF, EPUB, DOCX), busca su información
bibliográfica en internet, y los renombra automáticamente siguiendo un
formato configurable (personalizado o APA 7), manteniendo la colección
ordenada y consistente.

No requiere que el usuario final tenga Python instalado: se distribuye
como un `.exe` autocontenido generado con PyInstaller.

---

## Stack tecnológico

| Capa | Tecnología |
|---|---|
| Lenguaje | Python 3 |
| Interfaz gráfica | [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) (sobre Tkinter/ttk) |
| Tabla de resultados | `ttk.Treeview` (Tkinter clásico, estilizado a medida) |
| Arrastrar y soltar | `tkinterdnd2` |
| Peticiones HTTP | `requests` |
| Lectura de PDF | `pypdf` (fallback a `PyPDF2`) |
| Lectura de EPUB/DOCX | `zipfile` + `xml.etree.ElementTree` (librería estándar, sin dependencias extra) |
| Empaquetado a `.exe` | `PyInstaller` (`--onefile --windowed`) |
| Ícono de ventana/exe | Incrustado en base64 (`icono_datos.py`), generado con `Pillow` |

**Ninguna librería de IA/ML** — toda la lógica es determinística: reglas de
texto, expresiones regulares y llamadas a APIs bibliográficas públicas.

---

## Fuentes de datos bibliográficos

| Fuente | Uso |
|---|---|
| **Google Books API** | Búsqueda de libros por texto libre |
| **Open Library API** | Búsqueda directa por ISBN (la más confiable) |
| **Crossref API** | Artículos académicos, tesis, ponencias (con DOI) |
| Metadatos del propio archivo | Título/autor embebidos en el PDF/EPUB/DOCX |
| Nombre de archivo | Como último recurso, o como pista de alta confianza si sigue un patrón reconocible |

---

## Cómo decide qué nombre ponerle a cada archivo

1. **ISBN** encontrado en el texto o metadatos → Open Library (máxima confianza)
2. **Patrón reconocible en el nombre de archivo** (`Título - Autor`,
   `Autor, Título`, `Título_Autor`) → se usa directo, la búsqueda solo
   completa año/editorial si valida bien
3. **Búsqueda por texto** en Google Books + Crossref, pero **validando**
   que el resultado se parezca de verdad al título de referencia (para
   evitar falsos positivos con títulos cortos/genéricos)
4. Si nada de lo anterior funciona → metadatos del archivo o nombre de
   archivo limpio, marcado como "SIN DATOS" para revisión manual

---

## Funcionalidades principales

- Soporte multi-formato: PDF, EPUB, DOCX (mantiene la extensión original)
- Múltiples autores con formato APA 7 automático (`y`, comas, iniciales)
- Plantilla de nomenclatura configurable con placeholders (`{autores}`,
  `{titulo}`, `{anio}`, `{revista}`, `{doi}`, etc.) + presets (APA 7 libro/artículo)
- Edición en la celda (estilo planilla), con Tab para pasar al siguiente campo
- Tabla ordenable por columna, con colores por estado (verde/amarillo/rojo) y zebra
- Búsqueda manual con selección entre varios resultados, o carga 100% manual
- Editor de autores múltiples
- Arrastrar y soltar carpetas
- Detección y limpieza de nombres de archivo desprolijos (MAYÚSCULAS,
  guiones, guiones bajos, prefijos numéricos)
- Tema oscuro en toda la interfaz, incluida la barra de título de Windows
- Ícono y ventana "Acerca de" personalizados con la identidad de marca

---

## Estructura de archivos del proyecto

```
CanonicaDL/
├── app.py              # Toda la app (~2100 líneas): lógica + interfaz
├── icono_datos.py       # Ícono y logo incrustados en base64
├── icono.ico / icono.png  # Archivos de ícono sueltos (para --icon de PyInstaller)
├── requirements.txt      # Dependencias de Python
├── generar_exe.bat       # Compila el .exe (Windows)
├── ejecutar_sin_compilar.bat  # Corre la app directo con Python, sin compilar
└── README.md            # Historial de versiones y guía de uso
```

**Diseño interno de `app.py`:** la lógica de negocio (búsquedas, parseo de
nombres, formato APA, extracción de metadatos) está separada en funciones
puras al principio del archivo, sin ninguna dependencia de la interfaz
gráfica — la clase `RenombradorApp` y los diálogos (`AcercaDeDialog`,
`EditorAutoresDialog`, `DialogoBusquedaManual`) solo llaman a esas
funciones y muestran los resultados.

---

## Cómo se genera el ejecutable

```
PyInstaller --onefile --windowed --name CanonicaDL --icon=icono.ico
            --collect-all tkinterdnd2 --collect-all customtkinter app.py
```

Un solo archivo `.exe` (~40-50 MB), sin instalador, sin necesitar Python
en la máquina del usuario final.
