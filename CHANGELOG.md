# Registro de cambios

Historial de versiones de CanónicaDL, de la más antigua a la más reciente.

## Novedades

### 1. Limpieza automática de nombres de archivo
Cuando no hay datos de internet disponibles, el nombre del archivo (por ej.
`Manifesto-Antropofago-e-Manifesto-Da-Poesia-Pau.pdf`) se limpia
automáticamente reemplazando guiones/underscores por espacios y aplicando
mayúsculas iniciales (`Manifesto Antropofago e Manifesto Da Poesia Pau`),
tanto para usarlo como búsqueda como para el título de respaldo.

### 2. Carga de datos 100% manual
Clic derecho sobre una fila → **"Buscar manualmente..."**. Ahí podés:
- Buscar por texto y elegir entre varios resultados de Google Books, **o**
- Completar vos mismo los campos Nombre / Apellido / Título / Año (sección
  de abajo del diálogo) y aplicarlos directamente, sin depender de ninguna
  búsqueda en internet.

### 3. Formato de nomenclatura elegible y combinable
Arriba de la tabla hay una sección **"Formato de nomenclatura"** con:
- Un desplegable de formatos predefinidos, incluyendo **APA 7 (sin
  editorial)** tanto para libros como para artículos.
- Un campo de plantilla editable donde podés combinar libremente estos
  campos: `{nombre} {apellido} {inicial} {titulo} {anio} {editorial}
  {revista} {volumen} {numero} {paginas} {doi}`.
- Botón **"Aplicar a todas las filas"** para recalcular los nombres ya
  analizados con el nuevo formato.

Ejemplos de plantilla:
- `{apellido}, {inicial} ({anio}). {titulo}.` → `García Márquez, G. (1985). El amor en los tiempos del cólera.`
- `{titulo} ({anio})` → `El amor en los tiempos del cólera (1985)`

Los campos de artículo (revista, volumen, número, páginas, DOI) no se
completan automáticamente por búsqueda (Google Books no los provee de forma
confiable); si necesitás ese formato, cargalos manualmente por ahora vía
el diálogo de búsqueda manual y doble clic en la celda.

## Novedades (v3)

### Detección mejorada
- Si el propio nombre de archivo ya trae el patrón **"Título - Autor"** o
  **"Título_Autor Apellido"** (por ej. `Arte en la Red - Jesús Carrillo.pdf`
  o `Digital-Art_Christiane Paul.pdf`), la app ahora lo reconoce y separa
  automáticamente el título del autor, en vez de dejar todo mezclado en el
  título. Se marca como **"SIN CONFIRMAR (detectado en nombre de archivo)"**
  para que lo revises rápido.
- Se descartan los metadatos "basura" que a veces trae un PDF exportado
  desde Word (por ej. `Microsoft Word - archivo.doc` como título), que antes
  se colaban como título real.
- Los guiones bajos (`_`) y guiones (`-`) del nombre de archivo ahora se
  limpian siempre que se usa como respaldo, no solo en algunos casos.
- Un archivo problemático (PDF corrupto, error de red puntual, etc.) ya no
  traba el análisis del resto de la carpeta: queda marcado como "ERROR" y
  la app sigue con los demás.

### Limitación conocida
Nombres de archivo tipo `Autor_Palabra_Palabra_Palabra_Del_Titulo.pdf` (todo
separado por guion bajo, sin un delimitador claro entre autor y título) son
difíciles de separar de forma automática y confiable. En esos casos la app
intenta buscar el texto completo en Google Books (mejor que antes, porque ya
no usa metadatos basura), pero si no encuentra nada vas a tener que revisar
esa fila a mano con clic derecho → "Buscar manualmente...".

## Novedades (v4)

### Más fuentes de búsqueda
- Ahora se consulta también **Crossref** (base de datos académica libre:
  artículos, ensayos, actas de congresos, capítulos, todo lo que tenga DOI),
  además de Google Books y Open Library. Esto ayuda mucho con textos que no
  son "libros" con ficha propia, que es donde Google Books solía fallar.
- En el diálogo de búsqueda manual, cada resultado ahora indica de qué
  fuente viene: `[Google Books] ...` o `[Crossref] ...`.
- Nuevo botón **"Abrir en Google 🔗"** en la búsqueda manual: si ninguna de
  las dos APIs encuentra el libro/artículo (puede pasar con textos que
  circulan como PDF suelto sin estar indexados en ninguna base académica),
  abre una búsqueda de Google normal en tu navegador para que confirmes los
  datos ahí y los cargues en la sección manual de abajo.

## Novedades (v5)

### Validación de resultados (menos falsos positivos)
Antes, la app confiaba en el primer resultado que le daba Google Books o
Crossref, aunque no tuviera nada que ver con el archivo real (pasaba mucho
con títulos cortos/genéricos tipo "Arte en la Red", que hacen "match" con
cualquier texto que hable de arte).

Ahora cada resultado se valida comparando las palabras del título encontrado
contra un título de referencia:
- Si el título de referencia es **muy corto o genérico** (menos de 3
  palabras con contenido), la app **no confía en ningún resultado
  automático** y prefiere dejarlo marcado para que lo revises, en vez de
  asignar algo incorrecto.
- Si el nombre de archivo ya trae un patrón "Título - Autor" confiable, ese
  dato **nunca se pisa** con un resultado de búsqueda ambiguo: la búsqueda
  solo se usa (si valida bien) para completar año/editorial/DOI.

También se corrigió un bug donde, si Crossref no tenía el año de
publicación, aparecía literalmente el texto **"None"** en vez de dejarlo
vacío.

## Novedades (v6)

### Ícono personalizado
Se agregó `icono.ico` e `icono.png` a la carpeta - se usan tanto para el
ícono de la ventana de la app como para el ícono del `.exe` generado. Si
alguna vez querés cambiarlo, solo reemplazá esos dos archivos (el `.ico`
debe incluir varias resoluciones; herramientas online como icoconvert.com
lo generan a partir de un `.png`).

## Novedades (v7) — corrección definitiva del ícono
La versión anterior dependía de `--add-data` de PyInstaller para empaquetar
`icono.ico`/`icono.png` dentro del `.exe`, y en algunas instalaciones de
PyInstaller ese paso no funciona bien (el archivo queda afuera del paquete).

Ahora el ícono está **incrustado directamente en el código** como texto
base64 (`icono_datos.py`), así que no depende de ningún paso de empaquetado
de archivos: siempre está disponible, tanto corriendo `app.py` directo como
en el `.exe` compilado.

## Si el ícono se ve bien en la barra de tareas pero NO en el Explorador

Esto es el **caché de íconos de Windows** (guarda una "foto" del ícono de
cada archivo y no siempre la actualiza). Tres formas de solucionarlo, de
más simple a más a fondo:

1. **Más simple**: copiá o mové el `.exe` a otra carpeta, o renombralo.
   Windows va a tener que generar el ícono de nuevo para ese "archivo nuevo".
2. **Reiniciar el Explorador**: Ctrl+Shift+Esc (Administrador de tareas) →
   buscá "Explorador de Windows" → clic derecho → Reiniciar.
3. **Forzar el borrado completo del caché** (si lo anterior no alcanza):
   abrí el Símbolo del sistema como administrador y pegá:
   ```
   taskkill /f /im explorer.exe
   DEL /A /Q "%localappdata%\IconCache.db"
   DEL /A /F /Q "%localappdata%\Microsoft\Windows\Explorer\iconcache*"
   start explorer.exe
   ```

## Novedades (v8) — Tema oscuro
Toda la interfaz (ventana, tabla, botones, campos de texto, diálogos) ahora
usa una paleta oscura por defecto, con el azul del ícono como color de
acento. Si en algún momento querés volver al tema claro, buscá la función
`aplicar_tema_oscuro()` en `app.py` y comentá la línea que la llama en
`RenombradorApp.__init__` (o pedime que te agregue un botón para
alternar entre claro/oscuro).

## Novedades (v9) — Múltiples autores
Clic derecho sobre una fila → **"Editar autores..."** abre una ventana donde
podés cargar tantos autores como necesites (2, 3, 5, los que sean), en el
orden en que deben aparecer en la cita.

El formato APA 7 ahora arma automáticamente la lista completa de autores
según las reglas oficiales:
- 1 autor: `Apellido, I.`
- 2 autores: `Apellido1, I1. y Apellido2, I2.`
- 3-20 autores: `Apellido1, I1., Apellido2, I2., ... y ApellidoN, IN.`
- 21+ autores: los primeros 19, puntos suspensivos, y el último

Esto se controla con el nuevo placeholder `{autores}` en la plantilla de
nomenclatura (ya viene usado en los presets de APA 7). Si tu plantilla
personalizada usa `{nombre}`/`{apellido}`, esos siguen refiriéndose solo al
primer autor (por compatibilidad con formatos que no necesitan listarlos a
todos).

Además, la búsqueda automática (Google Books/Crossref) ahora reconoce y
guarda TODOS los autores que encuentre, no solo el primero.

## Novedades (v10) — Edición más rápida
- **Edición directa en la celda**: doble clic sobre Nombre, Apellido, Título
  o Año abre un campo de texto justo ahí encima (tipo planilla), en vez de
  una ventana emergente. Enter o Tab confirma, Escape cancela.
- **Acceso rápido a la búsqueda manual**: apareció una columna con un
  ícono 🔍 al lado de "Nuevo nombre" - un clic ahí abre la búsqueda manual
  para esa fila. Doble clic sobre la celda de "Nuevo nombre" hace lo mismo.

## Novedades (v11) — Elegir autores fácil
Arriba de la tabla, al lado del "Preset", hay un nuevo checkbox
**"Incluir todos los autores"**:
- **Marcado** (default): el nombre incluye a todos los autores cargados,
  con formato APA (`Pérez, J., Gómez, M. y Rodríguez, L.`).
- **Desmarcado**: el nombre usa solo el primer autor
  (`Pérez, J.`).

Cambia automáticamente la plantilla y recalcula todos los nombres, sin
tener que editar el texto de la plantilla a mano.

## Novedades (v12) — Mejoras estéticas y de flujo de trabajo

- **Colores por estado**: la columna Estado ahora se ve en verde (OK),
  amarillo (SIN CONFIRMAR) o rojo (SIN DATOS/ERROR) para detectar de un
  vistazo qué filas necesitan revisión.
- **Filas en zebra**: fondo alternado entre filas para seguir mejor una fila
  larga con la vista.
- **Ordenar la tabla**: clic en cualquier encabezado (Archivo original,
  Nombre, Apellido, Título, Año, Nuevo nombre, Estado) ordena por esa
  columna; otro clic invierte el orden. Aparece una flechita ▲/▼ indicando
  el orden activo.
- **Arrastrar y soltar la carpeta**: además de "Elegir carpeta...", podés
  arrastrar la carpeta directo a la ventana. Requiere el paquete
  `tkinterdnd2` (ya está en `requirements.txt`); si por algún motivo no se
  instala bien, la app sigue funcionando normal con el botón de elegir
  carpeta, solo que sin esta opción extra.
- **Tab entre campos al editar**: mientras editás Nombre, Apellido, Título
  o Año en la celda, Tab guarda y pasa directo al siguiente campo (Shift+Tab
  va para atrás), sin tener que hacer doble clic en cada uno.
- **Copiar el nombre original**: doble clic sobre la celda de "Archivo
  original" copia ese texto al portapapeles.
- **Abrir el archivo original**: nueva columna con ícono 📂 al lado del
  nombre de archivo - un clic abre el PDF con el visor predeterminado de
  Windows.

## Novedades (v13) — Ajustes de plantilla y detalles de UI

- El preset **"Personalizado"** ahora arma el nombre con
  **Apellido primero y nombre completo** (no inicial):
  - 1 autor: `Apellido, Nombre - Título (Año)`
  - Varios: `Apellido1, Nombre1 y Apellido2, Nombre2 - Título (Año)`
  - Nuevo placeholder `{autores_completo}` para esto (a diferencia de
    `{autores}`, que sigue siendo formato APA con iniciales).
- El checkbox **"Incluir todos los autores"** ahora arranca **desmarcado**.
- El botón de análisis ahora dice **"▶ Analizar PDFs"**.

## Novedades (v14) — Barra de título oscura (Windows)
En Windows 10/11, la barra de título (donde están minimizar/cerrar) ahora
también se pone en modo oscuro, usando la API de DWM de Windows vía
`ctypes` (Tkinter no lo soporta nativamente). En Windows más viejos, o si
por algún motivo la API no está disponible, la barra vuelve a su color
normal sin romper nada más de la app.

## Novedades (v15) — Soporte para EPUB y DOCX
Además de PDF, la app ahora también analiza y renombra archivos **EPUB** y
**DOCX**. No hace falta ninguna librería nueva: ambos formatos son en
realidad archivos ZIP con sus metadatos (título, autor, y en el caso de
EPUB a veces el ISBN) en formato XML adentro, así que se leen con lo que ya
trae Python.

- Al elegir una carpeta, ahora se detectan `.pdf`, `.epub` y `.docx` juntos.
- Cada archivo se renombra **manteniendo su extensión original** (un EPUB
  sigue siendo `.epub`, no se fuerza a `.pdf`).
- La búsqueda de ISBN también funciona en EPUB (a veces viene en los
  metadatos del propio archivo, sin ni siquiera necesitar buscarlo en el
  texto).

## Novedades (v16) — Rebranding a CanónicaDL
- La app pasa a llamarse **CanónicaDL**, con ícono nuevo (el isotipo de la
  "C" tipo estante de libros).
- Barra de marca nueva arriba de todo: **"CanónicaDL - Toolkit"** a la
  izquierda, **"[seb@lies / 2026]"** a la derecha.
- El ejecutable ahora se llama **`CanonicaDL.exe`**.
- El botón de análisis dice **"▶ Analizar Archivos"** (ya no dice
  "Analizar PDFs", porque también soporta EPUB y DOCX).

## Novedades (v17) — Ventana "Acerca de"
Arriba a la derecha, junto a la firma, hay un nuevo botón **"ⓘ Acerca de"**
que abre una ventana con el logo completo, una descripción de qué es la
app y un link a sebastianalies.com (clickeable, abre en el navegador).

## Novedades (v18) — Correcciones importantes

- **Bug corregido: el estado no se actualizaba tras editar a mano.** Si
  completabas Nombre/Apellido/Título en la celda, el dato quedaba bien pero
  la fila seguía en rojo "SIN DATOS" (por eso parecía que la app "no
  reconocía nada" aunque los datos estuvieran ahí). Ahora, al completar los
  campos a mano, el estado pasa a "OK (manual)".
- **Prefijos numéricos ya no rompen la detección.** Archivos numerados
  como `4_Marradi...`, `12-Piovani...` (numeración personal al principio
  del nombre) confundían al detector de autor, porque cualquier segmento
  con dígitos se descartaba como "no es un nombre". Ahora ese prefijo se
  quita antes de analizar.
- **Nuevo patrón reconocido: "Autor, Título".** Antes solo se reconocía
  `Título - Autor` o con guion bajo. Ahora también funciona con coma, por
  ejemplo `Ruth Sautu, Manual de metodologia.pdf`.

### Limitación conocida (sin resolver todavía)
Nombres de archivo con **más de un autor unidos por "y"** dentro del mismo
segmento (ej. `Marradi Archenti y Piovani - Metodología...`) todavía no se
separan automáticamente, porque la palabra "y" también es muy común en
títulos y no hay forma confiable de distinguir ambos casos solo con el
nombre del archivo. En esos casos, cargá los autores a mano con clic
derecho → "Editar autores...".

## Novedades (v19) — Migración a CustomTkinter
Toda la interfaz se reconstruyó con **CustomTkinter** (look más moderno:
botones redondeados, mejor tipografía, controles más prolijos), manteniendo
el 100% de la lógica de negocio sin cambios (búsquedas, parseo de nombres,
formato APA, renombrado, etc. - nada de eso se tocó).

**Qué sigue igual (por elección técnica, no por limitación):**
- La **tabla de resultados** sigue siendo `ttk.Treeview`, porque
  CustomTkinter no tiene un widget de tabla propio con columnas ordenables,
  colores por fila y edición en la celda como el que ya teníamos. Está
  estilizada para combinar con el resto.
- El **menú contextual** (clic derecho) y la **lista de resultados** del
  diálogo de búsqueda manual siguen siendo widgets clásicos de Tk, por el
  mismo motivo (CustomTkinter no los reemplaza).

**Nueva dependencia:** `customtkinter` (se agregó a `requirements.txt` y
al comando de PyInstaller en `generar_exe.bat`). El `.exe` va a pesar más
que antes (CustomTkinter agrega su propio motor de temas e íconos) - es un
cambio esperado, no un error.

## Novedades (v20) — Ocho mejoras de flujo de trabajo

1. **Quitar filas con Delete** - seleccioná una o varias filas y tocá
   Delete (o clic derecho → "Quitar de la lista") para sacarlas de la
   tabla. No borra los archivos del disco, solo los saca de la lista.
2. **Editar autores en varias filas a la vez** - seleccioná varias filas,
   clic derecho → "Editar autores...", y los autores que cargues se
   aplican a todas las seleccionadas.
3. **Columna "Año" centrada.**
4. **Nueva columna "Formato"** entre "Archivo original" y el ícono 📂,
   mostrando la extensión (PDF/EPUB/DOCX).
5. **Contador de archivos**: apenas elegís una carpeta (por botón o
   arrastrando), aparece cuántos archivos compatibles hay ahí, antes de
   tocar "Analizar Archivos".
6. **Navegación tipo planilla**: las flechas ← → cambian qué columna se
   edita con Enter (se marca con ▸ en el encabezado), ↑ ↓ cambian de fila.
7. **Toda la ayuda unificada en "Acerca de"**, ahora con pestañas (Info /
   Atajos / Plantillas) en vez de estar repartida en textos sueltos por
   toda la ventana.
8. **Guardar presets propios**: botón "💾 Guardar como preset..." guarda
   tu plantilla actual con un nombre (aparece con ★ en el desplegable), y
   queda guardada en disco entre sesiones. Botón 🗑 para eliminarlos.

## Notas

- La app nunca sobrescribe archivos: si ya existe un archivo con el nuevo
  nombre, agrega un "(1)", "(2)", etc.
- No modifica el contenido de los PDFs, solo el nombre del archivo.
- Si un libro no tiene ISBN legible (por ejemplo, escaneos de mala calidad),
  la detección por texto puede fallar y conviene revisar manualmente esa fila.
