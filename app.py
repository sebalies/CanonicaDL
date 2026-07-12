#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CanónicaDL - Toolkit
-------------------------------------------------
Analiza los archivos PDF/EPUB/DOCX de una carpeta, busca información
bibliográfica en internet (Open Library / Google Books / Crossref) y los
renombra siguiendo un formato configurable (personalizado o APA 7),
manteniendo la extensión original de cada archivo.

Autor: Claude (Anthropic) para Sebalies
"""

import os
import re
import sys
import json
import queue
import base64
import zipfile
import tempfile
import threading
import subprocess
import unicodedata
import webbrowser
import xml.etree.ElementTree as ET
from urllib.parse import quote_plus
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

import requests

try:
    from pypdf import PdfReader
except ImportError:
    from PyPDF2 import PdfReader  # fallback por si está instalada la versión vieja

try:
    from icono_datos import ICONO_ICO_BASE64, ICONO_PNG_BASE64, LOGO_COMPLETO_BASE64
except ImportError:
    ICONO_ICO_BASE64 = ""
    ICONO_PNG_BASE64 = ""
    LOGO_COMPLETO_BASE64 = ""

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_DISPONIBLE = True
except ImportError:
    DND_DISPONIBLE = False

import customtkinter as ctk


def ruta_recurso(nombre_archivo):
    """
    Devuelve la ruta a un recurso (ícono, etc.), funcionando tanto corriendo
    el script directamente como empaquetado en un .exe con PyInstaller
    (donde los archivos agregados con --add-data se extraen a una carpeta
    temporal accesible vía sys._MEIPASS).
    """
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, nombre_archivo)


def ruta_junto_al_ejecutable(nombre_archivo):
    """
    Devuelve una ruta persistente junto al .exe real (o al script, en modo
    desarrollo). A diferencia de ruta_recurso(), esto NO apunta a la carpeta
    temporal de PyInstaller (que se borra al cerrar la app), así que sirve
    para archivos que necesitamos que sobrevivan después de cerrar la app
    (ej. un log de errores).
    """
    if getattr(sys, "frozen", False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, nombre_archivo)


# --------------------------------------------------------------------------
# Utilidades de extracción y búsqueda
# --------------------------------------------------------------------------

ISBN_REGEX = re.compile(
    r'(?:ISBN(?:-1[03])?:?\s*)?((?:97[89][-\s]?)?\d{1,5}[-\s]?\d{1,7}[-\s]?\d{1,7}[-\s]?[\dXx])'
)


def extraer_texto_y_metadatos_pdf(ruta_pdf, max_paginas=8):
    """Devuelve (texto_extraido, metadatos_dict) de las primeras páginas de un PDF."""
    texto = ""
    metadatos = {}
    try:
        reader = PdfReader(ruta_pdf)
        n_paginas = min(max_paginas, len(reader.pages))
        for i in range(n_paginas):
            try:
                texto += reader.pages[i].extract_text() or ""
                texto += "\n"
            except Exception:
                continue
        if reader.metadata:
            metadatos = {
                "titulo": (reader.metadata.title or "").strip(),
                "autor": (reader.metadata.author or "").strip(),
            }
    except Exception:
        texto = ""
        metadatos = {}
    return texto, metadatos


def extraer_texto_y_metadatos_epub(ruta_epub):
    """
    Devuelve (texto, metadatos) de un EPUB. Un EPUB es un .zip que contiene
    un archivo .opf con los metadatos (título, autor, identificadores como
    ISBN) en formato Dublin Core - no hace falta ninguna librería extra,
    con zipfile + xml de la librería estándar alcanza.
    """
    texto = ""
    metadatos = {"titulo": "", "autor": ""}
    try:
        with zipfile.ZipFile(ruta_epub, "r") as z:
            container = ET.fromstring(z.read("META-INF/container.xml"))
            ns_contenedor = {"c": "urn:oasis:names:tc:opendocument:xmlns:container"}
            rootfile = container.find(".//c:rootfile", ns_contenedor)
            if rootfile is None:
                return texto, metadatos
            ruta_opf = rootfile.get("full-path")

            opf = ET.fromstring(z.read(ruta_opf))
            ns_opf = {
                "opf": "http://www.idpf.org/2007/opf",
                "dc": "http://purl.org/dc/elements/1.1/",
            }
            titulo_elem = opf.find(".//dc:title", ns_opf)
            autor_elem = opf.find(".//dc:creator", ns_opf)
            if titulo_elem is not None and titulo_elem.text:
                metadatos["titulo"] = titulo_elem.text.strip()
            if autor_elem is not None and autor_elem.text:
                metadatos["autor"] = autor_elem.text.strip()

            # Los identificadores (a veces incluyen el ISBN) se agregan al
            # "texto" para que buscar_isbn() los encuentre igual que en un PDF
            for identificador in opf.findall(".//dc:identifier", ns_opf):
                if identificador.text:
                    texto += f" {identificador.text.strip()} "
    except Exception:
        pass
    return texto, metadatos


def extraer_texto_y_metadatos_docx(ruta_docx):
    """
    Devuelve (texto, metadatos) de un .docx. También es un .zip: los
    metadatos están en docProps/core.xml (mismo esquema Dublin Core que un
    EPUB) y el texto del documento en word/document.xml.
    """
    texto = ""
    metadatos = {"titulo": "", "autor": ""}
    try:
        with zipfile.ZipFile(ruta_docx, "r") as z:
            try:
                core = ET.fromstring(z.read("docProps/core.xml"))
                ns_core = {
                    "dc": "http://purl.org/dc/elements/1.1/",
                }
                titulo_elem = core.find("dc:title", ns_core)
                autor_elem = core.find("dc:creator", ns_core)
                if titulo_elem is not None and titulo_elem.text:
                    metadatos["titulo"] = titulo_elem.text.strip()
                if autor_elem is not None and autor_elem.text:
                    metadatos["autor"] = autor_elem.text.strip()
            except KeyError:
                pass

            try:
                ns_w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
                doc = ET.fromstring(z.read("word/document.xml"))
                fragmentos = [t.text for t in doc.iter(f"{ns_w}t") if t.text]
                texto = " ".join(fragmentos)[:8000]  # alcanza para buscar ISBN/portada
            except KeyError:
                pass
    except Exception:
        pass
    return texto, metadatos


def extraer_texto_y_metadatos(ruta_archivo, max_paginas=8):
    """Despacha a la función de extracción correcta según la extensión del archivo."""
    extension = os.path.splitext(ruta_archivo)[1].lower()
    if extension == ".pdf":
        return extraer_texto_y_metadatos_pdf(ruta_archivo, max_paginas)
    elif extension == ".epub":
        return extraer_texto_y_metadatos_epub(ruta_archivo)
    elif extension == ".docx":
        return extraer_texto_y_metadatos_docx(ruta_archivo)
    return "", {}


def buscar_isbn(texto):
    """Busca un ISBN plausible dentro del texto extraído."""
    candidatos = ISBN_REGEX.findall(texto)
    for c in candidatos:
        limpio = re.sub(r'[-\s]', '', c)
        if len(limpio) in (10, 13):
            return limpio
    return None


def consultar_openlibrary_por_isbn(isbn):
    """Consulta Open Library por ISBN. Devuelve dict con titulo/autor/anio o None."""
    try:
        url = f"https://openlibrary.org/api/books?bibkeys=ISBN:{isbn}&format=json&jscmd=data"
        r = requests.get(url, timeout=8)
        r.raise_for_status()
        data = r.json()
        clave = f"ISBN:{isbn}"
        if clave not in data:
            return None
        info = data[clave]
        titulo = info.get("title", "")
        autores = info.get("authors", [])
        autor = autores[0]["name"] if autores else ""
        editorial = ""
        publishers = info.get("publishers", [])
        if publishers:
            editorial = publishers[0].get("name", "")
        anio = ""
        fecha_pub = info.get("publish_date", "")
        m = re.search(r'\d{4}', fecha_pub)
        if m:
            anio = m.group(0)
        if titulo:
            return {"titulo": titulo, "autor": autor, "anio": anio, "editorial": editorial}
    except Exception:
        pass
    return None


def consultar_google_books(query):
    """Consulta Google Books con una query de texto libre. Devuelve dict o None."""
    resultados = consultar_google_books_multiples(query, max_resultados=1)
    return resultados[0] if resultados else None


def _candidato_base(**kwargs):
    """Crea un dict de candidato con todos los campos estandarizados (vacíos por defecto)."""
    base = {
        "titulo": "", "autor": "", "autores": [], "anio": "", "editorial": "",
        "revista": "", "volumen": "", "numero": "", "paginas": "", "doi": "",
        "tipo": "Libro", "fuente_nombre": "",
    }
    base.update(kwargs)
    return base


def consultar_google_books_multiples(query, max_resultados=5):
    """Consulta Google Books y devuelve una lista de candidatos (dicts)."""
    candidatos = []
    try:
        url = "https://www.googleapis.com/books/v1/volumes"
        params = {"q": query, "maxResults": max_resultados}
        r = requests.get(url, params=params, timeout=8)
        r.raise_for_status()
        data = r.json()
        items = data.get("items", [])
        for item in items:
            info = item.get("volumeInfo", {})
            titulo = info.get("title", "")
            if not titulo:
                continue
            subtitulo = info.get("subtitle", "")
            if subtitulo:
                titulo = f"{titulo}: {subtitulo}"
            autores_raw = info.get("authors", [])
            autor = autores_raw[0] if autores_raw else ""
            autores_lista = []
            for nombre_completo in autores_raw:
                n, a = dividir_nombre_apellido(nombre_completo)
                autores_lista.append({"nombre": n, "apellido": a})
            editorial = info.get("publisher", "")
            fecha = info.get("publishedDate", "")
            anio = ""
            m = re.search(r'\d{4}', fecha)
            if m:
                anio = m.group(0)
            candidatos.append(_candidato_base(
                titulo=titulo, autor=autor, autores=autores_lista, anio=anio, editorial=editorial,
                tipo="Libro", fuente_nombre="Google Books",
            ))
    except Exception:
        pass
    return candidatos


def consultar_crossref_multiples(query, max_resultados=5):
    """
    Consulta Crossref (base de datos académica libre: artículos, ensayos,
    capítulos, actas de congresos, etc. con DOI). Muy útil para textos que
    no son "libros" con ficha propia y por eso Google Books no encuentra.
    """
    candidatos = []
    try:
        url = "https://api.crossref.org/works"
        params = {"query": query, "rows": max_resultados}
        headers = {"User-Agent": "RenombradorPDF/1.0 (mailto:no-reply@example.com)"}
        r = requests.get(url, params=params, headers=headers, timeout=8)
        r.raise_for_status()
        data = r.json()
        items = data.get("message", {}).get("items", [])
        for item in items:
            titulos = item.get("title", [])
            titulo = titulos[0] if titulos else ""
            if not titulo:
                continue
            autores_raw = item.get("author", [])
            autor = ""
            autores_lista = []
            for a in autores_raw:
                apellido = a.get("family", "")
                nombre = a.get("given", "")
                autores_lista.append({"nombre": nombre, "apellido": apellido})
            if autores_raw:
                primero = autores_raw[0]
                apellido = primero.get("family", "")
                nombre = primero.get("given", "")
                if apellido and nombre:
                    autor = f"{apellido}, {nombre}"
                else:
                    autor = apellido or nombre

            anio = ""
            for campo_fecha in ("published-print", "published-online", "issued", "created"):
                partes_fecha = item.get(campo_fecha, {}).get("date-parts", [[]])
                if partes_fecha and partes_fecha[0] and partes_fecha[0][0]:
                    anio = str(partes_fecha[0][0])
                    break

            contenedor = item.get("container-title", [])
            revista = contenedor[0] if contenedor else ""
            volumen = item.get("volume", "")
            numero = item.get("issue", "")
            paginas = item.get("page", "")
            doi = item.get("DOI", "")
            editorial = item.get("publisher", "")

            candidatos.append(_candidato_base(
                titulo=titulo, autor=autor, autores=autores_lista, anio=anio, editorial=editorial,
                revista=revista, volumen=volumen, numero=numero, paginas=paginas, doi=doi,
                tipo="Artículo", fuente_nombre="Crossref",
            ))
    except Exception:
        pass
    return candidatos


def consultar_multiples_fuentes(query, max_resultados=5):
    """Combina resultados de Google Books y Crossref para una misma búsqueda."""
    return consultar_google_books_multiples(query, max_resultados) + consultar_crossref_multiples(query, max_resultados)


def abrir_busqueda_en_navegador(query):
    """Abre una búsqueda de Google en el navegador predeterminado."""
    url = f"https://www.google.com/search?q={quote_plus(query)}"
    webbrowser.open(url)


def primera_linea_util(texto, min_letras=6, max_lineas=25):
    """Devuelve la primera línea 'con pinta de título' del texto extraído."""
    lineas = [l.strip() for l in texto.splitlines() if l.strip()]
    for l in lineas[:max_lineas]:
        letras = sum(c.isalpha() for c in l)
        if letras >= min_letras and not l.lower().startswith(("isbn", "copyright", "©")):
            return l
    return ""


CONECTORES_MINUSCULA = {
    "y", "e", "o", "u", "de", "del", "la", "las", "el", "los",
    "en", "a", "con", "sin", "por", "para", "un", "una",
}


def _es_todo_mayusculas(texto):
    letras = [c for c in texto if c.isalpha()]
    return len(letras) > 0 and all(c.isupper() for c in letras)


def titlecase_inteligente(texto):
    """
    Convierte un texto tipo 'Manifesto-Antropofago-e-Manifesto-Da-Poesia-Pau'
    en 'Manifesto Antropofago e Manifesto Da Poesia Pau': separa palabras
    pegadas por guiones/underscores y capitaliza, dejando en minúscula los
    conectores comunes (salvo si son la primera palabra).

    Si el texto original viene TODO EN MAYÚSCULAS (ej. 'ARTE-EN-LA-RED'),
    se interpreta como "gritado" y se pasa a formato oración normal
    ('Arte en la red') en vez de poner mayúscula en cada palabra.
    """
    texto = re.sub(r'[_\-]+', ' ', texto)
    texto = re.sub(r'\s+', ' ', texto).strip()

    if _es_todo_mayusculas(texto):
        texto_minuscula = texto.lower()
        return texto_minuscula[:1].upper() + texto_minuscula[1:]

    palabras = texto.split(" ")
    resultado = []
    for i, palabra in enumerate(palabras):
        if not palabra:
            continue
        minuscula = palabra.lower()
        if i > 0 and minuscula in CONECTORES_MINUSCULA:
            resultado.append(minuscula)
        else:
            resultado.append(minuscula[:1].upper() + minuscula[1:])
    return " ".join(resultado)


def _quitar_prefijo_numerico(texto):
    """
    Quita prefijos de numeración personal al principio del nombre de archivo,
    ej. '4_Marradi...' -> 'Marradi...', '12-Piovani...' -> 'Piovani...'.
    Sin esto, ese número interfiere con la detección de autor (el chequeo
    de "esto tiene pinta de nombre" rechaza cualquier segmento con dígitos).
    """
    return re.sub(r'^\d+[\s_\-]+', '', texto)


def limpiar_nombre_para_busqueda(nombre_archivo):
    """Limpia el nombre de archivo (sin extensión) para usarlo como query o título de respaldo."""
    base = os.path.splitext(nombre_archivo)[0]
    base = _quitar_prefijo_numerico(base)
    return titlecase_inteligente(base)


PARTICULAS_NOMBRE = {"de", "del", "von", "van", "der", "di", "da"}

PALABRAS_DE_TITULO = {
    # palabras que casi nunca forman parte de un nombre de persona,
    # pero sí son muy comunes en títulos (español e inglés)
    "en", "la", "las", "el", "los", "y", "e", "o", "u", "a", "al",
    "con", "sin", "por", "para", "un", "una",
    "the", "of", "an", "and", "or", "in", "on", "for", "to", "with", "from", "by",
}


def formatear_nombre_persona(texto):
    """
    Da formato prolijo a un nombre de persona (Título en cada palabra),
    incluso si viene TODO EN MAYÚSCULAS o todo en minúsculas.
    A diferencia de titlecase_inteligente, acá SIEMPRE se capitaliza
    cada palabra (los nombres no llevan "formato oración").
    """
    texto = re.sub(r'[_\-]+', ' ', texto)
    texto = re.sub(r'\s+', ' ', texto).strip()
    palabras = texto.split(" ")
    resultado = []
    for p in palabras:
        if not p:
            continue
        p_lower = p.lower()
        if p_lower in PARTICULAS_NOMBRE:
            resultado.append(p_lower)
        else:
            resultado.append(p_lower[:1].upper() + p_lower[1:])
    return " ".join(resultado)


def _es_nombre_probable(segmento):
    """Heurística: ¿este texto tiene pinta de nombre de persona (y no de título)?"""
    segmento_limpio = re.sub(r'[_\-]+', ' ', segmento).strip()
    palabras = [p for p in segmento_limpio.split() if p]
    if not (2 <= len(palabras) <= 4):
        return False
    if any(c.isdigit() for c in segmento_limpio):
        return False
    for p in palabras:
        p_lower = p.lower()
        if p_lower in PALABRAS_DE_TITULO:
            return False
        if p_lower in PARTICULAS_NOMBRE:
            continue
        letras = re.sub(r'[^A-Za-zÀ-ÖØ-öø-ÿ]', '', p)
        if not letras:
            return False
        # Acepta "Jesús" (Title Case) o "JESÚS" (todo mayúsculas)
        formato_titlecase = letras[0].isupper() and (len(letras) == 1 or letras[1:].islower())
        formato_mayusculas = letras.isupper()
        if not (formato_titlecase or formato_mayusculas):
            return False
    return True


def separar_titulo_autor_de_filename(nombre_original):
    """
    Intenta separar Título y Autor a partir de patrones comunes en el nombre
    de archivo, probando en orden de confianza:
      1. 'Título - Autor'  (el más confiable: guion con espacios alrededor)
      2. 'Autor, Título'   (coma: acá el autor casi siempre va primero)
      3. 'Título_Autor'    (guion bajo: el autor suele quedar al final)
    Devuelve (titulo, autor) si encuentra un patrón confiable, o (None, None).
    """
    base = os.path.splitext(nombre_original)[0]
    base = _quitar_prefijo_numerico(base).strip()

    # 1) "Título - Autor" (por convención, si el segmento final "tiene
    #    pinta" de nombre de persona, se lo toma como autor)
    if " - " in base:
        a, b = (p.strip() for p in base.split(" - ", 1))
        if not (a and b):
            return None, None
        if _es_nombre_probable(b):
            return titlecase_inteligente(a), formatear_nombre_persona(b)
        if _es_nombre_probable(a):
            return titlecase_inteligente(b), formatear_nombre_persona(a)
        return None, None

    # 2) "Autor, Título" (con coma, la convención habitual pone el autor
    #    primero, ej. 'Ruth Sautu, Manual de metodologia')
    if ", " in base:
        a, b = (p.strip() for p in base.split(", ", 1))
        if not (a and b):
            return None, None
        if _es_nombre_probable(a):
            return titlecase_inteligente(b), formatear_nombre_persona(a)
        if _es_nombre_probable(b):
            return titlecase_inteligente(a), formatear_nombre_persona(b)
        return None, None

    # 3) "Título_Autor" (guion bajo, sin separador más claro)
    if "_" in base:
        izquierda, _, derecha = base.rpartition("_")
        if izquierda and derecha:
            if _es_nombre_probable(derecha):
                return titlecase_inteligente(izquierda), formatear_nombre_persona(derecha)
            if _es_nombre_probable(izquierda):
                return titlecase_inteligente(derecha), formatear_nombre_persona(izquierda)

    return None, None


MARCADORES_METADATO_BASURA = ("microsoft word", ".doc", ".docx", ".pdf", ".rtf")


def metadato_titulo_es_confiable(titulo, nombre_archivo):
    """Detecta títulos de metadata que en realidad son basura (ej. 'Microsoft Word - archivo.doc')."""
    if not titulo or not titulo.strip():
        return False
    titulo_lower = titulo.strip().lower()
    if any(marcador in titulo_lower for marcador in MARCADORES_METADATO_BASURA):
        return False
    base_archivo = os.path.splitext(nombre_archivo)[0].strip().lower()
    if titulo_lower == base_archivo:
        return False
    return True


def dividir_nombre_apellido(nombre_completo):
    """
    Intenta separar 'Nombre Apellido' en (nombre, apellido).
    Si viene en formato 'Apellido, Nombre' lo detecta y lo da vuelta.
    Heurística simple: si hay coma, ya viene separado.
    """
    nombre_completo = nombre_completo.strip()
    if not nombre_completo:
        return "", ""
    if "," in nombre_completo:
        partes = [p.strip() for p in nombre_completo.split(",", 1)]
        # Asumimos que ya viene como "Apellido, Nombre"
        apellido, nombre = partes[0], partes[1] if len(partes) > 1 else ""
        return nombre, apellido
    partes = nombre_completo.split()
    if len(partes) == 1:
        return "", partes[0]
    nombre = " ".join(partes[:-1])
    apellido = partes[-1]
    return nombre, apellido


def sanitizar_nombre_archivo(nombre):
    """Elimina caracteres no permitidos en nombres de archivo (Windows/Mac/Linux)."""
    nombre = unicodedata.normalize("NFC", nombre)
    nombre = re.sub(r'[\\/:*?"<>|]', "", nombre)
    nombre = re.sub(r'\s+', " ", nombre).strip()
    return nombre


# --------------------------------------------------------------------------
# Formatos de nomenclatura (combinables mediante plantilla)
# --------------------------------------------------------------------------

# Cada preset tiene dos variantes: "multi" (usa {autores}, con todos los
# autores en formato APA) y "solo1" (usa solo el primer autor). El checkbox
# "Incluir todos los autores" de la interfaz elige cuál de las dos usar,
# sin que el usuario tenga que tocar la plantilla a mano.
FORMATOS_PRESET_INFO = {
    "Personalizado: Apellido, Nombre - Título (Año)": {
        "multi": "{autores_completo} - {titulo} ({anio})",
        "solo1": "{apellido}, {nombre} - {titulo} ({anio})",
    },
    "APA 7 - Libro (sin editorial)": {
        "multi": "{autores} ({anio}). {titulo}.",
        "solo1": "{apellido}, {inicial} ({anio}). {titulo}.",
    },
    "APA 7 - Artículo (sin editorial)": {
        "multi": "{autores} ({anio}). {titulo}. {revista}, {volumen}({numero}), {paginas}.",
        "solo1": "{apellido}, {inicial} ({anio}). {titulo}. {revista}, {volumen}({numero}), {paginas}.",
    },
    "Solo Título (Año)": {
        "multi": "{titulo} ({anio})",
        "solo1": "{titulo} ({anio})",
    },
    "Apellido - Título": {
        "multi": "{autores} - {titulo}",
        "solo1": "{apellido} - {titulo}",
    },
}

# Mantenido por compatibilidad con cualquier código que espere el dict plano
# (usa la variante "multi" por defecto)
FORMATOS_PRESET = {nombre: info["multi"] for nombre, info in FORMATOS_PRESET_INFO.items()}
FORMATO_DEFECTO = "Personalizado: Apellido, Nombre - Título (Año)"

PLACEHOLDERS_DISPONIBLES = (
    "{autores} (todos, formato APA con iniciales) — {autores_completo} (todos, nombre completo) — "
    "{nombre} {apellido} {inicial} "
    "(solo el 1er autor) — {titulo} {anio} {editorial} {revista} {volumen} {numero} {paginas} {doi}"
)

NOMBRE_ARCHIVO_PRESETS = "presets_personalizados.json"
PREFIJO_PRESET_PERSONALIZADO = "★ "  # para distinguirlos de los presets de fábrica en el desplegable


def cargar_presets_personalizados():
    """Lee los presets guardados por el usuario desde un .json junto al ejecutable."""
    try:
        ruta = ruta_junto_al_ejecutable(NOMBRE_ARCHIVO_PRESETS)
        if not os.path.exists(ruta):
            return {}
        with open(ruta, "r", encoding="utf-8") as f:
            datos = json.load(f)
        return datos if isinstance(datos, dict) else {}
    except Exception:
        return {}


def guardar_presets_personalizados(presets):
    """Escribe el diccionario {nombre: plantilla} completo al .json."""
    try:
        ruta = ruta_junto_al_ejecutable(NOMBRE_ARCHIVO_PRESETS)
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump(presets, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def calcular_iniciales(nombre):
    """'Gabriel José' -> 'G. J.'"""
    partes = [p for p in nombre.split() if p]
    return " ".join(f"{p[0].upper()}." for p in partes)


def formatear_autores_apa(autores):
    """
    Formatea una lista de autores [{'nombre':.., 'apellido':..}, ...] siguiendo
    las reglas de APA 7 para múltiples autores:
    - 1 autor:  "Apellido, N."
    - 2 autores: "Apellido1, N1. & Apellido2, N2."
    - 3-20 autores: "Apellido1, N1., Apellido2, N2., ..., & ApellidoN, NN."
    - 21+ autores: primeros 19, "...", y el último (regla APA de "et al." extendido)
    """
    entradas = []
    for a in autores or []:
        apellido = (a.get("apellido") or "").strip()
        nombre = (a.get("nombre") or "").strip()
        inicial = calcular_iniciales(nombre)
        if apellido and inicial:
            entradas.append(f"{apellido}, {inicial}")
        elif apellido:
            entradas.append(apellido)
        elif nombre:
            entradas.append(nombre)

    if not entradas:
        return ""
    if len(entradas) == 1:
        return entradas[0]
    if len(entradas) == 2:
        return f"{entradas[0]} y {entradas[1]}"
    if len(entradas) <= 20:
        return ", ".join(entradas[:-1]) + " y " + entradas[-1]
    # 21+ autores: regla APA de listar 19, puntos suspensivos, y el último
    return ", ".join(entradas[:19]) + ", ... " + entradas[-1]


def formatear_autores_completo(autores):
    """
    Igual que formatear_autores_apa, pero usa el nombre completo de pila
    (no la inicial). Para el preset "Personalizado":
    - 1 autor:  "Apellido, Nombre"
    - 2 autores: "Apellido1, Nombre1 y Apellido2, Nombre2"
    - 3+ autores: "Apellido1, Nombre1, Apellido2, Nombre2 y ApellidoN, NombreN"
    """
    entradas = []
    for a in autores or []:
        apellido = (a.get("apellido") or "").strip()
        nombre = (a.get("nombre") or "").strip()
        if apellido and nombre:
            entradas.append(f"{apellido}, {nombre}")
        elif apellido:
            entradas.append(apellido)
        elif nombre:
            entradas.append(nombre)

    if not entradas:
        return ""
    if len(entradas) == 1:
        return entradas[0]
    if len(entradas) == 2:
        return f"{entradas[0]} y {entradas[1]}"
    return ", ".join(entradas[:-1]) + " y " + entradas[-1]


def _limpiar_artefactos_nombre(texto):
    """Limpia huecos que quedan cuando algún campo de la plantilla está vacío."""
    texto = re.sub(r'\(\s*\)', '', texto)          # paréntesis vacíos
    texto = re.sub(r',\s*,', ',', texto)             # comas dobles
    texto = re.sub(r',\s*-\s*', ' - ', texto)         # ", -" -> " - "
    texto = re.sub(r'\s{2,}', ' ', texto)             # espacios repetidos
    texto = re.sub(r'\s+([.,;:])', r'\1', texto)      # espacio antes de puntuación
    texto = re.sub(r'\.{2,}', '.', texto)             # puntos repetidos
    texto = re.sub(r'^[\s,\-\.]+', '', texto)         # basura colgante al inicio
    texto = re.sub(r'\s*\.\s*\.', '.', texto)         # ". ." -> "."
    texto = texto.strip(" -,")
    return texto.strip()


def construir_nuevo_nombre(r, template=None):
    """
    Construye el nombre de archivo final a partir del resultado analizado
    y una plantilla con placeholders, ej: '{apellido}, {inicial} ({anio}). {titulo}.'
    """
    if not template:
        template = FORMATOS_PRESET[FORMATO_DEFECTO]

    autores_lista = r.get("autores") or []
    if not autores_lista and (r.get("nombre") or r.get("apellido")):
        autores_lista = [{"nombre": r.get("nombre", ""), "apellido": r.get("apellido", "")}]

    valores = {
        "nombre": (r.get("nombre") or "").strip(),
        "apellido": (r.get("apellido") or "").strip(),
        "inicial": calcular_iniciales(r.get("nombre") or ""),
        "autores": formatear_autores_apa(autores_lista),
        "autores_completo": formatear_autores_completo(autores_lista),
        "titulo": (r.get("titulo") or "").strip() or "Sin titulo",
        "anio": (r.get("anio") or "").strip(),
        "editorial": (r.get("editorial") or "").strip(),
        "revista": (r.get("revista") or "").strip(),
        "volumen": (r.get("volumen") or "").strip(),
        "numero": (r.get("numero") or "").strip(),
        "paginas": (r.get("paginas") or "").strip(),
        "doi": (r.get("doi") or "").strip(),
    }
    try:
        nombre_generado = template.format(**valores)
    except Exception:
        # Si la plantilla tiene un placeholder inválido, usamos un formato de emergencia
        nombre_generado = f"{valores['apellido']} - {valores['titulo']} ({valores['anio']})"

    nombre_generado = _limpiar_artefactos_nombre(nombre_generado)
    if not nombre_generado:
        nombre_generado = valores["titulo"]
    if nombre_generado.endswith("."):
        nombre_generado = nombre_generado[:-1]

    # Mantener la extensión original del archivo (.pdf, .epub, .docx, etc.)
    nombre_original = r.get("nombre_original") or r.get("ruta") or ""
    extension = os.path.splitext(nombre_original)[1].lower() or ".pdf"

    return sanitizar_nombre_archivo(nombre_generado) + extension


# --------------------------------------------------------------------------
# Validación de resultados de búsqueda (evitar falsos positivos)
# --------------------------------------------------------------------------

PALABRAS_VACIAS_VALIDACION = (
    CONECTORES_MINUSCULA | PALABRAS_DE_TITULO | PARTICULAS_NOMBRE | {
        "que", "este", "esta", "estos", "estas", "sobre", "entre", "como",
        "desde", "hasta", "donde", "cuando", "sus", "mas", "muy",
        "this", "that", "are", "was", "were", "been", "have", "has", "had",
    }
)


def _palabras_significativas(texto):
    """Extrae palabras 'con contenido' (sin conectores) de un texto, para comparar similitud."""
    texto = unicodedata.normalize("NFKD", texto or "").encode("ascii", "ignore").decode("ascii")
    palabras = re.findall(r"[a-zA-Z]{3,}", texto.lower())
    return {p for p in palabras if p not in PALABRAS_VACIAS_VALIDACION}


def _mejor_candidato_validado(titulo_referencia, candidatos, umbral=0.7, min_palabras_referencia=3):
    """
    De una lista de candidatos, devuelve el que mejor coincide con el título
    de referencia, siempre que la coincidencia sea lo bastante fuerte.

    Si el título de referencia es muy corto/genérico (menos de
    `min_palabras_referencia` palabras con contenido, ej. "Arte en la Red"),
    NO se confía en ningún candidato: títulos así hacen "match" con
    cualquier cosa relacionada y eso lleva a resultados incorrectos.
    """
    if not titulo_referencia or not candidatos:
        return None
    palabras_ref = _palabras_significativas(titulo_referencia)
    if len(palabras_ref) < min_palabras_referencia:
        return None
    mejor, mejor_ratio = None, 0.0
    for c in candidatos:
        palabras_c = _palabras_significativas(c.get("titulo", ""))
        if not palabras_c:
            continue
        interseccion = palabras_ref & palabras_c
        ratio = len(interseccion) / min(len(palabras_ref), len(palabras_c))
        if ratio > mejor_ratio:
            mejor_ratio, mejor = ratio, c
    return mejor if mejor and mejor_ratio >= umbral else None


# --------------------------------------------------------------------------
# Análisis de un archivo (corre en hilo secundario)
# --------------------------------------------------------------------------

def analizar_archivo(ruta_pdf, template=None):
    """
    Devuelve dict con: nombre_original, nombre, apellido, titulo, anio,
    editorial, revista, volumen, numero, paginas, doi, nuevo_nombre, estado, fuente
    """
    nombre_original = os.path.basename(ruta_pdf)
    resultado = {
        "ruta": ruta_pdf,
        "nombre_original": nombre_original,
        "nombre": "",
        "apellido": "",
        "autores": [],
        "titulo": "",
        "anio": "",
        "editorial": "",
        "revista": "",
        "volumen": "",
        "numero": "",
        "paginas": "",
        "doi": "",
        "nuevo_nombre": "",
        "estado": "",
        "fuente": "",
    }

    texto, metadatos = extraer_texto_y_metadatos(ruta_pdf)
    nombre_archivo_limpio = limpiar_nombre_para_busqueda(nombre_original)
    titulo_arch, autor_arch = separar_titulo_autor_de_filename(nombre_original)

    metadato_titulo_ok = metadato_titulo_es_confiable(metadatos.get("titulo", ""), nombre_original)

    info = None
    fuente = ""

    # 1) Intentar por ISBN (más confiable: es un identificador único)
    isbn = buscar_isbn(texto)
    if isbn:
        info = consultar_openlibrary_por_isbn(isbn)
        if info:
            fuente = f"Open Library (ISBN {isbn})"

    # 2) Si el propio nombre de archivo ya trae "Título - Autor" reconocible,
    #    esa es la pista más confiable que tenemos: se usa directo (no se
    #    deja que una búsqueda ambigua la pise), y la búsqueda en internet
    #    solo se usa -si valida bien- para completar año/editorial/DOI.
    if not info and titulo_arch and autor_arch:
        info = _candidato_base(
            titulo=titulo_arch, autor=autor_arch, tipo="Libro", fuente_nombre="Nombre de archivo",
        )
        candidatos_extra = consultar_multiples_fuentes(f"{titulo_arch} {autor_arch}", max_resultados=4)
        mejor = _mejor_candidato_validado(titulo_arch, candidatos_extra)
        if mejor:
            info["anio"] = mejor.get("anio", "")
            info["editorial"] = mejor.get("editorial", "")
            info["revista"] = mejor.get("revista", "")
            info["volumen"] = mejor.get("volumen", "")
            info["numero"] = mejor.get("numero", "")
            info["paginas"] = mejor.get("paginas", "")
            info["doi"] = mejor.get("doi", "")
            info["fuente_nombre"] = f"Nombre de archivo + {mejor.get('fuente_nombre', '')}"
        fuente = info["fuente_nombre"]

    # 3) Si no hay ISBN ni patrón confiable en el nombre de archivo, probar
    #    varias consultas por texto, validando cada resultado contra un
    #    título de referencia antes de aceptarlo (para evitar falsos
    #    positivos con títulos cortos/genéricos).
    if not info:
        candidatas_query = []

        if metadato_titulo_ok:
            q = metadatos["titulo"]
            if metadatos.get("autor"):
                q += f" {metadatos['autor']}"
            candidatas_query.append(q)

        primera_linea = primera_linea_util(texto)
        if primera_linea:
            candidatas_query.append(primera_linea)

        if nombre_archivo_limpio:
            candidatas_query.append(nombre_archivo_limpio)

        titulo_referencia = (metadatos.get("titulo") if metadato_titulo_ok else None) or nombre_archivo_limpio

        for q in candidatas_query:
            candidatos = consultar_multiples_fuentes(q, max_resultados=4)
            mejor = _mejor_candidato_validado(titulo_referencia, candidatos)
            if mejor:
                info = mejor
                fuente = mejor.get("fuente_nombre", "")
                break

    if info and info.get("fuente_nombre") == "Nombre de archivo":
        # Guess de archivo sin validar por web: confiable pero sin confirmar
        nombre, apellido = dividir_nombre_apellido(info.get("autor", ""))
        resultado["nombre"] = nombre
        resultado["apellido"] = apellido
        resultado["autores"] = info.get("autores") or ([{"nombre": nombre, "apellido": apellido}] if (nombre or apellido) else [])
        resultado["titulo"] = info.get("titulo", "")
        resultado["anio"] = info.get("anio", "")
        resultado["estado"] = "SIN CONFIRMAR (detectado en nombre de archivo)"
        resultado["fuente"] = fuente
    elif info:
        nombre, apellido = dividir_nombre_apellido(info.get("autor", ""))
        resultado["nombre"] = nombre
        resultado["apellido"] = apellido
        resultado["autores"] = info.get("autores") or ([{"nombre": nombre, "apellido": apellido}] if (nombre or apellido) else [])
        resultado["titulo"] = info.get("titulo", "")
        resultado["anio"] = info.get("anio", "")
        resultado["editorial"] = info.get("editorial", "")
        resultado["revista"] = info.get("revista", "")
        resultado["volumen"] = info.get("volumen", "")
        resultado["numero"] = info.get("numero", "")
        resultado["paginas"] = info.get("paginas", "")
        resultado["doi"] = info.get("doi", "")
        resultado["estado"] = "OK"
        resultado["fuente"] = fuente
    else:
        # Sin resultados confiables: usar lo que haya en metadatos válidos
        # del PDF, y si no, el nombre de archivo ya limpio
        titulo_metadato = metadatos.get("titulo", "") if metadato_titulo_ok else ""
        autor_metadato = metadatos.get("autor", "") if metadato_titulo_ok else ""
        nombre, apellido = dividir_nombre_apellido(autor_metadato)
        resultado["nombre"] = nombre
        resultado["apellido"] = apellido
        resultado["autores"] = [{"nombre": nombre, "apellido": apellido}] if (nombre or apellido) else []
        resultado["titulo"] = titulo_metadato or nombre_archivo_limpio or nombre_original
        resultado["anio"] = ""
        resultado["estado"] = "SIN DATOS (revisar manualmente)"
        resultado["fuente"] = "—"

    resultado["nuevo_nombre"] = construir_nuevo_nombre(resultado, template)
    return resultado


# --------------------------------------------------------------------------
# Interfaz gráfica
# --------------------------------------------------------------------------

# --------------------------------------------------------------------------
# Tema oscuro
# --------------------------------------------------------------------------

TEMA_OSCURO = {
    "bg": "#1e1e1e",
    "bg_alt": "#252526",
    "bg_widget": "#2d2d30",
    "bg_entry": "#3c3c3c",
    "fg": "#e0e0e0",
    "fg_dim": "#9a9a9a",
    "accent": "#5a6fb0",     # azul, en sintonía con el ícono
    "accent_hover": "#6f85c4",
    "select_bg": "#3a4a78",
    "border": "#45454a",
}


def aplicar_tema_oscuro(style, root):
    """Aplica una paleta oscura a todos los widgets ttk (y al root, que es tk plano)."""
    t = TEMA_OSCURO
    style.theme_use("clam")  # base necesaria para poder personalizar colores en Windows

    root.configure(bg=t["bg"])

    style.configure(".", background=t["bg"], foreground=t["fg"], fieldbackground=t["bg_entry"],
                     bordercolor=t["border"], lightcolor=t["bg_widget"], darkcolor=t["bg_widget"])

    style.configure("TFrame", background=t["bg"])
    style.configure("TLabel", background=t["bg"], foreground=t["fg"])
    style.configure("TLabelframe", background=t["bg"], foreground=t["fg"], bordercolor=t["border"])
    style.configure("TLabelframe.Label", background=t["bg"], foreground=t["fg"])

    style.configure("TButton", background=t["bg_widget"], foreground=t["fg"],
                     bordercolor=t["border"], focuscolor=t["accent"])
    style.map("TButton",
              background=[("active", t["accent_hover"]), ("pressed", t["accent"])],
              foreground=[("disabled", t["fg_dim"])])

    style.configure("TEntry", fieldbackground=t["bg_entry"], foreground=t["fg"],
                     insertcolor=t["fg"], bordercolor=t["border"])
    style.map("TEntry", fieldbackground=[("readonly", t["bg_widget"])])

    style.configure("TCombobox", fieldbackground=t["bg_entry"], background=t["bg_widget"],
                     foreground=t["fg"], arrowcolor=t["fg"], bordercolor=t["border"])
    style.map("TCombobox",
              fieldbackground=[("readonly", t["bg_entry"])],
              foreground=[("readonly", t["fg"])])
    root.option_add("*TCombobox*Listbox.background", t["bg_widget"])
    root.option_add("*TCombobox*Listbox.foreground", t["fg"])
    root.option_add("*TCombobox*Listbox.selectBackground", t["select_bg"])

    style.configure("TCheckbutton", background=t["bg"], foreground=t["fg"])
    style.map("TCheckbutton", background=[("active", t["bg"])])

    style.configure("Treeview", background=t["bg_widget"], fieldbackground=t["bg_widget"],
                     foreground=t["fg"], bordercolor=t["border"], borderwidth=0)
    style.configure("Treeview.Heading", background=t["bg_alt"], foreground=t["fg"], relief="flat")
    style.map("Treeview.Heading", background=[("active", t["accent_hover"])])
    style.map("Treeview",
              background=[("selected", t["select_bg"])],
              foreground=[("selected", "#ffffff")])

    style.configure("TScrollbar", background=t["bg_widget"], troughcolor=t["bg"],
                     arrowcolor=t["fg"], bordercolor=t["border"])
    style.map("TScrollbar", background=[("active", t["accent_hover"])])

    style.configure("TProgressbar", background=t["accent"], troughcolor=t["bg_widget"],
                     bordercolor=t["border"])

    style.configure("TSeparator", background=t["border"])


def widget_oscuro_kwargs():
    """Kwargs de color para widgets tk 'planos' con foco (ej. Listbox)."""
    t = TEMA_OSCURO
    return {
        "bg": t["bg_entry"], "fg": t["fg"],
        "selectbackground": t["select_bg"], "selectforeground": "#ffffff",
        "highlightthickness": 1, "highlightbackground": t["border"], "highlightcolor": t["accent"],
    }


def menu_oscuro_kwargs():
    """Kwargs de color para tk.Menu (no acepta las mismas opciones que Listbox/Entry)."""
    t = TEMA_OSCURO
    return {
        "bg": t["bg_widget"], "fg": t["fg"],
        "activebackground": t["accent"], "activeforeground": "#ffffff",
        "borderwidth": 0,
    }


def configurar_apariencia_ctk():
    """Configura CustomTkinter para que coincida con la paleta TEMA_OSCURO."""
    # CustomTkinter activa automáticamente un manejo de "DPI awareness" en
    # Windows que puede hacer que los diálogos nativos de Tk (elegir
    # carpeta, avisos, etc.) se rendericen invisibles o fuera de pantalla.
    # Lo desactivamos para que esos diálogos sigan funcionando normal.
    try:
        ctk.deactivate_automatic_dpi_awareness()
    except Exception:
        pass

    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    t = TEMA_OSCURO
    # Sobreescribimos algunos valores del tema por defecto de CTk para que
    # combine con nuestra paleta (en vez del azul/gris genérico de CTk)
    ctk.ThemeManager.theme["CTkFrame"]["fg_color"] = [t["bg"], t["bg"]]
    ctk.ThemeManager.theme["CTkFrame"]["top_fg_color"] = [t["bg_widget"], t["bg_widget"]]
    ctk.ThemeManager.theme["CTkButton"]["fg_color"] = [t["accent"], t["accent"]]
    ctk.ThemeManager.theme["CTkButton"]["hover_color"] = [t["accent_hover"], t["accent_hover"]]
    ctk.ThemeManager.theme["CTkEntry"]["fg_color"] = [t["bg_entry"], t["bg_entry"]]
    ctk.ThemeManager.theme["CTkEntry"]["border_color"] = [t["border"], t["border"]]
    ctk.ThemeManager.theme["CTkOptionMenu"]["fg_color"] = [t["bg_widget"], t["bg_widget"]]
    ctk.ThemeManager.theme["CTkOptionMenu"]["button_color"] = [t["accent"], t["accent"]]
    ctk.ThemeManager.theme["CTkOptionMenu"]["button_hover_color"] = [t["accent_hover"], t["accent_hover"]]
    ctk.ThemeManager.theme["CTkCheckBox"]["fg_color"] = [t["accent"], t["accent"]]
    ctk.ThemeManager.theme["CTkCheckBox"]["hover_color"] = [t["accent_hover"], t["accent_hover"]]
    ctk.ThemeManager.theme["CTkProgressBar"]["progress_color"] = [t["accent"], t["accent"]]
    ctk.ThemeManager.theme["CTkScrollableFrame"]["fg_color"] = [t["bg"], t["bg"]]


if DND_DISPONIBLE:
    class _BaseApp(ctk.CTk, TkinterDnD.DnDWrapper):
        """CustomTkinter + soporte de arrastrar y soltar (tkinterdnd2)."""
        def __init__(self):
            super().__init__()
            self.TkdndVersion = TkinterDnD._require(self)
else:
    _BaseApp = ctk.CTk


class RenombradorApp(_BaseApp):
    def __init__(self):
        configurar_apariencia_ctk()
        super().__init__()
        self.title("CanónicaDL - Toolkit")
        self.geometry("1150x600")
        self.minsize(900, 500)

        self.style = ttk.Style(self)
        aplicar_tema_oscuro(self.style, self)

        self.carpeta = tk.StringVar()
        self.incluir_subcarpetas = tk.BooleanVar(value=False)
        self.resultados = {}  # item_id -> dict
        self.cola = queue.Queue()
        self.incluir_todos_autores = tk.BooleanVar(value=False)
        self.presets_personalizados = cargar_presets_personalizados()
        self.formato_var = tk.StringVar(value=FORMATOS_PRESET_INFO[FORMATO_DEFECTO]["solo1"])
        self.preset_var = tk.StringVar(value=FORMATO_DEFECTO)

        self._construir_ui()
        self._configurar_drag_and_drop()
        self.after(150, self._procesar_cola)

        # Actualiza el contador de archivos encontrados cada vez que cambia
        # la carpeta o se activa/desactiva "incluir subcarpetas"
        self.carpeta.trace_add("write", self._actualizar_conteo_carpeta)
        self.incluir_subcarpetas.trace_add("write", self._actualizar_conteo_carpeta)

        # El ícono se configura al final, con la ventana ya armada, y se
        # fuerza su aplicación con update_idletasks(): en algunas versiones
        # de Windows/Tk, si se hace demasiado temprano (antes de que la
        # ventana tenga su handle nativo listo), la llamada "funciona" sin
        # error pero el ícono no queda aplicado visualmente.
        self._configurar_icono()
        self._configurar_barra_titulo_oscura()
        self.update_idletasks()
        # Reintento con un pequeño delay: en algunos casos Windows necesita
        # que la ventana ya esté mapeada en pantalla para que el ícono de
        # la barra de tareas (y la barra de título oscura) tomen efecto.
        self.after(300, self._configurar_icono)
        self.after(300, self._configurar_barra_titulo_oscura)

    # ---------------------------- UI ----------------------------
    def _configurar_drag_and_drop(self):
        """Permite arrastrar una carpeta a la ventana para cargarla (requiere tkinterdnd2)."""
        if not DND_DISPONIBLE:
            return
        try:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._al_soltar_carpeta)
        except Exception:
            pass

    def _al_soltar_carpeta(self, event):
        try:
            rutas = self.tk.splitlist(event.data)
        except Exception:
            rutas = [event.data] if event.data else []
        if not rutas:
            return
        ruta = rutas[0]
        if os.path.isdir(ruta):
            self.carpeta.set(ruta)
        elif os.path.isfile(ruta):
            self.carpeta.set(os.path.dirname(ruta))
        else:
            return
        self.etiqueta_estado.configure(text=f"Carpeta cargada: {self.carpeta.get()}")

    def _mostrar_acerca_de(self):
        AcercaDeDialog(self)

    def _configurar_barra_titulo_oscura(self):
        """
        En Windows 10/11, pone la barra de título (donde están los botones
        minimizar/cerrar) en modo oscuro usando la API de DWM. Tkinter no
        tiene soporte nativo para esto, así que se llama directo a la DLL
        de Windows con ctypes. En otros sistemas operativos no hace nada.
        """
        if not sys.platform.startswith("win"):
            return
        try:
            import ctypes
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            valor = ctypes.c_int(1)  # 1 = activar modo oscuro
            # El atributo cambió de número entre versiones de Windows 10;
            # probamos las dos posibles.
            for atributo in (20, 19):
                resultado = ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd, atributo, ctypes.byref(valor), ctypes.sizeof(valor)
                )
                if resultado == 0:  # 0 = éxito
                    break
        except Exception:
            pass

    def _configurar_icono(self):
        """
        Carga el ícono de la app desde datos incrustados en base64 (ver
        icono_datos.py), en vez de depender de archivos externos empaquetados
        con --add-data. Funciona igual en modo desarrollo y en el .exe.
        """
        # iconphoto: Tk soporta PNG en base64 directo con el parámetro data=
        try:
            if ICONO_PNG_BASE64:
                self._icono_img = tk.PhotoImage(data=ICONO_PNG_BASE64, format="png")  # referencia viva
                self.iconphoto(True, self._icono_img)
        except Exception:
            pass

        # iconbitmap necesita un archivo .ico real: lo escribimos a un
        # archivo temporal a partir de los bytes decodificados.
        try:
            if ICONO_ICO_BASE64:
                datos_ico = base64.b64decode(ICONO_ICO_BASE64)
                temp = tempfile.NamedTemporaryFile(suffix=".ico", delete=False)
                temp.write(datos_ico)
                temp.close()
                self._icono_ico_temp = temp.name  # referencia para que no se borre antes de tiempo
                try:
                    self.iconbitmap(default=self._icono_ico_temp)
                except Exception:
                    try:
                        self.iconbitmap(self._icono_ico_temp)
                    except Exception:
                        pass
        except Exception:
            pass

    def _construir_ui(self):
        t = TEMA_OSCURO

        # Barra superior de marca: nombre a la izquierda, firma + info a la derecha
        barra_marca = ctk.CTkFrame(self, height=36, corner_radius=0, fg_color=t["bg_alt"])
        barra_marca.pack(fill="x")
        barra_marca.pack_propagate(False)
        ctk.CTkLabel(
            barra_marca, text="CanónicaDL - Toolkit", font=ctk.CTkFont(size=13, weight="bold"),
        ).pack(side="left", padx=12)
        ctk.CTkLabel(
            barra_marca, text="[seb@lies / 2026]", text_color=t["fg_dim"],
            font=ctk.CTkFont(family="Consolas", size=11),
        ).pack(side="right", padx=12)
        boton_info = ctk.CTkLabel(
            barra_marca, text="ⓘ Acerca de", text_color=t["accent_hover"], cursor="hand2",
        )
        boton_info.pack(side="right", padx=10)
        boton_info.bind("<Button-1>", lambda e: self._mostrar_acerca_de())

        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", padx=10, pady=10)

        ctk.CTkLabel(top, text="Carpeta:").pack(side="left")
        ctk.CTkEntry(top, textvariable=self.carpeta, width=480).pack(side="left", padx=5)
        ctk.CTkButton(top, text="Elegir carpeta...", command=self._elegir_carpeta, width=130).pack(side="left")
        ctk.CTkCheckBox(top, text="Incluir subcarpetas", variable=self.incluir_subcarpetas).pack(
            side="left", padx=15
        )
        ctk.CTkButton(top, text="▶ Analizar Archivos", command=self._iniciar_analisis, width=150).pack(
            side="left", padx=10
        )
        if DND_DISPONIBLE:
            ctk.CTkLabel(top, text="(o arrastrá la carpeta a esta ventana)", text_color=t["fg_dim"]).pack(
                side="left", padx=10
            )

        fila_conteo = ctk.CTkFrame(self, fg_color="transparent")
        fila_conteo.pack(fill="x", padx=10)
        self.etiqueta_conteo = ctk.CTkLabel(fila_conteo, text="", text_color=t["accent_hover"])
        self.etiqueta_conteo.pack(side="left")

        # Sección de formato de nomenclatura (CTk no tiene LabelFrame nativo:
        # se simula con un CTkFrame con borde + un título arriba)
        formato_frame = ctk.CTkFrame(self, fg_color=t["bg_widget"], border_width=1, border_color=t["border"])
        formato_frame.pack(fill="x", padx=10, pady=(0, 5))
        ctk.CTkLabel(
            formato_frame, text="Formato de nomenclatura", font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(anchor="w", padx=12, pady=(10, 4))

        fila1 = ctk.CTkFrame(formato_frame, fg_color="transparent")
        fila1.pack(fill="x", padx=12)
        ctk.CTkLabel(fila1, text="Preset:").pack(side="left")
        self.combo_preset = ctk.CTkOptionMenu(
            fila1, values=self._lista_presets(), variable=self.preset_var,
            command=lambda valor: self._al_elegir_preset(), width=320,
        )
        self.combo_preset.pack(side="left", padx=5)

        self.check_todos_autores = ctk.CTkCheckBox(
            fila1, text="Incluir todos los autores", variable=self.incluir_todos_autores,
            command=self._al_cambiar_incluir_autores,
        )
        self.check_todos_autores.pack(side="left", padx=(15, 0))

        ctk.CTkButton(fila1, text="Aplicar a todas las filas", command=self._aplicar_formato_a_todos).pack(
            side="left", padx=10
        )

        fila1b = ctk.CTkFrame(formato_frame, fg_color="transparent")
        fila1b.pack(fill="x", padx=12, pady=(6, 0))
        ctk.CTkButton(
            fila1b, text="💾 Guardar como preset...", command=self._guardar_preset_actual, width=170,
        ).pack(side="left")
        ctk.CTkButton(
            fila1b, text="🗑 Eliminar preset actual", command=self._eliminar_preset_actual, width=170,
            fg_color="transparent", border_width=1, border_color=t["border"],
        ).pack(side="left", padx=5)

        fila2 = ctk.CTkFrame(formato_frame, fg_color="transparent")
        fila2.pack(fill="x", padx=12, pady=(8, 0))
        ctk.CTkLabel(fila2, text="Plantilla (editable):").pack(side="left")
        entrada_plantilla = ctk.CTkEntry(fila2, textvariable=self.formato_var)
        entrada_plantilla.pack(side="left", padx=5, fill="x", expand=True)
        entrada_plantilla.bind("<KeyRelease>", lambda e: self.preset_var.set("Personalizado (plantilla editada)"))

        ctk.CTkLabel(
            formato_frame,
            text="Podés combinar campos en la plantilla (ej. {autores}, {titulo}, {anio}...) — "
                 "lista completa en \"ⓘ Acerca de\"",
            text_color=t["fg_dim"], anchor="w",
        ).pack(anchor="w", padx=12, pady=(5, 10), fill="x")

        # Tabla de resultados (se mantiene ttk.Treeview: CustomTkinter no
        # tiene una tabla propia con columnas ordenables/edición/colores)
        columnas = ("original", "formato", "abrir", "nombre", "apellido", "titulo", "anio", "nuevo", "buscar", "estado")
        self.tabla = ttk.Treeview(self, columns=columnas, show="headings", selectmode="extended")
        encabezados = {
            "original": "Archivo original",
            "formato": "Formato",
            "abrir": "",
            "nombre": "Nombre",
            "apellido": "Apellido",
            "titulo": "Título",
            "anio": "Año",
            "nuevo": "Nuevo nombre",
            "buscar": "",
            "estado": "Estado",
        }
        anchos = {
            "original": 190, "formato": 70, "abrir": 30, "nombre": 100, "apellido": 110,
            "titulo": 250, "anio": 55, "nuevo": 300, "buscar": 30, "estado": 150,
        }
        columnas_ordenables = {"original", "formato", "nombre", "apellido", "titulo", "anio", "nuevo", "estado"}
        self._orden_actual = {"columna": None, "ascendente": True}
        self._encabezados_base = encabezados
        self._columnas_ordenables = columnas_ordenables
        for col in columnas:
            if col in columnas_ordenables:
                self.tabla.heading(col, text=encabezados[col], command=lambda c=col: self._ordenar_por_columna(c))
            else:
                self.tabla.heading(col, text=encabezados[col])
            self.tabla.column(col, width=anchos[col], anchor="w")
        self.tabla.column("buscar", anchor="center", stretch=False)
        self.tabla.column("abrir", anchor="center", stretch=False)
        self.tabla.column("formato", anchor="center", stretch=False)
        self.tabla.column("anio", anchor="center")
        self.tabla.heading("anio", anchor="center")

        # Colores de zebra (alternado) y por estado. El fondo lo pone la
        # zebra; el color de texto lo pone el estado (así no compiten entre sí).
        self.tabla.tag_configure("fila_par", background=t["bg_widget"])
        self.tabla.tag_configure("fila_impar", background=t["bg_alt"])
        self.tabla.tag_configure("estado_ok", foreground="#8fd19e")
        self.tabla.tag_configure("estado_sinconfirmar", foreground="#e0c26a")
        self.tabla.tag_configure("estado_error", foreground="#e08a8a")
        self.tabla.tag_configure("estado_neutral", foreground=t["fg"])

        self.tabla.pack(fill="both", expand=True, padx=10, pady=5)
        self.tabla.bind("<Double-1>", self._doble_clic_celda)
        self.tabla.bind("<Button-1>", self._clic_celda)
        self.tabla.bind("<Button-3>", self._mostrar_menu_contextual)
        self.tabla.bind("<Delete>", self._quitar_filas_seleccionadas)
        self.tabla.bind("<Left>", lambda e: self._mover_columna_activa(-1))
        self.tabla.bind("<Right>", lambda e: self._mover_columna_activa(1))
        self.tabla.bind("<Return>", self._enter_editar_celda_activa)
        self._columna_activa_idx = 0

        self.menu_contextual = tk.Menu(self, tearoff=0, **menu_oscuro_kwargs())
        self.menu_contextual.add_command(label="Buscar manualmente...", command=self._buscar_manualmente)
        self.menu_contextual.add_command(label="Editar autores...", command=self._editar_autores)
        self.menu_contextual.add_separator()
        self.menu_contextual.add_command(label="Quitar de la lista (Del)", command=self._quitar_filas_seleccionadas)

        # Scrollbar nativo de ttk (no CTk): la Treeview es un widget nativo,
        # y mezclarla dentro de un contenedor CTk (que usa Canvas por debajo
        # para dibujarse) puede causar problemas de renderizado en Windows.
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscroll=scrollbar.set)
        scrollbar.place(relx=1.0, rely=0.32, relheight=0.5, anchor="ne")

        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.pack(fill="x", padx=10, pady=10)

        self.progreso = ctk.CTkProgressBar(bottom, mode="determinate")
        self.progreso.set(0)
        self.progreso.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.etiqueta_estado = ctk.CTkLabel(bottom, text="Listo.")
        self.etiqueta_estado.pack(side="left", padx=10)

        ctk.CTkButton(bottom, text="Renombrar seleccionados/todos", command=self._renombrar).pack(
            side="right"
        )

        self._refrescar_encabezados()

        ctk.CTkLabel(
            self,
            text="Doble clic para editar, Enter/flechas para navegar, clic derecho para más opciones — "
                 "guía completa en \"ⓘ Acerca de\"",
            text_color=t["fg_dim"], anchor="w",
        ).pack(anchor="w", padx=10, pady=(0, 8), fill="x")

    def _valores_fila(self, r):
        """Tupla de valores para una fila de la tabla, en el mismo orden que 'columnas'."""
        extension = os.path.splitext(r.get("nombre_original", ""))[1].lstrip(".").upper()
        return (
            r["nombre_original"], extension, "📂", r["nombre"], r["apellido"], r["titulo"],
            r["anio"], r["nuevo_nombre"], "🔍", r["estado"],
        )

    def _tag_estado(self, estado):
        estado = estado or ""
        if estado.startswith("OK") or estado.startswith("RENOMBRADO"):
            return "estado_ok"
        if estado.startswith("SIN CONFIRMAR"):
            return "estado_sinconfirmar"
        if estado.startswith("SIN DATOS") or estado.startswith("ERROR"):
            return "estado_error"
        return "estado_neutral"

    def _tag_zebra_de(self, item_id):
        """Conserva la franja de zebra que ya tenía la fila (no cambia con ediciones)."""
        actuales = self.tabla.item(item_id, "tags") or ()
        for t in actuales:
            if t in ("fila_par", "fila_impar"):
                return t
        indice = self.tabla.index(item_id)
        return "fila_par" if indice % 2 == 0 else "fila_impar"

    def _actualizar_fila_tabla(self, item_id):
        r = self.resultados[item_id]
        tag_zebra = self._tag_zebra_de(item_id)
        tag_estado = self._tag_estado(r.get("estado", ""))
        self.tabla.item(item_id, values=self._valores_fila(r), tags=(tag_zebra, tag_estado))

    def _lista_presets(self):
        """Lista combinada: presets de fábrica + los guardados por el usuario (con ★)."""
        propios = [f"{PREFIJO_PRESET_PERSONALIZADO}{nombre}" for nombre in self.presets_personalizados]
        return list(FORMATOS_PRESET_INFO.keys()) + propios

    def _al_elegir_preset(self, event=None):
        self._actualizar_plantilla_desde_preset()
        self._aplicar_formato_a_todos()

    def _al_cambiar_incluir_autores(self):
        # Solo tiene sentido si hay un preset de fábrica seleccionado (si es
        # personalizado o si el usuario ya editó la plantilla a mano, no la pisamos)
        if self.preset_var.get() in FORMATOS_PRESET_INFO:
            self._actualizar_plantilla_desde_preset()
            self._aplicar_formato_a_todos()
        else:
            self.etiqueta_estado.configure(
                text="Este checkbox solo afecta a los presets de fábrica (no a los personalizados ni a una "
                     "plantilla editada a mano). Agregá {autores} o {autores_completo} vos mismo si lo necesitás."
            )

    def _actualizar_plantilla_desde_preset(self):
        nombre_preset = self.preset_var.get()
        info = FORMATOS_PRESET_INFO.get(nombre_preset)
        if info:
            variante = "multi" if self.incluir_todos_autores.get() else "solo1"
            self.formato_var.set(info[variante])
            return
        if nombre_preset.startswith(PREFIJO_PRESET_PERSONALIZADO):
            nombre_real = nombre_preset[len(PREFIJO_PRESET_PERSONALIZADO):]
            plantilla = self.presets_personalizados.get(nombre_real)
            if plantilla:
                self.formato_var.set(plantilla)

    def _guardar_preset_actual(self):
        template = self.formato_var.get().strip()
        if not template:
            messagebox.showinfo("Atención", "No hay ninguna plantilla para guardar.")
            return
        nombre = simpledialog.askstring(
            "Guardar preset", "Nombre para este preset:", parent=self,
        )
        if not nombre:
            return
        nombre = nombre.strip()
        if not nombre:
            return
        if nombre in self.presets_personalizados:
            if not messagebox.askyesno("Confirmar", f'Ya existe un preset llamado "{nombre}". ¿Reemplazarlo?'):
                return
        self.presets_personalizados[nombre] = template
        if not guardar_presets_personalizados(self.presets_personalizados):
            messagebox.showerror("Error", "No se pudo guardar el preset en el disco.")
            return
        self.combo_preset.configure(values=self._lista_presets())
        self.preset_var.set(f"{PREFIJO_PRESET_PERSONALIZADO}{nombre}")
        self.etiqueta_estado.configure(text=f'Preset "{nombre}" guardado.')

    def _eliminar_preset_actual(self):
        nombre_preset = self.preset_var.get()
        if not nombre_preset.startswith(PREFIJO_PRESET_PERSONALIZADO):
            messagebox.showinfo(
                "Atención", "Solo se pueden eliminar los presets guardados por vos (marcados con ★)."
            )
            return
        nombre_real = nombre_preset[len(PREFIJO_PRESET_PERSONALIZADO):]
        if not messagebox.askyesno("Confirmar", f'¿Eliminar el preset "{nombre_real}"?'):
            return
        self.presets_personalizados.pop(nombre_real, None)
        guardar_presets_personalizados(self.presets_personalizados)
        self.combo_preset.configure(values=self._lista_presets())
        self.preset_var.set(FORMATO_DEFECTO)
        self._actualizar_plantilla_desde_preset()
        self.etiqueta_estado.configure(text=f'Preset "{nombre_real}" eliminado.')

    def _aplicar_formato_a_todos(self):
        template = self.formato_var.get().strip()
        if not template:
            return
        for item_id, r in self.resultados.items():
            r["nuevo_nombre"] = construir_nuevo_nombre(r, template)
            self._actualizar_fila_tabla(item_id)

    def _ordenar_por_columna(self, columna):
        ascendente = True
        if self._orden_actual.get("columna") == columna:
            ascendente = not self._orden_actual.get("ascendente", True)
        self._orden_actual = {"columna": columna, "ascendente": ascendente}

        columnas_tabla = list(self.tabla["columns"])
        idx_col = columnas_tabla.index(columna)

        def clave(item_id):
            valores = self.tabla.item(item_id, "values")
            valor = valores[idx_col] if idx_col < len(valores) else ""
            if columna == "anio":
                try:
                    return (0, int(valor))
                except (ValueError, TypeError):
                    return (1, 0)  # los vacíos/no numéricos quedan al final
            return (0, str(valor).lower())

        items = list(self.tabla.get_children(""))
        items.sort(key=clave, reverse=not ascendente)
        for i, item_id in enumerate(items):
            self.tabla.move(item_id, "", i)

        self._recalcular_zebra()
        self._refrescar_encabezados()

    def _recalcular_zebra(self):
        for i, item_id in enumerate(self.tabla.get_children("")):
            tag_zebra = "fila_par" if i % 2 == 0 else "fila_impar"
            tags_actuales = [t for t in (self.tabla.item(item_id, "tags") or ()) if t not in ("fila_par", "fila_impar")]
            self.tabla.item(item_id, tags=tuple([tag_zebra] + tags_actuales))

    def _refrescar_encabezados(self):
        """Redibuja los encabezados combinando la flecha de orden (▲/▼) y la
        marca de columna activa para edición con teclado (▸)."""
        columna_ordenada = self._orden_actual.get("columna")
        ascendente = self._orden_actual.get("ascendente", True)
        idx_activo = getattr(self, "_columna_activa_idx", None)
        campo_activo = self._CAMPOS_EDICION_ORDEN[idx_activo][1] if idx_activo is not None else None
        for col in self._columnas_ordenables:
            texto = self._encabezados_base[col]
            if col == columna_ordenada:
                texto += " ▲" if ascendente else " ▼"
            if col == campo_activo:
                texto += " ▸"
            self.tabla.heading(col, text=texto)

    # ------------------------- Acciones -------------------------
    def _elegir_carpeta(self):
        carpeta = filedialog.askdirectory(title="Elegí la carpeta con los PDFs")
        if carpeta:
            self.carpeta.set(carpeta)

    EXTENSIONES_SOPORTADAS = (".pdf", ".epub", ".docx")

    def _buscar_archivos_en_carpeta(self, carpeta):
        """Lista los archivos soportados en una carpeta (sin mostrar ningún diálogo)."""
        archivos = []
        if not carpeta or not os.path.isdir(carpeta):
            return archivos
        try:
            if self.incluir_subcarpetas.get():
                for root, _, files in os.walk(carpeta):
                    for f in files:
                        if f.lower().endswith(self.EXTENSIONES_SOPORTADAS):
                            archivos.append(os.path.join(root, f))
            else:
                for f in os.listdir(carpeta):
                    if f.lower().endswith(self.EXTENSIONES_SOPORTADAS):
                        archivos.append(os.path.join(carpeta, f))
        except OSError:
            pass
        return archivos

    def _listar_pdfs(self):
        carpeta = self.carpeta.get()
        if not carpeta or not os.path.isdir(carpeta):
            messagebox.showwarning("Atención", "Elegí primero una carpeta válida.")
            return []
        return self._buscar_archivos_en_carpeta(carpeta)

    def _actualizar_conteo_carpeta(self, *args):
        carpeta = self.carpeta.get()
        if not carpeta or not os.path.isdir(carpeta):
            self.etiqueta_conteo.configure(text="")
            return
        archivos = self._buscar_archivos_en_carpeta(carpeta)
        if not archivos:
            self.etiqueta_conteo.configure(text="⚠ No se encontraron archivos PDF/EPUB/DOCX en esta carpeta")
        else:
            plural = "s" if len(archivos) != 1 else ""
            self.etiqueta_conteo.configure(text=f"📄 {len(archivos)} archivo{plural} encontrado{plural}, listo{plural} para analizar")

    def _iniciar_analisis(self):
        try:
            pdfs = self._listar_pdfs()
            if not pdfs:
                messagebox.showinfo("Sin resultados", "No se encontraron archivos PDF, EPUB o DOCX en la carpeta.")
                return

            self.tabla.delete(*self.tabla.get_children())
            self.resultados.clear()
            self._progreso_total = len(pdfs)
            self.progreso.set(0)
            self.etiqueta_estado.configure(text=f"Analizando 0/{len(pdfs)}...")

            hilo = threading.Thread(target=self._analizar_en_hilo, args=(pdfs, self.formato_var.get()), daemon=True)
            hilo.start()
        except Exception as e:
            import traceback
            try:
                ruta_log = ruta_junto_al_ejecutable("analizar_error.log")
                with open(ruta_log, "w", encoding="utf-8") as f:
                    f.write(traceback.format_exc())
            except Exception:
                pass
            messagebox.showerror("Error", f"No se pudo iniciar el análisis:\n{e}\n\n(Detalles en analizar_error.log)")

    def _analizar_en_hilo(self, pdfs, template):
        total = len(pdfs)
        for i, ruta in enumerate(pdfs, start=1):
            try:
                resultado = analizar_archivo(ruta, template)
            except Exception as e:
                nombre_archivo = os.path.basename(ruta)
                resultado = {
                    "ruta": ruta,
                    "nombre_original": nombre_archivo,
                    "nombre": "",
                    "apellido": "",
                    "titulo": nombre_archivo,
                    "anio": "",
                    "editorial": "",
                    "revista": "",
                    "volumen": "",
                    "numero": "",
                    "paginas": "",
                    "doi": "",
                    "nuevo_nombre": nombre_archivo,
                    "estado": f"ERROR: {e}",
                    "fuente": "—",
                }
            self.cola.put(("resultado", resultado, i, total))
        self.cola.put(("fin", None, total, total))

    def _procesar_cola(self):
        try:
            while True:
                tipo, dato, i, total = self.cola.get_nowait()
                if tipo == "resultado":
                    try:
                        self._agregar_fila(dato)
                    except Exception as e:
                        self._log_error_cola(f"_agregar_fila falló: {type(e).__name__}: {e}", dato)
                    fraccion = (i / total) if total else 0
                    self.progreso.set(fraccion)
                    self.etiqueta_estado.configure(text=f"Analizando {i}/{total}...")
                elif tipo == "fin":
                    self.etiqueta_estado.configure(text=f"Análisis completo. {total} archivo(s) procesado(s).")
        except queue.Empty:
            pass
        except Exception as e:
            self._log_error_cola(f"_procesar_cola falló: {type(e).__name__}: {e}", None)
        self.after(150, self._procesar_cola)

    def _log_error_cola(self, mensaje, dato):
        import traceback
        try:
            ruta_log = ruta_junto_al_ejecutable("analizar_error.log")
            with open(ruta_log, "a", encoding="utf-8") as f:
                f.write(f"{mensaje}\n")
                if dato is not None:
                    f.write(f"  dato: {dato}\n")
                f.write(traceback.format_exc())
                f.write("\n---\n")
        except Exception:
            pass

    def _agregar_fila(self, r):
        indice = len(self.tabla.get_children())
        tag_zebra = "fila_par" if indice % 2 == 0 else "fila_impar"
        tag_estado = self._tag_estado(r.get("estado", ""))
        item_id = self.tabla.insert("", "end", values=self._valores_fila(r), tags=(tag_zebra, tag_estado))
        self.resultados[item_id] = r

    def _doble_clic_celda(self, event):
        item_id = self.tabla.identify_row(event.y)
        columna = self.tabla.identify_column(event.x)
        if not item_id:
            return
        if columna == "#1":  # "Archivo original": copiar el texto
            self._copiar_texto(self.resultados[item_id]["nombre_original"])
            return
        campos_editables = {"#4": "nombre", "#5": "apellido", "#6": "titulo", "#7": "anio"}
        if columna in campos_editables:
            self._iniciar_edicion_inline(item_id, columna, campos_editables[columna])
        elif columna == "#8":  # "Nuevo nombre": doble clic abre la búsqueda manual
            self._abrir_busqueda_manual(item_id)

    def _clic_celda(self, event):
        columna = self.tabla.identify_column(event.x)
        item_id = self.tabla.identify_row(event.y)
        if not item_id:
            return
        if columna == "#3":  # ícono 📂: abrir el archivo original
            self._abrir_archivo_original(item_id)
        elif columna == "#9":  # ícono 🔍: buscar manualmente
            self._abrir_busqueda_manual(item_id)

    def _copiar_texto(self, texto):
        self.clipboard_clear()
        self.clipboard_append(texto)
        texto_anterior = self.etiqueta_estado.cget("text")
        self.etiqueta_estado.configure(text=f"Copiado: {texto}")
        self.after(1800, lambda: self.etiqueta_estado.configure(text=texto_anterior))

    def _abrir_archivo_original(self, item_id):
        if item_id not in self.resultados:
            return
        ruta = self.resultados[item_id].get("ruta")
        if not ruta or not os.path.exists(ruta):
            messagebox.showwarning("Atención", "No se encuentra el archivo (¿se movió o ya se renombró?).")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(ruta)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", ruta])
            else:
                subprocess.Popen(["xdg-open", ruta])
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")

    def _iniciar_edicion_inline(self, item_id, columna, campo):
        """Abre un Entry flotante justo encima de la celda para editarla en el lugar."""
        # Si ya hay una edición en curso, la confirmamos antes de abrir otra
        if getattr(self, "_entry_edicion", None) is not None:
            self._confirmar_edicion_inline()

        bbox = self.tabla.bbox(item_id, columna)
        if not bbox:
            return
        x, y, w, h = bbox

        valor_actual = self.resultados[item_id].get(campo, "")
        var = tk.StringVar(value=valor_actual)
        entry = tk.Entry(
            self.tabla, textvariable=var, relief="solid", borderwidth=1,
            bg=TEMA_OSCURO["bg_entry"], fg=TEMA_OSCURO["fg"],
            insertbackground=TEMA_OSCURO["fg"],
            highlightthickness=1, highlightcolor=TEMA_OSCURO["accent"],
            highlightbackground=TEMA_OSCURO["border"],
        )
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()
        entry.selection_range(0, "end")

        self._entry_edicion = entry
        self._entry_edicion_info = (item_id, campo, var)

        entry.bind("<Return>", lambda e: self._confirmar_edicion_inline())
        entry.bind("<KP_Enter>", lambda e: self._confirmar_edicion_inline())
        entry.bind("<Tab>", lambda e: self._tab_edicion_inline(item_id, retroceder=False))
        entry.bind("<Shift-Tab>", lambda e: self._tab_edicion_inline(item_id, retroceder=True))
        entry.bind("<Escape>", lambda e: self._cancelar_edicion_inline())
        entry.bind("<FocusOut>", lambda e: self._confirmar_edicion_inline())

    # Orden de navegación con Tab: Nombre -> Apellido -> Título -> Año
    _CAMPOS_EDICION_ORDEN = (("#4", "nombre"), ("#5", "apellido"), ("#6", "titulo"), ("#7", "anio"))

    def _tab_edicion_inline(self, item_id, retroceder=False):
        info_actual = getattr(self, "_entry_edicion_info", None)
        campo_actual = info_actual[1] if info_actual else None
        self._confirmar_edicion_inline()

        campos = [c for _, c in self._CAMPOS_EDICION_ORDEN]
        if campo_actual not in campos:
            return "break"
        idx = campos.index(campo_actual)
        idx_siguiente = idx - 1 if retroceder else idx + 1
        if 0 <= idx_siguiente < len(campos):
            columna_siguiente, campo_siguiente = self._CAMPOS_EDICION_ORDEN[idx_siguiente]
            self._iniciar_edicion_inline(item_id, columna_siguiente, campo_siguiente)
        return "break"

    def _confirmar_edicion_inline(self):
        info = getattr(self, "_entry_edicion_info", None)
        entry = getattr(self, "_entry_edicion", None)
        if not info or not entry:
            return
        item_id, campo, var = info
        nuevo_valor = var.get().strip()
        self._entry_edicion = None
        self._entry_edicion_info = None
        entry.destroy()

        if item_id in self.resultados:
            self.resultados[item_id][campo] = nuevo_valor
            r = self.resultados[item_id]
            self._reevaluar_estado_manual(r)
            r["nuevo_nombre"] = construir_nuevo_nombre(r, self.formato_var.get())
            self._actualizar_fila_tabla(item_id)

    def _reevaluar_estado_manual(self, r):
        """
        Si el usuario edita a mano un campo y la fila queda con título y
        autor completos, actualiza el estado para que deje de figurar en
        rojo/amarillo como si nada se hubiera resuelto.
        """
        estado_actual = r.get("estado", "")
        pendiente = (
            not estado_actual
            or estado_actual.startswith("SIN DATOS")
            or estado_actual.startswith("SIN CONFIRMAR")
        )
        if pendiente and r.get("titulo") and (r.get("nombre") or r.get("apellido")):
            r["estado"] = "OK (manual)"

    def _cancelar_edicion_inline(self):
        entry = getattr(self, "_entry_edicion", None)
        if entry:
            entry.destroy()
        self._entry_edicion = None
        self._entry_edicion_info = None

    def _mostrar_menu_contextual(self, event):
        item_id = self.tabla.identify_row(event.y)
        if not item_id:
            return
        seleccion_actual = self.tabla.selection()
        if item_id not in seleccion_actual:
            self.tabla.selection_set(item_id)
        self._item_menu_actual = item_id
        self.menu_contextual.tk_popup(event.x_root, event.y_root)

    def _mover_columna_activa(self, delta):
        maximo = len(self._CAMPOS_EDICION_ORDEN) - 1
        self._columna_activa_idx = max(0, min(maximo, getattr(self, "_columna_activa_idx", 0) + delta))
        self._refrescar_encabezados()
        return "break"

    def _enter_editar_celda_activa(self, event=None):
        item_id = self.tabla.focus()
        if not item_id:
            seleccion = self.tabla.selection()
            item_id = seleccion[0] if seleccion else None
        if not item_id:
            return "break"
        columna, campo = self._CAMPOS_EDICION_ORDEN[getattr(self, "_columna_activa_idx", 0)]
        self._iniciar_edicion_inline(item_id, columna, campo)
        return "break"

    def _quitar_filas_seleccionadas(self, event=None):
        seleccion = self.tabla.selection()
        if not seleccion:
            return
        for item_id in seleccion:
            self.resultados.pop(item_id, None)
            self.tabla.delete(item_id)
        self._recalcular_zebra()
        self.etiqueta_estado.configure(text=f"Se quitaron {len(seleccion)} archivo(s) de la lista.")

    def _abrir_busqueda_manual(self, item_id):
        if item_id not in self.resultados:
            return
        r = self.resultados[item_id]
        DialogoBusquedaManual(self, item_id, r, self._aplicar_seleccion_manual)

    def _buscar_manualmente(self):
        item_id = getattr(self, "_item_menu_actual", None)
        if not item_id:
            return
        self._abrir_busqueda_manual(item_id)

    def _aplicar_seleccion_manual(self, item_id, candidato):
        r = self.resultados[item_id]
        nombre, apellido = dividir_nombre_apellido(candidato.get("autor", ""))
        r["nombre"] = nombre
        r["apellido"] = apellido
        r["autores"] = candidato.get("autores") or ([{"nombre": nombre, "apellido": apellido}] if (nombre or apellido) else [])
        r["titulo"] = candidato.get("titulo", "")
        r["anio"] = candidato.get("anio", "")
        r["editorial"] = candidato.get("editorial", r.get("editorial", ""))
        r["revista"] = candidato.get("revista", r.get("revista", ""))
        r["volumen"] = candidato.get("volumen", r.get("volumen", ""))
        r["numero"] = candidato.get("numero", r.get("numero", ""))
        r["paginas"] = candidato.get("paginas", r.get("paginas", ""))
        r["doi"] = candidato.get("doi", r.get("doi", ""))
        r["estado"] = "OK (manual)"
        r["nuevo_nombre"] = construir_nuevo_nombre(r, self.formato_var.get())
        self._actualizar_fila_tabla(item_id)

    def _editar_autores(self):
        item_id = getattr(self, "_item_menu_actual", None)
        if not item_id or item_id not in self.resultados:
            return
        seleccion = self.tabla.selection()
        items_destino = list(seleccion) if item_id in seleccion and len(seleccion) > 1 else [item_id]
        r = self.resultados[item_id]
        EditorAutoresDialog(self, items_destino, r, self._aplicar_autores_editados)

    def _aplicar_autores_editados(self, items_destino, autores):
        primero = autores[0] if autores else {"nombre": "", "apellido": ""}
        for item_id in items_destino:
            r = self.resultados.get(item_id)
            if not r:
                continue
            r["autores"] = autores
            r["nombre"] = primero.get("nombre", "")
            r["apellido"] = primero.get("apellido", "")
            if r["estado"] in ("", "SIN DATOS (revisar manualmente)", "SIN CONFIRMAR (detectado en nombre de archivo)"):
                r["estado"] = "OK (manual)"
            r["nuevo_nombre"] = construir_nuevo_nombre(r, self.formato_var.get())
            self._actualizar_fila_tabla(item_id)

    def _renombrar(self):
        seleccion = self.tabla.selection()
        items = seleccion if seleccion else self.tabla.get_children()
        if not items:
            messagebox.showinfo("Sin datos", "No hay archivos analizados para renombrar.")
            return

        confirmacion = messagebox.askyesno(
            "Confirmar",
            f"Se van a renombrar {len(items)} archivo(s). ¿Continuar?",
        )
        if not confirmacion:
            return

        errores = []
        renombrados = 0
        for item_id in items:
            r = self.resultados.get(item_id)
            if not r:
                continue
            ruta_original = r["ruta"]
            carpeta = os.path.dirname(ruta_original)
            nuevo_nombre = r["nuevo_nombre"]
            ruta_nueva = os.path.join(carpeta, nuevo_nombre)

            if os.path.abspath(ruta_nueva) == os.path.abspath(ruta_original):
                continue

            # Evitar sobrescribir archivos existentes
            contador = 1
            base, ext = os.path.splitext(nuevo_nombre)
            while os.path.exists(ruta_nueva):
                ruta_nueva = os.path.join(carpeta, f"{base} ({contador}){ext}")
                contador += 1

            try:
                os.rename(ruta_original, ruta_nueva)
                r["ruta"] = ruta_nueva
                r["nombre_original"] = os.path.basename(ruta_nueva)
                r["estado"] = "RENOMBRADO"
                self._actualizar_fila_tabla(item_id)
                renombrados += 1
            except Exception as e:
                errores.append(f"{os.path.basename(ruta_original)}: {e}")

        mensaje = f"Se renombraron {renombrados} archivo(s)."
        if errores:
            mensaje += "\n\nErrores:\n" + "\n".join(errores)
        messagebox.showinfo("Resultado", mensaje)


class AcercaDeDialog(ctk.CTkToplevel):
    """Ventana "Acerca de": logo, descripción, y toda la ayuda de uso organizada por pestañas."""

    DESCRIPCION = (
        "Canónica es una aplicación de escritorio para organizar bibliotecas "
        "digitales. Permite renombrar automáticamente libros, artículos, tesis "
        "y otros documentos siguiendo una nomenclatura definida por el usuario, "
        "manteniendo toda la colección ordenada y consistente.\n\n"
        "Pensada para quienes manejan grandes cantidades de archivos académicos, "
        "Canónica simplifica el proceso de organizar, buscar y mantener una "
        "biblioteca personal. Su objetivo es transformar colecciones "
        "desordenadas en un archivo claro, uniforme y fácil de navegar, dejando "
        "que el usuario dedique menos tiempo a ordenar y más tiempo a leer e "
        "investigar."
    )
    LINK = "sebastianalies.com"

    ATAJOS = [
        ("🖱  Doble clic en Nombre/Apellido/Título/Año", "Editar ese campo en la celda"),
        ("⌨  Enter", "Editar el campo marcado con ▸ en el encabezado"),
        ("⌨  ← / →", "Mover cuál campo se edita con Enter"),
        ("⌨  ↑ / ↓", "Moverse entre filas"),
        ("⌨  Tab (durante edición)", "Guardar y pasar al siguiente campo"),
        ("⌨  Delete", "Quitar de la lista las filas seleccionadas"),
        ("🖱  Doble clic en \"Archivo original\"", "Copiar ese nombre al portapapeles"),
        ("🖱  Clic en 📂", "Abrir el archivo original"),
        ("🖱  Clic en 🔍", "Buscar manualmente ese archivo"),
        ("🖱  Doble clic en \"Nuevo nombre\"", "También abre la búsqueda manual"),
        ("🖱  Clic en un encabezado", "Ordenar la tabla por esa columna"),
        ("🖱  Clic derecho sobre una fila", "Menú: buscar, editar autores, quitar (funciona con varias filas seleccionadas a la vez)"),
        ("📁  Arrastrar una carpeta a la ventana", "Cargarla sin usar \"Elegir carpeta...\""),
    ]

    def __init__(self, parent):
        super().__init__(parent)
        t = TEMA_OSCURO
        self.title("Acerca de CanónicaDL")
        self.geometry("640x720")
        self.resizable(True, True)
        self.minsize(520, 550)

        # Cabecera con el logo completo, en fondo negro puro (igual que el logo)
        cabecera = ctk.CTkFrame(self, fg_color="#000000", corner_radius=0)
        cabecera.pack(fill="x")
        try:
            if LOGO_COMPLETO_BASE64:
                self._logo_img = tk.PhotoImage(data=LOGO_COMPLETO_BASE64, format="png")
                tk.Label(cabecera, image=self._logo_img, bg="#000000", borderwidth=0).pack(pady=12)
        except Exception:
            ctk.CTkLabel(
                cabecera, text="CanónicaDL", text_color="#ffffff",
                font=ctk.CTkFont(size=22, weight="bold"),
            ).pack(pady=25)

        tabs = ctk.CTkTabview(self, fg_color=t["bg"])
        tabs.pack(fill="both", expand=True, padx=15, pady=(10, 5))
        tab_info = tabs.add("Info")
        tab_atajos = tabs.add("Atajos")
        tab_plantillas = tabs.add("Plantillas")

        # --- Pestaña Info ---
        ctk.CTkLabel(
            tab_info, text=self.DESCRIPCION, wraplength=560, justify="left", anchor="w",
            font=ctk.CTkFont(size=12),
        ).pack(anchor="w", fill="x", padx=5, pady=(10, 0))

        ctk.CTkLabel(
            tab_info, text="Fuentes de datos bibliográficos: Google Books, Open Library (ISBN) y Crossref "
                           "(artículos académicos/DOI).",
            wraplength=560, justify="left", anchor="w", text_color=t["fg_dim"],
        ).pack(anchor="w", fill="x", padx=5, pady=(15, 0))

        link = ctk.CTkLabel(
            tab_info, text=self.LINK, text_color=t["accent_hover"], cursor="hand2",
            font=ctk.CTkFont(size=12, underline=True),
        )
        link.pack(anchor="w", padx=5, pady=(20, 0))
        link.bind("<Button-1>", lambda e: webbrowser.open(f"https://{self.LINK}"))

        # --- Pestaña Atajos ---
        scroll_atajos = ctk.CTkScrollableFrame(tab_atajos, fg_color="transparent")
        scroll_atajos.pack(fill="both", expand=True, padx=5, pady=5)
        for accion, descripcion in self.ATAJOS:
            fila = ctk.CTkFrame(scroll_atajos, fg_color="transparent")
            fila.pack(fill="x", pady=4)
            ctk.CTkLabel(
                fila, text=accion, font=ctk.CTkFont(size=12, weight="bold"),
                anchor="w", width=280, justify="left", wraplength=270,
            ).pack(side="left", anchor="n")
            ctk.CTkLabel(
                fila, text=descripcion, text_color=t["fg_dim"], anchor="w",
                justify="left", wraplength=250,
            ).pack(side="left", anchor="n", padx=(8, 0))

        # --- Pestaña Plantillas ---
        ctk.CTkLabel(
            tab_plantillas, text="Campos disponibles para armar tu propia plantilla de nomenclatura:",
            font=ctk.CTkFont(size=12, weight="bold"), anchor="w",
        ).pack(anchor="w", padx=5, pady=(10, 5), fill="x")

        campos_plantilla = [
            ("{autores}", "Todos los autores, formato APA con iniciales (García, J. y Pérez, M.)"),
            ("{autores_completo}", "Todos los autores, nombre completo (García, Juan y Pérez, María)"),
            ("{nombre} / {apellido}", "Nombre/apellido de pila del primer autor únicamente"),
            ("{inicial}", "Inicial del primer autor (G.)"),
            ("{titulo}", "Título del documento"),
            ("{anio}", "Año de publicación"),
            ("{editorial}", "Editorial (libros)"),
            ("{revista}", "Nombre de la revista (artículos)"),
            ("{volumen} / {numero} / {paginas}", "Datos de publicación de artículos"),
            ("{doi}", "DOI del artículo, si está disponible"),
        ]
        scroll_campos = ctk.CTkScrollableFrame(tab_plantillas, fg_color="transparent")
        scroll_campos.pack(fill="both", expand=True, padx=5, pady=5)
        for campo, descripcion in campos_plantilla:
            fila = ctk.CTkFrame(scroll_campos, fg_color="transparent")
            fila.pack(fill="x", pady=4)
            ctk.CTkLabel(
                fila, text=campo, font=ctk.CTkFont(size=12, weight="bold", family="Consolas"),
                anchor="w", width=190, justify="left",
            ).pack(side="left", anchor="n")
            ctk.CTkLabel(
                fila, text=descripcion, text_color=t["fg_dim"], anchor="w",
                justify="left", wraplength=340,
            ).pack(side="left", anchor="n", padx=(8, 0))

        ctk.CTkButton(self, text="Cerrar", command=self.destroy).pack(anchor="e", padx=15, pady=(0, 15))


class EditorAutoresDialog(ctk.CTkToplevel):
    """Ventana para cargar/editar la lista completa de autores de un archivo (2, 3 o más)."""

    def __init__(self, parent, items_destino, resultado_actual, callback_guardar):
        super().__init__(parent)
        t = TEMA_OSCURO
        self.items_destino = items_destino
        es_multiple = len(items_destino) > 1
        if es_multiple:
            self.title(f"Editar autores: {len(items_destino)} archivos seleccionados")
        else:
            self.title(f"Editar autores: {resultado_actual['nombre_original']}")
        self.geometry("540x480" if es_multiple else "540x460")
        self.minsize(480, 320)
        self.callback_guardar = callback_guardar
        self.filas = []  # lista de (frame, nombre_var, apellido_var)

        ctk.CTkLabel(self, text="Archivo:").pack(anchor="w", padx=10, pady=(10, 0))
        if es_multiple:
            ctk.CTkLabel(
                self, text=f"⚠ Se van a aplicar estos autores a los {len(items_destino)} archivos seleccionados",
                text_color="#e0c26a",
            ).pack(anchor="w", padx=10)
        else:
            ctk.CTkLabel(self, text=resultado_actual["nombre_original"], text_color=t["fg_dim"]).pack(
                anchor="w", padx=10
            )
        ctk.CTkLabel(
            self, text="Cargá los autores en el orden en que deben aparecer en la cita:",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(anchor="w", padx=10, pady=(12, 5))

        # Área scrolleable por si se cargan muchos autores
        self.contenedor = ctk.CTkScrollableFrame(self, fg_color=t["bg"])
        self.contenedor.pack(fill="both", expand=True, padx=10)

        autores_existentes = resultado_actual.get("autores") or []
        if not autores_existentes and (resultado_actual.get("nombre") or resultado_actual.get("apellido")):
            autores_existentes = [{
                "nombre": resultado_actual.get("nombre", ""),
                "apellido": resultado_actual.get("apellido", ""),
            }]
        if not autores_existentes:
            autores_existentes = [{"nombre": "", "apellido": ""}, {"nombre": "", "apellido": ""}]

        for a in autores_existentes:
            self._agregar_fila(a.get("nombre", ""), a.get("apellido", ""))

        ctk.CTkButton(self, text="+ Agregar otro autor", command=lambda: self._agregar_fila("", "")).pack(
            anchor="w", padx=10, pady=(8, 0)
        )

        ctk.CTkLabel(
            self, text="El orden importa: en formato APA el primer autor va Apellido, Inicial.",
            text_color=t["fg_dim"], wraplength=500, justify="left", anchor="w",
        ).pack(anchor="w", padx=10, pady=(8, 0), fill="x")

        botones_inferiores = ctk.CTkFrame(self, fg_color="transparent")
        botones_inferiores.pack(fill="x", padx=10, pady=10, side="bottom")
        ctk.CTkButton(botones_inferiores, text="Guardar", command=self._guardar).pack(side="right")
        ctk.CTkButton(
            botones_inferiores, text="Cancelar", command=self.destroy,
            fg_color="transparent", border_width=1, border_color=t["border"],
        ).pack(side="right", padx=5)

    def _agregar_fila(self, nombre="", apellido=""):
        fila = ctk.CTkFrame(self.contenedor, fg_color="transparent")
        fila.pack(fill="x", pady=3)
        nombre_var = tk.StringVar(value=nombre)
        apellido_var = tk.StringVar(value=apellido)
        numero = len(self.filas) + 1
        ctk.CTkLabel(fila, text=f"{numero}.", width=22).pack(side="left")
        ctk.CTkLabel(fila, text="Nombre:").pack(side="left")
        ctk.CTkEntry(fila, textvariable=nombre_var, width=110).pack(side="left", padx=(3, 10))
        ctk.CTkLabel(fila, text="Apellido:").pack(side="left")
        ctk.CTkEntry(fila, textvariable=apellido_var, width=130).pack(side="left", padx=(3, 10))
        ctk.CTkButton(
            fila, text="✕", width=28, command=lambda: self._quitar_fila(fila),
            fg_color="transparent", border_width=1, border_color=TEMA_OSCURO["border"],
        ).pack(side="left")
        self.filas.append((fila, nombre_var, apellido_var))

    def _quitar_fila(self, fila):
        self.filas = [f for f in self.filas if f[0] != fila]
        fila.destroy()
        self._renumerar()

    def _renumerar(self):
        for i, (fila, _, _) in enumerate(self.filas, start=1):
            fila.winfo_children()[0].configure(text=f"{i}.")

    def _guardar(self):
        autores = []
        for _, nombre_var, apellido_var in self.filas:
            n = nombre_var.get().strip()
            a = apellido_var.get().strip()
            if n or a:
                autores.append({"nombre": n, "apellido": a})
        if not autores:
            messagebox.showinfo("Atención", "Cargá al menos un autor.")
            return
        self.callback_guardar(self.items_destino, autores)
        self.destroy()


class DialogoBusquedaManual(ctk.CTkToplevel):
    """Ventana para buscar un libro manualmente en Google Books y elegir el resultado correcto."""

    def __init__(self, parent, item_id, resultado_actual, callback_aplicar):
        super().__init__(parent)
        t = TEMA_OSCURO
        self.title(f"Buscar manualmente: {resultado_actual['nombre_original']}")
        self.geometry("620x680")
        self.item_id = item_id
        self.callback_aplicar = callback_aplicar
        self.candidatos = []

        ctk.CTkLabel(self, text="Archivo:").pack(anchor="w", padx=10, pady=(10, 0))
        ctk.CTkLabel(self, text=resultado_actual["nombre_original"], text_color=t["fg_dim"]).pack(
            anchor="w", padx=10
        )

        frame_busqueda = ctk.CTkFrame(self, fg_color="transparent")
        frame_busqueda.pack(fill="x", padx=10, pady=10)

        self.query_var = tk.StringVar(
            value=f"{resultado_actual.get('titulo', '')} {resultado_actual.get('apellido', '')}".strip()
        )
        entrada = ctk.CTkEntry(frame_busqueda, textvariable=self.query_var)
        entrada.pack(side="left", fill="x", expand=True)
        entrada.bind("<Return>", lambda e: self._buscar())
        ctk.CTkButton(frame_busqueda, text="Buscar", command=self._buscar, width=80).pack(side="left", padx=5)
        ctk.CTkButton(frame_busqueda, text="Abrir en Google 🔗", command=self._abrir_en_navegador, width=150).pack(
            side="left", padx=(0, 5)
        )

        ctk.CTkLabel(
            self,
            text="Busca en Google Books y Crossref (académico). Si no aparece nada, probá "
                 "\"Abrir en Google\" y copiá los datos a mano abajo.",
            text_color=t["fg_dim"], wraplength=560, justify="left", anchor="w",
        ).pack(anchor="w", padx=10, fill="x")

        self.lista = tk.Listbox(self, height=10, **widget_oscuro_kwargs())
        self.lista.pack(fill="both", expand=True, padx=10, pady=(8, 5))

        botones = ctk.CTkFrame(self, fg_color="transparent")
        botones.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkButton(botones, text="Usar seleccionado", command=self._usar_seleccionado).pack(side="right")
        ctk.CTkButton(
            botones, text="Cancelar", command=self.destroy,
            fg_color="transparent", border_width=1, border_color=t["border"],
        ).pack(side="right", padx=5)

        # --- Sección de carga 100% manual, sin depender de la búsqueda ---
        ctk.CTkFrame(self, height=1, fg_color=t["border"]).pack(fill="x", padx=10, pady=(5, 10))

        ctk.CTkLabel(self, text="O cargá los datos vos mismo:", font=ctk.CTkFont(size=12, weight="bold")).pack(
            anchor="w", padx=10
        )

        form = ctk.CTkFrame(self, fg_color="transparent")
        form.pack(fill="x", padx=10, pady=(5, 10))
        form.columnconfigure(1, weight=1)
        form.columnconfigure(3, weight=1)

        self.nombre_manual = tk.StringVar(value=resultado_actual.get("nombre", ""))
        self.apellido_manual = tk.StringVar(value=resultado_actual.get("apellido", ""))
        self.titulo_manual = tk.StringVar(value=resultado_actual.get("titulo", ""))
        self.anio_manual = tk.StringVar(value=resultado_actual.get("anio", ""))

        ctk.CTkLabel(form, text="Nombre:").grid(row=0, column=0, sticky="w", padx=(0, 5), pady=2)
        ctk.CTkEntry(form, textvariable=self.nombre_manual).grid(row=0, column=1, sticky="ew", padx=(0, 15), pady=2)
        ctk.CTkLabel(form, text="Apellido:").grid(row=0, column=2, sticky="w", padx=(0, 5), pady=2)
        ctk.CTkEntry(form, textvariable=self.apellido_manual).grid(row=0, column=3, sticky="ew", pady=2)

        ctk.CTkLabel(form, text="Título:").grid(row=1, column=0, sticky="w", padx=(0, 5), pady=2)
        ctk.CTkEntry(form, textvariable=self.titulo_manual).grid(
            row=1, column=1, columnspan=2, sticky="ew", padx=(0, 15), pady=2
        )
        ctk.CTkLabel(form, text="Año:").grid(row=1, column=3, sticky="w", padx=(0, 5), pady=2)
        ctk.CTkEntry(form, textvariable=self.anio_manual, width=70).grid(row=1, column=4, sticky="w", pady=2)

        ctk.CTkButton(form, text="Aplicar datos manuales", command=self._usar_datos_manuales).grid(
            row=2, column=0, columnspan=5, sticky="e", pady=(8, 0)
        )

        entrada.focus_set()
        self._buscar()

    def _buscar(self):
        query = self.query_var.get().strip()
        if not query:
            return
        self.lista.delete(0, "end")
        self.lista.insert("end", "Buscando en Google Books y Crossref...")
        self.update_idletasks()
        self.candidatos = consultar_multiples_fuentes(query, max_resultados=6)
        self.lista.delete(0, "end")
        if not self.candidatos:
            self.lista.insert("end", "(sin resultados en ninguna fuente, probá 'Abrir en Google')")
            return
        for c in self.candidatos:
            autor = c.get("autor", "") or "(autor desconocido)"
            anio = c.get("anio", "") or "s/f"
            fuente = c.get("fuente_nombre", "")
            self.lista.insert("end", f"[{fuente}] {c['titulo']}  —  {autor}  ({anio})")

    def _abrir_en_navegador(self):
        query = self.query_var.get().strip()
        if not query:
            messagebox.showinfo("Atención", "Escribí algo en el campo de búsqueda primero.")
            return
        abrir_busqueda_en_navegador(query)

    def _usar_datos_manuales(self):
        candidato = {
            "titulo": self.titulo_manual.get().strip(),
            "autor": f"{self.apellido_manual.get().strip()}, {self.nombre_manual.get().strip()}".strip(", "),
            "anio": self.anio_manual.get().strip(),
        }
        if not candidato["titulo"] and not candidato["autor"]:
            messagebox.showinfo("Atención", "Completá al menos el título o el autor.")
            return
        self.callback_aplicar(self.item_id, candidato)
        self.destroy()

    def _usar_seleccionado(self):
        seleccion = self.lista.curselection()
        if not seleccion or not self.candidatos:
            messagebox.showinfo("Atención", "Elegí un resultado de la lista primero.")
            return
        indice = seleccion[0]
        if indice >= len(self.candidatos):
            return
        candidato = self.candidatos[indice]
        self.callback_aplicar(self.item_id, candidato)
        self.destroy()


if __name__ == "__main__":
    app = RenombradorApp()
    app.mainloop()
