#!/usr/bin/env python3
"""Automatiza la carga de Ingresos Brutos (AGIP) en la planilla de clientes.

Uso:
    python iibb_agip_scraper.py --crear-demo
    python iibb_agip_scraper.py --excel demo.xlsx --dry-run
    python iibb_agip_scraper.py --excel demo.xlsx --cuit 20111111112 --sin-headless
    python iibb_agip_scraper.py --excel "IIBB_ANUAL_2025.xlsx"

Este archivo no contiene ninguna credencial. El usuario de AGIP es el
CUIT/CUIL de cada cliente. La contrasena se resuelve asi, en orden: (1) si
la celda al lado del CUIT en la hoja "DIA N" tiene algo escrito, se prueba
primero, (2) si no funciona o no hay ninguna escrita, se prueban las
contrasenas por defecto (CONTRASENAS_POR_DEFECTO, mas abajo). Un cliente
solo se reporta como "fallido" si ninguna de las candidatas funciono.

Limitacion conocida: el login y la extraccion de datos de AGIP estan
escritos a partir de las instrucciones de la hoja DETALLE, sin haber podido
probarlos contra el sitio real. Es muy probable que el primer intento
necesite un ajuste de selectores en las funciones iniciar_sesion(),
ir_a_declaracion() y extraer_campos().
"""

from __future__ import annotations

import argparse
import logging
import random
import re
import shutil
import sys
import time
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.worksheet import Worksheet

# --------------------------------------------------------------------------
# Vocabulario del dominio (basado en las hojas "Lista de Clientes - IIBB",
# "DETALLE" y "DIA 1".."DIA 8" del archivo original)
# --------------------------------------------------------------------------

HOJA_CLIENTES = "Lista de Clientes - IIBB"
SUFIJO_2025 = " - 2025"
RE_HOJA_DIA_2024 = re.compile(r"^DIA\s+\d+$", re.I)
RE_HOJA_DIA_2025 = re.compile(r"^DIA\s+\d+\s*-\s*2025$", re.I)

BASE_URL = "https://www.agip.gob.ar"

MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO",
         "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]

CONCEPTOS = [
    "base imponible",
    "anticipo determinado",
    "retenciones",
    "retenciones bancarias",
    "percepciones",
    "impuestos internos",
    "pago a cuenta",
    "otros creditos",
    "saldo a favor",
    "importe a pagar(subtotal)",
    "intereses",
    "total pagado",
]

# Concepto -> texto de etiqueta a buscar en la pagina de la DDJJ de AGIP.
# La ruta completa dentro de e-Sicol (segun la hoja DETALLE) queda como
# comentario, de referencia para quien tenga que ajustar esto mirando el
# sitio real. "intereses" no tiene campo propio: sale de restar
# total_pagado - importe_a_pagar_subtotal (asi lo indica DETALLE).
UBICACION_AGIP = {
    "base imponible": "Base Imponible",                  # Rubro 1 - Determinacion del anticipo
    "anticipo determinado": "Valor",                      # Rubro 1 - Determinacion del anticipo (etiqueta generica, ver nota abajo)
    "retenciones": "Retenciones",                          # Retenciones / Agentes
    "retenciones bancarias": "Retenciones Bancarias",      # Retenciones / Bancarias
    "percepciones": "Percepciones",                        # Percepciones / Agentes
    "impuestos internos": "Impuestos Internos",            # Conceptos que no integran la base imponible
    "otros creditos": "Saldo a favor DDJJ periodo anterior",
    "saldo a favor": "Subtotal a favor del contribuyente",
    "importe a pagar(subtotal)": "Importe neto a ingresar",
    "total pagado": "Total importe actualizado",
    "alicuota": "Alicuota",                                # Rubro 1 - Determinacion del anticipo
}
# Nota: "Valor" es una etiqueta muy generica. Si extrae el numero incorrecto
# para "anticipo determinado", es el primer lugar donde hay que mirar.

UMBRAL_CUIT = 10 ** 10  # un CUIT/CUIL tiene 11 digitos

# Pausas con variacion aleatoria (jitter): un tiempo fijo identico en cada
# paso es, en si mismo, una firma de script. +/- el jitter de por medio
# para que el ritmo no sea perfectamente uniforme. Configurables por CLI
# (--pausa-accion / --pausa-clientes) sin tocar codigo.
PAUSA_ACCION = 1.0            # despues de cada click/navegacion
PAUSA_ACCION_JITTER = 0.4
PAUSA_ENTRE_CLIENTES = 4.0    # entre el cierre de un cliente y el login del siguiente
PAUSA_ENTRE_CLIENTES_JITTER = 2.0

# La mayoria de los clientes no tiene contrasena propia escrita en el Excel:
# usan una de estas dos por defecto. Si el bloque del cliente SI tiene una
# contrasena explicita al lado del CUIT, esa se prueba primero.
CONTRASENAS_POR_DEFECTO = ["magon130", "magon131"]

logger = logging.getLogger("iibb_agip")


# --------------------------------------------------------------------------
# Utilidades de texto
# --------------------------------------------------------------------------

def normalizar(texto) -> str:
    if texto is None:
        return ""
    texto = str(texto).strip().lower()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", texto)


MESES_NORM = [normalizar(m) for m in MESES]

PATRON_NUMERO = re.compile(r"-?\$?\s*([\d.]*\d(?:,\d+)?)")


def parsear_numero_ar(texto: Optional[str]) -> Optional[float]:
    """Convierte '$ 1.234,56' o '1234.56' a float. None si no encuentra numero."""
    if not texto:
        return None
    texto = texto.strip()
    if not texto or texto in {"-", "--"}:
        return None
    m = PATRON_NUMERO.search(texto)
    if not m:
        return None
    crudo = m.group(1)
    if "," in crudo:
        crudo = crudo.replace(".", "").replace(",", ".")
    try:
        return float(crudo)
    except ValueError:
        return None


# --------------------------------------------------------------------------
# Modelo: un bloque = los datos de un cliente dentro de una hoja "DIA N"
# --------------------------------------------------------------------------

@dataclass
class BloqueCliente:
    hoja: str
    anio: int
    nombre: str
    cuit: Optional[int]
    password: Optional[str]
    fila_encabezado: int
    columnas_mes: Dict[int, str]       # 1..12 -> "B".."M"
    filas_concepto: Dict[str, int]     # concepto normalizado -> fila
    fila_alicuota: Optional[int]


def parsear_hoja_dia(ws: Worksheet, anio: int) -> List[BloqueCliente]:
    """Recorre una hoja DIA-N y arma un BloqueCliente por cada cliente.

    No asume posiciones fijas de fila: cada concepto se identifica por el
    texto de la columna A (case/acentos insensible), porque en el archivo
    real el orden y la presencia de filas varia de cliente a cliente.
    """
    encabezados = [
        fila for fila in range(1, ws.max_row + 1)
        if normalizar(ws.cell(row=fila, column=2).value) == "enero"
    ]

    bloques: List[BloqueCliente] = []
    for i, fila_inicio in enumerate(encabezados):
        fila_fin = encabezados[i + 1] - 1 if i + 1 < len(encabezados) else ws.max_row
        nombre = ws.cell(row=fila_inicio, column=1).value or f"(sin nombre, fila {fila_inicio})"

        columnas_mes: Dict[int, str] = {}
        for col in range(2, ws.max_column + 1):
            texto = normalizar(ws.cell(row=fila_inicio, column=col).value)
            if texto in MESES_NORM:
                columnas_mes[MESES_NORM.index(texto) + 1] = get_column_letter(col)

        filas_concepto: Dict[str, int] = {}
        fila_total_pagado = None
        cuit_valor = None
        password = None

        for fila in range(fila_inicio + 1, fila_fin + 1):
            valor_a = ws.cell(row=fila, column=1).value
            texto_a = normalizar(valor_a)

            if texto_a in CONCEPTOS and texto_a not in filas_concepto:
                filas_concepto[texto_a] = fila
                if texto_a == "total pagado":
                    fila_total_pagado = fila
                continue

            if isinstance(valor_a, (int, float)) and valor_a >= UMBRAL_CUIT and cuit_valor is None:
                cuit_valor = int(valor_a)
                valor_b = ws.cell(row=fila, column=2).value
                if isinstance(valor_b, str) and valor_b.strip():
                    password = valor_b.strip()

        fila_alicuota = None
        if fila_total_pagado is not None:
            candidata = fila_total_pagado + 1
            if candidata <= fila_fin and normalizar(ws.cell(row=candidata, column=1).value) == "":
                fila_alicuota = candidata

        bloques.append(BloqueCliente(
            hoja=ws.title, anio=anio, nombre=str(nombre).strip(), cuit=cuit_valor,
            password=password, fila_encabezado=fila_inicio, columnas_mes=columnas_mes,
            filas_concepto=filas_concepto, fila_alicuota=fila_alicuota,
        ))
    return bloques


def indexar_workbook(wb) -> Dict[Tuple[int, int], BloqueCliente]:
    """Escanea todas las hojas DIA-N (2024) y DIA-N - 2025 y arma un indice
    (anio, cuit) -> BloqueCliente."""
    indice: Dict[Tuple[int, int], BloqueCliente] = {}
    for nombre in wb.sheetnames:
        nombre_limpio = nombre.strip()
        if RE_HOJA_DIA_2025.match(nombre_limpio):
            anio = 2025
        elif RE_HOJA_DIA_2024.match(nombre_limpio):
            anio = 2024
        else:
            continue
        for bloque in parsear_hoja_dia(wb[nombre], anio):
            if bloque.cuit:
                indice[(anio, bloque.cuit)] = bloque
            else:
                logger.warning("No pude identificar el CUIT de %r en %s (fila %s)",
                                bloque.nombre, nombre, bloque.fila_encabezado)
    return indice


# --------------------------------------------------------------------------
# Padron de clientes ("Lista de Clientes - IIBB") y generacion de hojas 2025
# --------------------------------------------------------------------------

def normalizar_dia(texto: str) -> str:
    m = re.search(r"\d+", texto)
    return f"DIA {m.group()}" if m else texto.strip()


def leer_padron(wb) -> Dict[int, List[Tuple[str, int, str]]]:
    """Lee la hoja 'Lista de Clientes - IIBB': para cada anio, la lista de
    (nombre, cuit, dia_asignado)."""
    ws = wb[HOJA_CLIENTES]
    padron: Dict[int, List[Tuple[str, int, str]]] = {2024: [], 2025: []}
    columnas = {2024: (1, 2), 2025: (4, 5)}
    for anio, (col_nombre, col_cuit) in columnas.items():
        dia_actual = "DIA 1"
        for fila in range(3, ws.max_row + 1):
            nombre = ws.cell(row=fila, column=col_nombre).value
            cuit = ws.cell(row=fila, column=col_cuit).value
            if isinstance(nombre, str) and normalizar(nombre).startswith("dia "):
                dia_actual = normalizar_dia(nombre)
                continue
            if nombre and isinstance(cuit, (int, float)):
                padron[anio].append((str(nombre).strip(), int(cuit), dia_actual))
    return padron


FUENTE_TITULO = Font(bold=True)
FORMATO_MONEDA = "#,##0.00"


def _escribir_bloque_vacio(ws: Worksheet, fila_inicio: int, nombre: str,
                            cuit: int, password: Optional[str] = None) -> int:
    """Escribe la estructura vacia de un cliente (encabezado + 12 conceptos +
    alicuota + CUIT/password) y devuelve la fila donde deberia empezar el
    siguiente bloque. El formato es simple y consistente, no una copia
    pixel-a-pixel de las hojas originales (que tienen estilos hechos a mano
    y ligeramente distintos entre si)."""
    ws.cell(row=fila_inicio, column=1, value=nombre).font = FUENTE_TITULO
    for i, mes in enumerate(MESES):
        ws.cell(row=fila_inicio, column=2 + i, value=mes).font = FUENTE_TITULO

    fila = fila_inicio + 1
    for concepto in CONCEPTOS:
        etiqueta = "Importe a pagar(subtotal)" if concepto == "importe a pagar(subtotal)" else concepto.title()
        ws.cell(row=fila, column=1, value=etiqueta)
        for col in range(2, 14):
            ws.cell(row=fila, column=col).number_format = FORMATO_MONEDA
        ws.cell(row=fila, column=14, value=f"=SUM(B{fila}:M{fila})").number_format = FORMATO_MONEDA
        fila += 1

    fila_alicuota = fila
    for col in range(2, 14):
        ws.cell(row=fila_alicuota, column=col).number_format = "0.00"
    fila += 1

    ws.cell(row=fila, column=1, value=cuit)
    if password:
        ws.cell(row=fila, column=2, value=password)
    fila += 2  # una fila en blanco de separacion antes del siguiente bloque
    return fila


def crear_hojas_2025(wb, padron: Dict[int, List[Tuple[str, int, str]]],
                      indice_2024: Dict[Tuple[int, int], BloqueCliente]) -> None:
    """Crea 'DIA N - 2025' para cada dia que aparezca en el padron 2025, con
    un bloque vacio por cliente. Si el cliente ya existia en 2024 y tenia una
    contrasena propia escrita, se copia (se prueba primero); si no, en el
    login se usan las contrasenas por defecto. Es idempotente: si la hoja ya
    existe, no la vuelve a crear."""
    dias = sorted({dia for _, _, dia in padron[2025]},
                  key=lambda d: int(re.search(r"\d+", d).group()))

    for dia in dias:
        nombre_hoja = f"{dia}{SUFIJO_2025}"
        if nombre_hoja in wb.sheetnames:
            continue
        ws = wb.create_sheet(nombre_hoja)
        ws.column_dimensions["A"].width = 32
        for col in range(2, 15):
            ws.column_dimensions[get_column_letter(col)].width = 14

        fila = 1
        clientes = [(n, c) for n, c, d in padron[2025] if d == dia]
        for nombre, cuit in clientes:
            previo = indice_2024.get((2024, cuit))
            password = previo.password if previo else None
            fila = _escribir_bloque_vacio(ws, fila, nombre, cuit, password)
        logger.info("Hoja %s creada con %d clientes", nombre_hoja, len(clientes))


# --------------------------------------------------------------------------
# Deteccion de meses pendientes
# --------------------------------------------------------------------------

def meses_pendientes(ws: Worksheet, bloque: BloqueCliente) -> List[int]:
    """Un mes se considera 'hecho' si Base Imponible y Total pagado ya
    tienen algo cargado (aunque sea 0). Si falta alguno de los dos, se
    considera pendiente y hay que ir a buscarlo a AGIP."""
    fila_bi = bloque.filas_concepto.get("base imponible")
    fila_tp = bloque.filas_concepto.get("total pagado")
    if not fila_bi or not fila_tp:
        return []
    pendientes = []
    for mes, col in sorted(bloque.columnas_mes.items()):
        v1 = ws[f"{col}{fila_bi}"].value
        v2 = ws[f"{col}{fila_tp}"].value
        if v1 in (None, "") or v2 in (None, ""):
            pendientes.append(mes)
    return pendientes


# --------------------------------------------------------------------------
# Guardado seguro
# --------------------------------------------------------------------------

def guardar_workbook(wb, ruta: Path, intentos: int = 5, espera: float = 2.0) -> None:
    """Guarda de forma atomica (escribe a un .tmp y reemplaza). En Windows,
    si el Excel esta abierto en otro programa (Excel, OneDrive
    sincronizando, etc.) el reemplazo falla con PermissionError; en vez de
    cortar toda la corrida, reintenta unas veces antes de rendirse."""
    tmp = ruta.with_suffix(f".tmp{ruta.suffix}")
    wb.save(tmp)
    for intento in range(1, intentos + 1):
        try:
            tmp.replace(ruta)
            return
        except PermissionError:
            if intento == intentos:
                tmp.unlink(missing_ok=True)
                raise RuntimeError(
                    f"No pude guardar '{ruta.name}': el archivo parece estar abierto en otro "
                    "programa (Excel, OneDrive sincronizando, etc). Cerralo y volve a correr "
                    "el script."
                ) from None
            logger.warning("'%s' esta bloqueado para escritura (intento %d/%d) -- "
                            "¿esta abierto en Excel? Reintento en %.0fs.",
                            ruta.name, intento, intentos, espera)
            time.sleep(espera)


# --------------------------------------------------------------------------
# Automatizacion del navegador (Playwright). Import perezoso para que
# --dry-run y --crear-demo funcionen sin tener playwright instalado.
# --------------------------------------------------------------------------

def _importar_playwright():
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise SystemExit(
            "Falta playwright. Instalalo con:\n"
            "  pip install playwright\n"
            "  playwright install chromium"
        ) from exc
    return sync_playwright


def pausar(base: Optional[float] = None, jitter: Optional[float] = None) -> None:
    """Espera ademas de las esperas automaticas de Playwright, con una
    variacion aleatoria +/- jitter para que el ritmo no sea perfectamente
    uniforme entre paso y paso. Ver PAUSA_ACCION / --pausa-accion."""
    base = PAUSA_ACCION if base is None else base
    jitter = PAUSA_ACCION_JITTER if jitter is None else jitter
    time.sleep(max(0.1, base + random.uniform(-jitter, jitter)))


def pausar_entre_clientes() -> None:
    """Espera mas larga entre el cierre de la sesion de un cliente y el
    login del siguiente. Ver PAUSA_ENTRE_CLIENTES / --pausa-clientes."""
    pausar(PAUSA_ENTRE_CLIENTES, PAUSA_ENTRE_CLIENTES_JITTER)


def _click_si_existe(page, patron_texto: str, tiempo: int) -> bool:
    """Intenta clickear un elemento por texto. Devuelve False tanto si no
    aparecio como si el click en si fallo (elemento tapado, se desprendio
    del DOM, etc.) -- nunca deja escapar la excepcion, para que un boton
    inesperado no tire abajo el procesamiento de todo un cliente."""
    loc = page.get_by_text(re.compile(patron_texto, re.I)).first
    try:
        loc.wait_for(state="visible", timeout=tiempo)
        loc.click(timeout=tiempo)
        pausar()
        return True
    except Exception:
        return False


# El enlace 'Clave Ciudad' de la home de AGIP abre claveciudad.agip.gob.ar
# en una PESTANA NUEVA (target=_blank). En esa pestana nueva hay OTRO
# enlace, tambien de texto 'Clave Ciudad' pero con onclick="toggleLogin()",
# que recien ahi despliega el formulario de usuario/contrasena. Selectores
# tomados del HTML real de la pagina (no son una adivinanza por texto).
SELECTOR_LINK_CLAVE_CIUDAD = 'a[href="https://claveciudad.agip.gob.ar/"]'
SELECTOR_LINK_TOGGLE_LOGIN = 'a[onclick*="toggleLogin"]'


def _intentar_login(page, cuit: int, password: str, tiempo_espera: int):
    """Un unico intento de login con una contrasena puntual. Devuelve la
    pagina donde quedo la sesion activa (la pestana nueva que abre Clave
    Ciudad) si el login funciono, o None si fallo en cualquier paso."""
    page.goto(BASE_URL, wait_until="domcontentloaded")
    pausar()

    try:
        with page.expect_popup(timeout=tiempo_espera) as popup_info:
            enlace = page.locator(SELECTOR_LINK_CLAVE_CIUDAD).first
            enlace.wait_for(state="visible", timeout=tiempo_espera)
            enlace.click()
        nueva_pagina = popup_info.value
    except Exception:
        logger.error("No se abrio la pestana de Clave Ciudad (%s) en %s",
                      SELECTOR_LINK_CLAVE_CIUDAD, BASE_URL)
        return None

    try:
        nueva_pagina.wait_for_load_state("domcontentloaded", timeout=tiempo_espera)
    except Exception:
        pass
    pausar()

    try:
        toggle = nueva_pagina.locator(SELECTOR_LINK_TOGGLE_LOGIN).first
        toggle.wait_for(state="visible", timeout=tiempo_espera)
        toggle.click()
        pausar()
    except Exception:
        logger.error("No encontre el segundo enlace Clave Ciudad (%s) en la pestana nueva",
                      SELECTOR_LINK_TOGGLE_LOGIN)
        nueva_pagina.close()
        return None

    campo_password = nueva_pagina.locator('input[type="password"]').first
    try:
        campo_password.wait_for(state="visible", timeout=tiempo_espera)
    except Exception:
        logger.error("No aparecio el campo de contrasena tras el segundo click (toggleLogin)")
        nueva_pagina.close()
        return None

    # El campo de usuario se busca dentro del mismo <form> si existe; si el
    # formulario revelado por toggleLogin no usa <form>, buscamos en toda
    # la pestana como respaldo.
    formulario = campo_password.locator("xpath=ancestor::form[1]")
    contenedor = formulario if formulario.count() > 0 else nueva_pagina
    campo_usuario = contenedor.locator('input[type="text"], input[type="tel"], input:not([type])').first
    campo_usuario.click()
    pausar(0.3, 0.2)
    campo_usuario.fill(str(cuit))
    pausar(0.4, 0.25)
    campo_password.click()
    pausar(0.25, 0.15)
    campo_password.fill(password)
    pausar(0.45, 0.25)

    boton = contenedor.get_by_role("button", name=re.compile(r"ingresar|entrar|iniciar", re.I))
    if boton.count() > 0:
        boton.first.click()
    else:
        campo_password.press("Enter")

    try:
        nueva_pagina.wait_for_load_state("networkidle", timeout=tiempo_espera)
    except Exception:
        pass
    pausar()

    # Si el login fallo, lo mas probable es que sigamos viendo el mismo
    # campo de contrasena (formulario no avanzo).
    sigue_en_login = nueva_pagina.locator('input[type="password"]').count() > 0
    if sigue_en_login:
        return None
    return nueva_pagina


def iniciar_sesion(page, cuit: int, candidatas: List[str], tiempo_espera: int):
    """Prueba cada contrasena candidata (la explicita del Excel, si la hay,
    y despues las 2 por defecto) hasta que una funcione. Devuelve
    (password_que_funciono, pagina_activa), o (None, None) si ninguna
    funciono. pagina_activa puede ser distinta de 'page': Clave Ciudad abre
    en una pestana nueva, y ahi es donde sigue el resto de la sesion."""
    for i, password in enumerate(candidatas):
        pagina_activa = _intentar_login(page, cuit, password, tiempo_espera)
        if pagina_activa:
            return password, pagina_activa
        logger.info("CUIT %s: contrasena candidata %d/%d no funciono", cuit, i + 1, len(candidatas))
    return None, None


def _texto_celda(locator_celda, tiempo: int) -> str:
    try:
        return locator_celda.locator(".x-grid-cell-inner").inner_text(timeout=tiempo).strip()
    except Exception:
        try:
            return locator_celda.inner_text(timeout=tiempo).strip()
        except Exception:
            return ""


def buscar_fila_ddjj(page, anio: int, mes_idx: int, tiempo_espera: int):
    """Busca, entre las filas de la grilla 'Declaraciones Juradas
    Presentadas' (ExtJS) YA renderizadas, la fila cuyo periodo (3ra
    columna, formato 'YYYY-MM') coincide con el mes buscado. La grilla no
    muestra el nombre del mes en español, sino el periodo en ese formato.

    Si hay 2 filas para el mismo periodo (Original + Rectificativa, ultima
    columna), se prioriza la Rectificativa por ser la que corrige/reemplaza
    a la original."""
    periodo_objetivo = f"{anio}-{mes_idx:02d}"
    filas = page.locator("tr.x-grid-row")
    encontradas = []
    for i in range(filas.count()):
        fila = filas.nth(i)
        celdas = fila.locator("td.x-grid-cell")
        if celdas.count() < 3:
            continue
        periodo = _texto_celda(celdas.nth(2), tiempo_espera)
        if periodo == periodo_objetivo:
            tipo = _texto_celda(celdas.last, tiempo_espera)
            encontradas.append((fila, tipo))
    if not encontradas:
        return None
    for fila, tipo in encontradas:
        if normalizar(tipo) == "rectificativa":
            return fila
    return encontradas[0][0]


def buscar_fila_ddjj_con_scroll(page, anio: int, mes_idx: int, tiempo_espera: int, intentos_scroll: int = 15):
    """Igual que buscar_fila_ddjj, pero si no la encuentra de entrada (la
    grilla tiene 110+ filas y puede no renderizar todas de una, tipico de
    ExtJS), va scrolleando de a poco y reintentando. Heuristica sin
    confirmar todavia contra el sitio real -- si no hace falta scrollear,
    simplemente encuentra la fila en el primer intento y no scrollea nada."""
    fila = buscar_fila_ddjj(page, anio, mes_idx, tiempo_espera)
    if fila:
        return fila
    for _ in range(intentos_scroll):
        try:
            page.mouse.wheel(0, 400)
        except Exception:
            break
        pausar(0.3, 0.15)
        fila = buscar_fila_ddjj(page, anio, mes_idx, tiempo_espera)
        if fila:
            return fila
    return None


def ir_a_declaracion(page, anio: int, mes_idx: int, tiempo_espera: int) -> bool:
    """Navega e-Sicol -> Declaraciones Juradas presentadas -> busca y abre
    la fila del periodo pedido."""
    if not _click_si_existe(page, r"e-?sicol", tiempo_espera):
        logger.error("No encontre el enlace 'e-Sicol'")
        return False
    page.wait_for_load_state("networkidle", timeout=tiempo_espera)

    _click_si_existe(page, r"declaraciones?\s+juradas\s+presentadas", tiempo_espera)
    page.wait_for_load_state("networkidle", timeout=tiempo_espera)

    fila = buscar_fila_ddjj_con_scroll(page, anio, mes_idx, tiempo_espera)
    if not fila:
        logger.warning("No encontre la fila de %s-%02d en la grilla de DDJJ", anio, mes_idx)
        return False
    try:
        fila.dblclick(timeout=tiempo_espera)
        pausar()
    except Exception:
        logger.error("Encontre la fila de %s-%02d pero no pude hacerle doble click", anio, mes_idx)
        return False
    page.wait_for_load_state("networkidle", timeout=tiempo_espera)
    return True


def expandir_secciones(page, tiempo: int) -> None:
    """Las Retenciones/Percepciones parecen estar en pestañas o acordeones
    plegados (asi lo describe DETALLE: 'Desplegar la pestaña...'). Intenta
    abrirlas; si ya estan abiertas o no existen como tales, no hace nada."""
    for texto in ("Retenciones", "Percepciones"):
        _click_si_existe(page, texto, tiempo)


def obtener_valor_por_etiqueta(page, etiqueta: str, tiempo: int) -> Optional[str]:
    """Busca un texto en la pagina y devuelve el numero que encuentra al
    lado (hermano siguiente, fila de tabla, o padre). Prueba primero
    coincidencia exacta (mas segura) y despues por substring."""
    patrones = [
        re.compile(rf"^\s*{re.escape(etiqueta)}\s*:?\s*$", re.I),
        re.compile(re.escape(etiqueta), re.I),
    ]
    xpaths = (
        "xpath=following-sibling::*[1]",
        "xpath=../following-sibling::*[1]",
        "xpath=ancestor::tr[1]//td[last()]",
        "xpath=parent::*",
    )
    for patron in patrones:
        loc = page.get_by_text(patron).first
        try:
            loc.wait_for(state="visible", timeout=tiempo)
        except Exception:
            continue
        for xp in xpaths:
            try:
                texto = loc.locator(xp).inner_text(timeout=800)
            except Exception:
                continue
            if parsear_numero_ar(texto) is not None:
                return texto
    return None


def extraer_campos(page, tiempo_espera: int) -> Dict[str, float]:
    """Extrae los valores de la DDJJ actualmente abierta. 'intereses' se
    calcula (no tiene campo propio), segun la nota de la hoja DETALLE."""
    expandir_secciones(page, tiempo_espera)
    valores: Dict[str, float] = {}

    for concepto, etiqueta in UBICACION_AGIP.items():
        if concepto in ("intereses", "total pagado"):
            continue
        texto = obtener_valor_por_etiqueta(page, etiqueta, tiempo_espera)
        numero = parsear_numero_ar(texto)
        if numero is not None:
            valores[concepto] = numero
        else:
            logger.warning("No encontre '%s' en la pagina", etiqueta)

    texto_total = obtener_valor_por_etiqueta(page, UBICACION_AGIP["total pagado"], tiempo_espera)
    texto_importe = obtener_valor_por_etiqueta(page, UBICACION_AGIP["importe a pagar(subtotal)"], tiempo_espera)
    total_pagina = parsear_numero_ar(texto_total)
    importe_subtotal = parsear_numero_ar(texto_importe)
    if total_pagina is not None:
        valores["total pagado"] = total_pagina
    if total_pagina is not None and importe_subtotal is not None:
        valores["intereses"] = round(total_pagina - importe_subtotal, 2)

    return valores


def escribir_valores(ws: Worksheet, bloque: BloqueCliente, mes: int, valores: Dict[str, float]) -> None:
    col = bloque.columnas_mes[mes]
    col_idx = column_index_from_string(col)
    for concepto, valor in valores.items():
        fila = bloque.filas_concepto.get(concepto)
        if fila:
            ws.cell(row=fila, column=col_idx, value=valor)
    if bloque.fila_alicuota and "alicuota" in valores:
        ws.cell(row=bloque.fila_alicuota, column=col_idx, value=valores["alicuota"])


def procesar_cliente(page, ws: Worksheet, bloque: BloqueCliente, pendientes: List[int],
                      candidatas: List[str], tiempo_espera: int, guardar_cb) -> bool:
    """Devuelve True si logro iniciar sesion (con alguna de las candidatas),
    False si ninguna contrasena funciono."""
    logger.info("Cliente %s (CUIT %s, %s): %d mes(es) pendientes, %d contrasena(s) a probar",
                bloque.nombre, bloque.cuit, bloque.anio, len(pendientes), len(candidatas))
    password_ok, pagina_activa = iniciar_sesion(page, bloque.cuit, candidatas, tiempo_espera)
    if not password_ok:
        logger.error("CUIT %s: ninguna contrasena funciono (probe %d)", bloque.cuit, len(candidatas))
        return False
    bloque.password = password_ok

    for mes in pendientes:
        try:
            if not ir_a_declaracion(pagina_activa, bloque.anio, mes, tiempo_espera):
                logger.warning("CUIT %s: no encontre la DDJJ de %s/%s", bloque.cuit, MESES[mes - 1], bloque.anio)
                continue
            valores = extraer_campos(pagina_activa, tiempo_espera)
            if not valores:
                logger.warning("CUIT %s: no se pudo extraer nada de %s/%s", bloque.cuit, MESES[mes - 1], bloque.anio)
                continue
            escribir_valores(ws, bloque, mes, valores)
            guardar_cb()
            logger.info("CUIT %s: %s/%s completado (%d campos)",
                        bloque.cuit, MESES[mes - 1], bloque.anio, len(valores))
        except Exception:
            logger.exception("Error procesando CUIT %s, mes %s/%s", bloque.cuit, mes, bloque.anio)

    return True


# --------------------------------------------------------------------------
# Planilla de ejemplo (para probar la logica de Excel sin tocar datos reales)
# --------------------------------------------------------------------------

def construir_planilla_demo(ruta: Path) -> None:
    wb = Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet(HOJA_CLIENTES)
    ws["E2"], ws["F2"], ws["G2"], ws["H2"] = "DDJJ HECHA", "COMPLETOS", "NO ESTA CERRADO", "NO HECHOS"
    ws["A3"], ws["D3"] = "INGRESOS BRUTOS 2024", "INGRESOS BRUTOS 2025"

    datos_2024 = [("DIA 1", None), ("Cliente Ejemplo Uno", 20111111112),
                  ("Cliente Ejemplo Dos", 20222222223), ("DIA 2", None),
                  ("Cliente Ejemplo Tres", 20333333334)]
    datos_2025 = [("DIA 1", None), ("Cliente Ejemplo Uno", 20111111112),
                  ("Cliente Nuevo 2025", 20444444445)]

    for i, (nombre, cuit) in enumerate(datos_2024, start=4):
        ws.cell(row=i, column=1, value=nombre)
        if cuit:
            ws.cell(row=i, column=2, value=cuit)
    for i, (nombre, cuit) in enumerate(datos_2025, start=4):
        ws.cell(row=i, column=4, value=nombre)
        if cuit:
            ws.cell(row=i, column=5, value=cuit)

    ws_dia1 = wb.create_sheet("DIA 1")
    fila = _escribir_bloque_vacio(ws_dia1, 1, "Cliente Ejemplo Uno", 20111111112, "clave-demo-1")
    for col in range(2, 8):  # enero..junio ya cargados -> deberian salir como "hechos"
        ws_dia1.cell(row=2, column=col, value=1000.0 * col)
        ws_dia1.cell(row=13, column=col, value=50.0 * col)
    _escribir_bloque_vacio(ws_dia1, fila, "Cliente Ejemplo Dos", 20222222223, "clave-demo-2")

    ws_dia2 = wb.create_sheet("DIA 2")
    _escribir_bloque_vacio(ws_dia2, 1, "Cliente Ejemplo Tres", 20333333334, "clave-demo-3")

    wb.create_sheet("DETALLE")
    wb.save(ruta)


def candidatas_password(bloque: BloqueCliente, es_cuit_filtrado: bool,
                         password_override: Optional[str]) -> List[str]:
    """Orden de intento: la contrasena explicita del Excel (si la hay), el
    --password de linea de comandos (solo si este cliente es el filtrado con
    --cuit), y despues las 2 contrasenas por defecto."""
    candidatas = []
    if bloque.password:
        candidatas.append(bloque.password)
    if es_cuit_filtrado and password_override and password_override not in candidatas:
        candidatas.append(password_override)
    for p in CONTRASENAS_POR_DEFECTO:
        if p not in candidatas:
            candidatas.append(p)
    return candidatas


def calcular_trabajo(wb, indice: Dict[Tuple[int, int], BloqueCliente], anios: set,
                      cuit_filtro: Optional[int], password_override: Optional[str]
                      ) -> List[Tuple[BloqueCliente, Worksheet, List[int], List[str]]]:
    """Arma la lista de (bloque, hoja, meses_pendientes, candidatas_password)
    a procesar, ordenada por anio (todos los de 2024 antes que cualquiera
    de 2025). Ya no salteamos por falta de contrasena explicita: se prueban
    las contrasenas por defecto en el momento del login."""
    trabajo = []
    for (anio, cuit), bloque in indice.items():
        if anio not in anios:
            continue
        if cuit_filtro and cuit != cuit_filtro:
            continue
        ws = wb[bloque.hoja]
        pendientes = meses_pendientes(ws, bloque)
        if not pendientes:
            continue
        candidatas = candidatas_password(bloque, cuit_filtro == cuit, password_override)
        trabajo.append((bloque, ws, pendientes, candidatas))
    trabajo.sort(key=lambda item: item[0].anio)
    return trabajo


def generar_resumen(wb, padron: Dict[int, List[Tuple[str, int, str]]],
                     indice: Dict[Tuple[int, int], BloqueCliente]) -> None:
    """Recorre TODO el padron (2024 y 2025, todos los clientes que figuran en
    'Lista de Clientes - IIBB') y lo compara contra lo que hay indexado en
    las hojas DIA-N, para reportar el estado real: completos, con meses
    pendientes, o ni siquiera ubicados en ninguna hoja."""
    completos: List[Tuple[int, str, int]] = []
    pendientes: List[Tuple[int, str, int, List[int]]] = []
    no_encontrados: List[Tuple[int, str, int, str]] = []

    for anio in (2024, 2025):
        for nombre, cuit, dia in padron.get(anio, []):
            bloque = indice.get((anio, cuit))
            if not bloque:
                no_encontrados.append((anio, nombre, cuit, dia))
                continue
            ws = wb[bloque.hoja]
            meses_falt = meses_pendientes(ws, bloque)
            if meses_falt:
                pendientes.append((anio, bloque.nombre, cuit, meses_falt))
            else:
                completos.append((anio, bloque.nombre, cuit))

    total = len(completos) + len(pendientes) + len(no_encontrados)
    logger.info("=" * 60)
    logger.info("RESUMEN: %d clientes en el padron (2024 + 2025)", total)
    logger.info("  Completos:            %d", len(completos))
    logger.info("  Con meses pendientes: %d", len(pendientes))
    logger.info("  No ubicados en hoja:  %d", len(no_encontrados))
    if pendientes:
        logger.info("--- Con meses pendientes ---")
        for anio, nombre, cuit, meses in pendientes:
            meses_txt = ", ".join(MESES[m - 1] for m in meses)
            logger.info("  [%s] %s (CUIT %s): %s", anio, nombre, cuit, meses_txt)
    if no_encontrados:
        logger.info("--- No ubicados (revisar esa fila/bloque a mano) ---")
        for anio, nombre, cuit, dia in no_encontrados:
            logger.info("  [%s] %s (CUIT %s) - %s", anio, nombre, cuit, dia)
    logger.info("=" * 60)


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def construir_argumentos() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--excel", type=Path, help="Ruta al archivo .xlsx real")
    parser.add_argument("--crear-demo", action="store_true", help="Genera demo.xlsx de ejemplo y termina")
    parser.add_argument("--dry-run", action="store_true", help="Solo analiza la planilla, no abre el navegador")
    parser.add_argument("--anios", default="2024,2025", help="Anios a procesar, ej: 2024,2025")
    parser.add_argument("--cuit", type=int, help="Procesar un unico CUIT (para probar)")
    parser.add_argument("--max-clientes", type=int, help="Limite de clientes a procesar en esta corrida")
    parser.add_argument("--sin-headless", action="store_true", help="Muestra el navegador (recomendado al probar)")
    parser.add_argument("--password", help="Contrasena a usar junto con --cuit si la celda todavia esta vacia")
    parser.add_argument("--timeout", type=int, default=15000, help="Timeout de Playwright en ms (default 15000)")
    parser.add_argument("--pausa-accion", type=float, default=PAUSA_ACCION,
                         help=f"Segundos de espera (promedio) tras cada click/navegacion (default {PAUSA_ACCION})")
    parser.add_argument("--pausa-clientes", type=float, default=PAUSA_ENTRE_CLIENTES,
                         help=f"Segundos de espera (promedio) entre cliente y cliente (default {PAUSA_ENTRE_CLIENTES})")
    return parser.parse_args()


def main() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(f"iibb_agip_{datetime.now():%Y%m%d_%H%M%S}.log", encoding="utf-8"),
        ],
    )

    args = construir_argumentos()

    global PAUSA_ACCION, PAUSA_ENTRE_CLIENTES
    PAUSA_ACCION = args.pausa_accion
    PAUSA_ENTRE_CLIENTES = args.pausa_clientes

    if args.crear_demo:
        ruta_demo = Path("demo.xlsx")
        construir_planilla_demo(ruta_demo)
        logger.info("Planilla de ejemplo creada en %s", ruta_demo.resolve())
        return

    if not args.excel:
        logger.error("Falta --excel <archivo.xlsx> (o --crear-demo para probar la logica sin datos reales)")
        sys.exit(1)

    ruta = args.excel
    anios = {int(a) for a in args.anios.split(",")}

    if args.dry_run:
        # Modo de solo lectura: no crea backup, no crea hojas 2025, no guarda
        # nada. Es seguro correrlo directo sobre el archivo real.
        logger.info("Modo --dry-run: no se modifica ni se guarda el archivo.")
        wb = load_workbook(ruta)
        padron = leer_padron(wb)
        indice = indexar_workbook(wb)
        generar_resumen(wb, padron, indice)
        return

    respaldo = ruta.with_name(ruta.stem + ".backup" + ruta.suffix)
    if not respaldo.exists():
        shutil.copy2(ruta, respaldo)
        logger.info("Backup creado en %s", respaldo)

    wb = load_workbook(ruta)
    padron = leer_padron(wb)
    indice = indexar_workbook(wb)

    crear_hojas_2025(wb, padron, indice)
    guardar_workbook(wb, ruta)
    indice = indexar_workbook(wb)  # re-indexa incluyendo las hojas 2025 recien creadas

    trabajo = calcular_trabajo(wb, indice, anios, args.cuit, args.password)
    logger.info("Clientes con meses pendientes: %d", len(trabajo))

    if args.max_clientes:
        trabajo = trabajo[: args.max_clientes]

    fallos_login: List[Tuple[int, str, int]] = []
    sync_playwright = _importar_playwright()
    with sync_playwright() as pw:
        navegador = pw.chromium.launch(headless=not args.sin_headless)
        try:
            for bloque, ws, pendientes, candidatas in trabajo:
                contexto = navegador.new_context()
                pagina = contexto.new_page()
                try:
                    ok = procesar_cliente(pagina, ws, bloque, pendientes, candidatas, args.timeout,
                                           guardar_cb=lambda: guardar_workbook(wb, ruta))
                    if not ok:
                        fallos_login.append((bloque.anio, bloque.nombre, bloque.cuit))
                except Exception:
                    logger.exception("Error inesperado con CUIT %s, sigo con el siguiente cliente", bloque.cuit)
                    fallos_login.append((bloque.anio, bloque.nombre, bloque.cuit))
                finally:
                    contexto.close()
                pausar_entre_clientes()
        finally:
            guardar_workbook(wb, ruta)
            navegador.close()

    if fallos_login:
        logger.warning("%d cliente(s) no se pudieron loguear con ninguna contrasena probada:",
                        len(fallos_login))
        for anio, nombre, cuit in fallos_login:
            logger.warning("  [%s] %s (CUIT %s)", anio, nombre, cuit)

    logger.info("Listo. Planilla actualizada: %s", ruta.resolve())
    indice_final = indexar_workbook(wb)
    generar_resumen(wb, padron, indice_final)


if __name__ == "__main__":
    main()
