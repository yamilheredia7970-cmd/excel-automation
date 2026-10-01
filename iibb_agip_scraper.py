#!/usr/bin/env python3
"""Automatiza la carga de Ingresos Brutos (AGIP) en la planilla de clientes.

Uso:
    python iibb_agip_scraper.py --crear-demo
    python iibb_agip_scraper.py --excel demo.xlsx --dry-run
    python iibb_agip_scraper.py --excel "IIBB ANUAL 2025-4.xlsx" --cuit 20934379447 --meses 1,2 --sin-headless
    python iibb_agip_scraper.py --excel "IIBB ANUAL 2025-4.xlsx"

Este archivo no contiene ninguna credencial. El usuario de AGIP es el
CUIT/CUIL de cada cliente. La contrasena se resuelve asi, en orden: (1) si
la celda al lado del CUIT en la hoja "DIA N" tiene algo escrito, se prueba
primero, (2) si no funciona o no hay ninguna escrita, se prueban las
contrasenas por defecto (CONTRASENAS_POR_DEFECTO, mas abajo). Un cliente
solo se reporta como "fallido" si ninguna de las candidatas funciono.

Por cada cliente y cada mes pendiente (primero todo 2024, despues 2025):
  1. abre la DDJJ de ese periodo; si tiene rectificativas, SOLO la ultima;
  2. confirma que lo abierto sea el anio/mes/version pedidos;
  3. lee cada seccion del arbol (Rubro 1, Rubro 2 y Liquidacion) y cierra
     cada ventana apenas la lee;
  4. completa en el Excel SOLO las celdas vacias de ese mes (nunca pisa lo
     que ya estaba cargado a mano), cada actividad en su propia fila, y
     guarda antes de pasar al mes siguiente.

Se puede cortar en cualquier momento (Ctrl+C) y volver a correr: retoma
desde los meses que siguen pendientes.

El navegador (Chromium) se usa siempre, porque e-Sicol es una aplicacion
que arma sus pantallas con JavaScript. Por defecto corre sin ventana; con
--sin-headless se ve todo lo que hace. En los dos casos, cada paso queda en
el log (iibb_agip_FECHA.log) y, si algo falla, se guarda en la carpeta
capturas_errores/ una captura de pantalla (.png) y el HTML de ese momento.

Pendiente de confirmar contra el sitio real: la seccion "Liquidacion del
Impuesto, Presentacion" (saldo a favor, importe a pagar y total pagado).
Se busca por las etiquetas que indica la hoja DETALLE; si no las
encuentra, el log muestra el texto de esas ventanas para poder ajustarlo.
"""

from __future__ import annotations

import argparse
import difflib
import logging
import random
import re
import shutil
import sys
import time
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
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

# Como escribe cada etiqueta la planilla original (para las hojas 2025).
ETIQUETAS_EXCEL = {
    "base imponible": "Base Imponible",
    "anticipo determinado": "anticipo determinado",
    "retenciones": "Retenciones",
    "retenciones bancarias": "Retenciones bancarias",
    "percepciones": "Percepciones",
    "impuestos internos": "Impuestos internos",
    "pago a cuenta": "Pago a cuenta",
    "otros creditos": "Otros Creditos",
    "saldo a favor": "Saldo a favor",
    "importe a pagar(subtotal)": "Importe a pagar(subtotal)",
    "intereses": "Intereses",
    "total pagado": "Total pagado",
}

# Convencion de la planilla original: estos conceptos quedan en blanco
# cuando AGIP dice 0 (ej. Retenciones bancarias de Burgos Jose en mayo).
# Base imponible, anticipo, importe a pagar, intereses y total pagado, en
# cambio, se escriben aunque sean 0.
CONCEPTOS_CERO_EN_BLANCO = {
    "retenciones", "retenciones bancarias", "percepciones", "impuestos internos",
    "pago a cuenta", "otros creditos", "saldo a favor",
}

UMBRAL_CUIT = 10 ** 10  # un CUIT/CUIL tiene 11 digitos

# Timeouts cortos y fijos para cosas que YA deberian estar en pantalla: si
# se usara el --timeout general (15s por defecto) en cada una, un nodo que
# no existe en alguna DDJJ puede sumar minutos enteros de espera.
TIMEOUT_NODO_MS = 6000
TIMEOUT_CONTENIDO_VENTANA_MS = 8000
MAX_MESES_SEGUIDOS_CON_ERROR = 3
MAX_CLIENTES_SEGUIDOS_SIN_AGIP = 3   # si AGIP no carga el login, se corta la corrida

# Cuando algo falla se guarda una captura de pantalla (y el HTML) de lo que
# mostraba el navegador en ese momento, para poder verlo aunque se corra
# sin ventana. De un mismo aviso repetido se guardan solo las primeras.
CARPETA_CAPTURAS = Path("capturas_errores")
MAX_CAPTURAS_POR_AVISO = 3
MAX_CAPTURAS_POR_ERROR = 20

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
# Utilidades de texto y numeros
# --------------------------------------------------------------------------

def normalizar(texto) -> str:
    if texto is None:
        return ""
    texto = str(texto).replace(" ", " ").strip().lower()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", texto)


def normalizar_etiqueta(texto) -> str:
    """Lleva una etiqueta de la columna A a la forma de CONCEPTOS, tolerando
    las variantes escritas a mano que aparecen en la planilla real."""
    t = normalizar(texto)
    t = re.sub(r"\s*\(\s*", "(", t)
    t = re.sub(r"\s*\)", ")", t)
    if t == "pagos a cuenta":
        return "pago a cuenta"
    if re.match(r"^conceptos? que no int\w* la base imponible", t):
        return "impuestos internos"
    return t


MESES_NORM = [normalizar(m) for m in MESES]


def mes_desde_texto(texto) -> Optional[int]:
    t = normalizar(texto)
    if t == "setiembre":
        return 9
    return MESES_NORM.index(t) + 1 if t in MESES_NORM else None


RE_NUMERO = re.compile(r"\d[\d.,]*")


def parsear_numero_ar(texto) -> Optional[float]:
    """'$ 1.234,56' -> 1234.56, '3%' -> 3.0, '-$10,00' -> -10.0. None si
    no hay numero. Si ya viene como numero, lo devuelve tal cual."""
    if texto is None or isinstance(texto, bool):
        return None
    if isinstance(texto, (int, float)):
        return float(texto)
    t = str(texto).replace(" ", " ").strip()
    if not t or t in {"-", "--"}:
        return None
    m = RE_NUMERO.search(t)
    if not m:
        return None
    crudo = m.group(0).rstrip(".,")
    negativo = "-" in t[:m.start()] or (t.startswith("(") and t.endswith(")"))
    if "," in crudo:
        crudo = crudo.replace(".", "").replace(",", ".")
    elif crudo.count(".") > 1:
        crudo = crudo.replace(".", "")
    elif "." in crudo:
        entero, decimales = crudo.split(".")
        if len(decimales) == 3:  # '1.234' en formato argentino son miles
            crudo = entero + decimales
    try:
        valor = float(crudo)
    except ValueError:
        return None
    return -valor if negativo else valor


def _vacia(valor) -> bool:
    return valor is None or (isinstance(valor, str) and not valor.strip())


def _levenshtein(a: str, b: str) -> int:
    previa = list(range(len(b) + 1))
    for i, ca in enumerate(a, start=1):
        actual = [i]
        for j, cb in enumerate(b, start=1):
            actual.append(min(previa[j] + 1, actual[j - 1] + 1, previa[j - 1] + (ca != cb)))
        previa = actual
    return previa[-1]


# --------------------------------------------------------------------------
# Modelo: un bloque = los datos de un cliente dentro de una hoja "DIA N"
# --------------------------------------------------------------------------

@dataclass
class BloqueActividad:
    """Filas de UNA actividad dentro del bloque del cliente. La primera es
    la del bloque principal; si el cliente tiene mas de una actividad, la
    planilla agrega abajo un sub-bloque por cada una (Base Imponible,
    anticipo determinado, alicuota y codigo), como indica la hoja DETALLE."""
    fila_base: int
    fila_anticipo: Optional[int] = None
    fila_alicuota: Optional[int] = None
    codigo: Optional[str] = None
    descripcion: Optional[str] = None


@dataclass
class BloqueCliente:
    hoja: str
    anio: int
    nombre: str
    cuit: Optional[int]
    password: Optional[str]
    fila_encabezado: int
    columnas_mes: Dict[int, str]       # 1..12 -> "B".."M"
    filas_concepto: Dict[str, int]     # concepto normalizado -> fila (bloque principal)
    fila_alicuota: Optional[int]
    fila_cuit: Optional[int] = None
    actividades: List[BloqueActividad] = field(default_factory=list)

    def filas_de_datos(self) -> List[int]:
        filas = set(self.filas_concepto.values())
        for act in self.actividades:
            filas.update(f for f in (act.fila_base, act.fila_anticipo, act.fila_alicuota) if f)
        if self.fila_alicuota:
            filas.add(self.fila_alicuota)
        return sorted(filas)


RE_CODIGO_TEXTO = re.compile(r"^\s*(\d{5,8})(?:\s+(\S.*))?$")


def _es_cuit(valor) -> bool:
    return isinstance(valor, (int, float)) and not isinstance(valor, bool) and valor >= UMBRAL_CUIT


def _codigo_actividad(valor_a, valor_b) -> Optional[Tuple[str, str]]:
    """Fila de codigo de actividad: A = codigo NAES (ej. 472130) y B = la
    descripcion. A veces esta todo junto como texto en A ('472160 Venta al
    por menor de ...')."""
    descripcion_b = valor_b.strip() if isinstance(valor_b, str) else ""
    if isinstance(valor_a, bool):
        return None
    if isinstance(valor_a, (int, float)) and 1000 <= valor_a < UMBRAL_CUIT and float(valor_a).is_integer():
        return str(int(valor_a)), descripcion_b
    if isinstance(valor_a, str):
        m = RE_CODIGO_TEXTO.match(valor_a)
        if m:
            return m.group(1), (m.group(2) or descripcion_b).strip()
    return None


def _password_de_celda(valor) -> Optional[str]:
    if isinstance(valor, str) and valor.strip():
        return valor.strip()
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return str(int(valor)) if float(valor).is_integer() else str(valor)
    return None


def _es_fila_alicuota(ws: Worksheet, fila: int, fila_fin: int) -> bool:
    """La fila de alicuota va justo debajo de 'Total pagado' (o del
    'anticipo determinado' de un sub-bloque) y no tiene etiqueta propia,
    aunque a veces tiene una nota escrita en A ('NO TIENE 2024')."""
    if fila > fila_fin:
        return False
    valor_a = ws.cell(row=fila, column=1).value
    if _es_cuit(valor_a) or _codigo_actividad(valor_a, ws.cell(row=fila, column=2).value):
        return False
    return not (isinstance(valor_a, str) and normalizar_etiqueta(valor_a) in CONCEPTOS)


def parsear_hoja_dia(ws: Worksheet, anio: int) -> List[BloqueCliente]:
    """Recorre una hoja DIA-N y arma un BloqueCliente por cada cliente.

    No asume posiciones fijas de fila: cada concepto se identifica por el
    texto de la columna A (case/acentos insensible), porque en el archivo
    real el orden y la presencia de filas varia de cliente a cliente. Un
    segundo 'Base Imponible' dentro del mismo cliente abre el sub-bloque
    de otra actividad (ej. Domingo Dominguez: tabaco arriba, bebidas abajo).
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
            mes = mes_desde_texto(ws.cell(row=fila_inicio, column=col).value)
            if mes and mes not in columnas_mes:
                columnas_mes[mes] = get_column_letter(col)

        filas_concepto: Dict[str, int] = {}
        sub_bloques: List[BloqueActividad] = []
        sub_actual: Optional[BloqueActividad] = None
        codigo_principal: Optional[Tuple[str, str]] = None
        cuit_valor = None
        fila_cuit = None
        password = None

        for fila in range(fila_inicio + 1, fila_fin + 1):
            valor_a = ws.cell(row=fila, column=1).value
            valor_b = ws.cell(row=fila, column=2).value
            etiqueta = normalizar_etiqueta(valor_a) if isinstance(valor_a, str) else ""

            if etiqueta in CONCEPTOS:
                if etiqueta == "base imponible" and etiqueta in filas_concepto:
                    sub_actual = BloqueActividad(fila_base=fila)
                    sub_bloques.append(sub_actual)
                elif (etiqueta == "anticipo determinado" and sub_actual is not None
                      and sub_actual.fila_anticipo is None):
                    sub_actual.fila_anticipo = fila
                elif etiqueta not in filas_concepto:
                    filas_concepto[etiqueta] = fila
                continue

            if _es_cuit(valor_a):
                if cuit_valor is None:
                    cuit_valor = int(valor_a)
                    fila_cuit = fila
                    password = _password_de_celda(valor_b)
                continue

            codigo = _codigo_actividad(valor_a, valor_b)
            if codigo:
                if sub_actual is not None:
                    if sub_actual.codigo is None:
                        sub_actual.codigo, sub_actual.descripcion = codigo
                elif codigo_principal is None:
                    codigo_principal = codigo

        fila_alicuota = None
        fila_total_pagado = filas_concepto.get("total pagado")
        if fila_total_pagado and _es_fila_alicuota(ws, fila_total_pagado + 1, fila_fin):
            fila_alicuota = fila_total_pagado + 1

        for sub in sub_bloques:
            candidata = (sub.fila_anticipo or sub.fila_base) + 1
            if _es_fila_alicuota(ws, candidata, fila_fin):
                sub.fila_alicuota = candidata

        actividades: List[BloqueActividad] = []
        if "base imponible" in filas_concepto:
            actividades.append(BloqueActividad(
                fila_base=filas_concepto["base imponible"],
                fila_anticipo=filas_concepto.get("anticipo determinado"),
                fila_alicuota=fila_alicuota,
                codigo=codigo_principal[0] if codigo_principal else None,
                descripcion=codigo_principal[1] if codigo_principal else None,
            ))
        actividades.extend(sub_bloques)

        bloques.append(BloqueCliente(
            hoja=ws.title, anio=anio, nombre=str(nombre).strip(), cuit=cuit_valor,
            password=password, fila_encabezado=fila_inicio, columnas_mes=columnas_mes,
            filas_concepto=filas_concepto, fila_alicuota=fila_alicuota,
            fila_cuit=fila_cuit, actividades=actividades,
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

# Mismos colores que ya usa "Lista de Clientes - IIBB" (leyenda E2:H2).
# El azul (DDJJ HECHA) no tiene relleno propio aca: es un estado que se
# asigna a mano y el script no lo toca.
RELLENO_COMPLETOS = PatternFill(fill_type="solid", fgColor="FF92D050")    # verde
RELLENO_NO_CERRADO = PatternFill(fill_type="solid", fgColor="FFFFE599")  # dorado
RELLENO_NO_HECHOS = PatternFill(fill_type="solid", fgColor="FFFF0000")   # rojo


def _escribir_fila_concepto(ws: Worksheet, fila: int, concepto: str) -> None:
    ws.cell(row=fila, column=1, value=ETIQUETAS_EXCEL[concepto])
    for col in range(2, 14):
        ws.cell(row=fila, column=col).number_format = FORMATO_MONEDA
    ws.cell(row=fila, column=14, value=f"=SUM(B{fila}:M{fila})").number_format = FORMATO_MONEDA


def _escribir_fila_codigo(ws: Worksheet, fila: int, actividad: Optional[Tuple[str, str]]) -> None:
    if not actividad:
        return
    codigo, descripcion = actividad
    ws.cell(row=fila, column=1, value=int(codigo) if codigo.isdigit() else codigo)
    if descripcion:
        ws.cell(row=fila, column=2, value=descripcion)


def _escribir_bloque_vacio(ws: Worksheet, fila_inicio: int, nombre: str, cuit: int,
                            password: Optional[str] = None,
                            actividades: Optional[List[Tuple[str, str]]] = None) -> int:
    """Escribe la estructura vacia de un cliente, con el mismo orden que usa
    la planilla original: encabezado, 12 conceptos, alicuota, codigo de
    actividad, CUIT/contrasena y, si tiene mas de una actividad, un
    sub-bloque por cada actividad extra (Base Imponible, anticipo
    determinado, alicuota y codigo). Devuelve la fila donde deberia
    empezar el siguiente cliente."""
    actividades = actividades or []
    ws.cell(row=fila_inicio, column=1, value=nombre).font = FUENTE_TITULO
    for i, mes in enumerate(MESES):
        ws.cell(row=fila_inicio, column=2 + i, value=mes).font = FUENTE_TITULO

    fila = fila_inicio + 1
    for concepto in CONCEPTOS:
        _escribir_fila_concepto(ws, fila, concepto)
        fila += 1

    fila += 1  # alicuota de la actividad principal
    _escribir_fila_codigo(ws, fila, actividades[0] if actividades else None)
    fila += 1

    ws.cell(row=fila, column=1, value=cuit)
    if password:
        ws.cell(row=fila, column=2, value=password)
    fila += 2

    for extra in actividades[1:]:
        _escribir_fila_concepto(ws, fila, "base imponible")
        _escribir_fila_concepto(ws, fila + 1, "anticipo determinado")
        _escribir_fila_codigo(ws, fila + 3, extra)  # fila + 2 = alicuota
        fila += 5
    return fila + 1


def _actividades_de_bloque(bloque: Optional[BloqueCliente]) -> List[Tuple[str, str]]:
    if not bloque:
        return []
    return [(a.codigo, a.descripcion or "") for a in bloque.actividades if a.codigo]


def _hoja_2025_vieja_y_vacia(ws: Worksheet) -> Tuple[bool, Dict[int, Optional[str]]]:
    """Detecta una hoja 'DIA N - 2025' generada por una version anterior de
    este script (sin fila de codigo de actividad entre la alicuota y el
    CUIT) que todavia no tiene NINGUN dato cargado en los meses. Esa hoja
    se puede regenerar sin perder nada. Devuelve tambien las contrasenas
    que ya tuviera escritas, para conservarlas."""
    bloques = parsear_hoja_dia(ws, 2025)
    if not bloques:
        return False, {}
    passwords = {b.cuit: b.password for b in bloques if b.cuit}
    for b in bloques:
        if not (b.fila_alicuota and b.fila_cuit == b.fila_alicuota + 1):
            return False, passwords
        for fila in b.filas_de_datos():
            for col in b.columnas_mes.values():
                if not _vacia(ws[f"{col}{fila}"].value):
                    return False, passwords
    return True, passwords


def crear_hojas_2025(wb, padron: Dict[int, List[Tuple[str, int, str]]],
                      indice_2024: Dict[Tuple[int, int], BloqueCliente]) -> None:
    """Crea 'DIA N - 2025' para cada dia que aparezca en el padron 2025, con
    un bloque vacio por cliente. Si el cliente ya existia en 2024, se copian
    sus actividades (asi los que tienen mas de una salen con su sub-bloque)
    y su contrasena propia si la tenia. Es idempotente: si la hoja ya existe
    y tiene datos, no la toca."""
    dias = sorted({dia for _, _, dia in padron[2025]},
                  key=lambda d: int(re.search(r"\d+", d).group()))

    for dia in dias:
        nombre_hoja = f"{dia}{SUFIJO_2025}"
        posicion = None
        passwords_previas: Dict[int, Optional[str]] = {}
        if nombre_hoja in wb.sheetnames:
            regenerar, passwords_previas = _hoja_2025_vieja_y_vacia(wb[nombre_hoja])
            if not regenerar:
                continue
            posicion = wb.sheetnames.index(nombre_hoja)
            wb.remove(wb[nombre_hoja])
            logger.info("Hoja %s tenia el formato viejo y ningun dato: la regenero", nombre_hoja)

        ws = wb.create_sheet(nombre_hoja, posicion)
        ws.column_dimensions["A"].width = 32
        for col in range(2, 15):
            ws.column_dimensions[get_column_letter(col)].width = 14

        fila = 1
        clientes = [(n, c) for n, c, d in padron[2025] if d == dia]
        for nombre, cuit in clientes:
            previo = indice_2024.get((2024, cuit))
            password = passwords_previas.get(cuit) or (previo.password if previo else None)
            fila = _escribir_bloque_vacio(ws, fila, nombre, cuit, password, _actividades_de_bloque(previo))
        logger.info("Hoja %s creada con %d clientes", nombre_hoja, len(clientes))


# --------------------------------------------------------------------------
# Deteccion de meses pendientes
# --------------------------------------------------------------------------

def meses_pendientes(ws: Worksheet, bloque: BloqueCliente) -> List[int]:
    """Un mes esta 'hecho' si ya tiene Total pagado (aunque sea 0) y alguna
    de sus actividades tiene Base Imponible. Con mas de una actividad, es
    normal que una quede en blanco en los meses en que no tuvo movimiento
    (Domingo Dominguez en enero: solo bebidas); eso no lo vuelve pendiente."""
    fila_tp = bloque.filas_concepto.get("total pagado")
    filas_base = [a.fila_base for a in bloque.actividades]
    if not fila_tp or not filas_base:
        return []
    pendientes = []
    for mes, col in sorted(bloque.columnas_mes.items()):
        sin_total = _vacia(ws[f"{col}{fila_tp}"].value)
        sin_base = all(_vacia(ws[f"{col}{f}"].value) for f in filas_base)
        if sin_total or sin_base:
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
# Datos extraidos de una DDJJ y su escritura en el Excel
# --------------------------------------------------------------------------

@dataclass
class Actividad:
    """Una fila de 'Informacion para el Calculo del impuesto' (Rubro 1)."""
    codigo: str
    descripcion: str
    base: Optional[float]
    alicuota: Optional[float]
    valor: Optional[float]   # columna "Valor" = anticipo determinado


@dataclass
class DatosDDJJ:
    tipo: str
    conceptos: Dict[str, float] = field(default_factory=dict)
    actividades: List[Actividad] = field(default_factory=list)
    avisos: List[str] = field(default_factory=list)


PALABRAS_VACIAS = {"venta", "ventas", "al", "por", "de", "del", "la", "las", "el", "los",
                   "en", "y", "e", "a", "con", "sin", "para", "otros", "otras", "ncp", "n", "c", "p"}


def _similitud_descripcion(a: str, b: str) -> float:
    def clave(texto: str) -> str:
        texto = normalizar(texto).replace("...", " ")
        return " ".join(p for p in re.findall(r"[a-z0-9]+", texto) if p not in PALABRAS_VACIAS)
    ka, kb = clave(a), clave(b)
    if not ka or not kb:
        return 0.0
    corta, larga = sorted((ka, kb), key=len)
    if len(corta) >= 10 and larga.startswith(corta):  # descripcion cortada con '...'
        return 0.95
    return difflib.SequenceMatcher(None, ka, kb).ratio()


def _puntaje(act: Actividad, destino: BloqueActividad) -> float:
    puntaje = 0.0
    if act.codigo and destino.codigo:
        if act.codigo == destino.codigo:
            puntaje += 100
        elif min(len(act.codigo), len(destino.codigo)) >= 5 and _levenshtein(act.codigo, destino.codigo) == 1:
            puntaje += 60  # error de tipeo en la planilla (ej. 4172200 en vez de 472200)
    if act.descripcion and destino.descripcion:
        similitud = _similitud_descripcion(act.descripcion, destino.descripcion)
        if similitud >= 0.75:
            puntaje += 40 * similitud
    return puntaje


def emparejar_actividades(agip: List[Actividad], excel: List[BloqueActividad]
                           ) -> Tuple[List[Tuple[Actividad, BloqueActividad]], List[Actividad]]:
    """Decide en que filas del Excel va cada actividad que muestra AGIP, por
    codigo (tolerando un digito de diferencia) y por descripcion. Devuelve
    (pares, actividades_sin_lugar)."""
    if not agip or not excel:
        return [], list(agip)
    if len(agip) == 1 and len(excel) == 1:
        return [(agip[0], excel[0])], []

    candidatos = sorted(((p, i, j) for i, a in enumerate(agip) for j, e in enumerate(excel)
                         if (p := _puntaje(a, e)) > 0), reverse=True)
    usados_agip: Set[int] = set()
    usados_excel: Set[int] = set()
    pares: List[Tuple[Actividad, BloqueActividad]] = []
    for _, i, j in candidatos:
        if i in usados_agip or j in usados_excel:
            continue
        pares.append((agip[i], excel[j]))
        usados_agip.add(i)
        usados_excel.add(j)

    # Un bloque del Excel sin codigo cargado (cliente nuevo en 2025) se queda
    # con la primera actividad que haya quedado sin lugar.
    libres = [j for j, e in enumerate(excel) if j not in usados_excel and not e.codigo]
    sueltas = [i for i in range(len(agip)) if i not in usados_agip]
    if len(libres) == 1 and sueltas:
        pares.append((agip[sueltas[0]], excel[libres[0]]))
        usados_agip.add(sueltas[0])

    return pares, [a for i, a in enumerate(agip) if i not in usados_agip]


def _fmt(valor) -> str:
    return "-" if valor is None else f"{valor:,.2f}"


def saldo_a_favor_anterior(wb, indice: Dict[Tuple[int, int], BloqueCliente],
                           bloque: BloqueCliente, mes: int) -> Optional[float]:
    """'Otros Creditos' de un mes = 'Saldo a favor' del mes anterior, si lo
    hay. Para enero se usa diciembre del anio anterior (2025 -> hoja 2024)."""
    if mes > 1:
        origen, columna = bloque, bloque.columnas_mes.get(mes - 1)
    else:
        origen = indice.get((bloque.anio - 1, bloque.cuit))
        columna = origen.columnas_mes.get(12) if origen else None
    fila = origen.filas_concepto.get("saldo a favor") if origen else None
    if not columna or not fila:
        return None
    valor = wb[origen.hoja][f"{columna}{fila}"].value
    if isinstance(valor, (int, float)) and not isinstance(valor, bool) and valor > 0.005:
        return float(valor)
    return None


def completar_otros_creditos(wb, indice: Dict[Tuple[int, int], BloqueCliente]) -> List[str]:
    """Completa cada 'Otros Creditos' vacio con el saldo a favor del mes
    anterior, en todos los clientes y meses (tambien los ya hechos). Nunca
    pisa una celda con algo escrito."""
    completados = []
    for bloque in indice.values():
        fila = bloque.filas_concepto.get("otros creditos")
        if not fila:
            continue
        ws = wb[bloque.hoja]
        for mes, columna in sorted(bloque.columnas_mes.items()):
            celda = ws[f"{columna}{fila}"]
            saldo = saldo_a_favor_anterior(wb, indice, bloque, mes) if _vacia(celda.value) else None
            if saldo is not None:
                celda.value = round(saldo, 2)
                completados.append(f"[{bloque.anio}] {bloque.nombre} {MESES[mes - 1]}: {_fmt(saldo)}")
    return completados


def escribir_valores(ws: Worksheet, bloque: BloqueCliente, mes: int, datos: DatosDDJJ,
                     saldo_anterior: Optional[float] = None) -> List[str]:
    """Completa las celdas VACIAS del mes con lo extraido de AGIP. Nunca
    pisa una celda que ya tenga algo: si AGIP dice otra cosa, lo devuelve
    como aviso para revisar a mano. 'Otros Creditos' es el saldo a favor del
    mes anterior (saldo_anterior) si lo hay; si no, lo que muestra AGIP en
    'Saldo a Favor DDJJ Periodo Anterior'. Devuelve la lista de avisos."""
    avisos: List[str] = []
    conceptos = dict(datos.conceptos)
    if saldo_anterior is not None:
        conceptos["otros creditos"] = saldo_anterior
    col = bloque.columnas_mes.get(mes)
    if not col:
        return [f"el bloque no tiene columna para {MESES[mes - 1]}"]
    col_idx = column_index_from_string(col)

    def poner(fila: Optional[int], valor: Optional[float], concepto: str, formato: Optional[str] = FORMATO_MONEDA):
        if fila is None or valor is None:
            return
        celda = ws.cell(row=fila, column=col_idx)
        if not _vacia(celda.value):
            if isinstance(celda.value, (int, float)) and abs(float(celda.value) - valor) > 0.01:
                avisos.append(f"{concepto}: la planilla ya tenia {celda.value} y AGIP dice {valor} (no se piso)")
            return
        if concepto in CONCEPTOS_CERO_EN_BLANCO and abs(valor) < 0.005:
            return
        celda.value = round(valor, 2)
        if formato and celda.number_format == "General":
            celda.number_format = formato

    # 1) Base imponible, anticipo y alicuota: cada actividad en SU fila.
    pares, sin_lugar = emparejar_actividades(datos.actividades, bloque.actividades)
    for act, destino in pares:
        poner(destino.fila_base, act.base, "base imponible")
        poner(destino.fila_anticipo, act.valor, "anticipo determinado")
        poner(destino.fila_alicuota, act.alicuota, "alicuota", formato=None)
    for act in sin_lugar:
        avisos.append(f"actividad {act.codigo} ({act.descripcion[:60]}) no tiene filas propias en el "
                      f"Excel: base {_fmt(act.base)}, alicuota {_fmt(act.alicuota)}, anticipo {_fmt(act.valor)}")
    if not datos.actividades and len(bloque.actividades) == 1:
        principal = bloque.actividades[0]
        poner(principal.fila_base, datos.conceptos.get("base imponible"), "base imponible")
        poner(principal.fila_anticipo, datos.conceptos.get("anticipo determinado"), "anticipo determinado")

    # 2) Conceptos del cliente (van solo en el bloque principal).
    for concepto in ("retenciones", "retenciones bancarias", "percepciones", "impuestos internos",
                     "pago a cuenta", "otros creditos", "saldo a favor",
                     "importe a pagar(subtotal)", "total pagado"):
        if concepto not in conceptos:
            continue
        valor = conceptos[concepto]
        fila = bloque.filas_concepto.get(concepto)
        if fila is None:
            if abs(valor) >= 0.005:
                avisos.append(f"{concepto} = {_fmt(valor)} pero el bloque no tiene fila '{concepto}'")
            continue
        poner(fila, valor, concepto)

    # 3) Intereses = Total pagado - Importe a pagar(subtotal) (hoja DETALLE).
    #    Como formula, igual que en la planilla original.
    fila_int = bloque.filas_concepto.get("intereses")
    fila_tp = bloque.filas_concepto.get("total pagado")
    fila_imp = bloque.filas_concepto.get("importe a pagar(subtotal)")
    if fila_int and fila_tp and fila_imp:
        celda = ws.cell(row=fila_int, column=col_idx)
        total = ws.cell(row=fila_tp, column=col_idx).value
        importe = ws.cell(row=fila_imp, column=col_idx).value
        if _vacia(celda.value) and isinstance(total, (int, float)) and isinstance(importe, (int, float)):
            celda.value = f"={col}{fila_tp}-{col}{fila_imp}"
            if celda.number_format == "General":
                celda.number_format = FORMATO_MONEDA
    return avisos


# --------------------------------------------------------------------------
# Lectura de las ventanas de e-Sicol (funciones puras sobre lo que devuelve
# JS_LEER_VENTANA, asi se pueden probar sin navegador)
# --------------------------------------------------------------------------

RE_FECHA = re.compile(r"^(\d{2})/(\d{2})/(\d{4})$")
RE_PERIODO = re.compile(r"^\d{4}-\d{2}$")
RE_MONTO_TOTAL = re.compile(r"monto\s+total")
RE_TITULO_DDJJ = re.compile(r"ano:\s*(\d{4}).*?mes:\s*([a-z]+).*?tipo:\s*(.+?)\s*$")
RE_MONTO_EN_TEXTO = re.compile(r"-?\s*\$\s*-?\s*\d[\d.]*(?:,\d+)?|-?\d[\d.]*,\d{2}\b")

ETIQUETAS_LIQUIDACION = {
    # Textos tal cual los nombra la hoja DETALLE.
    "saldo a favor": re.compile(r"subtotal\s+a\s+favor\s+del\s+contribuyente"),
    "importe a pagar(subtotal)": re.compile(r"importe\s+neto\s+a\s+ingresar"),
    "total pagado": re.compile(r"total\s+(de\s+)?importes?\s+actualizad"),
}


def numero_version(tipo: str) -> int:
    """'Original' -> 0, 'Rectificativa' -> 1, 'Rectificativa 2' -> 2."""
    t = normalizar(tipo)
    m = re.search(r"rectificativa\D*(\d+)", t)
    if m:
        return int(m.group(1))
    return 1 if "rectificativa" in t else 0


def parsear_titulo_ventana(titulo: str) -> Optional[Tuple[int, Optional[int], str]]:
    """'Percepciones de agentes. Año: 2024 - Mes: Enero - Tipo: Original'
    -> (2024, 1, 'original')."""
    m = RE_TITULO_DDJJ.search(normalizar(titulo))
    if not m:
        return None
    return int(m.group(1)), mes_desde_texto(m.group(2)), m.group(3)


def _es_fila_total(fila: dict) -> bool:
    return bool(fila.get("resumen")) or any(normalizar(c) == "total" for c in fila["celdas"])


def _montos(celdas: List[str]) -> List[float]:
    con_signo = [v for v in (parsear_numero_ar(c) for c in celdas if "$" in c) if v is not None]
    if con_signo:
        return con_signo
    return [v for v in (parsear_numero_ar(c) for c in celdas
                        if normalizar(c) != "total" and "%" not in c and not RE_FECHA.match(c))
            if v is not None]


def total_grilla(filas: List[dict]) -> Optional[float]:
    """Ultimo monto de la fila 'Total' (fila resumen de ExtJS)."""
    for fila in filas:
        if _es_fila_total(fila):
            montos = _montos(fila["celdas"])
            if montos:
                return montos[-1]
    return None


def suma_grilla(filas: List[dict]) -> Optional[float]:
    """Suma del ultimo monto de cada fila de detalle (sin la fila Total)."""
    valores = []
    for fila in filas:
        if not _es_fila_total(fila):
            montos = _montos(fila["celdas"])
            if montos:
                valores.append(montos[-1])
    return round(sum(valores), 2) if valores else None


def valor_de_campo(datos: dict, patron: re.Pattern) -> Optional[float]:
    """Valor de un campo de formulario por su etiqueta (ej. 'Monto total:'
    -> '$10.455,22'). En e-Sicol esos valores estan dentro de un <input>,
    no como texto suelto."""
    for campo in datos.get("campos", []):
        if patron.search(normalizar(campo.get("etiqueta"))):
            valor = parsear_numero_ar(campo.get("valor"))
            if valor is not None:
                return valor
    return None


def valor_en_texto(datos: dict, patron: re.Pattern) -> Optional[float]:
    """Ultimo recurso: 'etiqueta ... $X' en el texto suelto de la ventana."""
    texto = normalizar(datos.get("texto", ""))
    m = patron.search(texto)
    if m:
        monto = RE_MONTO_EN_TEXTO.search(texto[m.end():m.end() + 80])
        if monto:
            return parsear_numero_ar(monto.group(0))
    return None


def valor_en_fila(datos: dict, patron: re.Pattern) -> Optional[float]:
    """Ultimo monto de la fila de grilla que tiene la etiqueta buscada."""
    for fila in datos.get("filas", []):
        celdas = fila["celdas"]
        for k, celda in enumerate(celdas):
            if patron.search(normalizar(celda)):
                montos = _montos(celdas[k + 1:])
                if montos:
                    return montos[-1]
    return None


def valor_monto_total(datos: dict) -> Optional[float]:
    """Percepciones / Retenciones / Pagos a cuenta: 'Monto total', y si no
    esta, la fila Total de la grilla o la suma de su detalle."""
    for lector in (lambda: valor_de_campo(datos, RE_MONTO_TOTAL),
                   lambda: total_grilla(datos.get("filas", [])),
                   lambda: suma_grilla(datos.get("filas", [])),
                   lambda: valor_en_texto(datos, RE_MONTO_TOTAL)):
        valor = lector()
        if valor is not None:
            return valor
    return None


def valor_saldo_periodo_anterior(datos: dict) -> Optional[float]:
    """'Saldo a Favor DDJJ Periodo Anterior': grilla Anio/Cuota/Importe cuya
    fila de resumen (sin la palabra 'Total') trae el importe."""
    for lector in (lambda: total_grilla(datos.get("filas", [])),
                   lambda: suma_grilla(datos.get("filas", [])),
                   lambda: valor_de_campo(datos, RE_MONTO_TOTAL),
                   lambda: valor_en_texto(datos, RE_MONTO_TOTAL)):
        valor = lector()
        if valor is not None:
            return valor
    return None


def valor_impuestos_internos(datos: dict) -> Optional[float]:
    return valor_en_fila(datos, re.compile(r"^impuestos? internos?"))


def buscar_etiquetas(datos: dict, patrones: Dict[str, re.Pattern]) -> Dict[str, float]:
    encontrados: Dict[str, float] = {}
    for clave, patron in patrones.items():
        for lector in (valor_de_campo, valor_en_fila, valor_en_texto):
            valor = lector(datos, patron)
            if valor is not None:
                encontrados[clave] = valor
                break
    return encontrados


def _indice_encabezado(encabezados: List[str], patron: str) -> Optional[int]:
    for i, texto in enumerate(encabezados):
        if re.search(patron, normalizar(texto)):
            return i
    return None


def parsear_rubro1(datos: dict) -> Tuple[List[Actividad], Optional[float], Optional[float]]:
    """'Informacion para el Calculo del impuesto': una fila por actividad
    (Cod. / Descripcion / Base imponible / Alicuota / Valor) mas la fila
    Total. Devuelve (actividades, total_base, total_valor)."""
    encabezados = datos.get("encabezados", [])
    posiciones = {
        "codigo": _indice_encabezado(encabezados, r"^cod"),
        "descripcion": _indice_encabezado(encabezados, r"^descrip"),
        "base": _indice_encabezado(encabezados, r"base"),
        "alicuota": _indice_encabezado(encabezados, r"^alicuota"),
        "valor": _indice_encabezado(encabezados, r"^valor"),
    }
    actividades: List[Actividad] = []
    total_base = total_valor = None

    for fila in datos.get("filas", []):
        celdas = fila["celdas"]
        por_encabezado = len(encabezados) == len(celdas) and None not in posiciones.values()
        if _es_fila_total(fila):
            if por_encabezado:
                total_base = parsear_numero_ar(celdas[posiciones["base"]])
                total_valor = parsear_numero_ar(celdas[posiciones["valor"]])
            else:
                montos = _montos(celdas)
                if len(montos) >= 2:
                    total_base, total_valor = montos[0], montos[-1]
            continue

        if por_encabezado:
            codigo, descripcion, base, alicuota, valor = (celdas[posiciones[k]] for k in
                                                          ("codigo", "descripcion", "base", "alicuota", "valor"))
        elif len(celdas) == 5:
            codigo, descripcion, base, alicuota, valor = celdas
        else:
            codigo = next((c for c in celdas if re.fullmatch(r"\d{4,8}", c)), "")
            alicuota = next((c for c in celdas if "%" in c), "")
            dinero = [c for c in celdas if "$" in c]
            base, valor = (dinero[0], dinero[-1]) if len(dinero) >= 2 else ("", "")
            textos = [c for c in celdas if c and not re.search(r"\d", c)]
            descripcion = max(textos, key=len) if textos else ""
        codigo = codigo.strip()
        if not re.fullmatch(r"\d{4,8}", codigo):
            continue
        actividades.append(Actividad(
            codigo=codigo, descripcion=descripcion.strip(), base=parsear_numero_ar(base),
            alicuota=parsear_numero_ar(alicuota), valor=parsear_numero_ar(valor)))
    return actividades, total_base, total_valor


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


_capturas = {"contexto": "", "cantidad": {}}  # contexto: "CUIT_periodo" que se esta procesando


def guardar_captura(page, motivo: str, es_error: bool = False) -> None:
    """Guarda en CARPETA_CAPTURAS la captura de pantalla (.png) y el HTML de
    lo que muestra el navegador en este momento. Sirve para ver que paso
    aunque se corra sin --sin-headless. Nunca corta la corrida si falla."""
    clave = re.sub(r"[^a-z0-9]+", "-", normalizar(motivo)).strip("-")[:60] or "error"
    tope = MAX_CAPTURAS_POR_ERROR if es_error else MAX_CAPTURAS_POR_AVISO
    cantidad = _capturas["cantidad"].get(clave, 0)
    if cantidad >= tope:
        return
    _capturas["cantidad"][clave] = cantidad + 1
    nombre = "_".join(p for p in (f"{datetime.now():%Y%m%d_%H%M%S}", _capturas["contexto"], clave) if p)
    try:
        CARPETA_CAPTURAS.mkdir(parents=True, exist_ok=True)
        (CARPETA_CAPTURAS / f"{nombre}.html").write_text(page.content(), encoding="utf-8")
        page.screenshot(path=str(CARPETA_CAPTURAS / f"{nombre}.png"), timeout=10000)
        logger.info("  Captura de pantalla guardada: %s", CARPETA_CAPTURAS / f"{nombre}.png")
    except Exception as exc:
        logger.info("  No pude guardar la captura de pantalla (%s)",
                    str(exc).splitlines()[0][:120] if str(exc) else type(exc).__name__)


def _esperar_red(page, tiempo: int) -> None:
    try:
        page.wait_for_load_state("networkidle", timeout=tiempo)
    except Exception:
        pass


SELECTOR_CARGANDO = "#cargando:visible, .x-mask-msg:visible"


def esperar_sin_carga(page, tiempo: int) -> None:
    """Espera a que desaparezca el 'Cargando...' de e-Sicol (#cargando es
    un id real de la pagina, oculto cuando no esta cargando nada)."""
    limite = time.monotonic() + tiempo / 1000
    while time.monotonic() < limite:
        try:
            if page.locator(SELECTOR_CARGANDO).count() == 0:
                return
        except Exception:
            return
        time.sleep(0.2)


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


# El enlace 'Clave Ciudad' de la home de AGIP solo abre
# claveciudad.agip.gob.ar en una pestana nueva, asi que se entra directo
# ahi. En esa pagina hay OTRO enlace 'Clave Ciudad' (onclick="toggleLogin()")
# que despliega el formulario de usuario/contrasena. Selectores tomados del
# HTML real de la pagina.
URL_CLAVE_CIUDAD = "https://claveciudad.agip.gob.ar/"
SELECTOR_LINK_CLAVE_CIUDAD = 'a[href="https://claveciudad.agip.gob.ar/"]'
SELECTOR_LINK_TOGGLE_LOGIN = 'a[onclick*="toggleLogin"]'
TIMEOUT_CARGA_PAGINA_MS = 30000


class LoginNoDisponible(Exception):
    """No aparecio el formulario de Clave Ciudad: AGIP caida o muy lenta."""


def _mostrar_formulario_login(pagina, tiempo_espera: int) -> bool:
    campo = pagina.locator('input[type="password"]').first
    try:
        if not campo.is_visible():
            pagina.locator(SELECTOR_LINK_TOGGLE_LOGIN).first.click(timeout=tiempo_espera)
            pausar()
        campo.wait_for(state="visible", timeout=tiempo_espera)
        return True
    except Exception:
        return False


def _abrir_login(page, tiempo_espera: int):
    """Deja a la vista el formulario de Clave Ciudad y devuelve la pestana
    donde quedo. Si entrando directo no aparece, prueba el camino de antes:
    home de AGIP -> enlace 'Clave Ciudad' (que abre otra pestana)."""
    carga = max(tiempo_espera, TIMEOUT_CARGA_PAGINA_MS)
    try:
        page.goto(URL_CLAVE_CIUDAD, wait_until="domcontentloaded", timeout=carga)
        pausar()
        if _mostrar_formulario_login(page, tiempo_espera):
            return page
    except Exception:
        pass
    try:
        page.goto(BASE_URL, wait_until="domcontentloaded", timeout=carga)
        pausar()
        with page.expect_popup(timeout=tiempo_espera) as popup:
            page.locator(SELECTOR_LINK_CLAVE_CIUDAD).first.click(timeout=tiempo_espera)
        pestana = popup.value
        pestana.wait_for_load_state("domcontentloaded", timeout=carga)
        pausar()
        if _mostrar_formulario_login(pestana, tiempo_espera):
            return pestana
    except Exception:
        pass
    raise LoginNoDisponible("no cargo el formulario de Clave Ciudad (AGIP puede estar caida o muy lenta)")


def _intentar_login(page, cuit: int, password: str, tiempo_espera: int):
    """Un unico intento de login con una contrasena puntual. Devuelve la
    pestana donde quedo la sesion activa, o None si la contrasena no
    funciono. Si ni siquiera aparece el formulario lanza LoginNoDisponible
    (ahi no tiene sentido probar otra contrasena)."""
    nueva_pagina = _abrir_login(page, tiempo_espera)
    campo_password = nueva_pagina.locator('input[type="password"]').first

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
    funciono. Si AGIP no carga el login, deja pasar LoginNoDisponible."""
    for i, password in enumerate(candidatas):
        pagina_activa = _intentar_login(page, cuit, password, tiempo_espera)
        if pagina_activa:
            return password, pagina_activa
        logger.info("CUIT %s: contrasena candidata %d/%d no funciono", cuit, i + 1, len(candidatas))
    return None, None


# --------------------------------------------------------------------------
# e-Sicol: lista "Declaraciones Juradas Presentadas" (grilla ExtJS)
# --------------------------------------------------------------------------

SELECTOR_FILAS_GRILLA = "tr.x-grid-row"
RE_GRUPO_PRESENTADAS = r"declaraciones?\s+juradas\s+presentadas"

# Lee TODAS las filas de una vez (un solo viaje al navegador en vez de uno
# por celda). ExtJS deja en el DOM las filas del grupo aunque este plegado,
# asi que se pueden leer antes de desplegarlo.
JS_CELDAS_FILA = r"""
(tr) => {
  const limpiar = (s) => (s || '').replace(/ /g, ' ').replace(/\s+/g, ' ').trim();
  const celdas = Array.from(tr.children)
    .filter((td) => td.classList && td.classList.contains('x-grid-cell'))
    .map((td) => limpiar(td.innerText || td.textContent));
  const visible = !!(tr.offsetWidth || tr.offsetHeight || tr.getClientRects().length);
  return {celdas, visible};
}
"""
JS_LEER_FILAS_GRILLA = f"(filas) => filas.map({JS_CELDAS_FILA.strip()})"


@dataclass
class FilaDDJJ:
    indice: int      # posicion dentro de page.locator("tr.x-grid-row")
    estado: str      # "Presentada"
    periodo: str     # "2024-02"
    fecha: str       # "19/03/2024"
    tipo: str        # "Original" / "Rectificativa 2"
    visible: bool

    @property
    def version(self) -> int:
        return numero_version(self.tipo)


def _orden_fecha(fecha: str) -> Tuple[int, int, int]:
    m = RE_FECHA.match(fecha or "")
    return (int(m.group(3)), int(m.group(2)), int(m.group(1))) if m else (0, 0, 0)


def _es_presentada(fila: "FilaDDJJ") -> bool:
    return normalizar(fila.estado).startswith("presentad")


def filas_ddjj_desde_celdas(crudas: List[dict]) -> List[FilaDDJJ]:
    filas = []
    for indice, cruda in enumerate(crudas):
        celdas = cruda["celdas"]
        periodo = next((c for c in celdas if RE_PERIODO.match(c)), None)
        if not periodo:
            continue  # filas del arbol u otras grillas
        estado = next((c for c in celdas if "presentad" in normalizar(c)),
                      celdas[1] if len(celdas) > 1 else "")
        fecha = next((c for c in celdas if RE_FECHA.match(c)), "")
        filas.append(FilaDDJJ(indice, estado, periodo, fecha, celdas[-1], cruda.get("visible", True)))
    return filas


def elegir_ddjj(filas: List[FilaDDJJ], periodo: str) -> Optional[FilaDDJJ]:
    """De todas las filas de ese periodo, la ULTIMA version presentada:
    'Rectificativa 2' le gana a 'Rectificativa 1', que le gana a
    'Original'. Las versiones anteriores no se usan nunca."""
    candidatas = [f for f in filas if f.periodo == periodo]
    if not candidatas:
        return None
    presentadas = [f for f in candidatas if _es_presentada(f)] or candidatas
    return max(presentadas, key=lambda f: (f.version, _orden_fecha(f.fecha), -f.indice))


def leer_lista_ddjj(page) -> Tuple[List[FilaDDJJ], Optional[int]]:
    """Filas de la grilla de DDJJ y el total que anuncia su encabezado
    ('Declaraciones Juradas Presentadas: 110')."""
    try:
        crudas = page.locator(SELECTOR_FILAS_GRILLA).evaluate_all(JS_LEER_FILAS_GRILLA)
    except Exception:
        crudas = []
    total = None
    try:
        for titulo in page.locator(".x-grid-group-title").all_inner_texts():
            if "presentada" in normalizar(titulo):
                m = re.search(r"(\d+)\s*$", titulo.strip())
                if m:
                    total = int(m.group(1))
    except Exception:
        pass
    return filas_ddjj_desde_celdas(crudas), total


@dataclass
class ResultadoLista:
    elegida: Optional[FilaDDJJ]
    periodos: Set[str]
    completa: bool   # se leyeron tantas filas como anuncia el encabezado


class ErrorNavegacion(Exception):
    """No se pudo llegar a la lista de DDJJ o abrir la DDJJ elegida."""


def ir_a_declaracion(page, anio: int, mes_idx: int, tiempo_espera: int) -> ResultadoLista:
    """Navega e-Sicol -> Declaraciones Juradas Presentadas -> abre (doble
    click) la ultima version presentada del periodo pedido.

    Antes que nada, recarga la pagina (equivalente a F5): cierra la DDJJ
    del mes anterior y todas sus ventanas, y deja la aplicacion en su
    estado inicial. La sesion (cookies) se mantiene al recargar."""
    try:
        page.reload(wait_until="networkidle", timeout=tiempo_espera)
    except Exception:
        pass
    pausar()

    if not _click_si_existe(page, r"e-?sicol", tiempo_espera):
        raise ErrorNavegacion("no encontre el enlace 'e-Sicol' (¿se cerro la sesion?)")
    _esperar_red(page, tiempo_espera)
    esperar_sin_carga(page, tiempo_espera)

    periodo = f"{anio}-{mes_idx:02d}"
    filas, total = leer_lista_ddjj(page)
    for _ in range(2):
        if any(f.visible for f in filas):
            break
        # Grupo plegado (o todavia sin filas): un click en el encabezado lo despliega.
        if not _click_si_existe(page, RE_GRUPO_PRESENTADAS, tiempo_espera):
            break
        _esperar_red(page, tiempo_espera)
        esperar_sin_carga(page, tiempo_espera)
        filas, total = leer_lista_ddjj(page)
        if not filas:
            break

    if not filas:
        # ids reales de la pantalla de e-Sicol: si estan, la aplicacion cargo
        # bien y el cliente simplemente no tiene ninguna DDJJ presentada.
        if page.locator("#ddjjTree, #app-header-logo-esicol-titulo").count() == 0:
            raise ErrorNavegacion("no cargo la lista de DDJJ presentadas")
        return ResultadoLista(None, set(), True)

    periodos = {f.periodo for f in filas}
    completa = total is not None and sum(_es_presentada(f) for f in filas) >= total
    elegida = elegir_ddjj(filas, periodo)

    if elegida is None and not completa:
        # Por si la grilla no tuviera todas las filas en el DOM de entrada.
        try:
            page.locator(".x-grid-group-title").first.hover(timeout=tiempo_espera)
        except Exception:
            pass
        for _ in range(10):
            try:
                page.mouse.wheel(0, 400)
            except Exception:
                break
            pausar(0.3, 0.15)
            filas, total = leer_lista_ddjj(page)
            periodos |= {f.periodo for f in filas}
            elegida = elegir_ddjj(filas, periodo)
            if elegida:
                break
        completa = total is not None and sum(_es_presentada(f) for f in filas) >= total

    if elegida is None:
        return ResultadoLista(None, periodos, completa)

    if not elegida.visible:
        _click_si_existe(page, RE_GRUPO_PRESENTADAS, tiempo_espera)
        _esperar_red(page, tiempo_espera)
        filas, _ = leer_lista_ddjj(page)
        elegida = elegir_ddjj(filas, periodo)
        if elegida is None or not elegida.visible:
            raise ErrorNavegacion(f"la fila de {periodo} esta en la lista pero no se deja ver")

    fila = page.locator(SELECTOR_FILAS_GRILLA).nth(elegida.indice)
    try:
        actual = fila.evaluate(JS_CELDAS_FILA)["celdas"]
    except Exception:
        actual = []
    if elegida.periodo not in actual or actual[-1] != elegida.tipo:
        raise ErrorNavegacion(f"la grilla cambio mientras buscaba la fila de {periodo}")

    try:
        fila.dblclick(timeout=tiempo_espera)
    except Exception as exc:
        raise ErrorNavegacion(f"encontre la fila de {periodo} pero no pude hacerle doble click") from exc
    pausar()
    _esperar_red(page, tiempo_espera)
    esperar_sin_carga(page, tiempo_espera)
    esperar_ddjj_abierta(page, tiempo_espera)
    return ResultadoLista(elegida, periodos, completa)


# --------------------------------------------------------------------------
# e-Sicol: arbol de la DDJJ abierta (#ddjjTree, id real)
#
# Estructura confirmada con el HTML real: todas las filas del arbol son
# <tr class="x-grid-row"> hermanas; la profundidad de cada nodo sale de la
# cantidad de <img class="x-tree-elbow..."> que tiene antes del texto. Las
# carpetas tienen <img class="x-tree-expander">, y las desplegadas llevan
# 'x-grid-tree-node-expanded' en el <tr>. La DDJJ abierta (la del doble
# click) marca TODAS sus filas con la clase 'filasDdJjSeleccionada' (o
# 'ddjjSeleccionada' la que tiene el foco). OJO: el arbol puede tener otras
# DDJJ desplegadas al mismo tiempo, y dentro de una misma DDJJ hay dos
# nodos 'Agentes' (Percepciones y Retenciones): por eso cada seccion se
# busca por su ruta completa, y solo dentro de la DDJJ seleccionada.
# --------------------------------------------------------------------------

SELECTOR_FILAS_ARBOL = "#ddjjTree tr.x-grid-row"

JS_LEER_ARBOL = r"""
(filas) => filas.map((tr) => {
  const tds = Array.from(tr.querySelectorAll('td'));
  const clases = tds.map((td) => td.className).join(' ');
  const nivel = Array.from(tr.querySelectorAll('img'))
    .filter((img) => /x-tree-elbow/.test(img.className)).length;
  const nodo = tr.querySelector('.x-tree-node-text') ||
    Array.from(tr.querySelectorAll('.x-grid-cell-inner span')).pop() || tds[0] || tr;
  const texto = (nodo.innerText || nodo.textContent || '').replace(/ /g, ' ').replace(/\s+/g, ' ').trim();
  return {
    texto,
    nivel,
    x: nodo.getBoundingClientRect().left,
    expandida: /x-grid-tree-node-expanded/.test(tr.className),
    expandible: !!tr.querySelector('img.x-tree-expander'),
    seleccionada: /filasDdJjSeleccionada|ddjjSeleccionada/.test(clases),
  };
})
"""

RE_NODO_VERSION = re.compile(r"^(original|rectificativa)")

RUBRO_1 = re.compile(r"^rubro 1\b")
RUBRO_2 = re.compile(r"^rubro 2\b")
RUTAS = {
    "info_calculo": [RUBRO_1, re.compile(r"^informacion para el calculo")],
    "conceptos_no_integran": [RUBRO_1, re.compile(r"^conceptos que no integran")],
    "saldo_favor_anterior": [RUBRO_2, re.compile(r"^saldo a favor ddjj periodo anterior")],
    "pagos_a_cuenta": [RUBRO_2, re.compile(r"^pagos? a cuenta")],
    "percepciones": [RUBRO_2, re.compile(r"^percepciones$"), re.compile(r"^agentes$")],
    "retenciones": [RUBRO_2, re.compile(r"^retenciones$"), re.compile(r"^agentes$")],
    "retenciones_bancarias": [RUBRO_2, re.compile(r"^retenciones$"), re.compile(r"^bancari")],
    "liquidacion": [re.compile(r"^liquidacion del impuesto")],
}

# Nunca se clickea un nodo con alguna de estas palabras (solo se leen
# secciones, no se ejecuta ninguna accion sobre la DDJJ).
PALABRAS_PELIGROSAS = ("eliminar", "rectificar", "presentar", "pagar", "vep", "generar",
                       "imprimir", "descargar", "enviar", "anular", "borrar")


def leer_arbol(page) -> List[dict]:
    try:
        filas = page.locator(SELECTOR_FILAS_ARBOL).evaluate_all(JS_LEER_ARBOL)
    except Exception:
        return []
    if len(filas) > 1 and len({f["nivel"] for f in filas}) <= 1:
        # Sin imagenes de sangria: la profundidad sale de la posicion en pantalla.
        xs = sorted({round(f["x"]) for f in filas})
        for f in filas:
            f["nivel"] = xs.index(round(f["x"]))
    return filas


def _coincide(texto: str, paso) -> bool:
    if isinstance(paso, re.Pattern):
        return bool(paso.search(normalizar(texto)))
    return normalizar(texto) == normalizar(paso)


def _descendientes(filas: List[dict], i: int) -> range:
    nivel = filas[i]["nivel"]
    j = i + 1
    while j < len(filas) and filas[j]["nivel"] > nivel:
        j += 1
    return range(i + 1, j)


def _hijos_directos(filas: List[dict], i: int) -> List[int]:
    descendientes = _descendientes(filas, i)
    if not descendientes:
        return []
    nivel_hijo = min(filas[k]["nivel"] for k in descendientes)
    return [k for k in descendientes if filas[k]["nivel"] == nivel_hijo]


def _ancestros(filas: List[dict], i: int) -> List[int]:
    resultado = []
    nivel = filas[i]["nivel"]
    for j in range(i - 1, -1, -1):
        if filas[j]["nivel"] < nivel:
            resultado.append(j)
            nivel = filas[j]["nivel"]
    return resultado


def raiz_ddjj(filas: List[dict]) -> Optional[int]:
    """Indice del nodo 'Original'/'Rectificativa N' de la DDJJ abierta."""
    primera = next((i for i, f in enumerate(filas) if f["seleccionada"]), None)
    if primera is None:
        return None
    if RE_NODO_VERSION.search(normalizar(filas[primera]["texto"])):
        return primera
    for j in _ancestros(filas, primera):
        if RE_NODO_VERSION.search(normalizar(filas[j]["texto"])):
            return j
    return primera


def ddjj_abierta(filas: List[dict]) -> Optional[Tuple[Optional[int], Optional[int], str]]:
    """(anio, mes, tipo) de la DDJJ abierta, leidos del propio arbol."""
    raiz = raiz_ddjj(filas)
    if raiz is None:
        return None
    anio = mes = None
    for j in _ancestros(filas, raiz):
        texto = filas[j]["texto"].strip()
        if mes is None:
            mes = mes_desde_texto(texto)
        if anio is None and re.fullmatch(r"\d{4}", texto):
            anio = int(texto)
    return anio, mes, filas[raiz]["texto"]


def esperar_ddjj_abierta(page, tiempo: int) -> List[dict]:
    limite = time.monotonic() + tiempo / 1000
    filas: List[dict] = []
    while time.monotonic() < limite:
        filas = leer_arbol(page)
        if raiz_ddjj(filas) is not None:
            return filas
        time.sleep(0.3)
    return filas


def _rango_ddjj(filas: List[dict]) -> range:
    raiz = raiz_ddjj(filas)
    return range(len(filas)) if raiz is None else _descendientes(filas, raiz)


def resolver_ruta(filas: List[dict], ruta: list) -> Tuple[Optional[int], Optional[int]]:
    """Busca la ruta dentro de la DDJJ seleccionada. Devuelve
    (indice_del_nodo, None) si lo encontro, (None, indice_a_desplegar) si
    hay que desplegar una carpeta primero, o (None, None) si no esta."""
    rango = _rango_ddjj(filas)
    for n, paso in enumerate(ruta):
        idx = next((i for i in rango if _coincide(filas[i]["texto"], paso)), None)
        if idx is None:
            return None, None
        if n == len(ruta) - 1:
            return idx, None
        if filas[idx]["expandible"] and not filas[idx]["expandida"]:
            return None, idx
        rango = _descendientes(filas, idx)
    return None, None


def _texto_nodo(fila):
    return fila.locator(".x-tree-node-text, .x-grid-cell-inner span").last


def _expandir_nodo(page, idx: int, tiempo: int) -> None:
    """Despliega la carpeta y espera a que aparezcan sus hijos (pueden
    venir del servidor): si se volviera a clickear el +/- antes de tiempo,
    la carpeta se plegaria de nuevo."""
    fila = page.locator(SELECTOR_FILAS_ARBOL).nth(idx)
    try:
        fila.locator("img.x-tree-expander").first.click(timeout=tiempo)
    except Exception:
        try:
            _texto_nodo(fila).dblclick(timeout=tiempo)
        except Exception:
            pass
    pausar(0.5, 0.2)
    esperar_sin_carga(page, tiempo)
    limite = time.monotonic() + TIMEOUT_NODO_MS / 1000
    while time.monotonic() < limite:
        filas = leer_arbol(page)
        if idx < len(filas) and filas[idx]["expandida"] and _descendientes(filas, idx):
            return
        time.sleep(0.25)


def ubicar_nodo(page, ruta: list, tiempo: int) -> Optional[int]:
    """Indice (en SELECTOR_FILAS_ARBOL) del nodo al final de la ruta,
    desplegando las carpetas intermedias que haga falta."""
    limite = time.monotonic() + min(tiempo, TIMEOUT_NODO_MS) / 1000
    despliegues = 0
    while True:
        idx, desplegar = resolver_ruta(leer_arbol(page), ruta)
        if idx is not None:
            return idx
        if desplegar is not None and despliegues < 2 * len(ruta):
            _expandir_nodo(page, desplegar, tiempo)
            despliegues += 1
            continue
        if time.monotonic() >= limite:
            return None
        time.sleep(0.3)


def _ruta_legible(ruta: list) -> str:
    return " > ".join(p.pattern.strip("^$").replace("\\b", "").replace("\\", "") if isinstance(p, re.Pattern)
                      else str(p) for p in ruta)


# --------------------------------------------------------------------------
# e-Sicol: ventanas de detalle. Cada nodo del arbol abre una ventana
# flotante (Ext.window.Window) titulada, por ejemplo, "Percepciones de
# agentes. Año: 2024 - Mes: Enero - Tipo: Original".
# --------------------------------------------------------------------------

SELECTOR_VENTANA = "div.x-window:visible"
SELECTOR_PANEL_DETALLE = "#panelContenedor"


def _texto_panel(page) -> str:
    try:
        panel = page.locator(SELECTOR_PANEL_DETALLE)
        return panel.inner_text(timeout=1000).strip() if panel.count() else ""
    except Exception:
        return ""

JS_VENTANA_CARGADA = r"""
(w) => {
  const visible = (el) => !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length);
  if (Array.from(w.querySelectorAll('.x-mask-msg, .x-mask-loading')).some(visible)) return false;
  return !!(w.querySelector('td.x-grid-cell') ||
            Array.from(w.querySelectorAll('input')).some((i) => i.value) ||
            /\$/.test(w.innerText || ''));
}
"""

JS_LEER_VENTANA = r"""
(w) => {
  const limpiar = (s) => (s || '').replace(/ /g, ' ').replace(/\s+/g, ' ').trim();
  const visible = (el) => !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length);
  const tituloEl = w.querySelector('.x-window-header-text, .x-header-text, .x-panel-header-text');
  const filas = [];
  w.querySelectorAll('tr').forEach((tr) => {
    const tds = Array.from(tr.children).filter((td) => td.classList && td.classList.contains('x-grid-cell'));
    if (!tds.length || !visible(tr)) return;
    filas.push({
      resumen: /summary/.test(tr.className),
      celdas: tds.map((td) => limpiar(td.innerText || td.textContent)),
    });
  });
  const campos = [];
  w.querySelectorAll('label').forEach((lb) => {
    if (!visible(lb)) return;
    let control = lb.htmlFor ? document.getElementById(lb.htmlFor) : null;
    if (!control) {
      const contenedor = lb.closest('.x-field, .x-form-item, tr, table');
      if (contenedor) control = contenedor.querySelector('input, textarea, .x-form-display-field');
    }
    let valor = '';
    if (control) {
      valor = (control.tagName === 'INPUT' || control.tagName === 'TEXTAREA')
        ? control.value : (control.innerText || control.textContent);
    }
    campos.push({etiqueta: limpiar(lb.innerText || lb.textContent), valor: limpiar(valor)});
  });
  const encabezados = Array.from(w.querySelectorAll('.x-column-header')).filter(visible)
    .map((h) => limpiar(h.innerText || h.textContent));
  return {
    titulo: limpiar(tituloEl ? (tituloEl.innerText || tituloEl.textContent) : ''),
    filas, campos, encabezados,
    texto: limpiar(w.innerText || w.textContent).slice(0, 4000),
  };
}
"""


def cerrar_ventanas(page, tiempo: int = 3000) -> int:
    """Cierra todas las ventanas de detalle abiertas (boton X de cada una)."""
    cerradas = 0
    for _ in range(15):
        try:
            ventanas = page.locator(SELECTOR_VENTANA)
            cantidad = ventanas.count()
        except Exception:
            return cerradas
        if cantidad == 0:
            return cerradas
        ventana = ventanas.nth(cantidad - 1)
        try:
            ventana.locator(".x-tool-close").first.click(timeout=tiempo)
        except Exception:
            try:
                texto = ventana.inner_text(timeout=1000)[:200].replace("\n", " ")
            except Exception:
                texto = "?"
            logger.warning("Una ventana no se dejo cerrar con la X (%r); pruebo con Escape", texto)
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
        cerradas += 1
        time.sleep(0.3)
    return cerradas


def esperar_ventana(page, tiempo: int):
    """La ventana que abrio el ultimo click, ya con su contenido cargado."""
    limite = time.monotonic() + tiempo / 1000
    ventana = None
    while time.monotonic() < limite:
        try:
            cantidad = page.locator(SELECTOR_VENTANA).count()
        except Exception:
            cantidad = 0
        if cantidad:
            ventana = page.locator(SELECTOR_VENTANA).nth(cantidad - 1)
            break
        time.sleep(0.2)
    if ventana is None:
        return None
    limite = time.monotonic() + TIMEOUT_CONTENIDO_VENTANA_MS / 1000
    while time.monotonic() < limite:
        try:
            if ventana.evaluate(JS_VENTANA_CARGADA):
                break
        except Exception:
            break
        time.sleep(0.25)
    pausar(0.4, 0.15)
    return ventana


def leer_seccion(page, ruta: list, tiempo: int, espera_ventana: Optional[int] = None,
                 cerrar: bool = True) -> Optional[dict]:
    """Abre la seccion del arbol, lee su ventana entera y la cierra (salvo
    cerrar=False, para sacarle una captura antes)."""
    cerrar_ventanas(page)
    idx = ubicar_nodo(page, ruta, tiempo)
    if idx is None:
        logger.warning("No encontre '%s' en el arbol de la DDJJ", _ruta_legible(ruta))
        guardar_captura(page, f"no encontre {_ruta_legible(ruta)}")
        return None
    panel_antes = _texto_panel(page)
    try:
        _texto_nodo(page.locator(SELECTOR_FILAS_ARBOL).nth(idx)).click(timeout=tiempo)
    except Exception:
        logger.warning("No pude clickear '%s' en el arbol", _ruta_legible(ruta))
        guardar_captura(page, f"no pude clickear {_ruta_legible(ruta)}")
        return None
    pausar()
    esperar_sin_carga(page, tiempo)
    ventana = esperar_ventana(page, espera_ventana or tiempo)
    if ventana is None:
        # Respaldo: si en vez de una ventana flotante el detalle aparecio en
        # el panel de la derecha (#panelContenedor, id real), se lee de ahi.
        panel_despues = _texto_panel(page)
        if panel_despues and panel_despues != panel_antes:
            logger.warning("'%s' no abrio ventana flotante; leo el panel de la derecha", _ruta_legible(ruta))
            ventana = page.locator(SELECTOR_PANEL_DETALLE)
        else:
            if espera_ventana is None:
                logger.warning("'%s' no abrio ninguna ventana", _ruta_legible(ruta))
                guardar_captura(page, f"sin ventana {_ruta_legible(ruta)}")
            return None
    try:
        datos = ventana.evaluate(JS_LEER_VENTANA)
    except Exception:
        datos = None
    if cerrar:
        cerrar_ventanas(page)
    return datos


def leer_liquidacion(page, tiempo: int) -> Tuple[Dict[str, float], List[str]]:
    """'Liquidacion del Impuesto, Presentacion' -> saldo a favor, importe a
    pagar (subtotal) y total pagado, buscando las etiquetas que indica la
    hoja DETALLE en la ventana de la carpeta y en las de sus sub-secciones.
    Devuelve (valores, textos_de_las_ventanas_para_diagnostico)."""
    raiz = RUTAS["liquidacion"]
    encontrados: Dict[str, float] = {}
    textos: List[str] = []

    rutas = []
    idx = ubicar_nodo(page, raiz, tiempo)
    if idx is None:
        logger.warning("No encontre '%s' en el arbol de la DDJJ", _ruta_legible(raiz))
        return encontrados, textos
    filas = leer_arbol(page)
    if filas[idx]["expandible"] and not filas[idx]["expandida"]:
        _expandir_nodo(page, idx, tiempo)
    hijos: List[str] = []
    limite = time.monotonic() + TIMEOUT_NODO_MS / 1000
    while time.monotonic() < limite:  # los hijos pueden tardar en aparecer
        filas = leer_arbol(page)
        idx, _ = resolver_ruta(filas, raiz)
        hijos = [filas[k]["texto"] for k in _hijos_directos(filas, idx)] if idx is not None else []
        if hijos or idx is None or not filas[idx]["expandible"]:
            break
        time.sleep(0.3)
    logger.info("Liquidacion: sub-secciones %s", hijos or "(ninguna)")
    for hijo in hijos:
        if any(p in normalizar(hijo) for p in PALABRAS_PELIGROSAS):
            continue
        rutas.append(raiz + [re.compile("^" + re.escape(normalizar(hijo)) + "$")])
    rutas.append(raiz)  # por ultimo, la carpeta en si (por si abre su propia ventana)

    for ruta in rutas:
        # Se cierra a mano despues: si la ventana no trae ninguna de las
        # etiquetas buscadas, conviene la captura con la ventana a la vista.
        datos = leer_seccion(page, ruta, tiempo, espera_ventana=4000 if ruta is raiz else None, cerrar=False)
        if not datos:
            continue
        textos.append(f"[{datos.get('titulo') or _ruta_legible(ruta)}] {datos.get('texto', '')[:700]}")
        faltan = {k: p for k, p in ETIQUETAS_LIQUIDACION.items() if k not in encontrados}
        nuevos = buscar_etiquetas(datos, faltan)
        if not nuevos:
            guardar_captura(page, f"liquidacion sin etiquetas {_ruta_legible(ruta[-1:])}")
        cerrar_ventanas(page)
        encontrados.update(nuevos)
        if len(encontrados) == len(ETIQUETAS_LIQUIDACION):
            break
    return encontrados, textos


def extraer_campos(page, anio: int, mes: int, elegida: FilaDDJJ, tiempo: int) -> Optional[DatosDDJJ]:
    """Extrae los valores de la DDJJ abierta. Devuelve None si lo que quedo
    abierto NO es el anio/mes/version pedidos (en ese caso no se escribe
    nada, para no cargar datos de otro mes)."""
    datos = DatosDDJJ(tipo=elegida.tipo)
    etiqueta = f"{MESES[mes - 1]}/{anio}"

    abierta = ddjj_abierta(leer_arbol(page))
    if abierta:
        anio_arbol, mes_arbol, tipo_arbol = abierta
        if (anio_arbol and anio_arbol != anio) or (mes_arbol and mes_arbol != mes):
            logger.error("Se abrio %s/%s en vez de %s: no escribo nada",
                          MESES[mes_arbol - 1] if mes_arbol else "?", anio_arbol or "?", etiqueta)
            return None
        if RE_NODO_VERSION.search(normalizar(tipo_arbol)) and numero_version(tipo_arbol) != elegida.version:
            logger.error("Se abrio '%s' en vez de '%s' (%s): no escribo nada", tipo_arbol, elegida.tipo, etiqueta)
            return None
    else:
        logger.warning("No pude confirmar en el arbol que DDJJ quedo abierta; lo verifico con el titulo")

    # Rubro 1 - Informacion para el Calculo del impuesto
    info = leer_seccion(page, RUTAS["info_calculo"], tiempo, cerrar=False)
    if info and not parsear_rubro1(info)[0] and parsear_rubro1(info)[1] is None:
        guardar_captura(page, "no pude leer informacion para el calculo")
    cerrar_ventanas(page)
    if info:
        titulo = parsear_titulo_ventana(info.get("titulo", ""))
        if titulo:
            anio_t, mes_t, tipo_t = titulo
            if anio_t != anio or (mes_t and mes_t != mes) or numero_version(tipo_t) != elegida.version:
                logger.error("La ventana dice '%s' y se pidio %s '%s': no escribo nada",
                              info.get("titulo"), etiqueta, elegida.tipo)
                return None
        elif not abierta:
            datos.avisos.append("no se pudo confirmar anio/mes/version de la DDJJ abierta")
        actividades, total_base, total_valor = parsear_rubro1(info)
        datos.actividades = actividades
        suma_base = sum(a.base or 0 for a in actividades) if actividades else None
        suma_valor = sum(a.valor or 0 for a in actividades) if actividades else None
        base = total_base if total_base is not None else suma_base
        valor = total_valor if total_valor is not None else suma_valor
        if base is not None:
            datos.conceptos["base imponible"] = base
        if valor is not None:
            datos.conceptos["anticipo determinado"] = valor
        if len(actividades) > 1:
            logger.info("  %d actividades: %s", len(actividades),
                        "; ".join(f"{a.codigo} base {_fmt(a.base)} al {_fmt(a.alicuota)}%" for a in actividades))
        if not actividades and base is None:
            datos.avisos.append("no pude leer la tabla de 'Informacion para el Calculo del impuesto'")
    else:
        datos.avisos.append("no pude abrir 'Informacion para el Calculo del impuesto' (base imponible, "
                            "anticipo y alicuota quedan sin cargar)")

    lectores = [
        ("impuestos internos", RUTAS["conceptos_no_integran"], valor_impuestos_internos),
        ("otros creditos", RUTAS["saldo_favor_anterior"], valor_saldo_periodo_anterior),
        ("pago a cuenta", RUTAS["pagos_a_cuenta"], valor_monto_total),
        ("percepciones", RUTAS["percepciones"], valor_monto_total),
        ("retenciones", RUTAS["retenciones"], valor_monto_total),
        ("retenciones bancarias", RUTAS["retenciones_bancarias"], valor_monto_total),
    ]
    # Un 0 que muestra AGIP se registra como 0 (y en el Excel queda en blanco,
    # como hace la planilla). Lo que NO se pudo leer va a "REVISAR A MANO",
    # con captura de la ventana, para no confundirlo con un 0 real.
    for concepto, ruta, lector in lectores:
        valor = None
        try:
            ventana = leer_seccion(page, ruta, tiempo, cerrar=False)
            if ventana is None:
                datos.avisos.append(f"no pude abrir '{_ruta_legible(ruta)}' ({concepto} queda sin cargar)")
            else:
                valor = lector(ventana)
                if valor is None and (ventana.get("encabezados") or ventana.get("filas")):
                    valor = 0.0  # la grilla esta, pero ese mes no tuvo movimientos
                if valor is None:
                    datos.avisos.append(f"no pude leer {concepto} en la ventana "
                                        f"'{ventana.get('titulo') or _ruta_legible(ruta)}'")
                    guardar_captura(page, f"no pude leer {concepto}")
        except Exception as exc:
            logger.warning("Error leyendo '%s': %s", _ruta_legible(ruta), exc)
            datos.avisos.append(f"error leyendo {concepto}: {str(exc).splitlines()[0][:100] if str(exc) else exc!r}")
        finally:
            cerrar_ventanas(page)
        if valor is not None:
            datos.conceptos[concepto] = valor

    try:
        liquidacion, textos = leer_liquidacion(page, tiempo)
    except Exception as exc:
        logger.warning("Error leyendo la Liquidacion: %s", exc)
        liquidacion, textos = {}, []
    datos.conceptos.update(liquidacion)
    faltan = [k for k in ETIQUETAS_LIQUIDACION if k not in liquidacion]
    if faltan:
        datos.avisos.append(f"Liquidacion: no encontre {', '.join(faltan)}")
        logger.warning("Liquidacion: no encontre %s. Texto de las ventanas que se abrieron "
                        "(pasame esto para ajustar la busqueda):", ", ".join(faltan))
        for texto in textos or ["(no se abrio ninguna ventana)"]:
            logger.warning("    %s", texto)
    return datos


def _resumen_valores(datos: DatosDDJJ) -> str:
    abreviaturas = [("base imponible", "base"), ("anticipo determinado", "anticipo"),
                    ("percepciones", "percep"), ("retenciones", "ret"), ("retenciones bancarias", "ret.banc"),
                    ("impuestos internos", "imp.int"), ("pago a cuenta", "pago cta"),
                    ("otros creditos", "otros cred"), ("saldo a favor", "saldo fav"),
                    ("importe a pagar(subtotal)", "a pagar"), ("total pagado", "total pagado")]
    return ", ".join(f"{corto} {_fmt(datos.conceptos[largo])}" for largo, corto in abreviaturas
                     if largo in datos.conceptos)


# --------------------------------------------------------------------------
# Reporte de la corrida
# --------------------------------------------------------------------------

@dataclass
class Reporte:
    errores: Set[Tuple[int, int]] = field(default_factory=set)       # (anio, cuit) -> rojo
    detalle_errores: List[str] = field(default_factory=list)
    sin_ddjj: Dict[Tuple[int, str, int], List[int]] = field(default_factory=dict)
    revisar: List[str] = field(default_factory=list)
    sin_procesar: List[str] = field(default_factory=list)   # la corrida se corto antes de llegar a ellos

    @staticmethod
    def _quien(bloque: BloqueCliente, mes: Optional[int]) -> str:
        cuando = f" {MESES[mes - 1]}" if mes else ""
        return f"[{bloque.anio}] {bloque.nombre} (CUIT {bloque.cuit}){cuando}"

    def error(self, bloque: BloqueCliente, mes: Optional[int], texto: str) -> None:
        self.errores.add((bloque.anio, bloque.cuit))
        self.detalle_errores.append(f"{self._quien(bloque, mes)}: {texto}")

    def falta_ddjj(self, bloque: BloqueCliente, mes: int) -> None:
        self.sin_ddjj.setdefault((bloque.anio, bloque.nombre, bloque.cuit), []).append(mes)

    def aviso(self, bloque: BloqueCliente, mes: int, texto: str) -> None:
        self.revisar.append(f"{self._quien(bloque, mes)}: {texto}")

    def imprimir(self) -> None:
        if self.sin_procesar:
            logger.info("--- NO SE LLEGARON A PROCESAR: la corrida se corto antes (%d) ---", len(self.sin_procesar))
            for linea in self.sin_procesar:
                logger.info("  %s", linea)
        if self.revisar:
            logger.info("--- REVISAR A MANO (%d) ---", len(self.revisar))
            for linea in self.revisar:
                logger.info("  %s", linea)
        if self.sin_ddjj:
            logger.info("--- Meses sin DDJJ presentada en AGIP (%d clientes) ---", len(self.sin_ddjj))
            for (anio, nombre, cuit), meses in self.sin_ddjj.items():
                logger.info("  [%s] %s (CUIT %s): %s", anio, nombre, cuit, ", ".join(MESES[m - 1] for m in meses))
        if self.detalle_errores:
            logger.info("--- Errores (%d) ---", len(self.detalle_errores))
            for linea in self.detalle_errores:
                logger.info("  %s", linea)


def procesar_cliente(page, ws: Worksheet, bloque: BloqueCliente, pendientes: List[int],
                      candidatas: List[str], tiempo_espera: int, guardar_cb, reporte: Reporte,
                      indice: Optional[Dict[Tuple[int, int], BloqueCliente]] = None) -> bool:
    """Devuelve True si logro iniciar sesion (con alguna de las candidatas),
    False si ninguna contrasena funciono. 'indice' se usa para buscar el
    saldo a favor de diciembre del anio anterior (Otros Creditos de enero)."""
    logger.info("Cliente %s (CUIT %s, %s): %d mes(es) pendientes, %d contrasena(s) a probar",
                bloque.nombre, bloque.cuit, bloque.anio, len(pendientes), len(candidatas))
    _capturas["contexto"] = f"{bloque.cuit}_{bloque.anio}"
    password_ok, pagina = iniciar_sesion(page, bloque.cuit, candidatas, tiempo_espera)
    if not password_ok:
        logger.error("CUIT %s: ninguna contrasena funciono (probe %d)", bloque.cuit, len(candidatas))
        reporte.error(bloque, None, "ninguna contrasena funciono")
        abiertas = [p for p in page.context.pages if not p.is_closed()]
        guardar_captura(abiertas[-1] if abiertas else page, "login fallido", es_error=True)
        return False
    bloque.password = password_ok

    periodos_existentes: Optional[Set[str]] = None
    meses_con_error_seguidos = 0
    for mes in pendientes:
        periodo = f"{bloque.anio}-{mes:02d}"
        etiqueta = f"{MESES[mes - 1]}/{bloque.anio}"
        _capturas["contexto"] = f"{bloque.cuit}_{periodo}"
        if periodos_existentes is not None and periodo not in periodos_existentes:
            logger.info("CUIT %s: %s no tiene DDJJ presentada en AGIP", bloque.cuit, etiqueta)
            reporte.falta_ddjj(bloque, mes)
            continue
        try:
            resultado = ir_a_declaracion(pagina, bloque.anio, mes, tiempo_espera)
            if resultado.completa:
                periodos_existentes = resultado.periodos
            if resultado.elegida is None:
                logger.warning("CUIT %s: no hay DDJJ presentada de %s", bloque.cuit, etiqueta)
                reporte.falta_ddjj(bloque, mes)
                meses_con_error_seguidos = 0
                continue
            elegida = resultado.elegida
            logger.info("CUIT %s: %s -> abro '%s' (presentada el %s)",
                        bloque.cuit, etiqueta, elegida.tipo, elegida.fecha or "?")
            datos = extraer_campos(pagina, bloque.anio, mes, elegida, tiempo_espera)
            if datos is None:
                reporte.error(bloque, mes, "la DDJJ que se abrio no era la pedida; no se escribio nada")
                guardar_captura(pagina, "ddjj equivocada", es_error=True)
                meses_con_error_seguidos += 1
            else:
                saldo_anterior = saldo_a_favor_anterior(ws.parent, indice or {}, bloque, mes)
                avisos = datos.avisos + escribir_valores(ws, bloque, mes, datos, saldo_anterior)
                guardar_cb()
                for aviso in avisos:
                    reporte.aviso(bloque, mes, aviso)
                logger.info("CUIT %s: %s guardado -> %s", bloque.cuit, etiqueta,
                            _resumen_valores(datos) or "sin valores")
                meses_con_error_seguidos = 0
        except Exception as exc:
            logger.exception("Error procesando CUIT %s, %s", bloque.cuit, etiqueta)
            reporte.error(bloque, mes, str(exc).splitlines()[0][:200] if str(exc) else type(exc).__name__)
            guardar_captura(pagina, f"error {type(exc).__name__}", es_error=True)
            meses_con_error_seguidos += 1
        finally:
            try:
                cerrar_ventanas(pagina)
            except Exception:
                pass
        if meses_con_error_seguidos >= MAX_MESES_SEGUIDOS_CON_ERROR:
            logger.error("CUIT %s: %d meses seguidos con error, paso al siguiente cliente",
                          bloque.cuit, meses_con_error_seguidos)
            break
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
                  ("Cliente Ejemplo Dos", 20222222223), ("Cliente Nuevo 2025", 20444444445)]

    for i, (nombre, cuit) in enumerate(datos_2024, start=4):
        ws.cell(row=i, column=1, value=nombre)
        if cuit:
            ws.cell(row=i, column=2, value=cuit)
    for i, (nombre, cuit) in enumerate(datos_2025, start=4):
        ws.cell(row=i, column=4, value=nombre)
        if cuit:
            ws.cell(row=i, column=5, value=cuit)

    ws_dia1 = wb.create_sheet("DIA 1")
    fila = _escribir_bloque_vacio(ws_dia1, 1, "Cliente Ejemplo Uno", 20111111112, "clave-demo-1",
                                  [("472130", "Venta al por menor de carnes rojas")])
    for col in range(2, 8):  # enero..junio ya cargados -> deberian salir como "hechos"
        ws_dia1.cell(row=2, column=col, value=1000.0 * col)
        ws_dia1.cell(row=13, column=col, value=50.0 * col)
    _escribir_bloque_vacio(ws_dia1, fila, "Cliente Ejemplo Dos", 20222222223, None,
                           [("471192", "Venta al por menor de tabaco"),
                            ("472200", "Venta al por menor de bebidas en comercios especializados")])

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
                      cuit_filtro: Optional[int], password_override: Optional[str],
                      meses_filtro: Optional[Set[int]] = None
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
        if meses_filtro:
            pendientes = [m for m in pendientes if m in meses_filtro]
        if not pendientes:
            continue
        candidatas = candidatas_password(bloque, cuit_filtro == cuit, password_override)
        trabajo.append((bloque, ws, pendientes, candidatas))
    trabajo.sort(key=lambda item: item[0].anio)
    return trabajo


def _se_puede_pintar(celda) -> bool:
    """True si el nombre no tiene color, es blanco, o ya tiene uno de los
    tres colores que usa este script (verde, dorado, rojo)."""
    relleno = celda.fill
    if relleno is None or relleno.fill_type is None:
        return True
    color = relleno.fgColor
    if color is None or color.type != "rgb" or not isinstance(color.rgb, str):
        return False
    return color.rgb.upper() in {"00000000", "FFFFFFFF", "FF92D050", "FFFFE599", "FFFF0000"}


def colorear_padron(wb, padron: Dict[int, List[Tuple[str, int, str]]],
                     indice: Dict[Tuple[int, int], BloqueCliente],
                     fallos: set) -> None:
    """Pinta el nombre de cada cliente en 'Lista de Clientes - IIBB' segun
    su estado, con los mismos colores de la leyenda ya existente (E2:H2):
    verde (COMPLETOS) si no le quedan meses pendientes, rojo (NO HECHOS)
    si tuvo algun error en esta corrida (prioridad sobre lo demas), dorado
    (NO ESTA CERRADO) si le quedan meses pendientes sin error puntual. Los
    clientes no ubicados en ninguna hoja no se tocan (no hay forma de
    saber su estado real). Solo se pintan nombres sin color o con alguno de
    estos tres colores: el azul (DDJJ HECHA) y cualquier otro color puesto a
    mano (amarillo, celeste, etc.) nunca se pisa."""
    ws = wb[HOJA_CLIENTES]
    columnas = {2024: (1, 2), 2025: (4, 5)}
    for anio, (col_nombre, col_cuit) in columnas.items():
        for fila in range(3, ws.max_row + 1):
            cuit_celda = ws.cell(row=fila, column=col_cuit).value
            if not isinstance(cuit_celda, (int, float)):
                continue
            cuit = int(cuit_celda)
            celda_nombre = ws.cell(row=fila, column=col_nombre)
            if not _se_puede_pintar(celda_nombre):
                continue

            if (anio, cuit) in fallos:
                celda_nombre.fill = RELLENO_NO_HECHOS
                continue

            bloque = indice.get((anio, cuit))
            if not bloque:
                continue
            pendientes = meses_pendientes(wb[bloque.hoja], bloque)
            celda_nombre.fill = RELLENO_COMPLETOS if not pendientes else RELLENO_NO_CERRADO


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
    parser.add_argument("--meses", help="Solo estos meses (1-12), ej: 1,2 (para probar)")
    parser.add_argument("--max-clientes", type=int, help="Limite de clientes a procesar en esta corrida")
    parser.add_argument("--sin-headless", action="store_true",
                         help="Muestra la ventana del navegador (por defecto corre sin ventana; los errores "
                              "igual quedan en el log y en capturas_errores/)")
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
    meses_filtro = {int(m) for m in args.meses.split(",")} if args.meses else None

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

    trabajo = calcular_trabajo(wb, indice, anios, args.cuit, args.password, meses_filtro)
    logger.info("Clientes con meses pendientes: %d", len(trabajo))

    if args.max_clientes:
        trabajo = trabajo[: args.max_clientes]

    reporte = Reporte()
    sin_agip_seguidos = 0
    intentados = 0
    sync_playwright = _importar_playwright()
    with sync_playwright() as pw:
        navegador = pw.chromium.launch(headless=not args.sin_headless)
        try:
            for n, (bloque, ws, pendientes, candidatas) in enumerate(trabajo, start=1):
                intentados = n
                logger.info("=== Cliente %d de %d ===", n, len(trabajo))
                contexto = navegador.new_context()
                pagina = contexto.new_page()
                try:
                    procesar_cliente(pagina, ws, bloque, pendientes, candidatas, args.timeout,
                                     lambda: guardar_workbook(wb, ruta), reporte, indice)
                    sin_agip_seguidos = 0
                except LoginNoDisponible as exc:
                    logger.error("CUIT %s: %s", bloque.cuit, exc)
                    reporte.error(bloque, None, str(exc))
                    guardar_captura(pagina, "agip no cargo el login", es_error=True)
                    sin_agip_seguidos += 1
                    if sin_agip_seguidos >= MAX_CLIENTES_SEGUIDOS_SIN_AGIP:
                        logger.error("AGIP no cargo el login con %d clientes seguidos: corto la corrida "
                                      "(quedan %d clientes sin procesar). Volve a correr el script mas tarde "
                                      "(retoma desde lo pendiente).",
                                      sin_agip_seguidos, len(trabajo) - n)
                        break
                except Exception as exc:
                    logger.exception("Error inesperado con CUIT %s, sigo con el siguiente cliente", bloque.cuit)
                    reporte.error(bloque, None, f"error inesperado: {exc}"[:200])
                finally:
                    contexto.close()
                pausar_entre_clientes()
        finally:
            guardar_workbook(wb, ruta)
            navegador.close()

    for bloque, _, pendientes, _ in trabajo[intentados:]:
        reporte.sin_procesar.append(f"[{bloque.anio}] {bloque.nombre} (CUIT {bloque.cuit}): "
                                    f"{len(pendientes)} mes(es) pendientes")

    completados = completar_otros_creditos(wb, indice)
    if completados:
        logger.info("--- 'Otros Creditos' completados con el saldo a favor del mes anterior (%d) ---",
                    len(completados))
        for linea in completados:
            logger.info("  %s", linea)

    logger.info("Listo. Planilla actualizada: %s", ruta.resolve())
    indice_final = indexar_workbook(wb)
    colorear_padron(wb, padron, indice_final, reporte.errores)
    guardar_workbook(wb, ruta)
    generar_resumen(wb, padron, indice_final)
    reporte.imprimir()


if __name__ == "__main__":
    main()
