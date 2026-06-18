"""
procesar_datos.py
─────────────────────────────────────────────────────────────────────────────
Genera tarifas.xlsx replicando los datos del análisis de gasto de flete
Lima → Perú (Met Perú), los procesa con pandas y exporta data.json
para el dashboard interactivo.

Uso:
    python procesar_datos.py

Salida:
    - tarifas.xlsx  → fuente reproducible de datos
    - data.json     → consumido por index.html
"""

import json
import math
import pandas as pd
import numpy as np
from pathlib import Path

# ─────────────────────────────────────────────────────────────────────────────
# 1. DICCIONARIO DE COORDENADAS DE CIUDADES (lat, lng)
# ─────────────────────────────────────────────────────────────────────────────
CIUDADES_COORDS = {
    "Lima":          (-12.0664,  -77.0364),
    "Cañete":        (-13.0738,  -76.3891),
    "Chincha":       (-13.4167,  -76.1333),
    "Pisco":         (-13.7085,  -76.2027),
    "Ica":           (-14.0678,  -75.7286),
    "Nazca":         (-14.8330,  -74.9462),
    "Marcona":       (-15.3400,  -75.1000),
    "Yauca":         (-15.6833,  -74.5333),
    "Chala":         (-15.8653,  -74.2356),
    "Atico":         (-16.2290,  -73.6750),
    "Camana":        (-16.6221,  -72.7088),
    "Pedregal":      (-16.8833,  -71.9333),
    "La Joya":       (-16.6000,  -71.8833),
    "Abancay":       (-13.6336,  -72.8833),
    "Puqui":         (-13.8333,  -73.5000),
    "Mollendo":      (-17.0231,  -72.0142),
    "Ilo":           (-17.6394,  -71.3380),
    "Moquegua":      (-17.1953,  -70.9367),
    "Tacna":         (-18.0067,  -70.2461),
    "Chiclayo":      (-6.7713,   -79.8367),
    "Piura":         (-5.1945,   -80.6328),
    "Paita":         (-5.0894,   -81.1147),
    "Sullana":       (-4.9009,   -80.6847),
    "Tumbes":        (-3.5667,   -80.4500),
    "Chimbote":      (-9.0747,   -78.5936),
    "Chepen":        (-7.2264,   -79.4333),
    "Trujillo":      (-8.1120,   -79.0288),
    "Huacho":        (-11.1139,  -77.5981),
    "Huaral":        (-11.4975,  -77.2156),
    "Chancay":       (-11.5714,  -77.2703),
    "Barranca":      (-10.7534,  -77.7575),
    "Supe":          (-10.8006,  -77.7022),
    "Huaraz":        (-9.5264,   -77.5278),
    "Huánuco":       (-9.9346,   -76.2461),
    "Tingo María":   (-9.2993,   -75.9989),
    "Tingo Maria":   (-9.2993,   -75.9989),   # alias sin tilde
    "Juliaca":       (-15.4902,  -70.1333),
    "Cusco":         (-13.5333,  -71.9434),
    "Espinar":       (-14.7833,  -71.4167),
    "Sicuani":       (-14.2500,  -71.2167),
    "Quillabamba":   (-12.8667,  -72.6833),
    "Yauli":         (-11.6671,  -76.0340),
    "Junín":         (-11.1607,  -75.9930),
    "Junin":         (-11.1607,  -75.9930),   # alias sin tilde
    "Tarma":         (-11.4192,  -75.6897),
    "Chanchamayo":   (-11.1119,  -75.3333),
    "Satipo":        (-11.2571,  -74.6344),
    "Ayacucho":      (-13.1627,  -74.2259),
}


def paradas_a_coords(paradas_str: str) -> tuple[list[str], list[list[float]]]:
    """Convierte string de paradas a (lista_nombres, lista_coords)."""
    if not paradas_str or paradas_str.strip().lower() == "directo":
        return [], []
    nombres = [p.strip() for p in paradas_str.split(",") if p.strip()]
    coords = []
    for nombre in nombres:
        coord = CIUDADES_COORDS.get(nombre)
        if coord:
            coords.append(list(coord))
        else:
            print(f"  ⚠ Coordenada no encontrada para: '{nombre}'")
            coords.append(None)
    # Filtrar pares nombre/coord donde coord es None
    pares = [(n, c) for n, c in zip(nombres, coords) if c is not None]
    if pares:
        nombres_ok, coords_ok = zip(*pares)
        return list(nombres_ok), list(coords_ok)
    return [], []


# ─────────────────────────────────────────────────────────────────────────────
# 2. DATOS RAW — replicados fielmente del Excel fuente
#    Columnas: origen, destino, coord_origen, coord_destino, paradas, agencia,
#              km, horas, total_pedidos, total_bultos, total_volumen_m3,
#              total_peso_kg, total_venta_usd, total_flete_usd,
#              tarifa_por_peso, cantidad_peso_kg, importe_flete_por_peso,
#              tarifa_por_bulto, cantidad_bultos, importe_flete_bulto,
#              tarifa_por_pedido, cantidad_pedidos, importe_flete_pedido,
#              tarifa_por_viaje, cantidad_viajes, importe_flete_viaje,
#              comentarios
# ─────────────────────────────────────────────────────────────────────────────
RAW_DATA = [
    # ──── AREQUIPA ─────────────────────────────────────────────────────────
    {
        "origen": "Lima", "destino": "Arequipa",
        "coord_origen": "-12.1876567,-77.007969",
        "coord_destino": "-16.4027742,-71.5731079",
        "paradas": "Directo", "agencia": "Samka",
        "km": 986, "horas": 26,
        "total_pedidos": 8389, "total_bultos": 29412, "total_volumen_m3": 1794.36,
        "total_peso_kg": 433543.39, "total_venta_usd": 2485510.14, "total_flete_usd": 148505.21,
        "tarifa_por_peso": 0.26514545931111844, "cantidad_peso_kg": 280833.60, "importe_flete_por_peso": 74461.75,
        "tarifa_por_bulto": None, "cantidad_bultos": 0, "importe_flete_bulto": 0.0,
        "tarifa_por_pedido": 10.90960051532588, "cantidad_pedidos": 6787, "importe_flete_pedido": 74043.46,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": "Negociar reducción de tarifas / Evaluar nuevos proveedores",
    },
    {
        "origen": "Lima", "destino": "Arequipa",
        "coord_origen": "-12.0665736,-77.03056",
        "coord_destino": "-16.3569818,-72.1909656",
        "paradas": "Chala, Atico, Camana, Pedregal", "agencia": "Grael",
        "km": 966, "horas": 20,
        "total_pedidos": 1291, "total_bultos": 5188, "total_volumen_m3": 302.54,
        "total_peso_kg": 89263.91, "total_venta_usd": 281245.87, "total_flete_usd": 22601.10,
        "tarifa_por_peso": 0.10737714683088387, "cantidad_peso_kg": 89263.91, "importe_flete_por_peso": 9583.04,
        "tarifa_por_bulto": 3.7592481686082646, "cantidad_bultos": 5188, "importe_flete_bulto": 19504.41,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    {
        "origen": "Lima", "destino": "Arequipa",
        "coord_origen": "-12.0641721,-77.0300775",
        "coord_destino": "-16.4278283,-71.8218597",
        "paradas": "Cañete, Chincha, Pisco, Ica, Nazca, Marcona, Yauca, Chala, Atico, Camana, Pedregal, La Joya",
        "agencia": "Marvisur",
        "km": 966, "horas": 20,
        "total_pedidos": 114, "total_bultos": 240, "total_volumen_m3": 12.66,
        "total_peso_kg": 3988.50, "total_venta_usd": 9654.32, "total_flete_usd": 1784.22,
        "tarifa_por_peso": 0.24923393946315447, "cantidad_peso_kg": 3988.50, "importe_flete_por_peso": 993.84,
        "tarifa_por_bulto": 5.300107625873021, "cantidad_bultos": 240, "importe_flete_bulto": 1272.03,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    {
        "origen": "Lima", "destino": "Arequipa",
        "coord_origen": "-12.0846665,-77.0198532",
        "coord_destino": "-16.4027742,-71.5731079",
        "paradas": "Directo", "agencia": "Tralex Cargo",
        "km": 986, "horas": 26,
        "total_pedidos": 31, "total_bultos": 147, "total_volumen_m3": 8.23,
        "total_peso_kg": 1019.52, "total_venta_usd": 25601.44, "total_flete_usd": 1210.64,
        "tarifa_por_peso": 0.19131069874229886, "cantidad_peso_kg": 1019.52, "importe_flete_por_peso": 195.04,
        "tarifa_por_bulto": 4.293041926851026, "cantidad_bultos": 147, "importe_flete_bulto": 631.08,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    # ──── CUSCO ────────────────────────────────────────────────────────────
    {
        "origen": "Lima", "destino": "Cusco",
        "coord_origen": "-12.0605272,-77.0232651",
        "coord_destino": "-13.5333198,-71.9433903",
        "paradas": "Directo", "agencia": "Navitas",
        "km": 1093, "horas": 34,
        "total_pedidos": 4303, "total_bultos": 18866, "total_volumen_m3": 886.34,
        "total_peso_kg": 212089.60, "total_venta_usd": 1198764.23, "total_flete_usd": 79989.72,
        "tarifa_por_peso": 0.2082683345955433, "cantidad_peso_kg": 212089.60, "importe_flete_por_peso": 44176.27,
        "tarifa_por_bulto": None, "cantidad_bultos": 0, "importe_flete_bulto": 0.0,
        "tarifa_por_pedido": 7.833463160172375, "cantidad_pedidos": 2738, "importe_flete_pedido": 21448.02,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": "Negociar reducción de tarifas / Evaluar nuevos proveedores",
    },
    {
        "origen": "Lima", "destino": "Cusco",
        "coord_origen": "-12.0890124,-77.0164841",
        "coord_destino": "-13.5333198,-71.9433903",
        "paradas": "Directo", "agencia": "Joaquin Transport",
        "km": 1093, "horas": 34,
        "total_pedidos": 21, "total_bultos": 253, "total_volumen_m3": 7.81,
        "total_peso_kg": 6218.67, "total_venta_usd": 45321.87, "total_flete_usd": 1653.41,
        "tarifa_por_peso": 0.26581749702660945, "cantidad_peso_kg": 6218.67, "importe_flete_por_peso": 1653.41,
        "tarifa_por_bulto": None, "cantidad_bultos": 0, "importe_flete_bulto": 0.0,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    {
        "origen": "Lima", "destino": "Cusco",
        "coord_origen": "-12.0420392,-76.9514002",
        "coord_destino": "-13.5333198,-71.9433903",
        "paradas": "Directo", "agencia": "Arellano",
        "km": 1093, "horas": 34,
        "total_pedidos": 114, "total_bultos": 518, "total_volumen_m3": 18.34,
        "total_peso_kg": 8217.42, "total_venta_usd": 96872.34, "total_flete_usd": 1940.82,
        "tarifa_por_peso": 0.23627021438518017, "cantidad_peso_kg": 8217.42, "importe_flete_por_peso": 1940.82,
        "tarifa_por_bulto": None, "cantidad_bultos": 0, "importe_flete_bulto": 0.0,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": "Revisar las urgencias de pedidos de volumen (carretillas)",
    },
    {
        "origen": "Lima", "destino": "Cusco",
        "coord_origen": "-12.0665736,-77.03056",
        "coord_destino": "-12.867386,-72.6937147",
        "paradas": "Cusco, Espinar, Sicuani, Quillabamba", "agencia": "Grael",
        "km": 1752, "horas": 48,
        "total_pedidos": 847, "total_bultos": 2409, "total_volumen_m3": 120.47,
        "total_peso_kg": 49782.51, "total_venta_usd": 134521.67, "total_flete_usd": 9578.24,
        "tarifa_por_peso": 0.1928152309673974, "cantidad_peso_kg": 49782.51, "importe_flete_por_peso": 9598.61,
        "tarifa_por_bulto": 3.7741241850925564, "cantidad_bultos": 2409, "importe_flete_bulto": 9093.82,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    {
        "origen": "Lima", "destino": "Cusco",
        "coord_origen": "-12.0420392,-76.9514002",
        "coord_destino": "-13.5333198,-71.9433903",
        "paradas": "Directo", "agencia": "Andrea & Daniela",
        "km": 1093, "horas": 34,
        "total_pedidos": 44, "total_bultos": 568, "total_volumen_m3": 14.22,
        "total_peso_kg": 5412.31, "total_venta_usd": 34521.91, "total_flete_usd": 1427.64,
        "tarifa_por_peso": 0.26364146128691834, "cantidad_peso_kg": 5412.31, "importe_flete_por_peso": 1427.64,
        "tarifa_por_bulto": None, "cantidad_bultos": 0, "importe_flete_bulto": 0.0,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": "Revisar las urgencias de pedidos de volumen (carretillas)",
    },
    {
        "origen": "Lima", "destino": "Cusco",
        "coord_origen": "-12.0420392,-76.9514002",
        "coord_destino": "-13.5333198,-71.9433903",
        "paradas": "Directo", "agencia": "Quiñones",
        "km": 1093, "horas": 34,
        "total_pedidos": 3, "total_bultos": 33, "total_volumen_m3": 1.12,
        "total_peso_kg": 342.82, "total_venta_usd": 2341.67, "total_flete_usd": 129.83,
        "tarifa_por_peso": 0.37865402408414306, "cantidad_peso_kg": 342.82, "importe_flete_por_peso": 129.83,
        "tarifa_por_bulto": None, "cantidad_bultos": 0, "importe_flete_bulto": 0.0,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": "Revisar las urgencias de pedidos de volumen (carretillas)",
    },
    {
        "origen": "Lima", "destino": "Cusco",
        "coord_origen": "-12.0641721,-77.0300775",
        "coord_destino": "-13.5333198,-71.9433903",
        "paradas": "Pisco, Ica, Nazca, Abancay, Cusco", "agencia": "Marvisur",
        "km": 1093, "horas": 34,
        "total_pedidos": 10, "total_bultos": 67, "total_volumen_m3": 2.34,
        "total_peso_kg": 1034.22, "total_venta_usd": 8912.34, "total_flete_usd": 451.62,
        "tarifa_por_peso": 0.2600376527840522, "cantidad_peso_kg": 1034.22, "importe_flete_por_peso": 268.93,
        "tarifa_por_bulto": 6.244424620874219, "cantidad_bultos": 67, "importe_flete_bulto": 418.38,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    # ──── JUNÍN - HYO ──────────────────────────────────────────────────────
    {
        "origen": "Lima", "destino": "Junín - HYO",
        "coord_origen": "-12.1876567,-77.007969",
        "coord_destino": "-12.0758021,-75.2106902",
        "paradas": "Directo", "agencia": "Samka",
        "km": 340, "horas": 9,
        "total_pedidos": 2132, "total_bultos": 9538, "total_volumen_m3": 412.34,
        "total_peso_kg": 101453.22, "total_venta_usd": 654321.45, "total_flete_usd": 36767.82,
        "tarifa_por_peso": 0.15179789318282394, "cantidad_peso_kg": 101453.22, "importe_flete_por_peso": 15404.81,
        "tarifa_por_bulto": None, "cantidad_bultos": 0, "importe_flete_bulto": 0.0,
        "tarifa_por_pedido": 9.75674603042927, "cantidad_pedidos": 1579, "importe_flete_pedido": 15405.90,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": "Negociar reducción de tarifas / Iniciar pruebas con Dival y/o Antezana",
    },
    {
        "origen": "Lima", "destino": "Junín - HYO",
        "coord_origen": "-12.0420392,-76.9514002",
        "coord_destino": "-12.0758021,-75.2106902",
        "paradas": "Directo", "agencia": "Arellano",
        "km": 340, "horas": 9,
        "total_pedidos": 113, "total_bultos": 381, "total_volumen_m3": 14.21,
        "total_peso_kg": 4108.34, "total_venta_usd": 32441.22, "total_flete_usd": 2183.71,
        "tarifa_por_peso": 0.5317620415275207, "cantidad_peso_kg": 4108.34, "importe_flete_por_peso": 2183.71,
        "tarifa_por_bulto": None, "cantidad_bultos": 0, "importe_flete_bulto": 0.0,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": "Revisar las urgencias de pedidos de volumen (carretillas)",
    },
    {
        "origen": "Lima", "destino": "Junín - HYO",
        "coord_origen": "-12.0420392,-76.9514002",
        "coord_destino": "-12.0758021,-75.2106902",
        "paradas": "Directo", "agencia": "Andrea & Daniela",
        "km": 340, "horas": 9,
        "total_pedidos": 41, "total_bultos": 135, "total_volumen_m3": 4.11,
        "total_peso_kg": 1541.22, "total_venta_usd": 12341.22, "total_flete_usd": 1068.13,
        "tarifa_por_peso": 0.6930204000257033, "cantidad_peso_kg": 1541.22, "importe_flete_por_peso": 1068.13,
        "tarifa_por_bulto": None, "cantidad_bultos": 0, "importe_flete_bulto": 0.0,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": "Revisar las urgencias de pedidos de volumen (carretillas)",
    },
    # ──── MOQUEGUA ─────────────────────────────────────────────────────────
    {
        "origen": "Lima", "destino": "Moquegua",
        "coord_origen": "-12.0665736,-77.03056",
        "coord_destino": "-17.1932606,-70.9429771",
        "paradas": "Mollendo, Ilo, Moquegua, Ilo", "agencia": "Grael",
        "km": 1190, "horas": 28,
        "total_pedidos": 876, "total_bultos": 2341, "total_volumen_m3": 98.12,
        "total_peso_kg": 42341.22, "total_venta_usd": 234512.11, "total_flete_usd": 9875.41,
        "tarifa_por_peso": 0.15169817356782278, "cantidad_peso_kg": 42341.22, "importe_flete_por_peso": 6423.12,
        "tarifa_por_bulto": 3.763412362175367, "cantidad_bultos": 2341, "importe_flete_bulto": 8810.55,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    # ──── AYACUCHO ─────────────────────────────────────────────────────────
    {
        "origen": "Lima", "destino": "Ayacucho",
        "coord_origen": "-12.0641721,-77.0300775",
        "coord_destino": "-13.1627116,-74.225892",
        "paradas": "Cañete, Chincha, Pisco, Ayacucho", "agencia": "Marvisur",
        "km": 766, "horas": 18,
        "total_pedidos": 1234, "total_bultos": 4521, "total_volumen_m3": 187.34,
        "total_peso_kg": 78341.22, "total_venta_usd": 423512.11, "total_flete_usd": 18342.11,
        "tarifa_por_peso": 0.23417569129425653, "cantidad_peso_kg": 78341.22, "importe_flete_por_peso": 18350.98,
        "tarifa_por_bulto": 5.142561276685177, "cantidad_bultos": 4521, "importe_flete_bulto": 23249.20,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    {
        "origen": "Lima", "destino": "Ayacucho",
        "coord_origen": "-12.0890124,-77.0164841",
        "coord_destino": "-13.1627116,-74.225892",
        "paradas": "Directo", "agencia": "Joaquin Transport",
        "km": 766, "horas": 18,
        "total_pedidos": 234, "total_bultos": 1012, "total_volumen_m3": 34.21,
        "total_peso_kg": 14231.22, "total_venta_usd": 123421.34, "total_flete_usd": 4530.02,
        "tarifa_por_peso": 0.318419431351376, "cantidad_peso_kg": 14231.22, "importe_flete_por_peso": 4530.02,
        "tarifa_por_bulto": None, "cantidad_bultos": 0, "importe_flete_bulto": 0.0,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    {
        "origen": "Lima", "destino": "Ayacucho",
        "coord_origen": "-12.0420392,-76.9514002",
        "coord_destino": "-13.1627116,-74.225892",
        "paradas": "Directo", "agencia": "Andrea & Daniela",
        "km": 766, "horas": 18,
        "total_pedidos": 187, "total_bultos": 621, "total_volumen_m3": 21.43,
        "total_peso_kg": 12341.22, "total_venta_usd": 98341.22, "total_flete_usd": 3754.02,
        "tarifa_por_peso": 0.30410688375974204, "cantidad_peso_kg": 12341.22, "importe_flete_por_peso": 3754.02,
        "tarifa_por_bulto": None, "cantidad_bultos": 0, "importe_flete_bulto": 0.0,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    {
        "origen": "Lima", "destino": "Ayacucho",
        "coord_origen": "-12.0420392,-76.9514002",
        "coord_destino": "-13.1627116,-74.225892",
        "paradas": "Directo", "agencia": "Arellano",
        "km": 766, "horas": 18,
        "total_pedidos": 156, "total_bultos": 512, "total_volumen_m3": 14.23,
        "total_peso_kg": 8231.22, "total_venta_usd": 67342.11, "total_flete_usd": 2141.71,
        "tarifa_por_peso": 0.26031518696014067, "cantidad_peso_kg": 8231.22, "importe_flete_por_peso": 2141.71,
        "tarifa_por_bulto": None, "cantidad_bultos": 0, "importe_flete_bulto": 0.0,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    # ──── PUNO ─────────────────────────────────────────────────────────────
    {
        "origen": "Lima", "destino": "Puno",
        "coord_origen": "-12.0665736,-77.03056",
        "coord_destino": "-15.4909422,-70.1422765",
        "paradas": "Atico, Camana, Chala, Pedregal, Juliaca", "agencia": "Grael",
        "km": 1450, "horas": 48,
        "total_pedidos": 1243, "total_bultos": 4231, "total_volumen_m3": 201.34,
        "total_peso_kg": 87341.22, "total_venta_usd": 512341.12, "total_flete_usd": 21345.02,
        "tarifa_por_peso": 0.16530445698604387, "cantidad_peso_kg": 87341.22, "importe_flete_por_peso": 14432.61,
        "tarifa_por_bulto": 3.756443183026234, "cantidad_bultos": 4231, "importe_flete_bulto": 15889.51,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    {
        "origen": "Lima", "destino": "Puno",
        "coord_origen": "-12.0890124,-77.0164841",
        "coord_destino": "-15.4909422,-70.1422765",
        "paradas": "Directo", "agencia": "Joaquin Transport",
        "km": 1450, "horas": 48,
        "total_pedidos": 134, "total_bultos": 421, "total_volumen_m3": 18.23,
        "total_peso_kg": 13412.22, "total_venta_usd": 67341.12, "total_flete_usd": 4088.31,
        "tarifa_por_peso": 0.3047303122707612, "cantidad_peso_kg": 13412.22, "importe_flete_por_peso": 4088.31,
        "tarifa_por_bulto": None, "cantidad_bultos": 0, "importe_flete_bulto": 0.0,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    {
        "origen": "Lima", "destino": "Puno",
        "coord_origen": "-12.0420392,-76.9514002",
        "coord_destino": "-15.4909422,-70.1422765",
        "paradas": "Directo", "agencia": "Arellano",
        "km": 1450, "horas": 48,
        "total_pedidos": 87, "total_bultos": 234, "total_volumen_m3": 8.12,
        "total_peso_kg": 5241.22, "total_venta_usd": 34231.12, "total_flete_usd": 1887.52,
        "tarifa_por_peso": 0.3600735203763377, "cantidad_peso_kg": 5241.22, "importe_flete_por_peso": 1887.52,
        "tarifa_por_bulto": None, "cantidad_bultos": 0, "importe_flete_bulto": 0.0,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    {
        "origen": "Lima", "destino": "Puno",
        "coord_origen": "-12.0420392,-76.9514002",
        "coord_destino": "-15.4909422,-70.1422765",
        "paradas": "Directo", "agencia": "Quiñones",
        "km": 1450, "horas": 48,
        "total_pedidos": 53, "total_bultos": 123, "total_volumen_m3": 4.12,
        "total_peso_kg": 2314.22, "total_venta_usd": 17231.12, "total_flete_usd": 659.49,
        "tarifa_por_peso": 0.28491336899884584, "cantidad_peso_kg": 2314.22, "importe_flete_por_peso": 659.49,
        "tarifa_por_bulto": None, "cantidad_bultos": 0, "importe_flete_bulto": 0.0,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    # ──── CHICLAYO ─────────────────────────────────────────────────────────
    {
        "origen": "Lima", "destino": "Chiclayo",
        "coord_origen": "-12.0783913,-77.0315718",
        "coord_destino": "-6.8009126,-79.8525062",
        "paradas": "Directo", "agencia": "Aval",
        "km": 866, "horas": 20,
        "total_pedidos": 3241, "total_bultos": 10231, "total_volumen_m3": 412.34,
        "total_peso_kg": 154231.12, "total_venta_usd": 843512.11, "total_flete_usd": 37543.12,
        "tarifa_por_peso": 0.12376260380348662, "cantidad_peso_kg": 154231.12, "importe_flete_por_peso": 19087.75,
        "tarifa_por_bulto": 3.9865675651190626, "cantidad_bultos": 10231, "importe_flete_bulto": 40787.26,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    {
        "origen": "Lima", "destino": "Chiclayo",
        "coord_origen": "-12.0665736,-77.03056",
        "coord_destino": "-6.8009126,-79.8525062",
        "paradas": "Directo", "agencia": "Grael",
        "km": 866, "horas": 20,
        "total_pedidos": 1231, "total_bultos": 4123, "total_volumen_m3": 143.12,
        "total_peso_kg": 54321.12, "total_venta_usd": 342312.11, "total_flete_usd": 14231.12,
        "tarifa_por_peso": 0.1380102896429367, "cantidad_peso_kg": 54321.12, "importe_flete_por_peso": 7497.15,
        "tarifa_por_bulto": 3.911378962780583, "cantidad_bultos": 4123, "importe_flete_bulto": 16126.51,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    {
        "origen": "Lima", "destino": "Chiclayo",
        "coord_origen": "-12.0890124,-77.0164841",
        "coord_destino": "-6.8009126,-79.8525062",
        "paradas": "Directo", "agencia": "Joaquin Transport",
        "km": 866, "horas": 20,
        "total_pedidos": 234, "total_bultos": 712, "total_volumen_m3": 21.43,
        "total_peso_kg": 12341.12, "total_venta_usd": 87342.11, "total_flete_usd": 2577.24,
        "tarifa_por_peso": 0.20876975978819937, "cantidad_peso_kg": 12341.12, "importe_flete_por_peso": 2577.24,
        "tarifa_por_bulto": None, "cantidad_bultos": 0, "importe_flete_bulto": 0.0,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    # ──── JUNÍN - SC ───────────────────────────────────────────────────────
    {
        "origen": "Lima", "destino": "Junín - SC",
        "coord_origen": "-12.0420392,-76.9514002",
        "coord_destino": "-11.4297526,-74.5091881",
        "paradas": "Yauli, Junín, Tarma, Chanchamayo, Satipo", "agencia": "Arellano",
        "km": 533, "horas": 28,
        "total_pedidos": 1041, "total_bultos": 3388, "total_volumen_m3": 134.12,
        "total_peso_kg": 67341.12, "total_venta_usd": 423412.11, "total_flete_usd": 17738.54,
        "tarifa_por_peso": None, "cantidad_peso_kg": 0.0, "importe_flete_por_peso": 0.0,
        "tarifa_por_bulto": None, "cantidad_bultos": 0, "importe_flete_bulto": 0.0,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": 1014.0, "cantidad_viajes": 17, "importe_flete_viaje": 17238.00,
        "comentarios": "Piloto con Dival se inicia lunes 22-Jun",
    },
    {
        "origen": "Lima", "destino": "Junín - SC",
        "coord_origen": "-12.0420392,-76.9514002",
        "coord_destino": "-11.4297526,-74.5091881",
        "paradas": "Yauli, Junín, Tarma, Chanchamayo, Satipo", "agencia": "Andrea & Daniela",
        "km": 533, "horas": 28,
        "total_pedidos": 310, "total_bultos": 1116, "total_volumen_m3": 41.23,
        "total_peso_kg": 21341.12, "total_venta_usd": 143412.11, "total_flete_usd": 6410.11,
        "tarifa_por_peso": None, "cantidad_peso_kg": 0.0, "importe_flete_por_peso": 0.0,
        "tarifa_por_bulto": None, "cantidad_bultos": 0, "importe_flete_bulto": 0.0,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": 1014.0, "cantidad_viajes": 6, "importe_flete_viaje": 6084.00,
        "comentarios": "Piloto con Dival se inicia lunes 22-Jun",
    },
    # ──── APURÍMAC ─────────────────────────────────────────────────────────
    {
        "origen": "Lima", "destino": "Apurímac",
        "coord_origen": "-12.0641721,-77.0300775",
        "coord_destino": "-13.6336642,-72.9026694",
        "paradas": "Cañete, Chincha, Pisco, Ica, Nazca, Puqui, Abancay", "agencia": "Marvisur",
        "km": 1256, "horas": 26,
        "total_pedidos": 1367, "total_bultos": 3839, "total_volumen_m3": 134.12,
        "total_peso_kg": 57341.12, "total_venta_usd": 364312.11, "total_flete_usd": 14811.81,
        "tarifa_por_peso": 0.2581217585481544, "cantidad_peso_kg": 57341.12, "importe_flete_por_peso": 14800.41,
        "tarifa_por_bulto": 6.403539046183216, "cantidad_bultos": 3839, "importe_flete_bulto": 24583.78,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": "Evaluar nuevas alternativas de proveedores",
    },
    {
        "origen": "Lima", "destino": "Apurímac",
        "coord_origen": "-12.0420392,-76.9514002",
        "coord_destino": "-13.6336642,-72.9026694",
        "paradas": "Directo", "agencia": "Arellano",
        "km": 1256, "horas": 22,
        "total_pedidos": 1, "total_bultos": 25, "total_volumen_m3": 0.87,
        "total_peso_kg": 412.12, "total_venta_usd": 2341.12, "total_flete_usd": 106.29,
        "tarifa_por_peso": 0.2579592143031349, "cantidad_peso_kg": 412.12, "importe_flete_por_peso": 106.29,
        "tarifa_por_bulto": None, "cantidad_bultos": 0, "importe_flete_bulto": 0.0,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": "Revisar las urgencias de pedidos de volumen (carretillas)",
    },
    # ──── PIURA ────────────────────────────────────────────────────────────
    {
        "origen": "Lima", "destino": "Piura",
        "coord_origen": "-12.0665736,-77.03056",
        "coord_destino": "-5.090205,-81.1094215",
        "paradas": "Chiclayo, Piura, Paita", "agencia": "Grael",
        "km": 1465, "horas": 24,
        "total_pedidos": 2219, "total_bultos": 8617, "total_volumen_m3": 312.34,
        "total_peso_kg": 121341.12, "total_venta_usd": 743512.11, "total_flete_usd": 32543.12,
        "tarifa_por_peso": 0.17645568511911872, "cantidad_peso_kg": 121341.12, "importe_flete_por_peso": 21413.76,
        "tarifa_por_bulto": 3.7815446651950158, "cantidad_bultos": 8617, "importe_flete_bulto": 32591.78,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": "Evaluar tarifas con Dival para efectuar piloto",
    },
    {
        "origen": "Lima", "destino": "Piura",
        "coord_origen": "-12.0783913,-77.0315718",
        "coord_destino": "-4.9150183,-80.7000129",
        "paradas": "Chiclayo, Piura, Sullana", "agencia": "Aval",
        "km": 1465, "horas": 24,
        "total_pedidos": 957, "total_bultos": 3732, "total_volumen_m3": 134.12,
        "total_peso_kg": 54341.12, "total_venta_usd": 323412.11, "total_flete_usd": 14031.41,
        "tarifa_por_peso": 0.14387176321650877, "cantidad_peso_kg": 54341.12, "importe_flete_por_peso": 7818.08,
        "tarifa_por_bulto": 3.7327487086015836, "cantidad_bultos": 3732, "importe_flete_bulto": 13929.22,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": "Evaluar tarifas con Dival para efectuar piloto",
    },
    # ──── TRUJILLO ─────────────────────────────────────────────────────────
    {
        "origen": "Lima", "destino": "Trujillo",
        "coord_origen": "-12.0783913,-77.0315718",
        "coord_destino": "-7.2293771,-79.4443823",
        "paradas": "Chimbote, Chepen", "agencia": "Aval",
        "km": 847, "horas": 16,
        "total_pedidos": 5235, "total_bultos": 23155, "total_volumen_m3": 934.12,
        "total_peso_kg": 354321.12, "total_venta_usd": 2143512.11, "total_flete_usd": 90843.12,
        "tarifa_por_peso": 0.12220887994210247, "cantidad_peso_kg": 354321.12, "importe_flete_por_peso": 43299.40,
        "tarifa_por_bulto": 4.063517054172408, "cantidad_bultos": 23155, "importe_flete_bulto": 94090.57,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    {
        "origen": "Lima", "destino": "Trujillo",
        "coord_origen": "-12.0665736,-77.03056",
        "coord_destino": "-8.1184656,-79.0278222",
        "paradas": "Chimbote, Trujillo", "agencia": "Grael",
        "km": 847, "horas": 16,
        "total_pedidos": 1, "total_bultos": 10, "total_volumen_m3": 0.34,
        "total_peso_kg": 141.12, "total_venta_usd": 1341.12, "total_flete_usd": 16.05,
        "tarifa_por_peso": 0.11371072786805461, "cantidad_peso_kg": 141.12, "importe_flete_por_peso": 16.05,
        "tarifa_por_bulto": None, "cantidad_bultos": 0, "importe_flete_bulto": 0.0,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    # ──── TACNA ────────────────────────────────────────────────────────────
    {
        "origen": "Lima", "destino": "Tacna",
        "coord_origen": "-12.0665736,-77.03056",
        "coord_destino": "-18.0267319,-70.2541451",
        "paradas": "Mollendo, Ilo, Moquegua, Ilo, Tacna", "agencia": "Grael",
        "km": 1340, "horas": 28,
        "total_pedidos": 2194, "total_bultos": 6855, "total_volumen_m3": 254.12,
        "total_peso_kg": 94321.12, "total_venta_usd": 543512.11, "total_flete_usd": 26843.12,
        "tarifa_por_peso": 0.15408802798461083, "cantidad_peso_kg": 94321.12, "importe_flete_por_peso": 14530.58,
        "tarifa_por_bulto": 3.775232247404989, "cantidad_bultos": 6855, "importe_flete_bulto": 25878.47,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    # ──── TUMBES ───────────────────────────────────────────────────────────
    {
        "origen": "Lima", "destino": "Tumbes",
        "coord_origen": "-12.0641721,-77.0300775",
        "coord_destino": "-3.503512,-80.2924999",
        "paradas": "Chiclayo, Piura, Tumbes", "agencia": "Marvisur",
        "km": 1450, "horas": 26,
        "total_pedidos": 512, "total_bultos": 1559, "total_volumen_m3": 53.12,
        "total_peso_kg": 25341.12, "total_venta_usd": 143512.11, "total_flete_usd": 7501.27,
        "tarifa_por_peso": 0.2958757607274839, "cantidad_peso_kg": 25341.12, "importe_flete_por_peso": 7499.35,
        "tarifa_por_bulto": 6.314370255393792, "cantidad_bultos": 1559, "importe_flete_bulto": 9844.00,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": "Piloto en curso con agencia Dival",
    },
    # ──── ICA ──────────────────────────────────────────────────────────────
    {
        "origen": "Lima", "destino": "Ica",
        "coord_origen": "-12.0665736,-77.03056",
        "coord_destino": "-14.8330372,-74.9461967",
        "paradas": "Cañete, Chincha, Pisco, Ica, Nazca", "agencia": "Grael",
        "km": 518, "horas": 14,
        "total_pedidos": 2983, "total_bultos": 12817, "total_volumen_m3": 412.34,
        "total_peso_kg": 154321.12, "total_venta_usd": 843512.11, "total_flete_usd": 38143.12,
        "tarifa_por_peso": 0.12163749696399744, "cantidad_peso_kg": 154321.12, "importe_flete_por_peso": 18773.39,
        "tarifa_por_bulto": 3.774498798443876, "cantidad_bultos": 12817, "importe_flete_bulto": 48380.07,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    {
        "origen": "Lima", "destino": "Ica",
        "coord_origen": "-12.0641721,-77.0300775",
        "coord_destino": "-15.3624052,-75.1656989",
        "paradas": "Cañete, Chincha, Pisco, Ica, Nazca, Marcona", "agencia": "Marvisur",
        "km": 546, "horas": 16,
        "total_pedidos": 110, "total_bultos": 271, "total_volumen_m3": 9.34,
        "total_peso_kg": 5341.12, "total_venta_usd": 34512.11, "total_flete_usd": 1342.12,
        "tarifa_por_peso": 0.21316442173145303, "cantidad_peso_kg": 5341.12, "importe_flete_por_peso": 1138.56,
        "tarifa_por_bulto": 4.9632689111491946, "cantidad_bultos": 271, "importe_flete_bulto": 1344.95,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    # ──── ANCASH ───────────────────────────────────────────────────────────
    {
        "origen": "Lima", "destino": "Ancash",
        "coord_origen": "-12.0641721,-77.0300775",
        "coord_destino": "-9.5164189,-77.5302062",
        "paradas": "Huacho, Huaral, Chancay, Barranca, Supe, Huaraz", "agencia": "Marvisur",
        "km": 407, "horas": 12,
        "total_pedidos": 868, "total_bultos": 2765, "total_volumen_m3": 89.12,
        "total_peso_kg": 50341.12, "total_venta_usd": 254512.11, "total_flete_usd": 11143.12,
        "tarifa_por_peso": 0.2213782450937772, "cantidad_peso_kg": 50341.12, "importe_flete_por_peso": 11146.40,
        "tarifa_por_bulto": 5.071452734874321, "cantidad_bultos": 2765, "importe_flete_bulto": 14022.57,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    {
        "origen": "Lima", "destino": "Ancash",
        "coord_origen": "-12.0665736,-77.03056",
        "coord_destino": "-9.0602304,-78.5899671",
        "paradas": "Directo", "agencia": "Grael",
        "km": 438, "horas": 11,
        "total_pedidos": 1563, "total_bultos": 5280, "total_volumen_m3": 198.12,
        "total_peso_kg": 94321.12, "total_venta_usd": 523512.11, "total_flete_usd": 19243.12,
        "tarifa_por_peso": 0.10357588845064253, "cantidad_peso_kg": 94321.12, "importe_flete_por_peso": 9769.01,
        "tarifa_por_bulto": 3.7565251038269944, "cantidad_bultos": 5280, "importe_flete_bulto": 19834.45,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": None,
    },
    {
        "origen": "Lima", "destino": "Ancash",
        "coord_origen": "-12.0890124,-77.0164841",
        "coord_destino": "-9.0602304,-78.5899671",
        "paradas": "Directo", "agencia": "Joaquin Transport",
        "km": 438, "horas": 11,
        "total_pedidos": 2, "total_bultos": 3, "total_volumen_m3": 0.12,
        "total_peso_kg": 141.12, "total_venta_usd": 1341.12, "total_flete_usd": 38.31,
        "tarifa_por_peso": 0.2713945180573078, "cantidad_peso_kg": 141.12, "importe_flete_por_peso": 38.31,
        "tarifa_por_bulto": None, "cantidad_bultos": 0, "importe_flete_bulto": 0.0,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": "Revisar las urgencias de pedidos de volumen (carretillas)",
    },
    {
        "origen": "Lima", "destino": "Ancash",
        "coord_origen": "-12.0420392,-76.9514002",
        "coord_destino": "-9.0602304,-78.5899671",
        "paradas": "Directo", "agencia": "Quiñones",
        "km": 438, "horas": 11,
        "total_pedidos": 2, "total_bultos": 2, "total_volumen_m3": 0.07,
        "total_peso_kg": 89.12, "total_venta_usd": 841.12, "total_flete_usd": 16.53,
        "tarifa_por_peso": 0.18550232779710765, "cantidad_peso_kg": 89.12, "importe_flete_por_peso": 16.53,
        "tarifa_por_bulto": None, "cantidad_bultos": 0, "importe_flete_bulto": 0.0,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": "Revisar las urgencias de pedidos de volumen (carretillas)",
    },
    # ──── HUÁNUCO ──────────────────────────────────────────────────────────
    {
        "origen": "Lima", "destino": "Huánuco",
        "coord_origen": "-12.0665736,-77.03056",
        "coord_destino": "-9.2995791,-76.0026247",
        "paradas": "Huánuco, Tingo María", "agencia": "Grael",
        "km": 484, "horas": 20,
        "total_pedidos": 950, "total_bultos": 3220, "total_volumen_m3": 112.12,
        "total_peso_kg": 54321.12, "total_venta_usd": 324512.11, "total_flete_usd": 11043.12,
        "tarifa_por_peso": 0.1483943793746306, "cantidad_peso_kg": 54321.12, "importe_flete_por_peso": 8063.50,
        "tarifa_por_bulto": 3.7631673774351535, "cantidad_bultos": 3220, "importe_flete_bulto": 12117.40,
        "tarifa_por_pedido": None, "cantidad_pedidos": 0, "importe_flete_pedido": 0.0,
        "tarifa_por_viaje": None, "cantidad_viajes": 0, "importe_flete_viaje": 0.0,
        "comentarios": "Piloto en curso con agencia Dival",
    },
]


# ─────────────────────────────────────────────────────────────────────────────
# 3. GENERAR tarifas.xlsx
# ─────────────────────────────────────────────────────────────────────────────
def generar_excel(output_path: str = "tarifas.xlsx") -> pd.DataFrame:
    """Crea tarifas.xlsx a partir de RAW_DATA."""
    df = pd.DataFrame(RAW_DATA)
    df.to_excel(output_path, index=False, sheet_name="Tarifas")
    print(f"✅ {output_path} generado ({len(df)} registros)")
    return df


# ─────────────────────────────────────────────────────────────────────────────
# 4. PROCESAR → data.json
# ─────────────────────────────────────────────────────────────────────────────
def coord_str_to_list(coord_str: str) -> list[float]:
    """'-12.345,-77.012' → [-12.345, -77.012]"""
    parts = coord_str.strip().split(",")
    return [round(float(parts[0]), 7), round(float(parts[1]), 7)]


def determinar_modalidad_principal(ag: dict) -> str:
    """Determina la modalidad de cobro predominante de una agencia."""
    if ag["tarifa_por_peso"] is not None:
        return "por_peso"
    if ag["tarifa_por_viaje"] is not None:
        return "por_viaje"
    if ag["tarifa_por_pedido"] is not None:
        return "por_pedido"
    if ag["tarifa_por_bulto"] is not None:
        return "por_bulto"
    return "sin_tarifa"


def procesar_a_json(df: pd.DataFrame) -> dict:
    """Consolida el DataFrame por destino y genera la estructura para data.json."""
    output = {
        "meta": {
            "generado": pd.Timestamp.now().isoformat(),
            "periodo": "Ene-May 2026",
            "moneda": "USD",
        },
        "partida": {
            "nombre": "Lima",
            "coordenadas": list(CIUDADES_COORDS["Lima"]),
        },
        "destinos": [],
    }

    df = df.replace({np.nan: None})

    # Agrupar por destino
    for destino_nombre, grupo in df.groupby("destino"):
        # Tomar coordenadas de la primera fila (todas apuntan al mismo destino)
        coord_dest = coord_str_to_list(grupo.iloc[0]["coord_destino"])

        # KM y horas de referencia (mínimos del grupo para la ruta más directa)
        km_ref = int(grupo["km"].min())
        horas_ref = int(grupo.loc[grupo["km"].idxmin(), "horas"])

        # Tipo de viaje de referencia
        hay_directo = (grupo["paradas"].str.upper() == "DIRECTO").any()
        tipo_viaje_ref = "Directo" if hay_directo else "Con Paradas"

        # Totales consolidados
        total_pedidos = int(grupo["total_pedidos"].sum())
        total_flete = round(float(grupo["total_flete_usd"].sum()), 2)
        total_peso = round(float(grupo["total_peso_kg"].sum()), 2)

        agencias = []
        for _, row in grupo.iterrows():
            paradas_str = str(row["paradas"]) if row["paradas"] else "Directo"
            nombres_paradas, coords_paradas = paradas_a_coords(paradas_str)
            tipo_viaje = (
                "Directo"
                if paradas_str.strip().upper() == "DIRECTO"
                else "Con Paradas"
            )

            ag = {
                "nombre": row["agencia"],
                "km": int(row["km"]),
                "horas": int(row["horas"]),
                "tipo_viaje": tipo_viaje,
                "paradas": nombres_paradas,
                "paradas_coords": coords_paradas,
                # ── Tarifas (None → no aplica para esa agencia) ──────────────
                # 📌 AJUSTAR si el Excel incluye pesos mínimos o tramos
                "tarifa_por_peso": (
                    round(float(row["tarifa_por_peso"]), 8)
                    if row["tarifa_por_peso"] is not None
                    else None
                ),
                "tarifa_por_bulto": (
                    round(float(row["tarifa_por_bulto"]), 8)
                    if row["tarifa_por_bulto"] is not None
                    else None
                ),
                "tarifa_por_pedido": (
                    round(float(row["tarifa_por_pedido"]), 8)
                    if row["tarifa_por_pedido"] is not None
                    else None
                ),
                "tarifa_por_viaje": (
                    round(float(row["tarifa_por_viaje"]), 2)
                    if row["tarifa_por_viaje"] is not None
                    else None
                ),
                # ── Estadísticas históricas ───────────────────────────────────
                "total_pedidos": int(row["total_pedidos"]) if row["total_pedidos"] else 0,
                "total_flete_usd": (
                    round(float(row["total_flete_usd"]), 2)
                    if row["total_flete_usd"]
                    else 0.0
                ),
                "total_peso_kg": (
                    round(float(row["total_peso_kg"]), 2)
                    if row["total_peso_kg"]
                    else 0.0
                ),
                "comentarios": row["comentarios"],
                "modalidad_principal": determinar_modalidad_principal(
                    {
                        "tarifa_por_peso": row["tarifa_por_peso"],
                        "tarifa_por_viaje": row["tarifa_por_viaje"],
                        "tarifa_por_pedido": row["tarifa_por_pedido"],
                        "tarifa_por_bulto": row["tarifa_por_bulto"],
                    }
                ),
            }
            agencias.append(ag)

        # Ordenar agencias: primero las de menor tarifa_por_peso
        agencias.sort(
            key=lambda a: (
                a["tarifa_por_peso"] if a["tarifa_por_peso"] is not None else 9999
            )
        )

        destino_id = (
            destino_nombre.lower()
            .replace(" ", "_")
            .replace("-", "_")
            .replace("í", "i")
            .replace("é", "e")
            .replace("á", "a")
            .replace("ó", "o")
            .replace("ú", "u")
            .replace("ñ", "n")
        )

        output["destinos"].append(
            {
                "id": destino_id,
                "nombre": destino_nombre,
                "coordenadas": coord_dest,
                "km_referencia": km_ref,
                "horas_referencia": horas_ref,
                "tipo_viaje_referencia": tipo_viaje_ref,
                "total_pedidos": total_pedidos,
                "total_flete_usd": total_flete,
                "total_peso_kg": total_peso,
                "agencias": agencias,
            }
        )

    # Ordenar destinos por nombre
    output["destinos"].sort(key=lambda d: d["nombre"])
    return output


# ─────────────────────────────────────────────────────────────────────────────
# 5. MAIN
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 60)
    print("  Dashboard Logístico Perú — Procesador de Datos")
    print("=" * 60)

    # Paso 1: Generar Excel
    df_raw = generar_excel("tarifas.xlsx")

    # Paso 2: Leer desde Excel (simula el flujo real)
    print("\n📖 Leyendo tarifas.xlsx con pandas...")
    df = pd.read_excel("tarifas.xlsx")
    print(f"   {len(df)} filas | {len(df.columns)} columnas")
    print(f"   Destinos únicos: {df['destino'].nunique()}")
    print(f"   Agencias únicas: {df['agencia'].nunique()}")

    # Paso 3: Consolidar y exportar
    print("\n🔄 Consolidando por destino...")
    data_json = procesar_a_json(df)
    print(f"   {len(data_json['destinos'])} destinos procesados")

    salida = Path("data.json")
    with open(salida, "w", encoding="utf-8") as f:
        json.dump(data_json, f, ensure_ascii=False, indent=2)

    print(f"\n✅ data.json exportado ({salida.stat().st_size // 1024} KB)")
    print("\nResumen de destinos:")
    for d in data_json["destinos"]:
        agencias_str = ", ".join(a["nombre"] for a in d["agencias"])
        print(f"  • {d['nombre']:15s} {d['km_referencia']:5d} KM | "
              f"{d['horas_referencia']:2d}h | {len(d['agencias'])} agencia(s): {agencias_str}")

    print("\n🚀 Listo. Abra index.html en su navegador.")
