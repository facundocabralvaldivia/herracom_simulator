"""
procesar_datos.py
─────────────────────────────────────────────────────────────────────────────
Lee fuente_tarifas.xlsx (Análisis de Gasto de Flete Por Provincia - Met Perú)
y exporta data.json para el dashboard interactivo.

Uso:
    python procesar_datos.py

Salida:
    - data.json  → consumido por index.html
"""

import json
from pathlib import Path

import numpy as np
import pandas as pd

FUENTE_EXCEL = "fuente_tarifas.xlsx"
SHEET = "Resumen"

# ─────────────────────────────────────────────────────────────────────────────
# COORDENADAS DE CIUDADES PARA PARADAS INTERMEDIAS (lat, lng)
# ─────────────────────────────────────────────────────────────────────────────
CIUDADES_COORDS = {
    "Lima": (-12.0664, -77.0364),
    "Cañete": (-13.0738, -76.3891),
    "Chincha": (-13.4167, -76.1333),
    "Pisco": (-13.7085, -76.2027),
    "Ica": (-14.0678, -75.7286),
    "Nazca": (-14.8330, -74.9462),
    "Marcona": (-15.3400, -75.1000),
    "Yauca": (-15.6833, -74.5333),
    "Chala": (-15.8653, -74.2356),
    "Atico": (-16.2290, -73.6750),
    "Camana": (-16.6221, -72.7088),
    "Pedregal": (-16.8833, -71.9333),
    "La Joya": (-16.6000, -71.8833),
    "Abancay": (-13.6336, -72.8833),
    "Puqui": (-13.8333, -73.5000),
    "Mollendo": (-17.0231, -72.0142),
    "Ilo": (-17.6394, -71.3380),
    "Moquegua": (-17.1953, -70.9367),
    "Tacna": (-18.0067, -70.2461),
    "Chiclayo": (-6.7713, -79.8367),
    "Piura": (-5.1945, -80.6328),
    "Paita": (-5.0894, -81.1147),
    "Sullana": (-4.9009, -80.6847),
    "Tumbes": (-3.5667, -80.4500),
    "Chimbote": (-9.0747, -78.5936),
    "Chepen": (-7.2264, -79.4333),
    "Trujillo": (-8.1120, -79.0288),
    "Huacho": (-11.1139, -77.5981),
    "Huaral": (-11.4975, -77.2156),
    "Chancay": (-11.5714, -77.2703),
    "Barranca": (-10.7534, -77.7575),
    "Supe": (-10.8006, -77.7022),
    "Huaraz": (-9.5264, -77.5278),
    "Huánuco": (-9.9346, -76.2461),
    "Tingo María": (-9.2993, -75.9989),
    "Tingo Maria": (-9.2993, -75.9989),
    "Juliaca": (-15.4902, -70.1333),
    "Cusco": (-13.5333, -71.9434),
    "Espinar": (-14.7833, -71.4167),
    "Sicuani": (-14.2500, -71.2167),
    "Quillabamba": (-12.8667, -72.6833),
    "Yauli": (-11.6671, -76.0340),
    "Junín": (-11.1607, -75.9930),
    "Junin": (-11.1607, -75.9930),
    "Tarma": (-11.4192, -75.6897),
    "Chanchamayo": (-11.1119, -75.3333),
    "Satipo": (-11.2571, -74.6344),
    "Ayacucho": (-13.1627, -74.2259),
    "Lambayeque": (-6.7012, -79.9060),
    "Olmos": (-5.9833, -79.7500),
}


def normalizar_paradas(paradas_str: str) -> str:
    """Ajustes de nombres compuestos en el Excel."""
    s = str(paradas_str).strip()
    s = s.replace("Piura Tumbes", "Piura, Tumbes")
    return s


def paradas_a_coords(paradas_str: str) -> tuple[list[str], list[list[float]]]:
    """Convierte string de paradas a (lista_nombres, lista_coords)."""
    paradas_str = normalizar_paradas(paradas_str)
    if not paradas_str or paradas_str.lower() == "directo":
        return [], []
    nombres = [p.strip() for p in paradas_str.split(",") if p.strip()]
    pares = []
    for nombre in nombres:
        coord = CIUDADES_COORDS.get(nombre)
        if coord:
            pares.append((nombre, list(coord)))
        else:
            print(f"  AVISO: Coordenada no encontrada para: '{nombre}'")
    if pares:
        n, c = zip(*pares)
        return list(n), list(c)
    return [], []


def coord_str_to_list(coord_str: str) -> list[float]:
    parts = str(coord_str).strip().split(",")
    return [round(float(parts[0]), 7), round(float(parts[1]), 7)]


def _num(val, decimals=8):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return None
    return round(float(val), decimals)


def _int(val):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return 0
    return int(val)


def cargar_excel(path: str = FUENTE_EXCEL) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=SHEET, header=1)
    df = df.replace({np.nan: None})
    return df


def procesar_a_json(df: pd.DataFrame) -> dict:
    col_grafico = [c for c in df.columns if "Gr" in c and "fico" in c][0]
    col_km = [c for c in df.columns if c.startswith("KM")][0]
    col_horas = [c for c in df.columns if "Horas" in c][0]
    col_tarifa_bulto = [c for c in df.columns if "Tarifa Por Bulto" in c][0]
    col_tarifa_pedido = [c for c in df.columns if "Tarifa Por Pedido" in c][0]
    col_tarifa_viaje = [c for c in df.columns if "Tarifa Por Viaje" in c][0]

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

    for destino_nombre, grupo in df.groupby("Punto de Llegada"):
        primera = grupo.iloc[0]
        coord_dest = coord_str_to_list(primera["Coordenadas Punto de Llegada"])
        nombre_grafico = str(primera[col_grafico]).strip()

        agencias = []
        for _, row in grupo.iterrows():
            paradas_str = str(row["Paradas"]) if row["Paradas"] else "Directo"
            nombres_paradas, coords_paradas = paradas_a_coords(paradas_str)
            tipo_viaje = (
                "Directo"
                if paradas_str.strip().upper() == "DIRECTO"
                else "Con Paradas"
            )

            agencias.append(
                {
                    "nombre": row["Agencia"],
                    "costo_kg_ponderado": _num(row["Costo por Kilo Ponderado USD"], 2),
                    "km": _int(row[col_km]),
                    "horas": _int(row[col_horas]),
                    "tipo_viaje": tipo_viaje,
                    "paradas": nombres_paradas,
                    "paradas_coords": coords_paradas,
                    "total_pedidos": _int(row["Total Pedidos"]),
                    "total_bultos": _int(row["Total Bultos"]),
                    "flete_venta_usd": _num(row["Flete /Venta ($ USD)"], 2),
                    "tarifa_por_peso": _num(row["Tarifa /Por Peso"], 2),
                    "tarifa_por_bulto": _num(row[col_tarifa_bulto], 2),
                    "tarifa_por_pedido": _num(row[col_tarifa_pedido], 2),
                    "tarifa_por_viaje": _num(row[col_tarifa_viaje], 2),
                    "coordenadas_llegada": coord_str_to_list(
                        row["Coordenadas Punto de Llegada"]
                    ),
                }
            )

        agencias.sort(
            key=lambda a: (
                a["costo_kg_ponderado"]
                if a["costo_kg_ponderado"] is not None
                else 9999
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
                "nombre_grafico": nombre_grafico,
                "coordenadas": coord_dest,
                "agencias": agencias,
            }
        )

    output["destinos"].sort(key=lambda d: d["nombre"])
    return output


if __name__ == "__main__":
    print("=" * 60)
    print("  Dashboard Logístico Perú — Procesador de Datos")
    print("=" * 60)

    fuente = Path(FUENTE_EXCEL)
    if not fuente.exists():
        raise FileNotFoundError(f"No se encontró {FUENTE_EXCEL}")

    print(f"\nLeyendo {FUENTE_EXCEL}...")
    df = cargar_excel()
    print(f"   {len(df)} filas | {df['Punto de Llegada'].nunique()} destinos")

    print("\nConsolidando por destino...")
    data_json = procesar_a_json(df)
    print(f"   {len(data_json['destinos'])} destinos procesados")

    salida = Path("data.json")
    with open(salida, "w", encoding="utf-8") as f:
        json.dump(data_json, f, ensure_ascii=False, indent=2)

    print(f"\ndata.json exportado ({salida.stat().st_size // 1024} KB)")
    for d in data_json["destinos"]:
        ags = ", ".join(a["nombre"] for a in d["agencias"])
        mejor = d["agencias"][0]
        print(
            f"  • {d['nombre_grafico']:6s} {d['nombre']:15s} | "
            f"mejor: {mejor['nombre']} (${mejor['costo_kg_ponderado']}/KG) | {ags}"
        )

    print("\nListo. Abra index.html en su navegador.")
