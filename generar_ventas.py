"""Genera un CSV de ventas inventadas para probar reporte.py.

Uso:
    python generar_ventas.py                       # crea datos/ventas.csv
    python generar_ventas.py --filas 800 --salida datos/otro.csv
"""
from __future__ import annotations

import argparse
import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

PRODUCTOS = {
    "Teclado mecánico": 850,
    "Mouse inalámbrico": 320,
    'Monitor 24"': 2900,
    "Audífonos": 540,
    "Webcam HD": 690,
    "Hub USB": 280,
}
# Cuánto se vende cada producto en relación con los demás
POPULARIDAD = [3, 6, 1, 4, 2, 5]
CLIENTES = [
    "Juan Pérez", "María López", "Carlos Ramírez", "Ana Torres", "Luis Hernández", "Sofía Díaz",
    "Pedro Castillo", "Elena Ruiz", "Jorge Salas", "Paula Vega", "Andrés Mora", "Camila Rojas",
]
VENDEDORES = ["Laura", "Diego", "Marcos"]
# Algunos meses venden más que otros (mayo, por ejemplo)
PESO_MES = {1: 0.8, 2: 0.9, 3: 1.0, 4: 1.0, 5: 1.5, 6: 1.1}


def generar(ruta: Path, filas: int = 450, semilla: int = 7) -> pd.DataFrame:
    rng = random.Random(semilla)
    dias = [date(2026, 1, 1) + timedelta(days=n) for n in range(181)]  # enero a junio
    fechas = rng.choices(dias, weights=[PESO_MES[d.month] for d in dias], k=filas)

    registros = []
    for fecha in sorted(fechas):
        producto = rng.choices(list(PRODUCTOS), weights=POPULARIDAD)[0]
        cantidad = rng.randint(1, 6)
        precio = PRODUCTOS[producto]
        registros.append(
            {
                "fecha": fecha.isoformat(),
                "cliente": rng.choice(CLIENTES),
                "producto": producto,
                "vendedor": rng.choice(VENDEDORES),
                "cantidad": cantidad,
                "precio_unitario": precio,
                "total": cantidad * precio,
            }
        )

    df = pd.DataFrame(registros)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(ruta, index=False, encoding="utf-8")
    return df


def main() -> None:
    parser = argparse.ArgumentParser(description="Genera ventas inventadas")
    parser.add_argument("--filas", type=int, default=450)
    parser.add_argument("--salida", type=Path, default=Path("datos/ventas.csv"))
    args = parser.parse_args()
    df = generar(args.salida, args.filas)
    print(f"Listo: {len(df)} ventas guardadas en {args.salida}")


if __name__ == "__main__":
    main()
