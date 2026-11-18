"""Pruebas de reporte.py. Se corren con:  pytest -q"""
import zipfile

import pandas as pd
import pytest
from openpyxl import load_workbook

import generar_ventas
import reporte


@pytest.fixture
def csv_ventas(tmp_path):
    ruta = tmp_path / "ventas.csv"
    generar_ventas.generar(ruta)
    return ruta


def test_el_reporte_cuadra_con_los_datos(csv_ventas, tmp_path):
    salida = tmp_path / "reporte.xlsx"
    k = reporte.generar_reporte(csv_ventas, salida)
    df = pd.read_csv(csv_ventas)

    assert k["ventas"] == pytest.approx(df["total"].sum())
    assert k["pedidos"] == len(df)

    libro = load_workbook(salida)
    assert libro.sheetnames == ["Resumen", "Por mes", "Por producto", "Por vendedor", "Por cliente", "Datos"]
    assert libro["Resumen"]["A5"].value == pytest.approx(df["total"].sum())  # tarjeta "Ventas totales"
    assert libro["Resumen"]["C5"].value == len(df)  # tarjeta "Pedidos"

    por_mes = pd.read_excel(salida, sheet_name="Por mes")
    assert len(por_mes) == 6
    assert por_mes["Ventas"].sum() == pytest.approx(df["total"].sum())
    assert por_mes["Pedidos"].sum() == len(df)

    por_producto = pd.read_excel(salida, sheet_name="Por producto")
    assert por_producto["% del total"].sum() == pytest.approx(1.0)
    assert por_producto["Ventas"].is_monotonic_decreasing  # el que más vende va primero


def test_incluye_las_dos_graficas(csv_ventas, tmp_path):
    salida = tmp_path / "reporte.xlsx"
    reporte.generar_reporte(csv_ventas, salida)
    with zipfile.ZipFile(salida) as archivo:
        graficas = [n for n in archivo.namelist() if n.startswith("xl/charts/chart")]
    assert len(graficas) == 2


def test_csv_sin_columnas_obligatorias_da_mensaje_claro(tmp_path):
    ruta = tmp_path / "malo.csv"
    pd.DataFrame({"fecha": ["2026-01-01"], "total": [100]}).to_csv(ruta, index=False)
    with pytest.raises(SystemExit, match="faltan"):
        reporte.cargar(ruta)


def test_filas_sin_fecha_valida_se_omiten(csv_ventas, tmp_path, capsys):
    df = pd.read_csv(csv_ventas)
    df.loc[0, "fecha"] = "no es fecha"
    ruta = tmp_path / "con_error.csv"
    df.to_csv(ruta, index=False)

    limpio = reporte.cargar(ruta)
    assert len(limpio) == len(df) - 1
    assert "omitieron 1" in capsys.readouterr().out
