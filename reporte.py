"""Genera un reporte de ventas en Excel (resumen, gráficas y tablas) a partir de un CSV.

En lugar de armar el reporte a mano cada semana, se corre un comando y se obtiene
un Excel con indicadores, gráficas nativas de Excel y las tablas que las alimentan.

Uso:
    python reporte.py datos/ventas.csv
    python reporte.py datos/ventas.csv --salida salida/reporte_ventas.xlsx
"""
from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

REQUERIDAS = ["fecha", "cliente", "producto", "vendedor", "cantidad", "precio_unitario", "total"]
MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
NOMBRES_BONITOS = {
    "fecha": "Fecha",
    "cliente": "Cliente",
    "producto": "Producto",
    "vendedor": "Vendedor",
    "cantidad": "Cantidad",
    "precio_unitario": "Precio unitario",
    "total": "Total",
}
COLUMNAS_DINERO = ("Ventas", "Precio unitario", "Total")


# ---------------------------------------------------------------- datos

def cargar(ruta: Path) -> pd.DataFrame:
    df = pd.read_csv(ruta)
    faltan = [c for c in REQUERIDAS if c not in df.columns]
    if faltan:
        raise SystemExit(f"Al CSV le faltan estas columnas: {', '.join(faltan)}")

    # Las fechas del CSV deben venir como AAAA-MM-DD; lo demás se trata como fecha inválida
    df["fecha"] = pd.to_datetime(df["fecha"], format="ISO8601", errors="coerce")
    sin_fecha = int(df["fecha"].isna().sum())
    if sin_fecha:
        print(f"Aviso: se omitieron {sin_fecha} filas sin fecha válida.")
        df = df.dropna(subset=["fecha"])
    if df.empty:
        raise SystemExit("No hay filas con datos para reportar.")
    return df.sort_values("fecha").reset_index(drop=True)


def tabla_por(df: pd.DataFrame, columna: str, titulo: str) -> pd.DataFrame:
    """Ventas, unidades y pedidos agrupados por una columna, de mayor a menor venta."""
    tabla = (
        df.groupby(columna)
        .agg(Ventas=("total", "sum"), Unidades=("cantidad", "sum"), Pedidos=("total", "size"))
        .sort_values("Ventas", ascending=False)
        .reset_index()
        .rename(columns={columna: titulo})
    )
    tabla["% del total"] = tabla["Ventas"] / tabla["Ventas"].sum()
    return tabla


def tabla_por_mes(df: pd.DataFrame) -> pd.DataFrame:
    """Lo mismo, pero por mes y en orden cronológico."""
    tabla = (
        df.groupby(df["fecha"].dt.to_period("M"))
        .agg(Ventas=("total", "sum"), Unidades=("cantidad", "sum"), Pedidos=("total", "size"))
        .reset_index()
    )
    tabla.insert(0, "Mes", tabla["fecha"].map(lambda p: f"{MESES[p.month - 1]} {p.year}"))
    tabla = tabla.drop(columns="fecha")
    tabla["% del total"] = tabla["Ventas"] / tabla["Ventas"].sum()
    return tabla


def indicadores(df, por_mes, por_producto, por_vendedor) -> dict:
    return {
        "ventas": float(df["total"].sum()),
        "pedidos": int(len(df)),
        "ticket": float(df["total"].mean()),
        "unidades": int(df["cantidad"].sum()),
        "mejor_mes": por_mes.loc[por_mes["Ventas"].idxmax(), "Mes"],
        "producto_lider": por_producto.iloc[0]["Producto"],
        "mejor_vendedor": por_vendedor.iloc[0]["Vendedor"],
        "desde": df["fecha"].min(),
        "hasta": df["fecha"].max(),
    }


# ---------------------------------------------------------------- Excel

def estilos(libro) -> dict:
    borde = {"border": 1, "border_color": "#C9D3E6", "bg_color": "#EEF2F9", "align": "center", "valign": "vcenter"}
    return {
        "encabezado": libro.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#1F3864", "valign": "vcenter"}),
        "dinero": libro.add_format({"num_format": "$#,##0.00"}),
        "porcentaje": libro.add_format({"num_format": "0.0%"}),
        "fecha": libro.add_format({"num_format": "dd/mm/yyyy", "align": "left"}),
        "titulo": libro.add_format({"bold": True, "font_size": 20, "font_color": "#1F3864"}),
        "subtitulo": libro.add_format({"font_color": "#595959"}),
        "nota": libro.add_format({"font_color": "#1F3864", "italic": True}),
        "kpi_titulo": libro.add_format({**borde, "font_size": 10, "font_color": "#595959"}),
        "kpi_dinero": libro.add_format({**borde, "bold": True, "font_size": 16, "font_color": "#1F3864", "num_format": "$#,##0.00"}),
        "kpi_entero": libro.add_format({**borde, "bold": True, "font_size": 16, "font_color": "#1F3864", "num_format": "#,##0"}),
    }


def escribir_tabla(escritor, nombre: str, tabla: pd.DataFrame, fmt: dict, filtro: bool = False) -> None:
    tabla.to_excel(escritor, sheet_name=nombre, index=False)
    hoja = escritor.sheets[nombre]
    for j, columna in enumerate(tabla.columns):
        hoja.write(0, j, columna, fmt["encabezado"])
        if columna in COLUMNAS_DINERO:
            hoja.set_column(j, j, 16, fmt["dinero"])
        elif columna == "% del total":
            hoja.set_column(j, j, 14, fmt["porcentaje"])
        elif columna == "Fecha":
            hoja.set_column(j, j, 12, fmt["fecha"])
        else:
            largo = max([len(str(columna))] + [len(str(v)) for v in tabla[columna].head(200)])
            hoja.set_column(j, j, min(max(largo + 2, 10), 40))
    hoja.freeze_panes(1, 0)
    if filtro:
        hoja.autofilter(0, 0, len(tabla), len(tabla.columns) - 1)


def armar_resumen(hoja, libro, fmt: dict, k: dict, por_mes: pd.DataFrame, por_producto: pd.DataFrame) -> None:
    hoja.hide_gridlines(2)
    hoja.set_landscape()
    hoja.fit_to_pages(1, 1)  # que se imprima completo en una sola hoja
    hoja.set_column("A:H", 14)
    hoja.set_row(4, 34)
    hoja.write("A1", "Reporte de ventas", fmt["titulo"])
    hoja.write("A2", f"Periodo: {k['desde']:%d/%m/%Y} al {k['hasta']:%d/%m/%Y}", fmt["subtitulo"])

    tarjetas = [
        ("Ventas totales", k["ventas"], "kpi_dinero"),
        ("Pedidos", k["pedidos"], "kpi_entero"),
        ("Ticket promedio", k["ticket"], "kpi_dinero"),
        ("Unidades vendidas", k["unidades"], "kpi_entero"),
    ]
    for i, (titulo, valor, estilo) in enumerate(tarjetas):
        columna = i * 2
        hoja.merge_range(3, columna, 3, columna + 1, titulo, fmt["kpi_titulo"])
        hoja.merge_range(4, columna, 4, columna + 1, valor, fmt[estilo])

    hoja.write(
        "A7",
        f"Mejor mes: {k['mejor_mes']}   |   Producto líder: {k['producto_lider']}   |   Mejor vendedor: {k['mejor_vendedor']}",
        fmt["nota"],
    )

    # Las gráficas leen directo de las hojas "Por mes" y "Por producto"
    n_meses = len(por_mes)
    columnas = libro.add_chart({"type": "column"})
    columnas.add_series(
        {
            "name": "Ventas",
            "categories": ["Por mes", 1, 0, n_meses, 0],
            "values": ["Por mes", 1, 1, n_meses, 1],
            "fill": {"color": "#2F5597"},
            "gap": 60,
        }
    )
    columnas.set_title({"name": "Ventas por mes", "name_font": {"size": 12}})
    columnas.set_legend({"none": True})
    columnas.set_y_axis({"num_format": "$#,##0", "major_gridlines": {"visible": True, "line": {"color": "#E0E0E0"}}})
    columnas.set_size({"width": 400, "height": 300})
    hoja.insert_chart("A9", columnas)

    n_productos = min(len(por_producto), 10)
    barras = libro.add_chart({"type": "bar"})
    barras.add_series(
        {
            "name": "Ventas",
            "categories": ["Por producto", 1, 0, n_productos, 0],
            "values": ["Por producto", 1, 1, n_productos, 1],
            "fill": {"color": "#2F5597"},
            "gap": 60,
        }
    )
    barras.set_title({"name": "Ventas por producto", "name_font": {"size": 12}})
    barras.set_legend({"none": True})
    # Que el producto que más vende quede arriba y los montos sigan abajo
    barras.set_y_axis({"reverse": True, "crossing": "max"})
    barras.set_x_axis({"num_format": "$#,##0", "major_gridlines": {"visible": True, "line": {"color": "#E0E0E0"}}})
    barras.set_size({"width": 400, "height": 300})
    hoja.insert_chart("E9", barras, {"x_offset": 12})


def generar_reporte(entrada: Path, salida: Path) -> dict:
    df = cargar(entrada)
    por_mes = tabla_por_mes(df)
    por_producto = tabla_por(df, "producto", "Producto")
    por_vendedor = tabla_por(df, "vendedor", "Vendedor")
    por_cliente = tabla_por(df, "cliente", "Cliente")
    k = indicadores(df, por_mes, por_producto, por_vendedor)

    salida.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(salida, engine="xlsxwriter", datetime_format="dd/mm/yyyy", date_format="dd/mm/yyyy") as escritor:
        libro = escritor.book
        fmt = estilos(libro)
        resumen = libro.add_worksheet("Resumen")  # se crea primero para que sea la primera pestaña

        escribir_tabla(escritor, "Por mes", por_mes, fmt)
        escribir_tabla(escritor, "Por producto", por_producto, fmt)
        escribir_tabla(escritor, "Por vendedor", por_vendedor, fmt)
        escribir_tabla(escritor, "Por cliente", por_cliente, fmt)
        escribir_tabla(escritor, "Datos", df.rename(columns=NOMBRES_BONITOS), fmt, filtro=True)
        armar_resumen(resumen, libro, fmt, k, por_mes, por_producto)
    return k


def main() -> None:
    parser = argparse.ArgumentParser(description="Genera un reporte de ventas en Excel")
    parser.add_argument("entrada", type=Path, nargs="?", default=Path("datos/ventas.csv"))
    parser.add_argument("--salida", type=Path, default=Path("salida/reporte_ventas.xlsx"))
    args = parser.parse_args()

    k = generar_reporte(args.entrada, args.salida)
    print(f"Listo: {args.salida}")
    print(f"  Ventas totales: ${k['ventas']:,.2f} en {k['pedidos']} pedidos")
    print(f"  Mejor mes: {k['mejor_mes']} | Producto líder: {k['producto_lider']} | Mejor vendedor: {k['mejor_vendedor']}")


if __name__ == "__main__":
    main()
