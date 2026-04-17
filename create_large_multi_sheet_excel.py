import random
from datetime import datetime, timedelta

import pandas as pd


random.seed(42)

OUTPUT_FILE = "stress_test_retail_operacion.xlsx"


def random_date(start_date: datetime, end_date: datetime) -> datetime:
    delta_days = (end_date - start_date).days
    return start_date + timedelta(days=random.randint(0, delta_days))


def choice_weighted(options):
    labels = [item[0] for item in options]
    weights = [item[1] for item in options]
    return random.choices(labels, weights=weights, k=1)[0]


def build_clientes(count: int) -> pd.DataFrame:
    ciudades = ["Bogota", "Medellin", "Cali", "Barranquilla", "Bucaramanga", "Pereira", "Cartagena"]
    segmentos = ["Corporativo", "Pyme", "Retail", "Gobierno"]
    estados = ["Activo", "Activo", "Activo", "Inactivo", "En riesgo"]
    rows = []
    for i in range(1, count + 1):
      created_at = random_date(datetime(2021, 1, 1), datetime(2025, 4, 1))
      rows.append({
          "Cliente_ID": f"C{i:05d}",
          "Nombre_Cliente": f"Cliente {i:05d}",
          "Ciudad": random.choice(ciudades),
          "Segmento": random.choice(segmentos),
          "Estado_Cliente": random.choice(estados),
          "Fecha_Alta": created_at.date(),
          "Score_Riesgo": random.randint(35, 98),
          "Limite_Credito": random.randint(5_000_000, 250_000_000),
      })
    return pd.DataFrame(rows)


def build_ventas(count: int) -> pd.DataFrame:
    canales = ["Online", "Tienda", "Mayorista", "Distribuidor"]
    categorias = ["Ferreteria", "Pinturas", "Electricos", "Hogar", "Construccion", "Jardineria"]
    asesores = [f"Asesor {i:02d}" for i in range(1, 31)]
    ciudades = ["Bogota", "Medellin", "Cali", "Barranquilla", "Pereira", "Cartagena"]
    rows = []
    for i in range(1, count + 1):
      unidades = random.randint(1, 40)
      precio = random.randint(12_000, 980_000)
      descuento = random.choice([0, 0, 0.03, 0.05, 0.08, 0.10, 0.12])
      subtotal = unidades * precio
      total = round(subtotal * (1 - descuento), 2)
      rows.append({
          "Venta_ID": f"V{i:06d}",
          "Fecha_Venta": random_date(datetime(2024, 1, 1), datetime(2025, 3, 31)).date(),
          "Cliente_ID": f"C{random.randint(1, 2500):05d}",
          "Ciudad": random.choice(ciudades),
          "Canal": random.choice(canales),
          "Categoria": random.choice(categorias),
          "Asesor": random.choice(asesores),
          "Unidades": unidades,
          "Precio_Unitario": precio,
          "Descuento": descuento,
          "Total_Venta": total,
          "Margen_Porcentaje": round(random.uniform(0.08, 0.42), 4),
      })
    return pd.DataFrame(rows)


def build_inventario(count: int) -> pd.DataFrame:
    categorias = ["Herramientas", "Pinturas", "Iluminacion", "Cables", "Plomeria", "Seguridad"]
    bodegas = ["Central", "Norte", "Sur", "Occidente"]
    estados = [("Disponible", 55), ("Stock bajo", 20), ("Agotado", 10), ("Reservado", 15)]
    rows = []
    for i in range(1, count + 1):
      stock = random.randint(0, 250)
      costo = random.randint(8_000, 550_000)
      precio = round(costo * random.uniform(1.12, 1.75), 2)
      rows.append({
          "SKU": f"SKU-{i:06d}",
          "Producto": f"Producto {i:06d}",
          "Categoria": random.choice(categorias),
          "Bodega": random.choice(bodegas),
          "Estado_Stock": choice_weighted(estados),
          "Stock_Actual": stock,
          "Punto_Reorden": random.randint(5, 40),
          "Costo_Unitario": costo,
          "Precio_Venta": precio,
          "Valor_Stock": round(stock * costo, 2),
      })
    return pd.DataFrame(rows)


def build_compras(count: int) -> pd.DataFrame:
    proveedores = [f"Proveedor {i:03d}" for i in range(1, 121)]
    categorias = ["Herramientas", "Pinturas", "Electricos", "Plomeria", "Seguridad", "Consumibles"]
    estados = ["Recibida", "En transito", "Parcial", "Cancelada"]
    rows = []
    for i in range(1, count + 1):
      valor = random.randint(400_000, 85_000_000)
      rows.append({
          "Compra_ID": f"OC-{i:06d}",
          "Fecha_Orden": random_date(datetime(2023, 1, 1), datetime(2025, 3, 31)).date(),
          "Proveedor": random.choice(proveedores),
          "Categoria": random.choice(categorias),
          "Estado_OC": random.choice(estados),
          "Tiempo_Entrega_Dias": random.randint(2, 45),
          "Valor_Orden": valor,
          "Desviacion_Presupuesto": round(random.uniform(-0.25, 0.35), 4),
      })
    return pd.DataFrame(rows)


def build_empleados(count: int) -> pd.DataFrame:
    areas = ["Ventas", "Operaciones", "Logistica", "Compras", "Finanzas", "Servicio al cliente", "TI"]
    sedes = ["Bogota", "Medellin", "Cali", "Barranquilla"]
    cargos = ["Analista", "Coordinador", "Jefe", "Auxiliar", "Especialista", "Supervisor"]
    tipos = ["Fijo", "Indefinido", "Temporal"]
    rows = []
    for i in range(1, count + 1):
      ingreso = random_date(datetime(2018, 1, 1), datetime(2025, 1, 1))
      rows.append({
          "Empleado_ID": f"E{i:05d}",
          "Nombre": f"Empleado {i:05d}",
          "Area": random.choice(areas),
          "Sede": random.choice(sedes),
          "Cargo": random.choice(cargos),
          "Tipo_Contrato": random.choice(tipos),
          "Salario_Mensual": random.randint(1_500_000, 18_000_000),
          "Antiguedad_Anios": round((datetime(2025, 4, 1) - ingreso).days / 365, 1),
          "Cumplimiento_KPI": round(random.uniform(0.55, 1.0), 4),
      })
    return pd.DataFrame(rows)


def build_tickets(count: int) -> pd.DataFrame:
    prioridades = ["Baja", "Media", "Alta", "Critica"]
    estados = ["Abierto", "En progreso", "Resuelto", "Escalado", "Cerrado"]
    canales = ["Telefono", "Correo", "Portal", "WhatsApp"]
    areas = ["Soporte POS", "Facturacion", "Inventario", "Logistica", "Comercial"]
    rows = []
    for i in range(1, count + 1):
      apertura = random_date(datetime(2024, 1, 1), datetime(2025, 4, 1))
      tiempo = random.randint(1, 240)
      rows.append({
          "Ticket_ID": f"T{i:06d}",
          "Fecha_Apertura": apertura.date(),
          "Canal_Entrada": random.choice(canales),
          "Area_Responsable": random.choice(areas),
          "Prioridad": random.choice(prioridades),
          "Estado_Ticket": random.choice(estados),
          "SLA_Horas": random.choice([4, 8, 12, 24, 48, 72]),
          "Horas_Resolucion": tiempo,
          "Cumplio_SLA": "Si" if tiempo <= 24 else "No",
          "Satisfaccion": random.randint(1, 5),
      })
    return pd.DataFrame(rows)


def main():
    sheets = {
        "Clientes": build_clientes(2500),
        "Ventas_Detalle": build_ventas(12000),
        "Inventario": build_inventario(6000),
        "Compras": build_compras(5000),
        "Empleados": build_empleados(1800),
        "Tickets_Soporte": build_tickets(9000),
    }

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"Archivo creado: {OUTPUT_FILE}")
    for sheet_name, df in sheets.items():
        print(f"- {sheet_name}: {len(df)} filas x {len(df.columns)} columnas")


if __name__ == "__main__":
    main()
