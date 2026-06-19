"""God-tier stress Excel. Diseñado para romper supuestos del pipeline.

Hojas (12):
  1. VENTAS_MASIVAS  — 5000 filas, 12 columnas, datos reales pero densos
  2. KPIs_MIXTOS     — columna heterogénea peor escenario (money + % + count + days en mismo col)
  3. NaN_BOMB        — 80% nulos, mix de tipos
  4. TEXTO_LARGO     — celdas con 800+ chars, observaciones reales
  5. UTF8_CAOS       — ñ, acentos, símbolos, emojis (¡🚀!), kanjis
  6. FECHAS_FORMATOS — datetime, date, string ISO, string "DD/MM/YYYY", Excel serial
  7. MULTI_CURRENCY  — montos en COP, USD, EUR mezclados (col + col separada de moneda)
  8. NEGATIVOS_CEROS — valores negativos, ceros, infinitos teóricos
  9. DUPLICADAS      — columnas con mismo nombre, headers en fila 3
  10. HOJA_VACIA     — 0 filas
  11. UNA_FILA       — 1 sola fila (edge: ¿se considera serie? ¿tabla?)
  12. TOTALES_INLINE — totales en medio de la tabla (no solo al final)
"""
from datetime import datetime, timedelta
from pathlib import Path
import random
import numpy as np
import pandas as pd

random.seed(1337)
np.random.seed(1337)

OUT = Path(__file__).resolve().parent / "stress_god.xlsx"

# ─────────────────────────────────────────────────────────────────────
# 1. VENTAS_MASIVAS — 5000 filas
# ─────────────────────────────────────────────────────────────────────
CIUDADES = ["BOGOTÁ", "MEDELLÍN", "CALI", "BARRANQUILLA", "CARTAGENA",
            "BUCARAMANGA", "PEREIRA", "SANTA MARTA", "MANIZALES", "IBAGUÉ",
            "POPAYÁN", "NEIVA", "VILLAVICENCIO", "PASTO", "ARMENIA"]
PRODUCTOS = [f"PROD-{i:04d}" for i in range(1, 41)]
CANALES = ["Tienda Física", "E-commerce", "Mayorista", "Distribuidor", "Marketplace"]
VENDEDORES = [f"Vendedor {i:03d}" for i in range(1, 21)]

base_date = datetime(2023, 1, 1)
df_ventas = pd.DataFrame({
    "ID_Venta": [f"VTA-{i:07d}" for i in range(1, 5001)],
    "Fecha": [base_date + timedelta(days=random.randint(0, 1000)) for _ in range(5000)],
    "Ciudad": np.random.choice(CIUDADES, 5000),
    "Producto": np.random.choice(PRODUCTOS, 5000),
    "Canal": np.random.choice(CANALES, 5000),
    "Vendedor": np.random.choice(VENDEDORES, 5000),
    "Unidades": np.random.poisson(3.5, 5000) + 1,
    "Precio_Unitario": np.round(np.random.lognormal(11, 0.8, 5000), 0),
    "Descuento_Pct": np.round(np.random.beta(2, 8, 5000), 3),
    "Comision_Pct": np.round(np.random.uniform(0.02, 0.12, 5000), 3),
    "Estado": np.random.choice(["FACTURADO", "PENDIENTE", "CANCELADO"],
                                 5000, p=[0.78, 0.17, 0.05]),
    "Notas": np.random.choice(["", "Cliente VIP", "Requiere seguimiento", "",
                                  "Pago a 30 días", "Confirmar entrega"], 5000),
})

# ─────────────────────────────────────────────────────────────────────
# 2. KPIs_MIXTOS — peor escenario heterogéneo
# ─────────────────────────────────────────────────────────────────────
df_kpis_mixtos = pd.DataFrame({
    "KPI": [
        "Ingresos Brutos", "Tasa Conversión", "Ticket Promedio",
        "Nuevos Clientes", "NPS Score", "Tiempo Entrega (h)",
        "Costo Adquisición", "Retención Mensual", "Quejas", "Margen Bruto",
        "Días Inventario", "ROI Marketing", "Reembolsos", "Productividad",
        "Satisfacción",
    ],
    "Q1_2024": [847_320_500, 0.034, 184_500, 1247, 67, 28.5,
                 45_200, 0.82, 89, 0.38, 21.4, 3.4, 156, 4.2, 0.78],
    "Q2_2024": [912_450_300, 0.041, 192_300, 1389, 71, 24.2,
                 42_800, 0.85, 76, 0.42, 18.9, 4.1, 142, 4.4, 0.81],
    "Q3_2024": [1_023_100_400, 0.038, 198_700, 1521, 73, 22.8,
                 41_500, 0.87, 68, 0.45, 17.3, 4.8, 119, 4.5, 0.83],
    "Q4_2024": [1_245_780_900, 0.052, 215_400, 1854, 76, 19.5,
                 38_200, 0.89, 58, 0.48, 15.1, 5.6, 102, 4.7, 0.86],
    "Objetivo_2025": [5_000_000_000, 0.06, 230_000, 7500, 80, 16.0,
                      35_000, 0.92, 50, 0.52, 12.0, 7.0, 80, 5.0, 0.90],
})

# ─────────────────────────────────────────────────────────────────────
# 3. NaN_BOMB — 80% nulos
# ─────────────────────────────────────────────────────────────────────
n = 200
df_nan = pd.DataFrame({
    "Codigo": [f"X-{i:04d}" if random.random() > 0.6 else None for i in range(n)],
    "Valor": [random.uniform(100, 100000) if random.random() > 0.8 else np.nan
               for _ in range(n)],
    "Categoria": [random.choice(["A", "B", "C"]) if random.random() > 0.85 else None
                    for _ in range(n)],
    "Fecha": [datetime(2024, random.randint(1, 12), random.randint(1, 28))
                if random.random() > 0.75 else pd.NaT for _ in range(n)],
    "Activo": [random.choice([True, False]) if random.random() > 0.7 else None
                 for _ in range(n)],
})

# ─────────────────────────────────────────────────────────────────────
# 4. TEXTO_LARGO — celdas con texto descriptivo enorme
# ─────────────────────────────────────────────────────────────────────
LOREM = ("La auditoría reveló que el proceso de aprobación de comisiones "
         "presenta tiempos de ejecución variables, con un promedio de 12 "
         "días hábiles entre la solicitud inicial y el desembolso. Se "
         "identificaron oportunidades de optimización en las etapas de "
         "validación documental y verificación de cumplimiento normativo. "
         "Los hallazgos sugieren que la implementación de controles "
         "automáticos podría reducir el ciclo en un 40%, además de mejorar "
         "la trazabilidad y reducir los reprocesos. Se recomienda priorizar "
         "la digitalización del flujo y la capacitación del equipo en las "
         "nuevas plataformas tecnológicas disponibles. ")
df_texto = pd.DataFrame({
    "ID_Caso": [f"CASO-{i:03d}" for i in range(1, 31)],
    "Resumen_Ejecutivo": [LOREM * random.randint(1, 3) for _ in range(30)],
    "Tags": [", ".join(random.sample(
        ["urgente","compliance","alto-impacto","recurrente","estructural",
         "operativo","legal","financiero","RH","TI"], random.randint(2, 5)))
        for _ in range(30)],
    "Severidad": np.random.choice(["BAJA","MEDIA","ALTA","CRÍTICA"], 30,
                                    p=[0.4,0.3,0.2,0.1]),
})

# ─────────────────────────────────────────────────────────────────────
# 5. UTF8_CAOS — caracteres extremos
# ─────────────────────────────────────────────────────────────────────
df_utf8 = pd.DataFrame({
    "Nombre": ["José Muñoz", "Mariëlle Süßer", "李明 (Li Ming)",
               "Владимир", "محمد علي", "François Naïveté",
               "Niño 🧒", "Café ☕", "Reciclaje ♻️", "Estrella ⭐"],
    "Comentario": ["¿Qué pasa?", "¡Excelente!", "—Sin observaciones—",
                    "«Citado»", "Test «con citas»", "100% válido",
                    "<script>alert(1)</script>", "SELECT * FROM users;",
                    "Tab\there", "Salto\nde línea"],
    "Valor_Numerico": [1234.56, -987.65, 0.001, 1_000_000, -1e6,
                        np.inf, -np.inf, np.nan, 0, 42],
})

# ─────────────────────────────────────────────────────────────────────
# 6. FECHAS_FORMATOS
# ─────────────────────────────────────────────────────────────────────
df_fechas = pd.DataFrame({
    "Evento": [f"Evento {i}" for i in range(1, 21)],
    "Fecha_Datetime": [datetime(2024, m, 15) for m in range(1, 13)] +
                       [datetime(2025, m, 1) for m in range(1, 9)],
    "Fecha_Excel_Serial": [45292 + i*30 for i in range(20)],  # Excel date numbers
    "Fecha_String_ISO": [f"2024-{m:02d}-15" for m in range(1, 13)] +
                         [f"2025-{m:02d}-01" for m in range(1, 9)],
    "Fecha_String_LATAM": [f"15/{m:02d}/2024" for m in range(1, 13)] +
                           [f"01/{m:02d}/2025" for m in range(1, 9)],
    "Magnitud": np.random.lognormal(10, 1, 20).round(0),
})

# ─────────────────────────────────────────────────────────────────────
# 7. MULTI_CURRENCY
# ─────────────────────────────────────────────────────────────────────
df_curr = pd.DataFrame({
    "Transaccion": [f"TXN-{i:04d}" for i in range(1, 51)],
    "Moneda": np.random.choice(["COP", "USD", "EUR"], 50, p=[0.6, 0.3, 0.1]),
    "Monto_Local": np.round(np.random.lognormal(12, 1.5, 50), 2),
    "Monto_USD": np.round(np.random.lognormal(8, 1.2, 50), 2),
    "Tasa_Cambio": np.round(np.random.uniform(3800, 4500, 50), 2),
    "Fecha": [datetime(2024, random.randint(1, 12), random.randint(1, 28))
                for _ in range(50)],
})

# ─────────────────────────────────────────────────────────────────────
# 8. NEGATIVOS_CEROS — edge cases numéricos
# ─────────────────────────────────────────────────────────────────────
df_neg = pd.DataFrame({
    "Mes": pd.date_range("2024-01-01", periods=12, freq="MS"),
    "Ingresos": [1_000_000, 850_000, 920_000, 0, 1_100_000, -50_000,
                  1_300_000, 1_450_000, 0, 1_280_000, 1_590_000, 1_810_000],
    "Egresos": [-800_000, -780_000, -890_000, -950_000, -1_020_000,
                 -1_050_000, -1_120_000, -1_180_000, -1_200_000,
                 -1_250_000, -1_310_000, -1_400_000],
    "Margen_Pct": [0.20, 0.08, 0.03, -1.0, 0.07, -21.0, 0.14, 0.18,
                    -100, 0.024, 0.176, 0.227],
    "Conteo": [0, 0, 0, 0, 5, 12, 45, 67, 89, 234, 312, 401],
})

# ─────────────────────────────────────────────────────────────────────
# 9. DUPLICADAS — columnas con mismo nombre, headers en fila 3
# ─────────────────────────────────────────────────────────────────────
# Estructura: 2 filas de "metadata" antes del header real
df_dup_raw = pd.DataFrame({
    "Col0": ["INFORME EJECUTIVO", "Periodo: 2024", "Region", "Norte",
              "Sur", "Centro", "Este", "Oeste"],
    "Col1": [None, None, "Ventas", 12500, 8900, 15300, 7800, 9100],
    "Col2": [None, None, "Costos", 8200, 5400, 10100, 4700, 5800],
    "Col3": [None, None, "Ventas", 14000, 9500, 16100, 8100, 10200],  # duplicate name
    "Col4": [None, None, "Costos", 9100, 6000, 10800, 5200, 6500],   # duplicate name
})

# ─────────────────────────────────────────────────────────────────────
# 10. HOJA_VACIA
# ─────────────────────────────────────────────────────────────────────
df_empty = pd.DataFrame()

# ─────────────────────────────────────────────────────────────────────
# 11. UNA_FILA
# ─────────────────────────────────────────────────────────────────────
df_one = pd.DataFrame({
    "Indicador_Critico": ["LIQUIDEZ"],
    "Valor": [2.34],
    "Umbral_Min": [1.5],
    "Estado": ["OK"],
})

# ─────────────────────────────────────────────────────────────────────
# 12. TOTALES_INLINE — total rows interspersed
# ─────────────────────────────────────────────────────────────────────
df_totals = pd.DataFrame({
    "Region": ["BOGOTÁ", "MEDELLÍN", "Subtotal Andina",
                "BARRANQUILLA", "CARTAGENA", "Subtotal Caribe",
                "PASTO", "Subtotal Sur",
                "TOTAL GENERAL"],
    "Ventas_Q1": [340_000, 220_000, 560_000,
                   180_000, 130_000, 310_000,
                   75_000, 75_000, 945_000],
    "Ventas_Q2": [380_000, 245_000, 625_000,
                   195_000, 142_000, 337_000,
                   82_000, 82_000, 1_044_000],
    "Crecimiento": [0.117, 0.114, 0.116,
                     0.083, 0.092, 0.087,
                     0.093, 0.093, 0.105],
})

# ─────────────────────────────────────────────────────────────────────
# Escribir todo
# ─────────────────────────────────────────────────────────────────────
OUT.parent.mkdir(exist_ok=True, parents=True)
with pd.ExcelWriter(OUT, engine="openpyxl") as xw:
    df_ventas.to_excel(xw, sheet_name="VENTAS_MASIVAS", index=False)
    df_kpis_mixtos.to_excel(xw, sheet_name="KPIs_MIXTOS", index=False)
    df_nan.to_excel(xw, sheet_name="NaN_BOMB", index=False)
    df_texto.to_excel(xw, sheet_name="TEXTO_LARGO", index=False)
    df_utf8.to_excel(xw, sheet_name="UTF8_CAOS", index=False)
    df_fechas.to_excel(xw, sheet_name="FECHAS_FORMATOS", index=False)
    df_curr.to_excel(xw, sheet_name="MULTI_CURRENCY", index=False)
    df_neg.to_excel(xw, sheet_name="NEGATIVOS_CEROS", index=False)
    df_dup_raw.to_excel(xw, sheet_name="DUPLICADAS", index=False, header=False)
    df_empty.to_excel(xw, sheet_name="HOJA_VACIA", index=False)
    df_one.to_excel(xw, sheet_name="UNA_FILA", index=False)
    df_totals.to_excel(xw, sheet_name="TOTALES_INLINE", index=False)

print(f"Generado: {OUT}")
print(f"Tamaño: {OUT.stat().st_size:,} bytes")
print(f"Hojas: 12 (VENTAS 5000 filas, KPIs 15, NaN 200, TEXTO 30, UTF8 10,")
print(f"           FECHAS 20, MULTI_CURR 50, NEG 12, DUP 8, EMPTY 0,")
print(f"           UNA 1, TOTALS 9)")
