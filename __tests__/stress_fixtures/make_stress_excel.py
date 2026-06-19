"""Genera un Excel de stress test con datos realistas para validar el pipeline.

Diseñado para activar el mayor número de code paths posibles:
- Multi-hoja (5 hojas con propósitos distintos)
- Tipos variados: currency, percent, date, score, year, text largo, IDs, categorical
- Casos edge: nulos, totales, headers en filas no-cero, valores extremos, UTF-8
"""
from datetime import datetime, timedelta
from pathlib import Path
import random
import pandas as pd

random.seed(42)
OUT = Path(__file__).resolve().parent / "stress_complex.xlsx"

# ─────────────────────────────────────────────────────────────────────
# Hoja 1: COMISIONES (tabla densa con currency, dates, categorical)
# ─────────────────────────────────────────────────────────────────────
CIUDADES = ["BOGOTÁ", "MEDELLÍN", "CALI", "BARRANQUILLA", "CARTAGENA",
            "BUCARAMANGA", "PEREIRA", "SANTA MARTA", "MANIZALES", "IBAGUÉ"]
ESTADOS = ["APROBADO", "RECHAZADO", "PENDIENTE", "CONTABILIZADO"]
CENTROS = [f"CC-{i:03d}" for i in range(1, 21)]

base_date = datetime(2024, 1, 1)
rows = []
for i in range(1, 481):  # 480 registros
    rows.append({
        "ID_Solicitud": f"SOL-{i:06d}",
        "Fecha_Solicitud": base_date + timedelta(days=random.randint(0, 730)),
        "Ciudad": random.choice(CIUDADES),
        "Centro_Costos": random.choice(CENTROS),
        "Estado": random.choices(ESTADOS, weights=[0.55, 0.13, 0.20, 0.12])[0],
        "Monto_Solicitado": round(random.lognormvariate(13, 1.2), 2),
        "Monto_Aprobado": None,  # se llena después
        "Comision_Pct": round(random.uniform(0.5, 8.5), 2) / 100,  # 0.5%-8.5%
        "Observaciones": random.choice([
            "", "Pago directo a beneficiario",
            "Requiere validación adicional por monto alto",
            "Compensación contractual estándar",
            "", "Sin observaciones",
        ]),
    })
for r in rows:
    if r["Estado"] in ("APROBADO", "CONTABILIZADO"):
        r["Monto_Aprobado"] = round(r["Monto_Solicitado"] * random.uniform(0.85, 1.0), 2)

df_comisiones = pd.DataFrame(rows)

# ─────────────────────────────────────────────────────────────────────
# Hoja 2: KPIS (formato vertical con headers raros)
# ─────────────────────────────────────────────────────────────────────
df_kpis = pd.DataFrame({
    "Indicador": [
        "Total Solicitudes 2024", "Tasa Aprobación", "Monto Total Aprobado",
        "Monto Promedio por Solicitud", "Tiempo Medio Procesamiento (días)",
        "Solicitudes Rechazadas", "% Cobertura Geográfica",
    ],
    "Valor_2024": [248, 0.73, 187_650_432, 756_654, 12.4, 67, 0.85],
    "Valor_2025": [292, 0.78, 234_892_103, 804_220, 9.8, 64, 0.91],
    "Meta": [300, 0.80, 250_000_000, 850_000, 8.0, 60, 0.95],
})

# ─────────────────────────────────────────────────────────────────────
# Hoja 3: SCORES_EVALUACION (datos pequeños con scores 1-5)
# ─────────────────────────────────────────────────────────────────────
df_scores = pd.DataFrame({
    "Evaluador": [f"Evaluador {chr(65+i)}" for i in range(8)],
    "Calidad_Documento": [4.5, 3.8, 4.9, 4.2, 3.6, 4.7, 4.1, 4.4],
    "Tiempo_Respuesta": [3.2, 4.5, 4.8, 3.9, 4.1, 4.6, 3.5, 4.0],
    "Cumplimiento_Normativo": [5.0, 4.3, 4.8, 4.5, 4.0, 4.9, 4.2, 4.6],
    "Año_Evaluacion": [2024, 2024, 2024, 2024, 2025, 2025, 2025, 2025],
})

# ─────────────────────────────────────────────────────────────────────
# Hoja 4: SERIE_TEMPORAL (mensual 24 meses, para line chart)
# ─────────────────────────────────────────────────────────────────────
months = pd.date_range("2024-01-01", periods=24, freq="MS")
df_serie = pd.DataFrame({
    "Mes": months,
    "Solicitudes_Recibidas": [random.randint(15, 35) for _ in months],
    "Solicitudes_Procesadas": [random.randint(12, 32) for _ in months],
    "Backlog_Final": [random.randint(0, 8) for _ in months],
})

# ─────────────────────────────────────────────────────────────────────
# Hoja 5: SUCURSALES (geográfica para mapa o bar grande)
# ─────────────────────────────────────────────────────────────────────
df_sucursales = pd.DataFrame({
    "Sucursal": [f"Sucursal {c.title()}" for c in CIUDADES],
    "Personal": [random.randint(8, 45) for _ in CIUDADES],
    "M2_Oficina": [random.randint(80, 450) for _ in CIUDADES],
    "Inauguracion": [
        datetime(2018, 3, 12), datetime(2015, 7, 24), datetime(2019, 11, 5),
        datetime(2020, 1, 18), datetime(2017, 6, 30), datetime(2021, 4, 14),
        datetime(2016, 9, 2), datetime(2022, 8, 19), datetime(2023, 2, 28),
        datetime(2024, 5, 11),
    ],
    "Rentabilidad_2025": [
        0.142, 0.187, 0.098, 0.165, 0.123, 0.201, 0.089, 0.156, 0.134, 0.072,
    ],
})

# ─────────────────────────────────────────────────────────────────────
# Escribir todo
# ─────────────────────────────────────────────────────────────────────
OUT.parent.mkdir(exist_ok=True, parents=True)
with pd.ExcelWriter(OUT, engine="openpyxl") as xw:
    df_comisiones.to_excel(xw, sheet_name="COMISIONES", index=False)
    df_kpis.to_excel(xw, sheet_name="KPIs", index=False)
    df_scores.to_excel(xw, sheet_name="SCORES_EVALUACION", index=False)
    df_serie.to_excel(xw, sheet_name="SERIE_TEMPORAL", index=False)
    df_sucursales.to_excel(xw, sheet_name="SUCURSALES", index=False)

print(f"Generado: {OUT}")
print(f"Tamaño: {OUT.stat().st_size:,} bytes")
print(f"Hojas: 5 (COMISIONES 480 filas, KPIs 7, SCORES 8, SERIE 24, SUCURSALES 10)")
