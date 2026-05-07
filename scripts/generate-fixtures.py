"""Generate deterministic test fixtures for the Excel→PPTX pipeline."""
from pathlib import Path
import random
import pandas as pd
import numpy as np

OUT = Path(__file__).resolve().parent.parent / "__tests__" / "fixtures"
OUT.mkdir(parents=True, exist_ok=True)

random.seed(42)
np.random.seed(42)

def ventas_simple():
    dates = pd.date_range("2024-01-01", periods=100, freq="D")
    cities = np.random.choice(["Bogotá", "Medellín", "Cali", "Cartagena", "Barranquilla"],
                               size=100, p=[0.5, 0.25, 0.15, 0.07, 0.03])
    totals = np.random.randint(50_000, 900_000, size=100)
    df = pd.DataFrame({
        "Fecha": dates,
        "Total": totals,
        "Ciudad": cities,
        "Vendedor": np.random.choice(["Ana","Luis","Marta","Pedro"], 100),
        "Producto": np.random.choice(["A","B","C"], 100),
    })
    df.to_excel(OUT / "ventas_simple.xlsx", sheet_name="Ventas", index=False)

def casi_vacio():
    df = pd.DataFrame({
        "Col1": [None, None, "X"],
        "Col2": [None, "Y", None],
        "Col3": [None, None, None],
    })
    df.to_excel(OUT / "casi_vacio.xlsx", sheet_name="Hoja1", index=False)

def enorme():
    n = 50_000
    df = pd.DataFrame({
        "Fecha": pd.date_range("2020-01-01", periods=n, freq="h"),
        "Valor": np.random.randn(n) * 1000 + 50_000,
        "Categoria": np.random.choice([f"C{i}" for i in range(20)], n),
    })
    df.to_excel(OUT / "enorme.xlsx", sheet_name="Datos", index=False)

def corrupto():
    (OUT / "corrupto.xlsx").write_bytes(b"PK\x03\x04 not a real xlsx file")

def dominio_raro():
    df = pd.DataFrame({
        "Empleado": [f"Emp_{i:03d}" for i in range(60)],
        "Departamento": np.random.choice(["RRHH","IT","Logistica","Salud","Educacion"], 60),
        "Antiguedad_anios": np.random.randint(1, 25, 60),
        "Salario": np.random.randint(1_500_000, 12_000_000, 60),
        "Genero": np.random.choice(["F","M","Otro"], 60, p=[0.45,0.5,0.05]),
    })
    df.to_excel(OUT / "dominio_raro.xlsx", sheet_name="Plantilla", index=False)

if __name__ == "__main__":
    ventas_simple()
    casi_vacio()
    enorme()
    corrupto()
    dominio_raro()
    print("Fixtures generated in", OUT)
