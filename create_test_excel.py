import pandas as pd

with pd.ExcelWriter('test_data.xlsx') as writer:
    # Sheet 1: Chart Data
    df_sales = pd.DataFrame({
        'Mes': ['Enero', 'Febrero', 'Marzo', 'Abril'],
        'Ventas': [120, 150, 180, 210]
    })
    df_sales.to_excel(writer, sheet_name='Ventas', index=False)

    # Sheet 2: Table Data
    df_team = pd.DataFrame({
        'ID': [1, 2, 3],
        'Nombre': ['Carlos', 'Ana', 'Luis'],
        'Rol': ['Líder', 'Ingeniera', 'Diseñador'],
        'Estado': ['Activo', 'Activo', 'Pendiente']
    })
    df_team.to_excel(writer, sheet_name='Equipo', index=False)

    # Sheet 3: Text Data
    df_notes = pd.DataFrame({
        'Observaciones': [
            'El proyecto avanza según lo planeado.',
            'Se requiere revisión de presupuesto.',
            'Próximo hito: 15 de Abril.'
        ]
    })
    df_notes.to_excel(writer, sheet_name='Notas', index=False)

print("test_data.xlsx creado con éxito.")
