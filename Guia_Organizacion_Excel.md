# Guía de Organización de Excel para PowerPoint

Para que tus presentaciones se generen correctamente y sin desorden, ahora puedes usar **etiquetas especiales** en las primeras filas de cada hoja de tu Excel.

## Etiquetas Disponibles

Puedes usar estas etiquetas en la **primera columna** de las primeras filas de cada hoja:

1.  **`TITLE: [Texto]`**: Define el título principal de la diapositiva.
    *   *Ejemplo:* `TITLE: Resumen de Comisiones 2024`
2.  **`SUBTITLE: [Texto]`**: Agrega un subtítulo o descripción debajo del título.
    *   *Ejemplo:* `SUBTITLE: Datos correspondientes al primer trimestre`
3.  **`TYPE: [Categoría]`**: Indica al programa cómo mostrar los datos. Las opciones son:
    *   `TABLE`: Muestra los datos en una tabla profesional (ideal para listas largas).
    *   `CHART`: Crea un gráfico de barras (ideal si tienes 2 columnas: una de texto y una de números).
    *   `TEXT`: Muestra los datos como una lista de puntos (bullets).

---

## Ejemplo de Estructura en una Hoja

| Celda A | Celda B | Descripción |
| :--- | :--- | :--- |
| **TITLE: Comisiones por Mes** | | (Opcional) Título de la diapositiva |
| **TYPE: CHART** | | Forza a que sea un gráfico |
| **SUBTITLE: Reporte Mensual** | | (Opcional) Texto secundario |
| **Mes** | **Valor** | **<-- Los encabezados de tabla empiezan aquí** |
| Enero | 1500 | Datos... |
| Febrero | 2200 | Datos... |

---

## Consejos para tu archivo "Comisiones V1.xlsx"

1.  **Una diapositiva por pestaña**: El programa creará una diapositiva por cada hoja que tenga datos en tu Excel.
2.  **Limpieza**: Asegúrate de que no haya filas vacías entre las etiquetas (`TITLE:`, `TYPE:`) y tus encabezados de datos.
3.  **Gráficos**: Si quieres un gráfico, asegúrate de tener solo 2 columnas de datos y usar la etiqueta `TYPE: CHART`.
4.  **Tablas Largas**: Si tu tabla tiene muchas filas, el programa las dividirá automáticamente en varias diapositivas manteniendo el título.

---

### ¿Cómo proceder ahora?
1. Abre tu archivo **Comisiones V1.xlsx**.
2. En cada hoja, inserta un par de filas al principio.
3. Escribe las etiquetas `TITLE:`, `TYPE:` y `SUBTITLE:` según necesites.
4. Guarda y sube el archivo al generador.
