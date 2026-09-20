# Laboratorio 11 — Introducción a la visualización de datos

**Materia:** Programación Para la Analítica de Datos

**Autor:** David Ospino

**Herramientas:** Excel y Python (Google Colab)

## Contenido del repositorio

| Archivo | Descripción |
|---|---|
| `Bike_Sales_Outlier_Lab.xlsx` | Dataset de ventas de bicicletas (diciembre 2021) usado en el Ejercicio 1 |
| `Lab11_Visualizacion_Datos.ipynb` | Notebook de Colab con el desarrollo completo en Python |
| `Excel Tabla Dinamica Outliers.png` | Captura de Excel: tabla dinámica, orden descendente, funciones GRANDE/PEQUEÑO y gráficos de dispersión |
| `Excel Detalle 19dic Pedido261765.png` | Captura de Excel: detalle de los pedidos del 19 de diciembre, con el pedido 261765 identificado |

---

## Ejercicio de práctica 1 — Interpretación de visualizaciones con valores atípicos

**Objetivo:** examinar el dataset de ventas de bicicletas en busca de valores atípicos, usando tabla dinámica, ordenamiento, gráfico de dispersión y las funciones GRANDE/PEQUEÑO.

### Paso 2 — Tabla dinámica

Cree una tabla dinámica con los campos **Date** y **Order_Quantity**, obteniendo la suma de unidades pedidas por día. Total general del mes: **197 unidades**.

### Paso 3 — Ordenar los datos

Al ordenar de mayor a menor la suma de `Order_Quantity`:

- **¿En qué fecha de diciembre se registró la mayor cantidad de ventas?** → **19/12/2021**
- **¿Cuál fue la cantidad de ventas?** → **43 unidades**
- **¿Qué entrada contribuye más a la suma en la tabla dinámica? ¿Qué número de pedido es el más responsable del valor atípico?** → **Sales_Order #261765**, con 11 unidades y $25.520 de ingreso (el mayor pedido individual de ese día).

![Detalle de los pedidos del 19 de diciembre — pedido 261765 identificado](<Excel Detalle 19dic Pedido261765.png>)

### Paso 4 — Gráfico de dispersión

El gráfico de dispersión de la suma diaria de `Order_Quantity` muestra al 19 de diciembre claramente separado del resto de los puntos: mientras los demás días se mueven entre 1 y 19 unidades, ese día alcanza 43. Esto lo confirma como valor atípico.

### Paso 5 — Funciones GRANDE y PEQUEÑO

| Excel | Resultado |
|---|---|
| `=GRANDE($E$4:$E$27; 1)` | 43 |
| `=GRANDE($E$4:$E$27; FILA($1:5))` | 43, 19, 13, 12, 11 |
| `=PEQUEÑO($E$4:$E$27; FILA($1:6))` | 1, 3, 3, 3, 4, 4 |

**¿Qué función devolvería los 6 valores más bajos?** → `=PEQUEÑO($E$4:$E$27; FILA($1:6))` .

![Tabla dinámica, funciones GRANDE/PEQUEÑO y gráficos de dispersión en Excel](<Excel Tabla Dinamica Outliers.png>)

### Tratamiento del valor atípico

Probe las dos alternativas propuestas por la guía:

1. **Eliminarlo:** quite la fila 72 del dataset (pedido #261765), trabajando siempre sobre una copia para poder investigar después su origen.
2. **Normalizarlo:** ajuste la cantidad del 19 de diciembre de 43 a **20**, justo por encima del segundo valor más alto (19).

### Gráfico de dispersión: Cantidad de órdenes vs. Ingresos

Genere además el gráfico solicitado en la sección de desarrollo en Python (`Order_Quantity` vs. `Revenue`).

**Interpretación:**

1. **Relación entre cantidad de órdenes e ingresos:** en general, a mayor cantidad de órdenes los ingresos tienden a subir, pero el crecimiento no es proporcional: hay pedidos pequeños que generan ingresos muy altos.
2. **Variabilidad:** para cantidades pequeñas (1 a 4 unidades) los ingresos se dispersan bastante, por la diferencia de precio entre productos.
3. **Valores atípicos:** el punto de la esquina superior derecha combina una cantidad de órdenes alta con un ingreso excepcional. Corresponde precisamente al pedido #261765 del 19 de diciembre, el mismo identificado en el Paso 3, y distorsiona la escala del gráfico.
4. **Análisis adicional:** cruzar estos datos con `Customer_Age`, `Country` o `Product_Category` ayudaría a entender por qué algunas órdenes generan más ingreso que otras.

### Preguntas de reflexión

**¿Qué factores determinan si un valor atípico debe o no considerarse en el análisis final?**

- **Origen del dato:** si proviene de un error de digitación o de medición, debe corregirse o eliminarse; si es una transacción real y verificable, debe conservarse.
- **Tamaño del conjunto de datos:** en un dataset grande, quitar un par de atípicos casi no afecta el resultado; en uno pequeño como este (24 días) sí puede cambiar las conclusiones.
- **Objetivo del análisis:** si se busca el comportamiento típico del negocio, el atípico estorba; si se busca detectar ventas excepcionales o fraude, el atípico es la información relevante.
- **Impacto sobre los estadísticos:** la media y la desviación estándar se ven muy afectadas por los atípicos; la mediana no.
- **Trazabilidad:** siempre trabajar sobre una copia y documentar qué se eliminó o ajustó, para poder investigar después la causa.

---

## Ejercicio de práctica 2 — Escenarios de visualización

### 1. Comparación de calificaciones de estudiantes en diferentes asignaturas

Gráfico de barras con `asignaturas` vs. `calificaciones`.

**Interpretación:** Arte (92) y Ciencias (90) presentan los mejores resultados, mientras que Historia (78) queda claramente por debajo del resto. La diferencia entre la materia más alta y la más baja es de 14 puntos, lo que sugiere revisar la metodología o la carga de contenido en Historia.

### 2. Comparación del tiempo de carga de diferentes páginas web

Gráfico de barras horizontales con `sitios` vs. `tiempos`.

**Interpretación:** el Sitio C es el más rápido (0.9 s), mientras que el Sitio D es el más lento (3.0 s) — más de tres veces la demora del Sitio C. Considerando que se recomienda un tiempo de carga por debajo de 2 segundos, los sitios B y D requieren optimización.

### 3. Relación entre horas de estudio y rendimiento académico

Gráfico de dispersión con línea de tendencia.

**Interpretación:** correlación positiva muy fuerte (r ≈ 0.99) entre horas de estudio y calificaciones. Sin embargo, la relación no es perfectamente lineal: entre 5 y 30 horas cada bloque de estudio suma 5 puntos, mientras que a partir de las 30 horas la ganancia se reduce a 2-3 puntos — un rendimiento decreciente.

### 4. Distribución de salarios en tres departamentos distintos

Diagrama de cajas (boxplot) por departamento.

**Interpretación:** IT concentra la mediana más alta (6600) y el rango más amplio, es decir, mayor dispersión entre salarios de entrada y senior. Recursos Humanos tiene la mediana más baja (4100) y Administración se ubica en un punto intermedio (5200) con la distribución más compacta. Ningún departamento presenta valores atípicos.

---

## Conclusiones del laboratorio

- Los valores atípicos pueden detectarse por varios caminos: ordenando los datos, con un gráfico de dispersión, o con funciones como GRANDE/PEQUEÑO (`nlargest`/`nsmallest` en pandas). La visualización suele ser el método más rápido para conjuntos grandes.
- En el dataset de ventas de bicicletas, el 19 de diciembre de 2021 destaca con 43 unidades frente a un máximo de 19 en los demás días.
- Cada tipo de gráfico responde a una pregunta distinta: barras para comparar categorías, barras horizontales cuando las etiquetas son largas, dispersión para relaciones entre dos variables numéricas y cajas para distribuciones y detección de atípicos.
- Antes de eliminar un valor atípico hay que investigar su origen: puede ser un error o puede ser el dato más valioso del conjunto.
