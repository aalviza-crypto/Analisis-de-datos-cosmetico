# Análisis de NSO de cosméticos (INVIMA)

`analizar_nso_cosmeticos.py` descarga el dataset abierto **ui32-p9f2** de datos.gov.co
("Información sobre RS y NSO vigentes y vencidas"), lo filtra al grupo COSMÉTICOS,
cuenta NSO únicas y las clasifica por **forma cosmética** y **uso/indicación**
a partir del nombre comercial del producto.

```bash
pip install pandas requests openpyxl
python analizar_nso_cosmeticos.py                  # descarga directa desde la API
python analizar_nso_cosmeticos.py --solo-vigentes  # solo NSO vigentes
python analizar_nso_cosmeticos.py --csv archivo.csv  # si descargaste el CSV a mano
```

Resultado: `resultados_nso_cosmeticos.xlsx` con las hojas Formas, Usos, Forma x Uso,
Titulares, Estado y Detalle NSO.

**Limitación:** el dataset no tiene un campo de "forma cosmética" ni de "uso"; ambos se
infieren por palabras clave en el nombre del producto. La fila "Sin clasificar" indica
la cobertura del clasificador; las reglas se pueden ampliar en las listas `FORMAS` y `USOS`.
