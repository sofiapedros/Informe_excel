# Maven_pizzas_excel

El repositorio contiene los siguientes ficheros:
- etl.py (programa a ejecutar)
- analisis_de_los_datos.py (programa a ejecutar)
- Analisis_datos.xml
- requirements.txt
- Datos limpios.csv
- Fechas_limpias.csv
- crear_excel.py (librería)
- final.csv
- Informe.xlsx
- data_dictionary.csv
- order_details.csv
- orders.csv
- pizza_types.csv
- pizzas.csv


### etl.py
"etl.py": Programa que, dado un conjunto de datos de una pizzería como ficheros csv, calcula los ingredientes que debe comprar esa pizzería en una semana, suponiendo que el número de ventas de la pizzería en una semana sea regular (utiliza la media de todos los pedidos). Además, supone que la cantidad de ingredientes de una pizza es 1 ración para la pizza pequeña, 2 para la mediana, 3 para la grande y 4 para tamaños superiores. Guarda la lista de los ingredientes a comprar con el número de raciones necesarias en un csv llamado "final.csv". Realiza esta funcionalidad como una etl: primero extrae los datos, después los tranforma , y por último obtiene un resultado en forma de un csv (final.csv) y crea un reporte en excel con todos los datos (Informe.xlsx). 
Los csv que recibe esta etl están corruptos, y, antes de realizar la predicción de compra para la semana siguiente, limpia los csv. El fichero 'orders.csv', del que se han modificado fechas y horas, al limpiarlo, se obtiene "Fechas_limpias.csv". Por otra parte, en el fichero de order_daetails.csv, al limpiarlo, se obtiene "Datos limpios.csv". En la limpieza de este último fichero, se ha partido de la suposición de que las pizzas que se piden se distribuyen de forma homogénea. 

### analisis_de_los_datos.py
"analisis_de_los_datos.py": programa que guarda en un fichero .xml (Analisis_datos.xml) un análisis de los datos proporcionados por la pizzería. Imprime número de NaN y nulls totales y por columna así como el
tipo de datos de cada columna.

### Analisis_datos.xml
Fichero de texto de salida al ejecutar analisis_de_los_datos.py

### requirements.txt
Librerías utilizadas

### crear_excel.py
Librería creada llamada en etl.py que crea un informe en excel a partir de todos los ficheros que tiene disponibles. Este libro de excel tiene tres hojas. La primera contiene una comparativa general de las pizzas más vendidas y las pizzas más baratas junto con las pizzas menos vendidas y las más caras a partir de gráficas realizadas con los datos. En la segunda aparecen los datos de las pizzas y las ventas realizadas por cada tipo. En la última hoja aparecen los datos de los ingredientes que se deben comprar para la próxima semana. 

### final.csv
Csv obtenido como salida al realizar la etl

### Infrome.xlsx
Informe de excel generado con todos los datos

### Fechas_limpias.csv
Resultado de limpiar el fichero 'orders.csv'

### Datos limpios.csv
Resultado de limpiar el fichero 'order_details.csv'

### Resto de ficheros csv
Ficheros csv sobre los que se han extraido los datos para hacer el análisis
