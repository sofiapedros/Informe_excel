'''
Vamos a crear un excel a partir de los ingredientes, calidad del dato
y los pedidos
'''

import pandas as pd
import xlsxwriter
from openpyxl.chart import BarChart

def crear_excel():
    # Leemos los ficheros sobre los que vamos a sacar información
    # y quitamos las columnas que no vamos a usar
    df_pedidos = pd.read_csv('order_details_limpio.csv',sep=",",encoding="LATIN_1")
    df_pedidos = df_pedidos.drop(['order_details_id'], axis=1)
    df_pedidos = df_pedidos.drop(['order_id'],axis = 1)
    df_pedidos = df_pedidos.drop('Unnamed: 0',axis = 1)

    # Organizo los pedidos por pizza y sumo la cantidad que se pide de cada
    # una
    df_pedidos = df_pedidos.groupby('pizza_id').sum().reset_index()

    # Los ordeno por cantidad pedida para poner los que más se venden
    # antes en el reporte
    df_pedidos_ord = df_pedidos.sort_values('quantity',ascending=False)

    # Leemos los ingredientes y quitamos columnas innecesarios
    ingredients = pd.read_csv('final.csv',sep=",",encoding="LATIN_1")
    ingredients = ingredients.drop('Unnamed: 0',axis = 1)

    # Informe sobre las pizzas y su precio
    pizzas = pd.read_csv('pizzas.csv',sep=",",encoding="LATIN_1")
    pizzas = pizzas.drop('pizza_type_id',axis=1)
    pizzas = pizzas.drop('size',axis=1)
    pizzas = pizzas.sort_values('price',ascending=False)

    # Escribimos en un excel
    with pd.ExcelWriter('Informe.xlsx',engine='xlsxwriter') as writer:

        # Escribimos datos en reporte
        df_pedidos_ord.to_excel(writer,sheet_name="reporte",index=False)
        pizzas.to_excel(writer,sheet_name="reporte",index=False,startcol = 12)
        hoja = 'reporte'

        # Pintamos gráficas sobre las pizzas pedidas en un año
        # Pizzas más vendidas:
        chart1 = writer.book.add_chart({'type':'bar'})
        chart1.add_series({
            'categories':f'={hoja}!$A$2:$A$6',
            'values':f'={hoja}!$B$2:$B$6',
            })

        # Pizzas menos vendidas
        chart1.set_title({'name':'Pizzas more sold in a year'})
        writer.sheets['reporte'].insert_chart('D2',chart1)

        chart1 = writer.book.add_chart({'type':'bar'})
        chart1.add_series({
            'categories':f'={hoja}!$A$87:$A$92',
            'values':f'={hoja}!$B$87:$B$92',
            })
        chart1.set_title({'name':'Pizzas less sold in a year'})
        writer.sheets['reporte'].insert_chart('D20',chart1)

        # Pizzas más baratas
        chart1 = writer.book.add_chart({'type':'bar'})
        chart1.add_series({
            'categories':f'={hoja}!$M$90:$M$95',
            'values':f'={hoja}!$N$90:$N$95',
            })
        chart1.set_title({'name':'Cheapest pizzas'})
        writer.sheets['reporte'].insert_chart('P2',chart1)

        # Pizzas más caras
        chart1 = writer.book.add_chart({'type':'bar'})
        chart1.add_series({
            'categories':f'={hoja}!$M$2:$M$7',
            'values':f'={hoja}!$N$2:$N$7',
            })
        chart1.set_title({'name':'Most expensive pizzas'})
        writer.sheets['reporte'].insert_chart('P20',chart1)

        # Escribimos los pedidos y los ingredientes en una hoja
        df_pedidos.to_excel(writer, sheet_name='Orders',engine = 'xlsxwriter')
        ingredients.to_excel(writer, sheet_name='Ingredients',engine = 'xlsxwriter')
    
        

