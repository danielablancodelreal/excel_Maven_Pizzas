import xlsxwriter
import pandas as pd

def crear_excel(nombre_excel):

    df_order_details, df_pedido_semanal_ingredientes, df_pedidos = extract()

    #Usamos la función de pandas ExcelWriter() para escribir los df en Excel

    with pd.ExcelWriter(nombre_excel,engine='xlsxwriter') as writer:
        
        #Insertamos los pedidos anuales usando una función de pandas

        df_pedidos.to_excel(writer,sheet_name='Hoja_1_Reporte',index=False)
        sheet_name='Hoja_1_Reporte'

        #Creamos un gráfico con Excel

        barras = writer.book.add_chart({'type':'bar'})
        barras.add_series({'categories':f'={sheet_name}!$B$2:$B$6',
            'values':f'={sheet_name}!$C$2:$C$6'})
        barras.set_title({'name':'Pizzas más vendidas al año'})
        writer.sheets['Hoja_1_Reporte'].insert_chart('E16',barras)

        #Insertamos un gráfico a partir de una imagen

        writer.sheets['Hoja_1_Reporte'].write('E2', 'Pedidos por hora del día:')
        writer.sheets['Hoja_1_Reporte'].insert_image('E3', 'pedidos_horas.png',{'x_scale': 0.7, 'y_scale': 0.7})
    
        #Creamos las hojas 2 y 3 con las tablas de ingredientes y pedidos

        df_pedido_semanal_ingredientes.to_excel(writer,sheet_name='Hoja_2_Ingredientes',index=False)
        df_pedidos.to_excel(writer,sheet_name='Hoja_3_Pedidos',index=False)


def extract():
    df_order_details = pd.read_csv('order_details.csv',sep=",",encoding="LATIN_1")
    df_pedido_semanal_ingredientes = pd.read_csv('pedido_semanal_ingredientes.csv',sep=",",encoding="LATIN_1")
    df_pedidos = pd.read_csv('pedidos.csv',sep=",",encoding="LATIN_1")

    return df_order_details, df_pedido_semanal_ingredientes, df_pedidos

if __name__ == "__main__":
    
    crear_excel('pizzas_excel.xlsx')

