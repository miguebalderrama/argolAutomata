import tkinter as tk
from tkinter import messagebox, ttk
from wsgiref import headers
import pandas as pd
from pandas import ExcelWriter
import xlrd 
import glob


def cleandata():
    folder_path = '../data/comprobantes/comprob'
    file_list = glob.glob(folder_path + "/*.xlsx")
    print(len(file_list))

    main_dataframe = pd.DataFrame(pd.read_excel(file_list[0],skiprows=1))
    main_dataframe.rename(columns={'Número Desde':'Número de comprobante',
 'Nro. Doc. Receptor':'CUIT del comprador',
 'Denominación Receptor':'Comprador (Apellido y Nombre / Razón Social',
 'Imp. Total':'MONTO TOTAL DEL COMPROBANTE EN PESOS',
 'Moneda':'Moneda de Emisión del Comprobante',
 'Tipo Cambio':'Tipo de Cambio del Comprobante',
 'Cód. Autorización':'Código de autorización',
 'IVA':'IVA (en Pesos por concepto facturado)',
 },
               inplace=True)
    main_dataframe.drop(['Número Hasta','Tipo Doc. Receptor'], axis = 'columns', inplace=True)
    for i in range(0,len(file_list)):
     data = pd.read_excel(file_list[i], skiprows=1)
     data.drop(['Número Hasta','Tipo Doc. Receptor'], axis = 'columns', inplace=True)

     data.rename(columns={'Número Desde':'Número de comprobante',
    'Nro. Doc. Receptor':'CUIT del comprador',
    'Denominación Receptor':'Comprador (Apellido y Nombre / Razón Social',
    'Imp. Total':'MONTO TOTAL DEL COMPROBANTE EN PESOS',
    'Moneda':'Moneda de Emisión del Comprobante',
    'Tipo Cambio':'Tipo de Cambio del Comprobante',
    'Cód. Autorización':'Código de autorización',
    'IVA':'IVA (en Pesos por concepto facturado)'},
                inplace=True)
     df = pd.DataFrame(data)
     main_dataframe = pd.concat([main_dataframe,df])
    main_dataframe.insert(0,'Grupo','') 
#header = main_dataframe.head(1)
    main_dataframe.replace({"$": "PESOS", "USD": "DÓLAR"}, inplace=True)
    main_dataframe['Tipo'] = main_dataframe.Tipo.apply(
    lambda x: x.split(' ')[0]
    )
    main_dataframe = main_dataframe.reindex(columns=['Grupo','Fecha','Tipo','Punto de Venta',
'Número de comprobante','CUIT del comprador','Comprador (Apellido y Nombre / Razón Social',
'MONTO TOTAL DEL COMPROBANTE EN PESOS','Moneda de Emisión del Comprobante','Tipo de Cambio del Comprobante',
'Código de autorización','IVA (en Pesos por concepto facturado)'])
    main_dataframe.insert(10,'Imputación (solo en N/C y N/D) Tipo, punto de venta y numero de comprobante relacionado.',"", allow_duplicates=False)
    main_dataframe.insert(11,'Tipo de Autorización',"CAE", allow_duplicates=False)
    main_dataframe.insert(13,'Código Producto o de servicio',"", allow_duplicates=False)
    main_dataframe.insert(14,'N° de Serie',"", allow_duplicates=False)
    main_dataframe.insert(15,'NCM',"", allow_duplicates=False)
    main_dataframe.insert(16,'Descripción',"", allow_duplicates=False)
    main_dataframe.insert(17,'cant.',"", allow_duplicates=False)
    main_dataframe.insert(18,'VALOR UNITARIO NETO EN PESOS',"", allow_duplicates=False)
    main_dataframe.insert(19,' VALOR NETO EN PESOS TOTAL POR CONCEPTO (Valor Unitario x Cantidad) (campo de cálculo automático)',"", allow_duplicates=False)
#main_dataframe.insert(20,'IVA (en Pesos por concepto facturado)',"", allow_duplicates=False)
    main_dataframe.insert(21,'PERCEPCION I.V.A. (en Pesos por concepto facturado)',"", allow_duplicates=False)
    main_dataframe.insert(22,"PERCEPCION ISIB (en Pesos por concepto facturado)","", allow_duplicates=False)
    main_dataframe.insert(23,'OTROS IMPUESTOS (en Pesos por concepto facturado)',"", allow_duplicates=False)
    main_dataframe.insert(24,'TOTAL (en Pesos por concepto facturado)(campo de cálculo automático)',"", allow_duplicates=False)



    print(main_dataframe)
#print(header)
    writer = ExcelWriter('Archivo.xlsx')
    main_dataframe.to_excel(writer, 'Hoja de datos', index=False)
    writer.close()
    messagebox.showinfo(message="¡Haz generado la tabla de comprobantes!", title="Informacion")

root = tk.Tk()
root.config(width=300, height=200)
root.title("Automata Argol")
boton = ttk.Button(text="Generar Tabla", command=cleandata)
botonSalir = ttk.Button(text = "Salir", command=root.destroy)
boton.place(x=60, y=50)
botonSalir.place(x=200, y=50)
root.mainloop()