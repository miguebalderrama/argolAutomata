import tkinter as tk
from tkinter import messagebox, ttk
from wsgiref import headers
import pandas as pd
from pandas import ExcelWriter
import xlrd 
import glob

def reldata():
    # folder_path = '../data/relaborales' Aca sube dos niveles
    folder_path = '../data/relaborales'
    file_list = glob.glob(folder_path + "/*.txt")
    print(len(file_list))
    main_dataframe = pd.DataFrame(pd.read_fwf(file_list[0], widths=[12, 60, 160,100,59,60]))
   
    for i in range(1,len(file_list)):
      data = pd.read_fwf(file_list[i], widths=[12, 60, 160,100,59,60]) #, skiprows=1)        
      df = pd.DataFrame(data)
      main_dataframe = pd.concat([main_dataframe,df], axis = 0)

    name = main_dataframe["APELLIDO Y NOMBRE"].str.split(expand=True)
    nrocolum= len(name.columns)
    print(name.columns.values)
    if nrocolum == 3:
    #name.set_axis(['last_name0', 'last_name1', 'first_0', 'first_1', 'first_2'], axis=1)
    #name.columns = ['last_name0', 'last_name1', 'first_0', 'first_1', 'first_2']
        name['APELLIDO'] = name[0] 
        name['NOMBRE'] = name[1].str.cat(name[2], sep = " ", na_rep = "")
        name.drop([0,1,2], axis = 'columns', inplace=True)
    #name['NOMBRE'] = name.NOMBRE.str.cat(name[1], sep = " ", na_rep = "")
    #main_dataframe = pd.concat([main_dataframe, name['NOMBRE'],name['APELLIDO']  ], axis=1)
    #print(name)
    #main_dataframe = main_dataframe.reindex(columns=['Grupo','Fecha','Tipo','Punto de Venta'])
    else :
         name['APELLIDO'] = name[0] 
         name['NOMBRE'] = name[1].str.cat(name[2], sep = " ", na_rep = "")
         name['NOMBRE'] = name.NOMBRE.str.cat(name[3], sep = " ", na_rep = "")
         name.drop([0,1,2,3], axis = 'columns', inplace=True)
        
    main_dataframe.insert(0,'GRUPO',"", allow_duplicates=False)
    main_dataframe.insert(1,'Nº LEGAJO',"", allow_duplicates=False)
   # main_dataframe = main_dataframe.reindex(columns=['GRUPO','Nº LEGAJO','APELLIDO','NOMBRE'])
    main_dataframe.insert(2,'APELLIDO',"", allow_duplicates=False)
    main_dataframe['APELLIDO'] = name['APELLIDO']
    main_dataframe.insert(3,'NOMBRE',"", allow_duplicates=False)
    main_dataframe['NOMBRE'] = name['NOMBRE']
    main_dataframe.drop(['APELLIDO Y NOMBRE'], axis = 'columns', inplace=True)
    

    main_dataframe.insert(4,'FECHA DE NACIMIENTO DD/MM/AAAA',"", allow_duplicates=False)
    #main_dataframe.insert(5,'Nº LEGAJO',"", allow_duplicates=False)
    main_dataframe.insert(6,'FECHA DE INGRESO',"", allow_duplicates=False)
    #main_dataframe.insert(7,'OBRA SOCIAL',"", allow_duplicates=False)
    main_dataframe.insert(8,'SITUACION DE REVISTA',"1", allow_duplicates=False)
    main_dataframe.insert(9,'ACTIVIDAD ECONÓMICA A LA QUE SE ENCUENTRA AFECTADO',"", allow_duplicates=False)
    main_dataframe['ACTIVIDAD ECONÓMICA A LA QUE SE ENCUENTRA AFECTADO'] = main_dataframe['ACTIVIDAD LABORAL']
    main_dataframe.drop(['ACTIVIDAD LABORAL'], axis = 'columns', inplace=True)

    main_dataframe.insert(10,'REMUNERACION MENSUAL',"", allow_duplicates=False) #	
    main_dataframe.insert(11,'INCLUIDO EN CCT',"SI", allow_duplicates=False)
    main_dataframe.insert(12,'CONVENIO APLICABLE',"", allow_duplicates=False)
    main_dataframe.insert(13,'CATEGORIA',"", allow_duplicates=False)
    main_dataframe.insert(14,'SECCION',"", allow_duplicates=False)
    main_dataframe.insert(15,'CALIFICACION PROFESIONAL',"", allow_duplicates=False)
    main_dataframe.insert(16,'PUESTO DESEMPEÑADO (Resolución SRT 244/06 Anexo II)',"", allow_duplicates=False)
    main_dataframe['PUESTO DESEMPEÑADO (Resolución SRT 244/06 Anexo II)'] = main_dataframe['PUESTO DESEMP.']
    main_dataframe.drop(['PUESTO DESEMP.'], axis = 'columns', inplace=True)
    main_dataframe.insert(17,'CODIGO ACTIVIDAD SICOSS',"", allow_duplicates=False)
    main_dataframe.insert(18,'CONDICION SICOSS',"", allow_duplicates=False)
    main_dataframe.insert(19,'MODALIDAD DE CONTRATACION SICOSS',"", allow_duplicates=False)
    main_dataframe.insert(20,'CODIGO DESCRIPCION DE PUESTO DE TRABAJO',"", allow_duplicates=False)

    writer = ExcelWriter('../data/formato_sec_rel_laborales.xlsx')
    main_dataframe.to_excel(writer, 'Tabla Relaciones Laborales', index=False)
    writer.close()
    messagebox.showinfo(message="¡Haz generado la tabla de relaciones laborales!", title="Finalizado")

def cleandata():
      
    #folder_path = '../data/comprobantes/comprob'
    folder_path = '../data/miscomprobantes'
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
    main_dataframe.replace({"$": "PESOS", "USD": "DÓLAR" , "€" : "EURO" , "$R" : "REAL"}, inplace=True)
    main_dataframe['Tipo'] = main_dataframe.Tipo.apply(
    lambda x: x.split(' ')[0]
    )

    datos = (main_dataframe['Imp. Neto Gravado']+main_dataframe['Imp. Neto No Gravado']+main_dataframe['Imp. Op. Exentas'])*main_dataframe['Tipo de Cambio del Comprobante']
    imp_total = main_dataframe['MONTO TOTAL DEL COMPROBANTE EN PESOS']*main_dataframe['Tipo de Cambio del Comprobante']
#columna = pd.DataFrame(datos)
#print(columna)
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
    main_dataframe.insert(17,'cant.',"1", allow_duplicates=False)
    main_dataframe.insert(18,'VALOR UNITARIO NETO EN PESOS', datos , allow_duplicates=False)
#main_dataframe['VALOR UNITARIO NETO EN PESOS'] =  "1"
    main_dataframe.insert(19,'VALOR NETO EN PESOS TOTAL POR CONCEPTO (Valor Unitario x Cantidad) (campo de cálculo automático)',"0", allow_duplicates=False)
#main_dataframe.insert(20,'IVA (en Pesos por concepto facturado)',"", allow_duplicates=False)
    main_dataframe.insert(21,'PERCEPCION I.V.A. (en Pesos por concepto facturado)',0, allow_duplicates=False)
    main_dataframe.insert(22,"PERCEPCION ISIB (en Pesos por concepto facturado)","", allow_duplicates=False)
    main_dataframe.insert(23,'OTROS IMPUESTOS (en Pesos por concepto facturado)',0, allow_duplicates=False)
    main_dataframe.insert(24,'TOTAL (en Pesos por concepto facturado)(campo de cálculo automático)',"", allow_duplicates=False)
    main_dataframe['MONTO TOTAL DEL COMPROBANTE EN PESOS'] = imp_total
   
    main_dataframe.at[0,'VALOR NETO EN PESOS TOTAL POR CONCEPTO (Valor Unitario x Cantidad) (campo de cálculo automático)'] = '=+IF(R2*S2=0;"";R2*S2)'
    main_dataframe['IVA (en Pesos por concepto facturado)'] = ""

    main_dataframe.at[0,'IVA (en Pesos por concepto facturado)'] = "=+(T2*0.21)"
    main_dataframe.at[0,'PERCEPCION ISIB (en Pesos por concepto facturado)'] = "=+(T2*0.015)"
    main_dataframe.at[0,'TOTAL (en Pesos por concepto facturado)(campo de cálculo automático)'] = '=+IF(SUM(T2:X2)=0;"";SUM(T2:X2))'
    
    writer = ExcelWriter('../data/formato_sec_industria.xlsx')
    main_dataframe.to_excel(writer, 'Tabla de Mis Comprobantes', index=False)
    writer.close()
    messagebox.showinfo(message="¡Haz generado la tabla de comprobantes!", title="Finalizado")


########### Parte de interfaz grafica #################################
root = tk.Tk()
root.config(width=400, height=250)
root.title("Automata Argol")
boton = ttk.Button(text="Generar tabla comprobantes", command=cleandata)
botonRel = ttk.Button(text="Generar tabla Rel. Laborales", command=reldata)

botonSalir = ttk.Button(text = "Salir", command=root.destroy)
boton.place(x=60, y=50)
botonRel.place(x=230, y=50)
botonSalir.place(x=180, y=90)
root.mainloop()