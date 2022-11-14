import tkinter as tk
from tkinter import messagebox, ttk
from wsgiref import headers
import pandas as pd
from pandas import ExcelWriter
import xlrd 
import glob

def reldata():
# folder_path = '../data/relaborales' Aca sube dos niveles
#folder_path = '../data/relaborales'
 i = 0
 apellido = nombre = cuil = obrasocial = modsicos = acteconomica = convenio = categoria = puesto = salario = ingreso = cct = califprof = ""
 var = 1
 main_dataframe = pd.DataFrame()

 main_dataframe = pd.DataFrame(columns=['GRUPO', 'Nº LEGAJO', 'APELLIDO', 'NOMBRE', 'FECHA DE NACIMIENTO DD/MM/AAAA', 'CUIL', 'FECHA DE INGRESO', 'OBRA SOCIAL', 'SITUACION DE REVISTA', 'ACTIVIDAD ECONOMICA A LA QUE SE ENCUENTRA AFECTADO', 'REMUNERACION MENSUAL', 'INCLUIDO EN CCT', 'CONVENIO APLICABLE', 'CATEGORIA', 'SECCION', 'CALIFICACION PROFESIONAL', 'PUESTO DESEMPEÑADO (Resolución SRT 244/06 Anexo II)', 'CODIGO ACTIVIDAD SICOSS', 'CONDICION SICOSS', 'MODALIDAD DE CONTRATACION SICOSS', 'CODIGO DESCRIPCION DE PUESTO DE TRABAJO'])
 with open('../data/relaborales/labodata.txt', "r") as f:
    for linea in f:        
               
        if i == 1:
            x = linea.split(" ")
            apellido = x[1]
            if len(x) >= 4:
                 nombre = x[2] + " " + x[3]
            else:
                 nombre = x[2]
            cuil = linea[9:22]
            cuil = cuil.replace("-", "")
        if i == 2:            
            x = linea.split(" ")
            obrasocial = x[1]
            obrasocial = obrasocial[-6:]
        if i == 3:
            modsicos = linea.split("-")
            modsicos = modsicos[0]
            modsicos = modsicos[14:17]
        if i == 5:
            x = linea.split(" ")
            acteconomica = x[0]
            acteconomica = acteconomica[10:16]
                        
        if i == 6:            
            convenio = linea.replace("Convenio:", "")
            if "9999/99" in convenio:
                cct = "NO"
            else:
                cct = "SI"
        if i == 7:
            categoria = linea.split(" ")[0]
            categoria = categoria[11:17]
        if i == 8:
            puesto = linea.split(" ")[0]
            puesto = puesto[7:11]
            califprof = linea.split("-")[1]
        if i == 10:
            x = linea.split(" ")
            salario = x[1]
            salario = salario[8:15]            
            
        if i == 13:
            ingreso = linea[-11:]
        if i == 20:
            i = 0
            main_dataframe = main_dataframe.append({'APELLIDO': apellido, 'NOMBRE':nombre, 'CUIL': cuil,'OBRA SOCIAL': obrasocial, 'MODALIDAD DE CONTRATACION SICOSS': modsicos,'ACTIVIDAD ECONOMICA A LA QUE SE ENCUENTRA AFECTADO': acteconomica,'SITUACION DE REVISTA': var,'CONVENIO APLICABLE': convenio, 'CATEGORIA':categoria, 'PUESTO DESEMPEÑADO (Resolución SRT 244/06 Anexo II)': puesto,'REMUNERACION MENSUAL': salario, 'FECHA DE INGRESO': ingreso, "INCLUIDO EN CCT": cct , "CALIFICACION PROFESIONAL": califprof, "SECCION" : califprof}, ignore_index=True)
         
        
        i = i + 1
 main_dataframe.to_excel("../data/formato_sec_relaborales.xlsx", index=False) 

 #writer = ExcelWriter('../data/formato_sec_rel_laborales.xlsx')
 #main_dataframe.to_excel(writer, 'Tabla Relaciones Laborales', index=False)
 #writer.close()
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
    main_dataframe['Tipo'] = main_dataframe.Tipo.apply(lambda x: x.split(' ')[0] )
    main_dataframe['Tipo'] = main_dataframe['Tipo'].astype('int64')

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