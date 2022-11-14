from wsgiref import headers
import pandas as pd
from pandas import ExcelWriter
from pip import main
import xlrd 
import glob


  
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
main_dataframe.to_excel(writer, 'Tabla', index=False)
writer.close()

