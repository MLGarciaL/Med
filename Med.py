import sys
def remedios(Archivo):
    import pdfplumber
    import numpy as np
    import pandas as pd

    #VARIABLES
    data=pd.read_excel(r'C:\Users\usuario\Desktop\Proyectos\Versiones anteriores\Planilla Med.xlsm')
    data=data[data['Activo']==1]
    data.fillna('', inplace=True)

    nombre = data.iloc[:, 0].tolist()
    remedios = data.iloc[:, 6].tolist()
    nextline1 = data.iloc[:, 7].tolist()
    comprimido = data.iloc[:, 8].tolist()
    nextline2 = data.iloc[:, 9].tolist()

              

    #EXTRACCIÓN INFORMACIÓN
    all_lines = []    
    with pdfplumber.open(Archivo) as pdf:
        for page in pdf.pages: 
            columns = [[] for _ in range(3)]  
            for table in page.extract_tables():
                for row in table:
                    for i, cell in enumerate(row[:3]):
                        if cell is not None:
                            cell_lines = cell.strip().split('\n')
                            columns[i].extend(cell_lines)
                        else:
                            columns[i].append('')
    
            lines = []
            for column in columns:
                lines.extend(column)

            lines = [line.replace("•", "") for line in lines]
    
            all_lines.append(lines)

    lineas= [item for sublist in all_lines for item in sublist]
    lineas= [linea for linea in lineas if 'Manual Farmacéutico Digital' not in linea and 'Ioma M.F.' not in linea and 'Pami' not in linea and 'IOMA' not in linea and 'VENTA VIGILADA' not in linea and 'precio' not in linea and 'Página' not in linea]
    lineas= [linea for linea in lineas if linea != ''] 

    #PROCESAMIENTO DATOS
    med=[]
    fechas=[]
    importe=[]
    
    for i in range(len(remedios)): 
        if nextline1[i] == '':
            indice = None
            busqueda = True
            h = 1
            for j in range(len(lineas)):
                if lineas[j] == remedios[i]:
                    indice = j
                    med.append(lineas[j])
                    break
            if indice is not None:
                while busqueda==True and j+h <= len(lineas):
                    if not lineas[indice + h].split()[0].isupper():
                        h +=1
                    else: 
                        busqueda=False
    
        else: 
            indice = None
            busqueda = True
            h = 1
            for j in range(len(lineas)):
                if lineas[j] == remedios[i] and lineas[j+1] == nextline1[i]:
                    indice = j
                    med.append(lineas[j])
                    break
            if indice is not None:
                while busqueda==True and j+h+1<= len(lineas):
                    if not lineas[indice + h+1].split()[0].isupper():
                        h +=1
                    else: 
                        busqueda=False
                    
        if indice is not None:                
            descripcion = False
            k=0
            if nextline2[i] == '':
                while descripcion == False and k<=h:                        
                    if  comprimido[i] not in lineas[indice+k]:
                        k+=1
                    else:
                        descripcion = True
                            
            else: 
                while descripcion == False and k<=h:                        
                    if comprimido[i] not in lineas[indice+k]:
                        k+=1                    
                    else:
                        if  nextline2[i] not in lineas[indice+k+1]:
                            k+=1
                        else:
                            descripcion = True
            
            if comprimido[i] not in lineas[indice+k]:
                k=0
                
            info=lineas[j+k]
            fechas.append(info[len(comprimido[i]) + 1: len(comprimido[i]) + 11])
            importe.append(info[len(comprimido[i]) + 11:])
    
            if fechas[i] != '':
                if not fechas[i][0].isdigit():
                    fechas[i]=''
                    importe[i]=''
                
        else:
            med.append('')
            fechas.append('')
            importe.append('')

    return nombre, fechas, importe
 
def nuevomes():
    import pandas as pd
    from datetime import datetime
    Archivo=input('Nombre Archivo con terminación .pdf: ')
    def validate_date(date_str):
        try:
            fecha = datetime.strptime(date_str, '%d-%m-%y')
            return fecha.strftime('%d-%m-%y')
        except ValueError:
            return None
    
    while True:
        fecha_str = input('Fecha dd-mm-aa: ')
        fecha = validate_date(fecha_str)
        if fecha:
            break
        else:
            print('Formato de fecha incorrecto.')

    rem, fechas, importe=remedios(Archivo)
    
    df=pd.DataFrame({'Remedio': rem, 'fecha':fechas, 'Importe':importe })
    df['Importe'] = df['Importe'].str.replace(',', '')
    Archivo='Back Up Med ' + fecha+'.xlsx'
    df.to_excel(Archivo, index=False)
    print('Proceso finalizado')

def actualizaciones():
    import os
    import pandas as pd
    from datetime import datetime
    
    Carpeta=input('Nombre Carpeta: ')

    def validate_date(date_str):
        try:
            fecha = datetime.strptime(date_str, '%d-%m-%y')
            return fecha
        except ValueError:
            return None
    
    while True:
        fecha_str = input('Fecha dd-mm-aa: ')
        fecha = validate_date(fecha_str)
        if fecha:
            break
        else:
            print('Formato de fecha incorrecto.')

    if os.path.exists(Carpeta):
        files = os.listdir(Carpeta)
        i=0
        for Archivo in files:
            rem, fechas, importe=remedios(Carpeta + '/' + Archivo)
            i+=1
            
            nombre_remedios = f"remedios{i}"
            nombre_fechas = f"fechas{i}"
            nombre_importe = f"importe{i}"
            
            globals()[nombre_remedios] = rem
            globals()[nombre_fechas] = fechas
            globals()[nombre_importe] = importe
            
        fechas0 = fechas1
        importe0 = importe1
        fechas0 = [datetime.strptime(date, '%d/%m/%Y') if date else '' for date in fechas0]
        
        for j in range(1,i+1):
            nombre_fechas = globals()[f"fechas{j}"]
            nombre_fechas = [datetime.strptime(date, '%d/%m/%Y') if date else '' for date in nombre_fechas]
            nombre_importe = globals()[f"importe{j}"]
            for k in range(0,len(rem)):
                if nombre_fechas[k]!='' and nombre_fechas[k] <= fecha:
                    fechas0[k]=nombre_fechas[k]
                    importe0[k]=nombre_importe[k]
        fechas0= [date.strftime('%d-%m-%y') if date else '' for date in fechas0]
        df=pd.DataFrame({'Remedio': rem, 'Fecha':fechas0, 'Importe':importe0})
        df['Importe'] = df['Importe'].str.replace(',', '')
        df['Importe'] = df['Importe'].str.replace('.', ',')
        fecha = fecha.strftime('%d-%m-%y')
        Archivo='Back Up Med ' + fecha+'.xlsx'
        df.to_excel(Archivo, index=False)
    else:
        print('No se encontró la carpeta')

    print('Proceso finalizado')

if __name__ == "__main__":
    if len(sys.argv) > 1:
        if sys.argv[1] == "nuevomes":
            nuevomes()
        elif sys.argv[1] == "actualizaciones":
            actualizaciones()
