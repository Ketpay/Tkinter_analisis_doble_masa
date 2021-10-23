##############################################################
#----------------------Libraries------------------------------
##############################################################

#Las librerias tkinter sirven para generar las vistas
#La libreria pandas permite generar archivos excel

from tkinter import*  # importar interfaz

from tkinter import messagebox  # pop ups mensajes

from tkinter import filedialog # guardar archivos

from tkinter.filedialog import asksaveasfile, askdirectory  #descargar archivos

import tkinter as tk  # iniciar interfaz

import pandas as pd  # generar archivos excel

import numpy as np # libreria de matrices 

import xlsxwriter  # escribir excels

from statistics import variance #libreria para haver varianza

import math #libreria para uso de funciones matematicas

import matplotlib.pyplot as plt # permite generar graficos

import os  #permite acceder a las fucciones principales de la pc - navegar por archivos de pc


##############################################################
#----------------------Variables globales---------------------
##############################################################
#variables que pasaran de funcion a funcion

global estaciones
global years
global years2

##############################################################
#----------------------Funciones------------------------------
##############################################################

#----------------------Descargar------------------------------
#Esta funcion permite descargar el formulario en excel basandose en los ajustes que se le da.

def descargar():
#----------------------variables globales------------------------------
    global estaciones
    global years
    global years2
    global frame_comandos
    #variables de la funcion
    dic={}
    lista=[]
    lista_estaciones=[]

    try:
        estaciones=int(estaciones.get())
        years=int(years.get())
        years2=int(years2.get())
        #condicion si uno o mas campos es menos o igual a cero a la hora de descargar el formato

        if estaciones <= 0 or years<= 0 or years2<= 0 :


            #cambios visuales del Tkinter
            messagebox.showerror(message="Uno o mas campos son valores negativos o son cero, porfavor llene todos los campos de ajustes con valores positivos", title="Error al generar formato")
            frame_comandos.destroy()
            frame_comandos=Frame()
            frame_comandos.pack()
            frame_comandos.config(bd=4,relief="groove",bg="#282923")
            frame_comandos.place(x="20",y="150")
            text_comandos_1=Label(frame_comandos,text="AJUSTES",font="Verdana 12 bold", fg="white", bg="#282923")
            text_comandos_1.grid(row=0,column=0,pady=5,padx=5)


            text_comandos_3=Label(frame_comandos,text="Numero de estaciones:",font="Verdana 10 bold", fg="white", bg="#282923",justify="left")
            text_comandos_3.grid(row=3,column=0,padx=5,pady=2)
            estaciones=tk.Entry(frame_comandos)
            estaciones.grid(row=3,column=1,padx=5,pady=2)

            text_comandos_4=Label(frame_comandos,text="Cantidad de años:",font="Verdana 10 bold", fg="white", bg="#282923",justify="left")
            text_comandos_4.grid(row=4,column=0,padx=5,pady=2)
            years=tk.Entry(frame_comandos)
            years.grid(row=4,column=1,padx=5,pady=2)
            text_comandos_5=Label(frame_comandos,text="Año inicial:",font="Verdana 10 bold", fg="white", bg="#282923",justify="left")
            text_comandos_5.grid(row=5,column=0,padx=5,pady=2)
            years2=tk.Entry(frame_comandos)
            years2.grid(row=5,column=1,padx=5,pady=2)

        else:

            try:
                file=asksaveasfile(defaultextension=".xlsx", initialfile="formato.xlsx", title="Guardar Formato",)

            except:
                messagebox.showerror(message="Ha ocurrido un error, verifique que el archivo no este siendo utilizado y vuelva a intentarlo.", title="Error al Guardar")

            if file!=None:

                # permite generar un excel con las estaciones y años dichos 

                for i in range(years):
                    lista.append(years2+i)
                    lista_estaciones.append('-')

                dic['Años']=lista

                lista=[]
                for i in range(estaciones):
                    numero='Estacion_'+str(i+1)
                    dic[numero]=lista_estaciones

                writer = pd.ExcelWriter(str(file.name), engine='xlsxwriter')

                df=pd.DataFrame(dic)
#----------------------creacion del archivo------------------------------
                df.to_excel(writer, sheet_name="Formato", index=False)
                workbook = writer.book

                worksheet = writer.sheets['Formato']

                header_format = workbook.add_format()
                a_format = workbook.add_format()

                header_format.set_bold()
                #estilo del excel
                header_format.set_font_size(12)
                header_format.set_italic()
                header_format.set_align('center')
                header_format.set_align('vcenter')
                a_format.set_align('center')
                a_format.set_align('vcenter')
                header_format.set_pattern(1)  
                header_format.set_bg_color('#d8e4c0')
                for i in range(estaciones+1):
                    worksheet.set_column(1, i, 14.11,a_format)
                for col_num, value in enumerate(df.columns.values):

                    worksheet.write(0, col_num, value, header_format)
                
                writer.close()

#----------------------limpiesa por descarga------------------------------

                messagebox.showinfo(message="Para evitar Errores al procesar, no modifique el formato descargado.", title="Información")
                frame_comandos.destroy()
                frame_comandos=Frame()
                frame_comandos.pack()
                frame_comandos.config(bd=4,relief="groove",bg="#282923")
                frame_comandos.place(x="20",y="150")
                text_comandos_1=Label(frame_comandos,text="AJUSTES",font="Verdana 12 bold", fg="white", bg="#282923")
                text_comandos_1.grid(row=0,column=0,pady=5,padx=5)


                text_comandos_3=Label(frame_comandos,text="Numero de estaciones:",font="Verdana 10 bold", fg="white", bg="#282923",justify="left")
                text_comandos_3.grid(row=3,column=0,padx=5,pady=2)
                estaciones=tk.Entry(frame_comandos)
                estaciones.grid(row=3,column=1,padx=5,pady=2)

                text_comandos_4=Label(frame_comandos,text="Cantidad de años:",font="Verdana 10 bold", fg="white", bg="#282923",justify="left")
                text_comandos_4.grid(row=4,column=0,padx=5,pady=2)
                years=tk.Entry(frame_comandos)
                years.grid(row=4,column=1,padx=5,pady=2)
                text_comandos_5=Label(frame_comandos,text="Año inicial:",font="Verdana 10 bold", fg="white", bg="#282923",justify="left")
                text_comandos_5.grid(row=5,column=0,padx=5,pady=2)
                years2=tk.Entry(frame_comandos)
                years2.grid(row=5,column=1,padx=5,pady=2)
            else:

                frame_comandos.destroy()
                frame_comandos=Frame()
                frame_comandos.pack()
                frame_comandos.config(bd=4,relief="groove",bg="#282923")
                frame_comandos.place(x="20",y="150")
                text_comandos_1=Label(frame_comandos,text="AJUSTES",font="Verdana 12 bold", fg="white", bg="#282923")
                text_comandos_1.grid(row=0,column=0,pady=5,padx=5)


                text_comandos_3=Label(frame_comandos,text="Numero de estaciones:",font="Verdana 10 bold", fg="white", bg="#282923",justify="left")
                text_comandos_3.grid(row=3,column=0,padx=5,pady=2)
                estaciones=tk.Entry(frame_comandos)
                estaciones.grid(row=3,column=1,padx=5,pady=2)

                text_comandos_4=Label(frame_comandos,text="Cantidad de años:",font="Verdana 10 bold", fg="white", bg="#282923",justify="left")
                text_comandos_4.grid(row=4,column=0,padx=5,pady=2)
                years=tk.Entry(frame_comandos)
                years.grid(row=4,column=1,padx=5,pady=2)
                text_comandos_5=Label(frame_comandos,text="Año inicial:",font="Verdana 10 bold", fg="white", bg="#282923",justify="left")
                text_comandos_5.grid(row=5,column=0,padx=5,pady=2)
                years2=tk.Entry(frame_comandos)
                years2.grid(row=5,column=1,padx=5,pady=2)
                pass







    except:
        messagebox.showerror(message="Error Campos vacios o con texto, porfavor llene los campos correctamente", title="Error al generar formato")
        frame_comandos.destroy()
        frame_comandos=Frame()
        frame_comandos.pack()
        frame_comandos.config(bd=4,relief="groove",bg="#282923")
        frame_comandos.place(x="20",y="150")
        text_comandos_1=Label(frame_comandos,text="AJUSTES",font="Verdana 12 bold", fg="white", bg="#282923")
        text_comandos_1.grid(row=0,column=0,pady=5,padx=5)


        text_comandos_3=Label(frame_comandos,text="Numero de estaciones:",font="Verdana 10 bold", fg="white", bg="#282923",justify="left")
        text_comandos_3.grid(row=3,column=0,padx=5,pady=2)
        estaciones=tk.Entry(frame_comandos)
        estaciones.grid(row=3,column=1,padx=5,pady=2)

        text_comandos_4=Label(frame_comandos,text="Cantidad de años:",font="Verdana 10 bold", fg="white", bg="#282923",justify="left")
        text_comandos_4.grid(row=4,column=0,padx=5,pady=2)
        years=tk.Entry(frame_comandos)
        years.grid(row=4,column=1,padx=5,pady=2)
        text_comandos_5=Label(frame_comandos,text="Año inicial:",font="Verdana 10 bold", fg="white", bg="#282923",justify="left")
        text_comandos_5.grid(row=5,column=0,padx=5,pady=2)
        years2=tk.Entry(frame_comandos)
        years2.grid(row=5,column=1,padx=5,pady=2)


 
    pass


#----------------------Abrir------------------------------
#Esta funcion permite abrir el archivo a procesar.

def abrir():

    #variables globales-------------------------------
    global boton_principal

    global boton_procesar 

    global file

    global frame_directorio

    global file_image2

    global borrar

    global frame_comandos

    global analisis
    global texto_advertencia
    global texto_advertencia2

    #analiza el archivo-----------------------------

    file = filedialog.askopenfilename(title="Abrir", filetypes=(
        ("Archivos .xlsx", "*.xlsx"), ("Todos los ficheros", "*.*"))) # abre archivo

    if file!="":
        archivo=file.split("/") 
        longitud=len(archivo)
        archivo="Archivo: "+ archivo[longitud-1]

        #Realiza el cambio de la vista------------------------
        frame_comandos.destroy()
        analisis=Frame()
        analisis.pack()
        analisis.config(bd=4,relief="groove",bg="#282923")
        analisis.place(x="20",y="150")

        analisis_text_1=Label(analisis,text="ANÁLISIS",font="Verdana 12 bold", fg="white", bg="#282923")
        analisis_text_1.grid(row=0,column=0,pady=5,padx=20)



        analisis_text_2=Label(analisis,text="Numero de estaciones:",font="Verdana 10 bold", fg="white", bg="#282923")
        analisis_text_2.grid(row=2,column=0,padx=20)
        analisis_text_2_2=Label(analisis,text=" Sin Datos ",font="Verdana 10 bold", fg="white", bg="#282923")
        analisis_text_2_2.grid(row=2,column=1,padx=0)

        analisis_text_3=Label(analisis,text="Numero de años:",font="Verdana 10 bold", fg="white", bg="#282923")
        analisis_text_3.grid(row=3,column=0,padx=20)
        analisis_text_3_2=Label(analisis,text=" Sin Datos ",font="Verdana 10 bold", fg="white", bg="#282923")
        analisis_text_3_2.grid(row=3,column=1,padx=0)

        analisis_text_4=Label(analisis,text="Datos analizados:",font="Verdana 10 bold", fg="white", bg="#282923")
        analisis_text_4.grid(row=4,column=0,padx=20)
        analisis_text_4_2=Label(analisis,text=" Sin Datos ",font="Verdana 10 bold", fg="white", bg="#282923")
        analisis_text_4_2.grid(row=4,column=1,padx=0)

        analisis_text_5=Label(analisis,text="Datos Faltantes:",font="Verdana 10 bold", fg="white", bg="#282923")
        analisis_text_5.grid(row=5,column=0,padx=20)
        analisis_text_5_2=Label(analisis,text=" Sin Datos ",font="Verdana 10 bold", fg="white", bg="#282923")
        analisis_text_5_2.grid(row=5,column=1,padx=0)
          
        texto_advertencia=Label(app,text="Advertencia: Archivos con muchas estaciones ",font="Verdana 10 bold", fg="white", bg="#282923")
        texto_advertencia.place(x="20", y="300")
        texto_advertencia2=Label(app,text="o años puede demorar para dar el resultado.",font="Verdana 10 bold", fg="white", bg="#282923")
        texto_advertencia2.place(x="20", y="320")

        frame_directorio.destroy()

        frame_directorio = Frame()

        frame_directorio.config(bd=4, relief="groove", bg="#282923")

        frame_directorio.place(x="80", y="100", width="450", height="30")

        directorio = Label(frame_directorio, bg="#282923",text=archivo, fg="white")

        directorio.place(x="0", y="0")

        boton_principal.destroy()

        boton_principal = Button(app, width="20", height="2", text="Procesar",

                                font="Verdana 10 bold", fg="white", cursor="hand2", bg="#282923", command=procesar)
        boton_principal.place(x="400", y="150")

        borrar = Button(app, image=file_image2, cursor="hand2",
                     bg="#282923", command=limpiar) 

        borrar.place(x="560", y="100")
   
    pass

#----------------------Procesar------------------------------
#Esta funcion permite procesar los datos y dar resultados
def procesar():

    #variables globales-------------------------------
#----------------------El proceso hace un barrido de todos los datos para analizarlo------------------------------
    global boton_carpeta

    global boton_guardar
    global boton_principal
    global texto_advertencia
    global texto_advertencia2
    global file
    global analisis
    global file_excel
    global filas
    global columnas
    global lista_porcentual2
    global columnas_grafico

    lista=[]
    lista_blanco=[]
    columnas_blanco=[]
    filas_blanco=[]
    columna_completa=[]
    lista_general=[]

    #Analisis doble masa-------------------------
    df = pd.read_excel (file)
    columnas=len(df.columns)
    columnas_grafico=len(np.copy(df.columns))

    #almacenamos en una lista las columnas

    for i in range (columnas):
        x = pd.DataFrame(df, columns= [df.columns[i]])
        x = np.array( x.values)
        copia=np.copy(x)
        lista.append(x)
        lista_general.append(copia)
    sumatoria=[]
    promedio=[]

    filas=len(lista[0])


    for i in range(columnas):
        x=lista[i]
        completa=False

        if i==0:
            completa=True   
            pass
        else:
            for y in range(filas):
                if np.isnan(x[y]):
                    completa=True
                    lista_blanco.append([y,i])
                    filas_blanco.append(y)
                    if i in columnas_blanco:
                        pass
                    else:
                        columnas_blanco.append(i)
                else:
                    pass


        if completa==False:
            columna_completa.append(i)


    lista_masa=[]
    sumatoria=[]
    promedio=[]
    lista_acumulada=[]
    if len(columna_completa)==0:
        messagebox.showerror(message="Tienes todas las estaciones con datos faltantes", title="Error al Procesar")

    elif len(columnas_blanco)==0:
        messagebox.showerror(message="No tienes datos Faltantes", title="Error al Procesar")
    else:

        for i in range (columnas):
            suma=0
            prom=0
            acumulada=[]
            if i==0:
                sumatoria.append("Sumatoria")
                promedio.append("Promedio")

            else:

                for y in range(filas):
                    if y in filas_blanco:
                        pass
                    else:
                        dato=lista[i]
                        dato=dato[y]

                        suma=round(suma+dato[0],3)

                        acumulada.append(suma)


                prom=round(suma/(filas-len(filas_blanco)), 2)

                acumulada.append(np.nan)
                acumulada.append(np.nan)
                lista_acumulada.append(acumulada)
                sumatoria.append(round(suma, 2))
                promedio.append(prom)

        lista_porcentual=[]
        lista_porcentual2=[]
        for i in range (len(lista_acumulada)):
            datos=lista_acumulada[i]
            porcentual=[]
            porcentual2=[]
            for y in range(len(datos)-2):  
                dato=datos[y]
                x=round((dato/sumatoria[i+1])*100,2)
                x_copia=x
                x=str(x)+"%"
                porcentual.append(x)
                porcentual2.append(x_copia)

            porcentual.append(np.nan)
            porcentual.append(np.nan)
            lista_porcentual.append(porcentual)
            lista_porcentual2.append(porcentual2)


        dic={}

        lista_excel=[]
        lista_auxiliar=[]

        
        for i in range((columnas*3)-2):


            lista_auxiliar=[]
            if i<columnas:
                numero='Estacion_'+str(i)
                datos=lista[i]
                for y in range(filas):
                    if y in filas_blanco:
                        pass
                    else:
                        dato=datos[y]
                        lista_auxiliar.append(dato[0])
                if i==0:
                    lista_auxiliar.append(sumatoria[i])
                    lista_auxiliar.append(promedio[i])
                    dic['Años']=lista_auxiliar
                else:
                    lista_auxiliar.append(sumatoria[i])
                    lista_auxiliar.append(promedio[i])
                    dic[numero]=lista_auxiliar

            else:
                if i<((columnas*2)-1):
                    numero='Acumulada_'+str(i-columnas+1)
                    dic[numero]=lista_acumulada[i-columnas]
                else:
                    numero='Porcentual_'+str(i-columnas-columnas+2)
                    dic[numero]=lista_porcentual[i-columnas-columnas+1]

        df1=pd.DataFrame(dic)



    #Recta de Regresión programa----------------
        #analizamos las columnas

        dic2={}
        dic4={}
        lista_dato=[]
        tiempo=[]
        tiempo2=[]
        
        # for i in range(filas):

        #     tiempo.append("Dato "+str(i+1))
        tiempo.append("Promedio")
        tiempo.append("Varianza")
        tiempo.append("Covarianza")
        tiempo.append("Coeficiente A")
        tiempo.append("Coeficiente B")

        tiempo2.append("Promedio")
        tiempo2.append("Varianza")
        tiempo2.append("Covarianza")
        tiempo2.append("Coeficiente A")
        tiempo2.append("Coeficiente B")
        tiempo2.append("Valor A")
        tiempo2.append("Valor B")
        tiempo2.append("Valor C")
        tiempo2.append("Valor L1")
        tiempo2.append("Valor M")
        while len(tiempo)<filas:
            tiempo=np.append(tiempo,"")
        while len(tiempo2)<filas:
            tiempo2=np.append(tiempo2,"")   
        columna_vacia=[]
        for i in range(filas):
            columna_vacia.append("")
        for i in range(len(columnas_blanco)):
            valor1=[]
            valor2=[]
            valor3=[]
            valor4=[]
            lista_regresion=[]
            lista_ortogonal=[]
            lista1=[]
            lista2=[]
            lista3=[]
            lista4=[]
            filas_borrar=[]
            varianza=[]
            
            lista_regresion.append(lista[columna_completa[0]])
            lista_regresion.append(lista[columnas_blanco[i]])
            for y in range(len(lista_regresion[1])):
                datos=lista_regresion[1]
                dato=datos[y]
                if np.isnan(dato):
                    filas_borrar.append(y)
                else:
                    dato=lista[columna_completa[0]]
                    dato2=lista[columnas_blanco[i]]
                    dato3=lista[0]
                    lista1.append(dato[y][0])
                    lista2.append(dato2[y][0])
                    # tiempo.append(dato3[y][0])
            promedio1=round(np.mean(lista1),3)
            promedio2=round(np.mean(lista2),3)
            covarianza=0
            for w in range(len(lista1)):
                dato=(lista1[w]-promedio1)*(lista2[w]-promedio2)
                covarianza=covarianza+dato
            covarianza=round(covarianza/len(lista1),3)


            varianza1=round(variance(lista1),3)
            varianza2=round(variance(lista2),3)

            # if len(lista1)<filas:
                # faltantes=filas-len(lista1)
            while len(lista1)<filas:
                lista1.append(np.nan)
                lista2.append(np.nan)


                    

            valor1.append(promedio1)
            valor2.append(promedio2)
            valor1.append(varianza1)
            valor2.append(varianza2)  
            valor1.append(covarianza)
            valor2.append("-")

            valor3.append(promedio1)
            valor4.append(promedio2)
            valor3.append(varianza1)
            valor4.append(varianza2)  
            valor3.append(covarianza)
            valor4.append("-")

            a=round(promedio2-covarianza/varianza1*promedio1,3)  
            b=round(covarianza/varianza1 ,3)
            valor_a=1
            valor_b=round(-1*(varianza1+varianza2),3)
            valor_c=round(varianza1*varianza2-covarianza*covarianza,3)
            valor_auxiliar=(-1*valor_b+math.sqrt(valor_b*valor_b-4*valor_a*valor_c))
            valor_l1=round(valor_auxiliar/(2*valor_a),3)
            valor_m=round(covarianza/(valor_l1-varianza2),3)
            a2=round(promedio2-valor_m*promedio1,3)
            b2=valor_m      
            lista_regresion=[]
            lista3=np.copy(lista1)
            # print(type(lista3[0]))
            agregar=[a2,b2,valor_a,valor_b,valor_c,valor_l1,valor_m]
            valor3=np.append(valor3,agregar)



            valor1.append(a)
            valor2.append("-")
            valor1.append(b)
            valor2.append("-")
            while len(valor1)<filas:
                valor1=np.append(valor1,"")
                valor2=np.append(valor2,"")
            lista4=np.copy(lista2)
            valor4=np.append(valor4,["-","-","-","-","-","-","-"])
            while len(valor3)<filas:
                valor3=np.append(valor3,"")
                valor4=np.append(valor4,"")
            # tiempo2=np.copy(tiempo)
            # tiempo2=np.append(tiempo2,["Valor A","Valor B","Valor C","Valor L1","Valor M"])
     


            lista_regresion.append(columna_vacia)
            lista_regresion.append(lista1)
            lista_regresion.append(lista2)
            lista_regresion.append(columna_vacia)
            lista_regresion.append(tiempo)
            lista_regresion.append(valor1)
            lista_regresion.append(valor2)
            # lista_regresion.append(columna_vacia)
            # print("lista1",len(lista1))
            # print("lista2",len(lista2))
            # print("tiempo",len(tiempo))
            # print("valor1",len(valor1))
            # print("valor2",len(valor2))



            lista_ortogonal.append(columna_vacia)
            lista_ortogonal.append(lista3)
            lista_ortogonal.append(lista4)
            lista_ortogonal.append(columna_vacia)
            lista_ortogonal.append(tiempo2)
            lista_ortogonal.append(valor3)
            lista_ortogonal.append(valor4)
            # lista_ortogonal.append(columna_vacia)
            # print("lista3",len(lista3))
            # print("lista4",len(lista4))
            # print("tiempo2",len(tiempo2))
            # print("valor3",len(valor3))
            # print("valor4",len(valor4))
   
            nombre0=str(i+1)+")"
            nombre1="Estacion_A_"+str(i+1)
            nombre2="Estacion_B_"+str(i+1)
            nombre3="Datos "+str(i+1)
            nombre4="-"+str(i+1)+"-"
            
            nombre5="Resultados_A"+str(i+1)
            nombre6="Resultados_B"+str(i+1)

            dic2[nombre0]=lista_regresion[0]
            dic2[nombre1]=lista_regresion[1]
            dic2[nombre2]=lista_regresion[2]
            dic2[nombre3]=lista_regresion[3]
            dic2[nombre4]=lista_regresion[4]
            dic2[nombre5]=lista_regresion[5]
            dic2[nombre6]=lista_regresion[6]
             



            dic4[nombre0]=lista_ortogonal[0]
            dic4[nombre1]=lista_ortogonal[1]
            dic4[nombre2]=lista_ortogonal[2]
            dic4[nombre3]=lista_ortogonal[3]
            dic4[nombre4]=lista_ortogonal[4]
            dic4[nombre5]=lista_ortogonal[5]
            dic4[nombre6]=lista_ortogonal[6]

            
            datos=lista[columnas_blanco[i]]
            datos2=lista_general[columnas_blanco[i]]

            datos_correcto=lista[columna_completa[0]]
            nueva_lista=[]

            for y in range(filas):
                dato=datos[y]
   
                dato_c=datos_correcto[y]
                if np.isnan(dato):
                    valor=round(a+b*dato_c[0],3)
                    valor2=round(a2+b2*dato_c[0],3)
                    datos[y]=valor
                    datos2[y]=valor2


        # print(dic2)
        # print(dic4)

        df3=pd.DataFrame(dic2)
        df5=pd.DataFrame(dic4)

        dic3={}
        dic5={}


        for i in range(columnas):
            datos=lista[i]
            datos2=lista_general[i]
            lista_aux=[]
            lista_aux2=[]
            if i ==0:
                for y in range(filas):
                    
                    dato=datos[y]
                    lista_aux.append(dato[0])
                dic3["Años"]=lista_aux
                dic5["Años"]=lista_aux
            else:
                for y in range(filas):
                    
                    dato=datos[y]
                    dato2=datos2[y]
                    lista_aux.append(dato[0])
                    lista_aux2.append(dato2[0])
                nombre= "Estacion_"+str(i)
                dic3[nombre]=lista_aux 
                dic5[nombre]=lista_aux2              


 
        df2=pd.DataFrame(dic3)
        df4=pd.DataFrame(dic5)

        analisis.destroy()
        analisis=Frame()
        analisis.pack()
        analisis.config(bd=4,relief="groove",bg="#282923")
        analisis.place(x="20",y="150")

        analisis_text_1=Label(analisis,text="ANÁLISIS",font="Verdana 12 bold", fg="white", bg="#282923")
        analisis_text_1.grid(row=0,column=0,pady=5,padx=20)



        analisis_text_2=Label(analisis,text="Numero de estaciones:",font="Verdana 10 bold", fg="white", bg="#282923")
        analisis_text_2.grid(row=2,column=0,padx=20)
        analisis_text_2_2=Label(analisis,text=(columnas-1),font="Verdana 10 bold", fg="white", bg="#282923")
        analisis_text_2_2.grid(row=2,column=1,padx=0)

        analisis_text_3=Label(analisis,text="Numero de años:",font="Verdana 10 bold", fg="white", bg="#282923")
        analisis_text_3.grid(row=3,column=0,padx=20)
        analisis_text_3_2=Label(analisis,text=filas,font="Verdana 10 bold", fg="white", bg="#282923")
        analisis_text_3_2.grid(row=3,column=1,padx=0)

        analisis_text_4=Label(analisis,text="Datos analizados:",font="Verdana 10 bold", fg="white", bg="#282923")
        analisis_text_4.grid(row=4,column=0,padx=20)
        analisis_text_4_2=Label(analisis,text=((columnas-1)*filas),font="Verdana 10 bold", fg="white", bg="#282923")
        analisis_text_4_2.grid(row=4,column=1,padx=0)

        analisis_text_5=Label(analisis,text="Datos Faltantes:",font="Verdana 10 bold", fg="white", bg="#282923")
        analisis_text_5.grid(row=5,column=0,padx=20)
        analisis_text_5_2=Label(analisis,text=(len(lista_blanco)),font="Verdana 10 bold", fg="white", bg="#282923")
        analisis_text_5_2.grid(row=5,column=1,padx=0)


        messagebox.showinfo(message="Proceso terminado!", title="Procesado") # msg box de Finalizado
        file_excel=asksaveasfile(defaultextension=".xlsx", initialfile="Resultados.xlsx", title="Guardar",)

 




        # dfs = {'Doble Masa':df1, 'Recta de Regresión':df2, 'Recta de Regresión Solucion':df3,'Correlación Ortogonal':df4,'Correlación Ortogonal Solucion':df5}

        # writer = pd.ExcelWriter(file_excel.name, engine='xlsxwriter')
        # for sheet_name in dfs.keys():

        #     dfs[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
            
        # writer.save()
        writer = pd.ExcelWriter(file_excel.name, engine='xlsxwriter')

        # df=pd.DataFrame(dic)

        df1.to_excel(writer, sheet_name="Doble Masa", index=False)
        workbook = writer.book

        worksheet = writer.sheets['Doble Masa']

        header_format = workbook.add_format()
        a_format = workbook.add_format()
        a_format.set_align('center')
        a_format.set_align('vcenter')
        #header_format.set_font_name('Bodoni MT Black')
        header_format.set_bold()

        header_format.set_font_size(12)
        header_format.set_italic()
        header_format.set_align('center')
        header_format.set_align('vcenter')
        header_format.set_pattern(1)  
        header_format.set_bg_color('#d8e4c0')
        columnas=len(df1.columns)
        for i in range(columnas):
            worksheet.set_column(1, i, 14.11,a_format)
        for col_num, value in enumerate(df1.columns.values):

            worksheet.write(0, col_num, value, header_format)


        df2.to_excel(writer, sheet_name="Recta de Regresión", index=False)
        workbook = writer.book

        worksheet = writer.sheets['Recta de Regresión']

        header_format = workbook.add_format()
        a_format = workbook.add_format()
        a_format.set_align('center')
        a_format.set_align('vcenter')
        #header_format.set_font_name('Bodoni MT Black')
        header_format.set_bold()

        header_format.set_font_size(12)
        header_format.set_italic()
        header_format.set_align('center')
        header_format.set_align('vcenter')
        header_format.set_pattern(1)  
        header_format.set_bg_color('#d8e4c0')
        columnas=len(df2.columns)
        for i in range(columnas):
            worksheet.set_column(1, i, 14.11,a_format)
        for col_num, value in enumerate(df2.columns.values):

            worksheet.write(0, col_num, value, header_format)

        df3.to_excel(writer, sheet_name="Recta de Regresión Solucion", index=False)
        workbook = writer.book

        worksheet = writer.sheets['Recta de Regresión Solucion']

        header_format = workbook.add_format()
        a_format = workbook.add_format()
        a_format.set_align('center')
        a_format.set_align('vcenter')
        #header_format.set_font_name('Bodoni MT Black')
        header_format.set_bold()
        header_format.set_align('center')
        header_format.set_align('vcenter')
        header_format.set_font_size(12)
        header_format.set_italic()

        header_format.set_pattern(1)  
        header_format.set_bg_color('#f5f9a4')
        columnas=len(df3.columns)
        for i in range(columnas):
            worksheet.set_column(1, i, 14.11,a_format)
        for col_num, value in enumerate(df3.columns.values):

            worksheet.write(0, col_num, value, header_format)

        df4.to_excel(writer, sheet_name="Correlación Ortogonal", index=False)
        workbook = writer.book

        worksheet = writer.sheets['Correlación Ortogonal']

        header_format = workbook.add_format()
        a_format = workbook.add_format()
        a_format.set_align('center')
        a_format.set_align('vcenter')
        #header_format.set_font_name('Bodoni MT Black')
        header_format.set_bold()

        header_format.set_font_size(12)
        header_format.set_italic()
        header_format.set_align('center')
        header_format.set_align('vcenter')
        header_format.set_pattern(1)  
        header_format.set_bg_color('#d8e4c0')
        columnas=len(df4.columns)
        for i in range(columnas):
            worksheet.set_column(1, i, 14.11,a_format)
        for col_num, value in enumerate(df4.columns.values):

            worksheet.write(0, col_num, value, header_format)

        df5.to_excel(writer, sheet_name="Correlación Ortogonal Solucion", index=False)
        workbook = writer.book

        worksheet = writer.sheets['Correlación Ortogonal Solucion']

        header_format = workbook.add_format()
        a_format = workbook.add_format()
        a_format.set_align('center')
        a_format.set_align('vcenter')
        #header_format.set_font_name('Bodoni MT Black')
        header_format.set_bold()

        header_format.set_font_size(12)
        header_format.set_italic()
        header_format.set_align('center')
        header_format.set_align('vcenter')
        header_format.set_pattern(1)  
        header_format.set_bg_color('#f5f9a4')
        columnas=len(df5.columns)
        for i in range(columnas):
            worksheet.set_column(1, i, 14.11,a_format)
        for col_num, value in enumerate(df5.columns.values):

            worksheet.write(0, col_num, value, header_format)


        writer.close()

  

    

    #Realiza el cambio de la vista-------------------------------
        boton_principal.destroy()
        texto_advertencia.destroy()
        texto_advertencia2.destroy()

        boton_graficos= Button(app, width="20", height="2", text="Ver graficos",
                                    font="Verdana 10 bold", fg="white", cursor="hand2", bg="#282923", command=graficos)
        boton_graficos.place(x="400", y="150")

        boton_guardar= Button(app, width="20", height="2", text="Abrir ubicacion",
                                        font="Verdana 10 bold", fg="white", cursor="hand2", bg="#282923", command=abrir_carpeta)
        boton_guardar.place(x="400", y="200")
  


    pass
#----------------------Graficos------------------------------
#Esta funcion permite ver los graficos de lo procesado

def graficos():
    #----------------------genera graficos------------------------------
    global file_excel
    
    global filas
    global columnas

    global lista_porcentual2
    global columnas_grafico

    leyenda=[]

    file=file_excel.name
    df=pd.read_excel (file)
    df2=pd.read_excel (file,sheet_name=1)
    year = pd.DataFrame(df2, columns= [df.columns[0]])
    df3=pd.read_excel (file,sheet_name=3)


    plt.figure(1)

    for i in range(len(lista_porcentual2)-1):

        x = lista_porcentual2[i]

        y = lista_porcentual2[i+1]

        plt.scatter(x, y)
        nombre="Porcentual " + str(i+1)+ " y " +str(i+2) 
        leyenda.append(nombre)

    plt.grid(True)
    plt.title("Gráfica de Doble masa")
    plt.legend(tuple(leyenda))


    leyenda2=[]

    plt.figure(2)

    for ii in range(1,columnas_grafico):
        y = pd.DataFrame(df2, columns= [df.columns[ii]])

        nombre="Estacion "+str(ii)
        leyenda2.append(nombre)
        plt.plot(year, y)



    plt.grid(True)
    plt.title("Recta de Regresión")
    plt.legend(tuple(leyenda2))
    plt.figure(3)
    for ii in range(1,columnas_grafico):

        w = pd.DataFrame(df3, columns= [df.columns[ii]])

        plt.plot(year, w)

    plt.grid(True)
    plt.title("Correlación Ortogonal")
    plt.legend(tuple(leyenda2))


    plt.show()


    pass


#----------------------Abrir carpeta------------------------------
#Esta funcion permite abrir la carpeta de ubicacion del archivo

def abrir_carpeta():
    #----------------------permite abrir la carpeta donde procesaste el archivo------------------------------
    global file_excel
    file=str(file_excel.name)
    file=file.split("/")
    # print(file)
    largo=len(file)
    ubicacion=""
    for i in range (largo-1):
        ubicacion=ubicacion+str(file[i])+"/"

    # print(ubicacion)
    os.startfile(ubicacion) 

    pass

#----------------------Guardar------------------------------
#Esta funcion permite guardar la informacion procesada

# def guardar():
#     global boton_carpeta



#     #Realiza el cambio de la vista   
#     boton_carpeta = Button(app, width="20", height="2", text="Abrir ubicacion",
#                                 font="Verdana 10 bold", fg="white", cursor="hand2", bg="#282923", command=procesar)

#     boton_carpeta.place(x="400", y="250")

#     pass
#----------------------Teoria------------------------------
#Esta funcion permite abrir una ventana y ver la teoria 

def teoria():

    #----------------------abre una nueva vista q nos muestra la teoria------------------------------
    global imagen_bg
    teoria = tk.Toplevel(app)

    teoria.title("Ver Teoria") # titulo
    teoria.geometry("550x300") # geometria inicial
    teoria.resizable(0, 0) # no e sposible  agrandar
    teoria.iconbitmap("icon.ico")
    # fondo=Label(teoria,image=imagen_bg).place(x=0,y=0)
    teoria.config(bg="#FFFFFF") #  background color

    texto=Label(teoria,text="TEORIA DEL ANÁLISIS DE DOBLE MASA",font="Verdana 14 bold", fg="black", bg="#ffffff")
    texto.place(x="20", y="15")
    texto=Label(teoria,text="El método de doble masa considera que, en una zona meteorológica",font="Verdana 10 bold", fg="black", bg="#ffffff")
    texto.place(x="20", y="100")
    texto=Label(teoria,text="homogénea, los valores de precipitación que ocurren en diferentes ",font="Verdana 10 bold", fg="black", bg="#ffffff")
    texto.place(x="20", y="120")
    texto=Label(teoria,text="puntos de esa zona en períodos anuales o estacionales guardan una ",font="Verdana 10 bold", fg="black", bg="#ffffff")
    texto.place(x="20", y="140")
    texto=Label(teoria,text="relación de proporcionalidad que puede representarse gráficamente.",font="Verdana 10 bold", fg="black", bg="#ffffff")
    texto.place(x="20", y="160")




    pass



#----------------------Acerca------------------------------
#Esta funcion permite mirar la informacion del creador

def acerca():
    global file_image3

    acerca = tk.Toplevel(app)

    acerca.title("Acerca de ...") # titulo
    acerca.geometry("460x400") # geometria inicial
    acerca.resizable(0, 0) # no e sposible  agrandar
    acerca.iconbitmap("icon.ico")
    acerca.config(bg="#282923") #  background color
    fondo=Label(acerca,image=file_image3).place(x=0,y=0)
    texto=Label(acerca,text="GRUPO LACHESIS",font="Verdana 14 bold", fg="white", bg="#282923")
    texto.place(x="20", y="100")
    texto=Label(acerca,text="Integrantes:",font="Verdana 10 bold", fg="white", bg="#282923")
    texto.place(x="20", y="130")
    texto=Label(acerca,text="- Samuel",font="Verdana 10 bold", fg="white", bg="#282923")
    texto.place(x="20", y="160")

    pass

#----------------------Salir------------------------------
#Esta funcion permite cerrar la app

def salir():
    app.destroy()
    pass

#----------------------Limpiar------------------------------
#Esta funcion limpia todo lo realizado

def limpiar():

    #variables globales --------------------------------------

    global frame_directorio
    global frame_comandos
    global analisis
    global boton_principal


    #limpiesa de botones-------------------------------------
    boton_principal.destroy()
    boton_principal = Button(app, width="20", height="2", text="Descargar Formato",
                                font="Verdana 10 bold", fg="white", cursor="hand2", bg="#282923", command=descargar)# genera boton de procesar
    boton_principal.place(x="400", y="150")
    try:
        global boton_graficos
        global boton_guardar

        boton_guardar.destroy()



        boton_graficos.destroy()
    except :
        pass
    try:
        global boton_carpeta
        boton_carpeta.destroy()
    except :
        pass
    try:
        global analisis
        global texto_advertencia
        global texto_advertencia2
        texto_advertencia.destroy()
        texto_advertencia2.destroy()
        analisis.destroy()
    except :
        pass
    try:
        global borrar
        borrar.destroy()
    except :
        pass

    #limpiar directorio-------------------------------------
    frame_directorio.destroy()
    frame_directorio = Frame()
    frame_directorio.config(bd=4, relief="groove", bg="#282923")
    frame_directorio.place(x="80", y="100", width="450", height="30")
    directorio = Label(frame_directorio, bg="#282923",
                   text="Archivo:  ...  ", fg="white")
    directorio.place(x="0", y="0")

    #limpiar ajustes --------------------------------------
    frame_comandos.destroy()
    frame_comandos=Frame()
    frame_comandos.pack()
    frame_comandos.config(bd=4,relief="groove",bg="#282923")
    frame_comandos.place(x="20",y="150")
    text_comandos_1=Label(frame_comandos,text="AJUSTES",font="Verdana 12 bold", fg="white", bg="#282923")
    text_comandos_1.grid(row=0,column=0,pady=5,padx=5)
    text_comandos_3=Label(frame_comandos,text="Numero de estaciones:",font="Verdana 10 bold", fg="white", bg="#282923",justify="left")
    text_comandos_3.grid(row=3,column=0,padx=5,pady=2)
    estaciones=tk.Entry(frame_comandos)
    estaciones.grid(row=3,column=1,padx=5,pady=2)

    text_comandos_4=Label(frame_comandos,text="Cantidad de años:",font="Verdana 10 bold", fg="white", bg="#282923",justify="left")
    text_comandos_4.grid(row=4,column=0,padx=5,pady=2)
    years=tk.Entry(frame_comandos)
    years.grid(row=4,column=1,padx=5,pady=2)
    text_comandos_5=Label(frame_comandos,text="Año inicial:",font="Verdana 10 bold", fg="white", bg="#282923",justify="left")
    text_comandos_5.grid(row=5,column=0,padx=5,pady=2)
    years2=tk.Entry(frame_comandos)
    years2.grid(row=5,column=1,padx=5,pady=2)


    
    pass



##############################################################
#----------------------INICIO de la APP-----------------------
##############################################################


# declaramos propiedades de la pagina principal
app = Tk() # Creamos la App
app.title("App calculo doble masa - V.1.0") # titulo
app.geometry("610x400") # geometria inicial
app.resizable(0, 0) # no e sposible  agrandar
app.iconbitmap("icon.ico")
imagen_bg=PhotoImage(file="bg2.png")
fondo=Label(app,image=imagen_bg).place(x=0,y=0)
# app.config(bg="#282923") #  background color

# ---------------------------------------------------------------
#                   Menu
# ---------------------------------------------------------------


barraMenu = Menu(app)
mnuOpciones = Menu(barraMenu)
mnuInicio = Menu(barraMenu)
mnuAYUDA = Menu(barraMenu)
submenu = Menu(mnuOpciones, tearoff=0)

#Menu inicio----------------------------------
mnuInicio=Menu(barraMenu,tearoff=0)
mnuInicio.add_command(label = "Abrir",command=abrir)


mnuInicio.add_separator()
mnuInicio.add_command(label = "Salir",command=salir)

#Menu opciones------------------------------------

mnuOpciones=Menu(barraMenu,tearoff=0)
mnuOpciones.add_command(label = "Limpiar",command=limpiar)


#Menu ayuda--------------------------------------

mnuAYUDA=Menu(barraMenu,tearoff=0)

mnuAYUDA.add_command(label = "Ver Teoria",command=teoria)

mnuAYUDA.add_separator()
mnuAYUDA.add_command(label = "Acerca de ...",command=acerca)

#inicio de los menus-----------------------------
barraMenu.add_cascade(label = "Inicio", menu = mnuInicio)
barraMenu.add_cascade(label = "Opciones", menu = mnuOpciones)
barraMenu.add_cascade(label = "Ayuda", menu = mnuAYUDA)

app.config(menu = barraMenu)



# ------------------------------------------------------------------------------
#   Descripcion - Titulo
# ------------------------------------------------------------------------------

texto=Label(app,text="Calculo de doble masa, Recta de Regresión y",font="Verdana 12 bold", fg="white", bg="#282923")
texto.place(x="20", y="15")
texto=Label(app,text="Correlación Ortogonal ",font="Verdana 12 bold", fg="white", bg="#282923")
texto.place(x="20", y="35")
texto=Label(app,text="Descarga el formato  modificando los ajustes y procesalo ..",font="Verdana 10 bold", fg="white", bg="#282923")
texto.place(x="20", y="70")


# ------------------------------------------------------------------------------
#   Boton DE ABRIR ARCHIVO y nombre del archivo
# ------------------------------------------------------------------------------
file_image3 = PhotoImage(file="logo.png") 
file_image2 = PhotoImage(file="refresh.png") 
file_image = PhotoImage(file="file_32.png") 
file_image3 = file_image3.subsample(2, 2)
boton_abrir = Button(app, image=file_image, cursor="hand2",
                     bg="#282923", command=abrir) 
boton_abrir.place(x="20", y="100")

frame_directorio = Frame()
frame_directorio.config(bd=4, relief="groove", bg="#282923")
frame_directorio.place(x="80", y="100", width="450", height="30")
directorio = Label(frame_directorio, bg="#282923",
                   text="Archivo:  ...  ", fg="white")
directorio.place(x="0", y="0")

# ------------------------------------------------------------------------------
#   Botones
# ------------------------------------------------------------------------------


boton_principal = Button(app, width="20", height="2", text="Descargar Formato",
                                font="Verdana 10 bold", fg="white", cursor="hand2", bg="#282923", command=descargar)
boton_principal.place(x="400", y="150")

boton_teoria= Button(app, width="20", height="2", text="Ver teoria",
                                font="Verdana 10 bold", fg="white", cursor="hand2", bg="#282923", command=teoria)
boton_teoria.place(x="400", y="300")


# ------------------------------------------------------------------------------
#   frame de Ajustes
# ------------------------------------------------------------------------------

frame_comandos=Frame()
frame_comandos.pack()
frame_comandos.config(bd=4,relief="groove",bg="#282923")
frame_comandos.place(x="20",y="150")
text_comandos_1=Label(frame_comandos,text="AJUSTES",font="Verdana 12 bold", fg="white", bg="#282923")
text_comandos_1.grid(row=0,column=0,pady=5,padx=5)

text_comandos_3=Label(frame_comandos,text="Numero de estaciones:",font="Verdana 10 bold", fg="white", bg="#282923",justify="left")
text_comandos_3.grid(row=3,column=0,padx=5,pady=2)
estaciones=tk.Entry(frame_comandos)
estaciones.grid(row=3,column=1,padx=5,pady=2)

text_comandos_4=Label(frame_comandos,text="Cantidad de años:",font="Verdana 10 bold", fg="white", bg="#282923",justify="left")
text_comandos_4.grid(row=4,column=0,padx=5,pady=2)
years=tk.Entry(frame_comandos)
years.grid(row=4,column=1,padx=5,pady=2)
text_comandos_5=Label(frame_comandos,text="Año inicial:",font="Verdana 10 bold", fg="white", bg="#282923",justify="left")
text_comandos_5.grid(row=5,column=0,padx=5,pady=2)
years2=tk.Entry(frame_comandos)
years2.grid(row=5,column=1,padx=5,pady=2)


# ------------------------------------------------------------------------------
#   Fin de App
# ------------------------------------------------------------------------------
app.mainloop()
