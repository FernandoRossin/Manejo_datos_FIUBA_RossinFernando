
'''
Es necesario instalar: 
        -Python
        -Pandas: pip install pandas ( en cmd)
        -Pypxl:pip install openpyxl (en cmd)
'''

import sys
import tkinter
import pandas as pd
from tkinter.filedialog import askopenfilename,askopenfilenames 
from tkinter import *
import shutil

def cambiar_archivos():
    
    ubicacion_archivos = askopenfilenames()

    if ubicacion_archivos == '':
        pass
    else:  
        for i in ubicacion_archivos:
            direcc = str(i)
            respuesta= ''
            seleccion_archivo(direcc,respuesta)
            archivo_creado()

def get_nombre():
        global newname
        newname = text.get()
        res.destroy()
        seleccion_archivo(ubicacion,newname)
        archivo_creado()
        

def cambiar_nombre():
    global ubicacion
    ubicacion = askopenfilename()
    if ubicacion == '':
        pass
    else:  
        global res
        res = Toplevel(root)
        res.title('SAL_to_Excel by Fernando Rossin')
        res.geometry('500x100')
        x = root.winfo_x()
        y = root.winfo_y()
        res.geometry("+%d+%d" %(x,y+150))
        res.resizable(width=0,height=0)
        res.focus()
        res.grab_set()
        res.transient(master=root)
        mensaje2 = Label(res,text='Renombre archivo (opcional):')
        mensaje2.place(x=10,y=20)
        global text
        text=Entry(res,width=30)
        text.place(x=190,y=20)
        res.bind('<Return>',lambda event:get_nombre())
        btnRead=Button(res, height=1, width=15, text="Convertir")
        btnRead.place(x=80,y=50)
        btnRead.bind('<Button-1>',lambda event:get_nombre())
        btncan=Button(res, height=1, width=15, text="Cancelar",command=cancelar)
        btncan.place(x=280,y=50)
        res.wait_window(res)


def seleccion_archivo(filename,respuesta):
    
    df = pd.read_csv(filename, header= 0,encoding= 'unicode_escape')
    filenamext = filename.split(".")
    filenamespl = filenamext[0]
    filenamespl = filenamespl.rsplit("/",1)
    ruta = filenamespl[0] +'/'
    
    if respuesta == '':
        respuesta = filenamespl[1] + (".xlsx")
    else:
        respuesta = respuesta + (".xlsx")
    
    df.to_excel(respuesta, index=False)

    source = respuesta
    destination = ruta + respuesta
    shutil.move(source,destination)

def cancelar():
    res.destroy()

def salir():
    sys.exit(1)

def archivo_creado():
    arcreado = tkinter.Tk()
    arcreado.title('SAL_to_Excel by Fernando Rossin')
    # arcreado.geometry('300x200+400+250')
    arcreado.geometry('300x200')
    x = root.winfo_x()
    y = root.winfo_y()
    arcreado.geometry("+%d+%d" %(x+100,y+50))
    arcreado.config(background='black')
    titulo1 = tkinter.Label(arcreado,text='Archivo',font='Helvetica 30',fg='red', bg='black')
    titulo1.place(relx=0.5,y=70, anchor=CENTER)
    titulo2 = tkinter.Label(arcreado,text='Convertido',font='Helvetica 30',fg='red', bg='black')
    titulo2.place(relx=0.5,y=120, anchor=CENTER)
    arcreado.after(1000,lambda:arcreado.destroy())
    

        
newname = ''
root = tkinter.Tk()
root.title('SAL_to_Excel by Fernando Rossin')
root.geometry('500x300+300+200')
root.resizable(width=0,height=0)
titulo = tkinter.Label(root,text='SAL to Excel',font='Helvetica 50',fg='red',bg='black')
titulo.place(relx=0.5,y=60, anchor=CENTER)
mensaje = tkinter.Label(root,text='Seleccione el/los archivo que desea convertir:',font='Helvetica 12')
mensaje.place(relx=0.5,y=160,anchor=CENTER)
boton1 = tkinter.Button(root,text='archivo',command=cambiar_nombre,width=15)
boton1.place(relx=0.25,y=200,anchor=CENTER)
boton3 = tkinter.Button(root,text='selección múltiple',command=cambiar_archivos,width=15)
boton3.place(relx=0.75,y=200,anchor=CENTER)
boton2 = tkinter.Button(root,text='Salir',command=salir,width=15)
boton2.place(relx=0.5,y=280, anchor= CENTER)

root.mainloop()







































