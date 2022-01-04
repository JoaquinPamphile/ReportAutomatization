import pandas as pd
from tkinter import *
from tkinter import filedialog
import sys
import os

def openFile():
    global salida
    global path
    try:
        filepath = filedialog.askopenfilename()

        Label(text="Archivo cargado: ", bg="light gray", fg="black", width="250", height="1", font=("Calibri", 15)).pack()
        Label(text="").pack()

        Label(text=filepath, bg="light gray", fg="black", width="250", height="1", font=("Calibri", 15)).pack()
        Label(text="").pack()

        i = 0
        fechas = []
        gravados = []
        exentos = []
        ivas = []
        totales = []
        primeras = []
        ultimas = []
        df = pd.read_excel(filepath)
        fechas_unicas = df['Fecha'].unique()
        for fu in fechas_unicas:
            fechas.append(fu)
            gravado = 0
            exento = 0
            iva = 0
            total = 0
            z = 0
            for f in df['Fecha']:
                if f == fu:
                    if z == 0 and 'FC B' in df['Comprobante'][i]:
                        primero = df['Comprobante'][i]
                        z = z + 1
                    if z != 0 and 'FC B' in df['Comprobante'][i]:
                        ultimo = df['Comprobante'][i]
                    gravado = gravado + df['Gravado'][i]
                    exento = exento + df['Exento'][i]
                    iva = iva + df['IVA'][i]
                    total = total + df['Total'][i]
                    i = i + 1
            gravados.append(gravado)
            exentos.append(exento)
            ivas.append(iva)
            totales.append(total)
            primeras.append(primero)
            ultimas.append(ultimo)

        salida = pd.DataFrame()
        salida['Fecha'] = fechas
        salida['Gravado'] = gravados
        salida['Exento'] = exentos
        salida['IVA'] = ivas
        salida['Total'] = totales
        salida['Primer Comprobante Dia'] = primeras
        salida['Ultimo Comprobante Dia'] = ultimas

        path = Entry(pantalla, width="70", bd=5)
        path.insert(0,"Ingrese nombre de archivo de salida...")
        path.configure(state=DISABLED)
        path.pack()
        path.bind('<Button-1>', on_click)
        Label(text="").pack()
        Button(text="Guardar", bg="LavenderBlush4", height="3", width="30",command=saveFile).pack()
    except:
        python = sys.executable
        os.execl(python, python, * sys.argv)


def on_click(event):
    path.configure(state=NORMAL)
    path.delete(0, END)

    # make the callback only work once
    path.unbind('<Button-1>')

def saveFile():
    salida.to_excel(path.get()+'.xlsx', header=True, index=False)

    Label(text="").pack()
    Label(text="Archivo Generado: ", bg="light gray", fg="black", width="250", height="1", font=("Calibri", 15)).pack()
    Label(text="").pack()

    Label(text=path.get(), bg="light gray", fg="black", width="250", height="1", font=("Calibri", 15)).pack()
    Label(text="").pack()

    Button(text="Salir", bg="LavenderBlush4", height="3", width="30", command=exit).pack()

def exit():
    sys.exit()
    
#INTERFAZ GRAFICA
def menu_pantalla():
    global pantalla
    pantalla = Tk()
    pantalla.geometry("700x520")
    pantalla.title("Bienvenida Irene!")
    pantalla.configure(background="lavender")

    Label(text="Por favor seleccione el archivo", bg="light gray", fg="black", width="250", height="1", font=("Calibri", 15)).pack()
    Label(text="").pack()

    Button(text="Seleccionar file", bg="LavenderBlush4", height="3", width="30", command=openFile).pack()
    Label(text="").pack()

    pantalla.mainloop()

menu_pantalla()