# -*- coding: utf-8 -*-
"""
Created on Sun Jul 11 10:58:07 2021

@author: Felipe
"""

import xlwings as xw
import tkinter as tk
from xlwings.constants import DeleteShiftDirection
from tkinter import *
from idlelib.tooltip import Hovertip



class CustomHovertip(Hovertip):
    def showcontents(self):
        label = tk.Label(
            self.tipwindow, text=f' "{self.text}" ', justify=tk.LEFT,
            bg="#151515", fg="#ffffff", relief=tk.SOLID, borderwidth=1,
            font=("Times New Roman", 12)
            )
        label.pack()
        
        #La "f" es un literal String (desde Python 3 es así) y permite agregar un texto proveniente desde otra variable dentro de las llaves {}
        
      




raiz = Tk()

#raiz.iconbitmap("imagen_Iexcel.ico")
raiz.title("Ventana de Configuración")
#raiz.config(bg="linen")
#raiz.config(bd=10)
#raiz.geometry("400x300")
#raiz.resizable(False,False)

miMarco = Frame(raiz, bg="linen")
miMarco.pack()



def codigoLimpieza():
    
    
    libroEnEliminacion = libro_eliminar.get()
    hojaEnEliminacion = hoja_eliminar.get()
    columnasEnEliminacion = columnas_eliminar.get()
    filasEnEliminacion = filas_eliminar.get()
    
    
    limpiador = xw.App()
    libro_trabajo = limpiador.books.open(libroEnEliminacion) 
    hoja_trabajo = libro_trabajo.sheets[hojaEnEliminacion] 
    hoja_trabajo.range(columnasEnEliminacion).api.Delete(DeleteShiftDirection.xlShiftToLeft)
    hoja_trabajo.range(filasEnEliminacion).api.Delete(DeleteShiftDirection.xlShiftUp)
    
    libro_eliminar.set(" ")
    hoja_eliminar.set(" ")
    columnas_eliminar.set(" ")
    filas_eliminar.set(" ")
    
    
    libro_trabajo.save()
    limpiador.kill()
    
  
        
        
        
    
    
        
        
#-----------------------------------------------------------

libro_eliminar = StringVar()
hoja_eliminar = StringVar()
columnas_eliminar = StringVar()
filas_eliminar = StringVar()



libroEliminar=Entry(miMarco, textvariable = libro_eliminar, width=40)
libroEliminar.grid(row=1, column=2, padx=10, pady=10)
CustomHovertip(libroEliminar, text="Ejemplo de ruta al archivo ( agregar \ a la ruta ejemplo )=C:\\Users\\Usuario\\Carpeta1\\Carpeta2\\Archivo.xlsx", hover_delay=500)

hojaEliminar=Entry(miMarco, textvariable = hoja_eliminar, width=40)
hojaEliminar.grid(row=2, column=2, padx=10, pady=10)
CustomHovertip(hojaEliminar, text="No requiere formatos especiales, sólo el nombre que lleve la hoja de cálculo dentro del archivo", hover_delay=500)

columnasEliminar=Entry(miMarco, textvariable = columnas_eliminar, width=40)
columnasEliminar.grid(row=3, column=2, padx=10, pady=10)
CustomHovertip(columnasEliminar, text="Ejemplo de formato de rango de columnas a eliminar=B:D", hover_delay=500)

filasEliminar=Entry(miMarco, textvariable = filas_eliminar, width=40)
filasEliminar.grid(row=4, column=2, padx=10, pady=10)
CustomHovertip(filasEliminar, text="Ejemplo de formato de rango de filas a eliminar=1:13", hover_delay=500)



tituloPrograma=Label(miMarco, text="Modificador Archivos Excel", font=("Arial", 15))
tituloPrograma.grid(row=0, column=0, columnspan=3, padx=5, pady=10)
tituloPrograma.config(fg="navy", bg="linen")


libroPorEliminar=Label(miMarco, text="Libro de trabajo:")
libroPorEliminar.grid(row=1, column=0, sticky="w", padx=5, pady=5)
libroPorEliminar.config(fg="blue2", bg="linen")


hojaPorEliminar=Label(miMarco, text="Hoja a trabajar:")
hojaPorEliminar.grid(row=2, column=0, sticky="w", padx=5, pady=5)
hojaPorEliminar.config(fg="blue2", bg="linen")

columnasPorEliminar=Label(miMarco, text="Columnas a eliminar:")
columnasPorEliminar.grid(row=3, column=0, sticky="w", padx=5, pady=5)
columnasPorEliminar.config(fg="blue2", bg="linen")

filasPorEliminar=Label(miMarco, text="Filas a eliminar:")
filasPorEliminar.grid(row=4, column=0, sticky="w", padx=5, pady=5)
filasPorEliminar.config(fg="blue2", bg="linen")








botonLibro=Button(miMarco, text="Ejecutar Modificación", command=codigoLimpieza, bg="midnight blue", fg="snow", font=("Helvetica", "12") )
#botonLibro.pack()
botonLibro.grid(row=5, column=0, padx=10, pady=30, columnspan=3)










raiz.mainloop()

