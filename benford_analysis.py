# Graficas
import matplotlib
matplotlib.use('TkAgg')
# Interfaz Grafica
import tkinter as tk
from tkinter import filedialog 
from tkinter import *
from tkinter import ttk
# Archivos, calculos y otros
import pandas as pd
from benfordslaw import benfordslaw
import string
import getpass
import os
import csv
import re
# Estadisticas
from statistics import multimode



# Función para manejar el evento de clic en el botón "Seleccionar archivo"
def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path) 

    # Obtener los nombres de las columnas del archivo seleccionado
    if file_path:
        df = pd.read_excel(file_path)
        column_names = df.columns.tolist()
        columns_combobox['values'] = column_names 

# -------------------------------------------------------------------------------------------------------------------------- 


# Función para calcular la distribución de Benford para una columna
def calculate_benford_distribution(data_column):
    
    # Obtener el usuario del equipo
    username = getpass.getuser()
    # Ruta de acceso donde se guardara las imagenes
    directorio = f"/Users/{username}/Desktop/Resultados Ley Benford"
    # Creacion de la carpeta donde se guadara los resultados de no existir
    if not os.path.exists(directorio):
        os.makedirs(directorio)
    
    bl = benfordslaw(alpha=0.05)
    bl.fit(data_column.values)

    # titles = f"Grafica Ley de primer digito: {columns[current_column_index]}"
    titles = f"Grafica Ley de primer digito: {columns}"
    out = titles.translate(str.maketrans('', '', string.punctuation))
    path = f"{directorio}/{out}.png"

    # Generar grafica
    fig , _ =bl.plot(title=out, barcolor=[0.4, 0.1, 0.5], fontsize=16, barwidth=0.6)

    # Guardar la figura en un archivo
    fig.savefig(path)


# -------------------------------------------------------------------------------------------------------------------------- 


# Función para exportar datos a un archivo CSV
def export_data():
    global columns

    # Obtener el usuario del equipo
    username = getpass.getuser()
    titulo = columns
    titulo_limpio = titulo.translate(str.maketrans('', '', string.punctuation))
    # Ruta de acceso donde se guardará el archivo CSV
    file_path = f"/Users/{username}/Desktop/Resultados Ley Benford/datos_{titulo_limpio}.csv"

    # Capturar el contenido del widget Text
    text_content = result_text.get("1.0", tk.END)

    # Separar el contenido por líneas
    lines = text_content.split("\n")

    # Eliminar líneas en blanco
    lines = [line.strip() for line in lines if line.strip()]

    # Crear una lista para almacenar las filas de datos
    data_rows = []

    # Procesar cada línea para obtener los valores de las columnas
    for line in lines:
        # Eliminar las tabulaciones adicionales y dejar solo una tabulación
        line = re.sub(r'\t+', '\t', line)

        # Separar la línea en columnas
        columns = line.split("\t")

        # Agregar la fila de datos a la lista
        data_rows.append(columns)

    # Escribir el contenido en el archivo CSV
    with open(file_path, "w", newline="") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerows(data_rows)

    # Mostrar un mensaje de éxito
    result_text.insert(tk.END, "\n¡Datos exportados exitosamente a un archivo CSV!")


# -------------------------------------------------------------------------------------------------------------------------- 


# Función para manejar el evento de clic en el botón "Aplicar ley de Benford"
def apply_benford_law():

    global columns, data_frame

    file_path = file_entry.get()
    columns=columns_combobox.get()

    # Cargar el archivo Excel
    try:
        data_frame = pd.read_excel(file_path)
    except Exception as e:
        result_text.delete("1.0", tk.END)
        result_text.insert(tk.END, f"Error al cargar el archivo Excel: {str(e)}")
        return
    

    # Filtrar valores no numéricos o en blanco
    invalid_rows = data_frame[columns].apply(lambda x: not isinstance(x, (float, int)))
    filtered_data_frame = data_frame[pd.to_numeric(data_frame[columns], errors='coerce').notnull()]


    # Mostrar información sobre valores no numéricos o en blanco
    result_text.insert(tk.END, f"Columna: {columns}\n\n")
    result_text.insert(tk.END, "NORMALIZACION DE DATOS\n\n")
    result_text.insert(tk.END, f"   Total de filas iniciales\t\t\t{data_frame.shape[0]}\n")
    result_text.insert(tk.END, f"   Numero de filas invalidas\t\t\t{invalid_rows.sum()}\n\n")

    data_column = filtered_data_frame[columns]  

    # Mostrar información de resumen estadistico
    result_text.insert(tk.END, "RESUMEN ESTADISTICO\n\n")
    result_text.insert(tk.END, f"   Cantidad de datos\t\t\t{round(data_column.shape[0],3)}\n")
    result_text.insert(tk.END, f"   Maximo\t\t\t{round(data_column.max(),3)}\n")
    result_text.insert(tk.END, f"   Minimo\t\t\t{round(data_column.min(),3)}\n") 
    result_text.insert(tk.END, f"   Media\t\t\t{round(data_column.mean(),3)}\n")
    result_text.insert(tk.END, f"   Desviacion estandar\t\t\t{round(data_column.std(),3)}\n") 
    result_text.insert(tk.END, f"   Moda\t\t\t{multimode(data_column)}\n")

    # Crear el botón de exportar
    export_button = tk.Button(window, text="Exportar datos", command=export_data, bg='#EEEEEE', fg='#000000' ,font=('Segoe UI Semibold', 12))
    export_button.pack()

    calculate_benford_distribution(data_column)


# --------------------------------------------------------------------------------------------------------------------------                


# Crear la ventana de la interfaz gráfica
window = tk.Tk()
window.title("Análisis de Benford")
window.configure(background='#2C2C2C')

# Etiqueta y campo de entrada para el archivo de Excel
file_label = tk.Label(window, text="Archivo de Excel:", bg='#2C2C2C' ,fg='#FFFFFF' ,font=('Segoe UI Semibold', 18))
file_label.pack()
file_entry = tk.Entry(window, width=70,bg='#1C1C1C',fg='#FFFFFF')
file_entry.pack(ipadx=2,ipady=2)
file_button = tk.Button(window, text="Seleccionar archivo", height=0, command=select_file, bg='#EEEEEE', fg='#000000' ,font=('Segoe UI Semibold', 12))
file_button.pack(pady=10, ipady=0)

# Define the style for combobox widget
style= ttk.Style()
style.theme_use('clam')
style.configure("TCombobox", fieldbackground= "#504C54", background= "#504C54", fg='#FFFFFF' )

# Etiqueta y campo de entrada para las columnas
columns_label = tk.Label(window, text="Nombre de la Columna a analizar:\n", height=2,bg='#2C2C2C'  , fg='#FFFFFF'  ,font=('Segoe UI Semibold', 18))
columns_label.pack(pady=(15,0))

columns_combobox = ttk.Combobox(window, width=48,height=0, values=[], font=('Segoe UI Semibold', 12))
columns_combobox.pack(pady=(0,0))

# Botón para aplicar la ley de Benford
apply_button = tk.Button(window, text="Aplicar ley de Benford", command=apply_benford_law, bg='#EEEEEE', fg='#000000' ,font=('Segoe UI Semibold', 12))
apply_button.pack(pady=10)

# Crear un widget de texto para mostrar el resultado
result_text = tk.Text(window, height=20, width=60, bg='#1C1C1C', fg='#FFFFFF' ,font=('Segoe UI Semibold', 14))
result_text.pack()


# Iniciar la interfaz gráfica
window.mainloop()