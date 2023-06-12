# Graficas
import matplotlib
matplotlib.use('TkAgg')
# Interfaz Grafica
import tkinter as tk
from tkinter import filedialog 
from tkinter import *
# Archivos, calculos y otros
import pandas as pd
from benfordslaw import benfordslaw
import string
import getpass
import os


current_column_index = 0


# Función para manejar el evento de clic en el botón "Seleccionar archivo"
def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path)

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

    titles = f"Grafica Ley de primer digito: {columns[current_column_index]}"
    out = titles.translate(str.maketrans('', '', string.punctuation))
    path = f"{directorio}/{out}.png"

    # Generar grafica
    fig , _ =bl.plot(title=out, barcolor=[0.4, 0.1, 0.5], fontsize=16, barwidth=0.6)
    # Guardar la figura en un archivo
    fig.savefig(path)

# -------------------------------------------------------------------------------------------------------------------------- 

# Función para generar la siguiente gráfica
def next_graph():
    global current_column_index
    current_column_index += 1

    invalid_rows = data_frame[columns[current_column_index]].apply(lambda x: not isinstance(x, (float, int)))
    filtered_data_frame = data_frame[pd.to_numeric(data_frame[columns[current_column_index]], errors='coerce').notnull()]

    # Mostrar información sobre valores no numéricos o en blanco
    result_text.insert(tk.END, f"Columna: {columns[current_column_index]}\n\n")
    result_text.insert(tk.END, "Información sobre valores no numéricos o en blanco:\n")
    result_text.insert(tk.END, f"Total de filas: {data_frame.shape[0]}\n")
    result_text.insert(tk.END, f"Filas con valores no numéricos o en blanco: {invalid_rows.sum()}\n\n")

    data_column = filtered_data_frame[columns[current_column_index]]

    if current_column_index <= len(columns):
        calculate_benford_distribution(data_column)

# -------------------------------------------------------------------------------------------------------------------------- 

# Función para manejar el evento de clic en el botón "Aplicar ley de Benford"
def apply_benford_law():

    global columns, data_frame

    file_path = file_entry.get()
    columns_input = columns_entry.get()
    columns = [col.strip() for col in columns_input.split(",")]

    # Cargar el archivo Excel
    try:
        data_frame = pd.read_excel(file_path)
    except Exception as e:
        result_text.delete("1.0", tk.END)
        result_text.insert(tk.END, f"Error al cargar el archivo Excel: {str(e)}")
        return
    
    # Botón para generar la siguiente gráfica
    if(len(columns) > 1):
        next_button = tk.Button(window, text="Siguiente resultado", command=next_graph)
        next_button.pack()

    # Filtrar valores no numéricos o en blanco
    invalid_rows = data_frame[columns[current_column_index]].apply(lambda x: not isinstance(x, (float, int)))
    filtered_data_frame = data_frame[pd.to_numeric(data_frame[columns[current_column_index]], errors='coerce').notnull()]

    # Mostrar información sobre valores no numéricos o en blanco
    result_text.insert(tk.END, f"Columna: {columns[current_column_index]}\n\n")
    result_text.insert(tk.END, "Información sobre valores no numéricos o en blanco:\n")
    result_text.insert(tk.END, f"Total de filas: {data_frame.shape[0]}\n")
    result_text.insert(tk.END, f"Filas con valores no numéricos o en blanco: {invalid_rows.sum()}\n\n")

    data_column = filtered_data_frame[columns[current_column_index]]

    calculate_benford_distribution(data_column)

# --------------------------------------------------------------------------------------------------------------------------                

# Crear la ventana de la interfaz gráfica
window = tk.Tk()
window.title("Análisis de Benford")

# Etiqueta y campo de entrada para el archivo de Excel
file_label = tk.Label(window, text="Archivo de Excel:")
file_label.pack()
file_entry = tk.Entry(window, width=50)
file_entry.pack()
file_button = tk.Button(window, text="Seleccionar archivo", command=select_file)
file_button.pack()

# Etiqueta y campo de entrada para las columnas
columns_label = tk.Label(window, text="Nombre de la/s Columna/s a analizar:\n(Separar con , sin espacios)")
columns_label.pack()
columns_entry = tk.Entry(window, width=50)
columns_entry.pack()

# Botón para aplicar la ley de Benford
apply_button = tk.Button(window, text="Aplicar ley de Benford", command=apply_benford_law)
apply_button.pack()

# Crear un widget de texto para mostrar el resultado
result_text = tk.Text(window, height=10, width=50)
result_text.pack()

# Iniciar la interfaz gráfica
window.mainloop()