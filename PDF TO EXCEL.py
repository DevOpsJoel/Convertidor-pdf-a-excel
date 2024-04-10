import PyPDF2
import pandas as pd
import os
from tqdm import tqdm
import tkinter as tk
from tkinter import filedialog

def extraer_texto_desde_pdf(ruta_pdf):
    texto_extraido = ""
    
    with open(ruta_pdf, 'rb') as archivo_pdf:
        lector_pdf = PyPDF2.PdfReader(archivo_pdf)
        
        for pagina_num in range(len(lector_pdf.pages)):
            pagina = lector_pdf.pages[pagina_num]
            texto_extraido += pagina.extract_text()
    
    return texto_extraido

def convertir_a_excel(texto, ruta_excel):
    try:
        lineas = texto.split('\n')
        
        # Divide cada linea en palabras y elimina lineas vacias
        datos = [linea.split() for linea in lineas if linea.strip()]

        # Determina el número máximo de columnas
        max_col = max(len(linea) for linea in datos)

        # Rellena cada fila para que tenga el mismo número de columnas
        datos_rellenados = [fila + [''] * (max_col - len(fila)) for fila in datos]

        # Crea un DataFrame de pandas
        df = pd.DataFrame(datos_rellenados, columns=[f"Columna{i+1}" for i in range(max_col)])

        # Consolida los nombres en una columna
        if 'Columna15' in df.columns and 'Columna16' in df.columns and 'Columna17' in df.columns:
            df['Nombres'] = df['Columna15'] + ' ' + df['Columna16'] + ' ' + df['Columna17']
            # Elimina las columnas individuales de nombre
            df = df.drop(['Columna15', 'Columna16', 'Columna17'], axis=1)

        # Inicializa la barra de progreso
        total_paginas = len(df)
        with tqdm(total=total_paginas, desc="Procesando") as pbar:
            # Escribe el DataFrame en un archivo Excel en la carpeta de salida
            df.to_excel(ruta_excel, index=False)
            pbar.update(total_paginas)  # Actualiza la barra de progreso al 100%
        
        resultado.config(text=f"Texto extraído del PDF y guardado en {ruta_excel}", fg="green")

    except PermissionError:
        resultado.config(text=f"Error: Permiso denegado para escribir en {ruta_excel}. Asegúrate de que el archivo no esté abierto y que tengas permisos de escritura.", fg="red")
    except Exception as e:
        resultado.config(text=f"Error inesperado: {e}", fg="red")

def seleccionar_archivo():
    ruta_pdf = filedialog.askopenfilename(title="Seleccionar archivo PDF", filetypes=[("Archivos PDF", "*.pdf")])
    if ruta_pdf:
        entrada_ruta_pdf.delete(0, tk.END)
        entrada_ruta_pdf.insert(0, ruta_pdf)

def convertir():
    ruta_del_pdf = entrada_ruta_pdf.get()
    
    if not os.path.exists(ruta_del_pdf):
        resultado.config(text="Error: El archivo PDF no existe. Intente nuevamente.", fg="red")
    else:
        resultado.config(text="")
        nombre_pdf = os.path.basename(ruta_del_pdf)
        nombre_sin_extension = os.path.splitext(nombre_pdf)[0]

        # Crea la carpeta de salida si no existe
        carpeta_salida = os.path.join(os.path.dirname(ruta_del_pdf), "convertidos")
        os.makedirs(carpeta_salida, exist_ok=True)

        # Genera la ruta para el archivo Excel en la carpeta de salida
        ruta_del_excel = os.path.join(carpeta_salida, f"{nombre_sin_extension}.xlsx")
        texto_extraido_del_pdf = extraer_texto_desde_pdf(ruta_del_pdf)

        if texto_extraido_del_pdf:
            convertir_a_excel(texto_extraido_del_pdf, ruta_del_excel)

# Crear la interfaz gráfica pa que se vea bonito
ventana = tk.Tk()
ventana.title("Conversor PDF a Excel")

# Mensaje de bienvenida
bienvenida_label = tk.Label(ventana, text="¡Buen día Janeth!", font=("Arial", 14, "bold"), fg="blue")
bienvenida_label.pack(pady=10)

# Etiqueta e input para la ruta del archivo PDF
etiqueta_ruta_pdf = tk.Label(ventana, text="Ruta del archivo PDF:")
etiqueta_ruta_pdf.pack(pady=5)
entrada_ruta_pdf = tk.Entry(ventana, width=50)
entrada_ruta_pdf.pack(pady=5)
boton_seleccionar = tk.Button(ventana, text="Seleccionar Archivo", command=seleccionar_archivo)
boton_seleccionar.pack(pady=5)

# Botón para iniciar la conversión
boton_convertir = tk.Button(ventana, text="Convertir", command=convertir)
boton_convertir.pack(pady=10)

# Mostrar resultados
resultado = tk.Label(ventana, text="", fg="black")
resultado.pack(pady=10)

# Iniciar el bucle principal de la interfaz gráfica
ventana.mainloop()
