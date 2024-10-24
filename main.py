import argparse
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from factories.reporte_factory import ReporteFactory

# Función para generar el reporte
def ejecutar_proceso(archivo_entrada, archivo_salida):
    try:
        # Crear el reporte mediante la fábrica
        reporte = ReporteFactory.crear_reporte('acompañamientos')
        reporte.generar(archivo_entrada, archivo_salida)
        print(f"Reporte generado correctamente en {archivo_salida}")
    except Exception as e:
        print(f"Error generando el reporte: {str(e)}")

# Función para manejar el modo línea de comandos
def modo_comando():
    parser = argparse.ArgumentParser(description="Generar reporte de acompañamientos.")
    parser.add_argument('--input', type=str, help="Ruta del archivo de entrada (opcional)")
    parser.add_argument('--output', type=str, help="Ruta del archivo de salida (opcional)")
    args = parser.parse_args()

    # Si no se especifica archivo de entrada, buscar uno en el directorio actual
    if args.input is None:
        archivos = [f for f in os.listdir('.') if f.endswith('.xlsx') and 'Gestion_Septiembre' in f]
        if archivos:
            archivo_entrada = archivos[0]  # Toma el primer archivo que coincide
        else:
            raise FileNotFoundError("No se encontró un archivo de entrada por defecto.")
    else:
        archivo_entrada = args.input

    # Si no se especifica archivo de salida, usar un nombre por defecto
    archivo_salida = args.output if args.output else "reporte.xlsx"

    # Ejecutar el proceso
    ejecutar_proceso(archivo_entrada, archivo_salida)

# Función para seleccionar archivo de entrada usando la GUI
def seleccionar_archivo_entrada():
    ruta_archivo = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
    if ruta_archivo:
        entrada_var.set(ruta_archivo)

# Función para seleccionar archivo de salida usando la GUI
def seleccionar_archivo_salida():
    ruta_salida = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivos de Excel", "*.xlsx")])
    if ruta_salida:
        salida_var.set(ruta_salida)

# Función para ejecutar el proceso desde la GUI
def ejecutar():
    ruta_archivo = entrada_var.get()
    ruta_salida = salida_var.get()
    if ruta_archivo and ruta_salida:
        try:
            ejecutar_proceso(ruta_archivo, ruta_salida)
            messagebox.showinfo("Éxito", f"Reporte generado correctamente en {ruta_salida}")
        except Exception as e:
            messagebox.showerror("Error", f"Error generando el reporte: {str(e)}")
    else:
        messagebox.showwarning("Advertencia", "Debe seleccionar tanto un archivo de entrada como un archivo de salida.")

# Función para iniciar la interfaz gráfica
def iniciar_gui():
    global entrada_var, salida_var
    app = tk.Tk()
    app.title("Generador de Reportes Mensuales")
    app.geometry("400x250")

    entrada_var = tk.StringVar()
    salida_var = tk.StringVar()

    tk.Label(app, text="Archivo de entrada:").pack(pady=5)
    tk.Entry(app, textvariable=entrada_var, width=50).pack(pady=5)
    tk.Button(app, text="Seleccionar archivo", command=seleccionar_archivo_entrada).pack(pady=5)

    tk.Label(app, text="Archivo de salida:").pack(pady=5)
    tk.Entry(app, textvariable=salida_var, width=50).pack(pady=5)
    tk.Button(app, text="Seleccionar ruta", command=seleccionar_archivo_salida).pack(pady=5)

    tk.Button(app, text="Ejecutar", command=ejecutar).pack(pady=20)

    app.mainloop()

# Punto de entrada principal
if __name__ == "__main__":
    # Si se proporcionan argumentos desde la línea de comandos, ejecutar en modo comando
    if len(os.sys.argv) > 1:
        modo_comando()
    else:
        # Si no se proporcionan argumentos, iniciar la GUI
        iniciar_gui()
