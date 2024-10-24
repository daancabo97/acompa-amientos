import pandas as pd

class ProcesadorAcompanamientos:
    def procesar(self, archivo_entrada):
        # Leer archivo de entrada
        df = pd.read_excel(archivo_entrada, sheet_name='Acompañamientos')

        # Limpiar espacios en los nombres de las columnas
        df.columns = df.columns.str.strip()

        # Imprimir las columnas disponibles para depuración
        print("Columnas disponibles en el archivo de entrada:", df.columns)

        # Limpiar los espacios en blanco en la columna "Asignado a"
        if 'Asignado a' not in df.columns:
            raise KeyError(f"La columna 'Asignado a' no se encuentra en el archivo. Columnas encontradas: {df.columns}")

        df['Asignado a'] = df['Asignado a'].str.strip()

        # Agrupar por la columna "Asignado a" y contar los acompañamientos
        conteo_acompanamientos = df.groupby('Asignado a').size().reset_index(name='Cantidad')

        # Calcular el total de acompañamientos
        total_acompanamientos = conteo_acompanamientos['Cantidad'].sum()

        # Añadir la fila con el total
        conteo_acompanamientos.loc[len(conteo_acompanamientos)] = ['Total Acompañamiento', total_acompanamientos]

        return conteo_acompanamientos
