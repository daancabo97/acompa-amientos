
import pandas as pd
from services.procesador import ProcesadorAcompanamientos
from services.graficador import Graficador
from utils.utils import ajustar_ancho_columnas

class ExportadorReporteAcompanamientos:
    def generar(self, archivo_entrada, archivo_salida):
        # Procesar datos
        procesador = ProcesadorAcompanamientos()
        df_acompanamientos = procesador.procesar(archivo_entrada)

        # Exportar a Excel
        with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
            df_acompanamientos.to_excel(writer, sheet_name='Acompañamientos', index=False)
            
            # Ajustar el ancho de las columnas
            ajustar_ancho_columnas(writer, 'Acompañamientos')

            # Crear la gráfica de barras
            workbook = writer.book
            worksheet = workbook['Acompañamientos']
            
            data_range = {
                'min_col': 2,
                'min_row': 1,
                'max_col': 2,
                'max_row': len(df_acompanamientos)
            }
            categoria_range = {
                'min_col': 1,
                'min_row': 2,
                'max_col': 1,
                'max_row': len(df_acompanamientos)
            }

            graficador = Graficador()
            graficador.generar_grafica_barras(worksheet, data_range, categoria_range, 'Acompañamientos por Persona', 'E5')
