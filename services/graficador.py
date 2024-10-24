
from openpyxl.chart import BarChart, Reference

class Graficador:
    def generar_grafica_barras(self, worksheet, data_range, categoria_range, titulo, celda):
        chart = BarChart()
        chart.title = titulo
        chart.style = 10
        chart.y_axis.title = 'Casos'

        data = Reference(worksheet, min_col=data_range['min_col'], min_row=data_range['min_row'],
                         max_col=data_range['max_col'], max_row=data_range['max_row'])

        categories = Reference(worksheet, min_col=categoria_range['min_col'], min_row=categoria_range['min_row'],
                               max_col=categoria_range['max_col'], max_row=categoria_range['max_row'])

        chart.add_data(data, titles_from_data=True)
        chart.set_categories(categories)
        chart.shape = 4
        chart.width = 30  # Ajustar el ancho de la gráfica
        chart.height = 15  # Ajustar la altura de la gráfica
        worksheet.add_chart(chart, celda)
