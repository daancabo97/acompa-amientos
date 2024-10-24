
from services.exportador import ExportadorReporteAcompanamientos

class ReporteFactory:
    @staticmethod
    def crear_reporte(tipo):
        if tipo == 'acompañamientos':
            return ExportadorReporteAcompanamientos()
        else:
            raise ValueError(f"Tipo de reporte no soportado: {tipo}")
