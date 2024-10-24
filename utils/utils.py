def ajustar_ancho_columnas(writer, sheet_name):
    workbook = writer.book
    worksheet = workbook[sheet_name]
    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter
        for celda in col:
            try:
                if len(str(celda.value)) > max_length:
                    max_length = len(str(celda.value))
            except Exception:  # Capturar excepciones correctamente
                pass
        ajustar_anchura = (max_length + 2)
        worksheet.column_dimensions[column].width = ajustar_anchura
