import os
import re
from flask import Flask, request, render_template, send_file
import pdfplumber
import pandas as pd
from werkzeug.utils import secure_filename
import shutil
import zipfile
import matplotlib
matplotlib.use('Agg')  # Use non-interactive backend
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.backends.backend_pdf import PdfPages
import numpy as np

app = Flask(__name__)

# Crear directorios para archivos temporales
# En producción, usar directorio temporal del sistema si es necesario
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'tmp')
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    # Asegurar permisos de escritura
    os.chmod(UPLOAD_FOLDER, 0o777)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

def cleanup_old_files():
    """Limpia archivos temporales antiguos"""
    for filename in os.listdir(UPLOAD_FOLDER):
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        try:
            if os.path.isfile(filepath):
                os.unlink(filepath)
            elif os.path.isdir(filepath):
                shutil.rmtree(filepath)
        except Exception as e:
            print(f"Error limpiando archivo {filepath}: {e}")

def process_pdf(pdf_path):
    import re
    import pdfplumber
    import pandas as pd

    with pdfplumber.open(pdf_path) as pdf:
        # Extract table from first page with data
        for page in pdf.pages:
            tables = page.extract_tables()
            if tables:
                raw_table = tables[0]
                break

        # Convert to DataFrame
        df_raw = pd.DataFrame(raw_table)

        # Process headers
        header = df_raw.iloc[0].tolist()
        df_data = df_raw.iloc[1:].reset_index(drop=True)

        # Process data into lists
        lists = []
        for col in range(df_data.shape[1]):
            values = '\n'.join(df_data.iloc[:, col].values)
            values_list = [x.strip() for x in values.strip().split('\n') if x.strip() != '']
            lists.append(values_list)

        # Get maximum rows
        n_rows = max(len(l) for l in lists)

        # Fill shorter lists
        lists = [l + ['']*(n_rows-len(l)) for l in lists]

        # Create final DataFrame
        df_final = pd.DataFrame({header[i]: lists[i] for i in range(len(header))})

        # Extract totals using regex
        text = ""
        for page in pdf.pages:
            text += page.extract_text() + "\n"

        # Patrón regex para encontrar los totales
        pattern = r'(Subtotal Cotización|Bonificación|Subtotal Neto|IVA|Total Cotización)\s*:\s*([\d.,]+)'
        matches = re.findall(pattern, text)

        # Convertimos los resultados en diccionario
        resultados = {label: float(valor.replace(',', '').replace('.', '', valor.count('.')-1)) for label, valor in matches}

        # Calcular porcentaje de bonificación correctamente
        if "Bonificación" in resultados and "Subtotal Cotización" in resultados:
            try:
                pct_bonificacion = resultados["Bonificación"] / resultados["Subtotal Cotización"]
            except ZeroDivisionError:
                pct_bonificacion = 0
        else:
            pct_bonificacion = 0

        # Convert numeric columns
        df_final["Cantidad"] = pd.to_numeric(df_final["Cantidad"], errors="coerce")
        df_final["Precio_Unit"] = pd.to_numeric(df_final["Precio Unit"], errors="coerce")
        df_final["% Desc."] = pd.to_numeric(df_final["% Desc."], errors="coerce").fillna(0)
        df_final["% IVA"] = pd.to_numeric(df_final["% IVA"], errors="coerce").fillna(0)
        df_final["Importe"] = pd.to_numeric(df_final["Importe"], errors="coerce").fillna(0)

        # Calcular Precio Lista (sin bonificación ni descuento)
        df_final["Precio_Lista"] = df_final["Cantidad"] * df_final["Precio_Unit"]

        # Calcular Precio Neto (aplicando bonificación sobre el precio con descuento)
        df_final["Precio Neto"] = df_final["Importe"] * (1 - pct_bonificacion)

        # Calcular Precio Neto Unitario
        df_final["Precio Neto Unitario"] = df_final["Precio Neto"] / df_final["Cantidad"]

        # Calcular IVA: Precio Neto * (%IVA/100)
        df_final["IVA Calculado"] = df_final["Precio Neto"] * (df_final["% IVA"] / 100)

        # Calcular Precio con Impuestos
        df_final["Precio con Impuestos"] = df_final["Precio Neto"] + df_final["IVA Calculado"]

        # Para mostrar en la tabla, el precio unitario neto
        df_final["Precio"] = df_final["Precio Neto Unitario"]

        # Calcular Subtotal: suma de Precio Neto
        subtotal = round(pd.to_numeric(df_final["Precio Neto"], errors="coerce").sum(), 4)

        # Calcular IVA por tipo
        precio_neto_21 = pd.to_numeric(df_final.loc[df_final["% IVA"] == 21, "Precio Neto"].replace("", 0), errors="coerce").fillna(0)
        iva_21 = round((precio_neto_21 * 0.21).sum(), 4)
        precio_neto_105 = pd.to_numeric(df_final.loc[df_final["% IVA"] == 10.5, "Precio Neto"].replace("", 0), errors="coerce").fillna(0)
        iva_105 = round((precio_neto_105 * 0.105).sum(), 4)

        # Calcular Total: suma de Precio con Impuestos
        total = round(pd.to_numeric(df_final["Precio con Impuestos"], errors="coerce").sum(), 4)

        # Crear DataFrame resumen
        resumen_df = pd.DataFrame({
            "Concepto": ["Subtotal", "IVA 21%", "IVA 10.5%", "Total"],
            "Importe": [subtotal, iva_21, iva_105, total]
        })

        print(pct_bonificacion)
        # Crear DataFrame de resultado final
        df_result = df_final[["Descripción Artículo", "Desc. Adicional", "Cantidad", "Precio", "% IVA", "Precio Neto"]].copy()
        df_result["Cantidad"] = df_result["Cantidad"].round(4)
        df_result["Precio"] = df_result["Precio"].round(4)
        df_result["% IVA"] = df_result["% IVA"].round(4)
        df_result["Precio Neto"] = df_result["Precio Neto"].round(4)

        return df_result, resumen_df

def generate_excel(df_result, resumen_df, razon_social, cuit, nro_cotizacion, fecha):
    # Create Excel file in our tmp directory
    excel_filename = f"temp_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
    
    # Write to Excel with formatting
    with pd.ExcelWriter(excel_path, engine="xlsxwriter") as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('Hoja1')
        
        # Ocultar líneas de grilla
        worksheet.hide_gridlines(2)  # Ocultar líneas de grilla en pantalla e impresión
        
        # Configurar altura de filas para el header
        worksheet.set_row(0, 35)  # Altura aumentada para el texto más grande
        worksheet.set_row(1, 45)  # Altura aumentada para el logo más grande
        
        # Crear franja azul para el header
        header_bg_format = workbook.add_format({
            'bg_color': '#14324B'
        })
        
        # Aplicar fondo azul a las primeras dos filas
        for col in range(12):  # A hasta L
            worksheet.write(0, col, '', header_bg_format)
            worksheet.write(1, col, '', header_bg_format)
        
        # Crear header con el nombre de la empresa y cotización
        # Formato para el nombre de la empresa
        company_name_format = workbook.add_format({
            'bold': True,
            'font_size': 28,  # Tamaño más grande para el nombre
            'font_name': 'Montserrat',  # Tipografía más moderna
            'align': 'center',
            'valign': 'vcenter',
            'font_color': 'white',
            'bg_color': '#14324B'
        })
        
        # Formato para "Cotización"
        subtitle_format = workbook.add_format({
            'font_size': 16,  # Tamaño más pequeño para el subtítulo
            'font_name': 'Montserrat',
            'align': 'center',
            'valign': 'vcenter',
            'font_color': 'white',
            'bg_color': '#14324B'
        })
        
        # Insertar textos
        worksheet.merge_range('A1:L1', 'Acquatrade Sudamericana S.A', company_name_format)
        worksheet.merge_range('A2:L2', 'Cotización', subtitle_format)
        
        # Configurar anchos de columna uniformes después del logo
        worksheet.set_column('A:L', 15)    # Todas las columnas con el mismo ancho
        
        # Formats para la tabla (con bordes)
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#14324B',
            'font_color': 'white',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })
        
        cell_format = workbook.add_format({
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'text_wrap': True
        })
        
        alt_row_format = workbook.add_format({
            'border': 1,
            'bg_color': '#F0F0F0',
            'align': 'left',
            'valign': 'vcenter',
            'text_wrap': True
        })
        
        # Formatos SIN bordes para información general
        label_format = workbook.add_format({
            'bold': True,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        value_format = workbook.add_format({
            'align': 'left',
            'valign': 'vcenter'
        })
        
        total_label_format = workbook.add_format({
            'bold': True,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        total_value_format = workbook.add_format({
            'bold': True,
            'align': 'right',
            'valign': 'vcenter',
            'num_format': '#,##0.00'
        })
        
        # Agregar espacio después del header
        worksheet.set_row(2, 5)  # Fila vacía pequeña (después de las 2 filas del header)
        worksheet.set_row(3, 5)  # Fila vacía pequeña
        
        # Client information (sin bordes, con mejor espaciado)
        info_row = 4
        worksheet.write(info_row, 0, 'Razón social:', label_format)
        worksheet.write(info_row, 1, razon_social, value_format)
        worksheet.write(info_row, 4, 'N° de cotiz:', label_format)
        worksheet.write(info_row, 5, nro_cotizacion, value_format)
        
        worksheet.write(info_row + 1, 0, 'CUIT:', label_format)
        worksheet.write(info_row + 1, 1, cuit, value_format)
        worksheet.write(info_row + 1, 4, 'Fecha:', label_format)
        worksheet.write(info_row + 1, 5, fecha, value_format)
        
        # Agregar más espacio antes de la tabla
        worksheet.set_row(6, 10)  # Fila vacía antes de la tabla
        
        # Start table at row 8 (ajustado por el nuevo header)
        start_row = 8
        
        # Write column headers
        for col, value in enumerate(df_result.columns):
            worksheet.write(start_row - 1, col, value, header_format)
        
        # Write data with alternating row colors (solo la tabla tiene bordes)
        for row_idx, row in enumerate(df_result.values):
            row_format = alt_row_format if row_idx % 2 else cell_format
            for col_idx, value in enumerate(row):
                worksheet.write(row_idx + start_row, col_idx, value, row_format)
        
        # Agregar espacio después de la tabla
        worksheet.set_row(start_row + len(df_result), 10)
        
        # Calcular posición para totales (lado derecho)
        total_start_col = len(df_result.columns) - 2  # Dos columnas desde la derecha
        total_row = start_row + len(df_result) + 2
        
        # Write totals en el lado derecho (sin bordes)
        for idx, (concepto, importe) in enumerate(resumen_df.values):
            worksheet.write(total_row + idx, total_start_col, concepto, total_label_format)
            worksheet.write(total_row + idx, total_start_col + 1, importe, total_value_format)
        
        # Agregar espacio antes de las condiciones
        conditions_row = total_row + len(resumen_df) + 3
        conditions = [
            "Condiciones comerciales:",
            "• Validez de la presente cotización: 3 días corridos.",
            "• Precios expresados en dólares.",
            "• Forma de pago: a convenir.",
            "• Entrega: sujeta a disponibilidad de stock."
        ]
        
        condition_format = workbook.add_format({
            'align': 'left',
            'valign': 'vcenter'
        })
        
        condition_title_format = workbook.add_format({
            'bold': True,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        for idx, condition in enumerate(conditions):
            format_to_use = condition_title_format if idx == 0 else condition_format
            worksheet.write(conditions_row + idx, 0, condition, format_to_use)
        
        # Footer (sin bordes)
        footer_row = conditions_row + len(conditions) + 2
        footer_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'italic': True
        })
        worksheet.write(footer_row, 0, 'www.acquatrade.com', footer_format)
        
        # Adjust column widths para que no haya celdas cortadas
        for idx, col in enumerate(df_result.columns):
            # Calcular el ancho máximo necesario
            header_width = len(str(col))
            max_content_width = 0
            
            # Revisar el contenido de cada celda en la columna
            for value in df_result[col]:
                content_width = len(str(value))
                if content_width > max_content_width:
                    max_content_width = content_width
            
            # Usar el mayor entre header y contenido, con un mínimo de 10 y máximo de 50
            optimal_width = max(header_width, max_content_width) + 3
            final_width = min(max(optimal_width, 10), 50)
            
            worksheet.set_column(idx, idx, final_width)
        
        # Ajustar altura de filas de la tabla para texto largo
        for row_idx in range(len(df_result)):
            worksheet.set_row(start_row + row_idx, 20)  # Altura mínima para texto
    
    return excel_path, excel_filename

def generate_pdf(df_result, resumen_df, razon_social, cuit, nro_cotizacion, fecha):
    # Create PDF file in our tmp directory
    pdf_filename = f"temp_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_filename)
    
    # Set up matplotlib to not display plots
    plt.ioff()
    
    # Create PDF with matplotlib - simpler approach
    with PdfPages(pdf_path) as pdf:
        fig = plt.figure(figsize=(11.7, 8.3))  # A4 landscape
        ax = fig.add_subplot(111)
        ax.set_xlim(0, 10)
        ax.set_ylim(0, 8)
        ax.axis('off')
        
        # Header background - single rectangle
        header_rect = patches.Rectangle((0, 7.2), 10, 0.8, linewidth=0, 
                                      facecolor='#14324B', alpha=1)
        ax.add_patch(header_rect)
        
        # Company name centered
        ax.text(5, 7.75, 'Acquatrade Sudamericana S.A', 
                horizontalalignment='center', verticalalignment='center',
                fontsize=18, color='white', weight='bold')
        
        # Agregar el subtítulo "Cotización"
        ax.text(5, 7.4, 'Cotización', 
               horizontalalignment='center', verticalalignment='center',
               fontsize=14, color='white', weight='normal')
        
        # Client information - cleaner spacing
        ax.text(0.5, 6.8, f'Razón social: {razon_social}', fontsize=10, weight='bold')
        ax.text(0.5, 6.6, f'CUIT: {cuit}', fontsize=10, weight='bold')
        
        ax.text(6.5, 6.8, f'N° de cotiz: {nro_cotizacion}', fontsize=10, weight='bold')
        ax.text(6.5, 6.6, f'Fecha: {fecha}', fontsize=10, weight='bold')
        
        # Table - simplified approach
        table_y = 6.2
        headers = ['Descripción Artículo', 'Desc. Adicional', 'Cantidad', 'Precio', '% IVA', 'Precio Neto']
        
        # Simple column setup
        col_widths = [2.5, 1.5, 1, 1, 0.8, 1.2]
        col_starts = [0.5, 3, 4.5, 5.5, 6.5, 7.3]
        
        # Table header
        header_height = 0.35
        for i, (header, start, width) in enumerate(zip(headers, col_starts, col_widths)):
            # Header background
            rect = patches.Rectangle((start, table_y), width, header_height, 
                                   facecolor='#14324B', edgecolor='black', linewidth=1)
            ax.add_patch(rect)
            
            # Header text
            ax.text(start + width/2, table_y + header_height/2, header, 
                   ha='center', va='center', fontsize=8, color='white', weight='bold')
        
        # Table data
        row_height = 0.3
        max_rows = min(len(df_result), 14)  # More rows
        
        for row_idx, row_data in enumerate(df_result.values[:max_rows]):
            y_pos = table_y - header_height - (row_idx * row_height)
            
            # No alternate row colors - clean white background for all rows
            
            # Cell data
            for col_idx, (value, start, width) in enumerate(zip(row_data, col_starts, col_widths)):
                # Cell border - lighter and cleaner
                border = patches.Rectangle((start, y_pos), width, row_height, 
                                         facecolor='white', edgecolor='lightgray', linewidth=0.3)
                ax.add_patch(border)
                
                # Cell text
                text_value = str(value) if pd.notna(value) else ""
                
                # Truncate if too long
                if len(text_value) > int(width * 10):
                    text_value = text_value[:int(width * 10)-3] + "..."
                
                # Align numbers right, text center
                if col_idx >= 2:  # Numeric columns
                    ax.text(start + width - 0.05, y_pos + row_height/2, text_value, 
                           ha='right', va='center', fontsize=7)
                else:
                    ax.text(start + width/2, y_pos + row_height/2, text_value, 
                           ha='center', va='center', fontsize=7)
        
        # Totals section - more space from table
        totals_y = table_y - header_height - (max_rows * row_height) - 0.8
        
        ax.text(7.5, totals_y + 0.5, 'TOTALES', fontsize=12, weight='bold')
        
        for i, (concepto, importe) in enumerate(resumen_df.values):
            y = totals_y - (i * 0.25)
            ax.text(7.0, y, f'{concepto}:', fontsize=9, weight='bold')
            ax.text(9.5, y, f'${importe:,.2f}', fontsize=9, ha='right', weight='bold')
        
        # Conditions - adjusted spacing
        cond_y = totals_y - 1.8
        conditions = [
            "Condiciones comerciales:",
            "• Validez de la presente cotización: 3 días corridos.",
            "• Precios expresados en dólares.",
            "• Forma de pago: a convenir.",
            "• Entrega: sujeta a disponibilidad de stock."
        ]
        
        for i, condition in enumerate(conditions):
            weight = 'bold' if i == 0 else 'normal'
            ax.text(0.5, cond_y - (i * 0.18), condition, fontsize=9, weight=weight)
        
        # Footer
        ax.text(5, 0.3, 'www.acquatrade.com', ha='center', fontsize=9, style='italic')
        
        # Save with white background
        pdf.savefig(fig, bbox_inches='tight', dpi=300, facecolor='white', edgecolor='none')
        plt.close(fig)
    
    return pdf_path, pdf_filename

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        # Limpiar archivos temporales antiguos
        cleanup_old_files()
        
        if 'file' not in request.files:
            return 'No file uploaded', 400
        
        file = request.files['file']
        if file.filename == '':
            return 'No file selected', 400
        
        if not file.filename.lower().endswith('.pdf'):
            return 'Only PDF files are allowed', 400
        
        # Get form data
        razon_social = request.form.get('razon_social', '')
        cuit = request.form.get('cuit', '')
        nro_cotizacion = request.form.get('nro_cotizacion', '')
        fecha = request.form.get('fecha', '')
        
        # Create a unique filename
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Process PDF
        df_result, resumen_df = process_pdf(filepath)
        
        # Generate Excel and PDF
        excel_path, excel_filename = generate_excel(df_result, resumen_df, razon_social, cuit, nro_cotizacion, fecha)
        pdf_path, pdf_filename = generate_pdf(df_result, resumen_df, razon_social, cuit, nro_cotizacion, fecha)
        
        # Clean up original PDF file
        os.unlink(filepath)
        
        # Rename files to match original PDF name
        base_filename = os.path.splitext(filename)[0]
        final_excel_filename = base_filename + '.xlsx'
        final_pdf_filename = base_filename + '.pdf'
        final_excel_path = os.path.join(app.config['UPLOAD_FOLDER'], final_excel_filename)
        final_pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], final_pdf_filename)
        
        # Move the generated files to final names
        shutil.move(excel_path, final_excel_path)
        shutil.move(pdf_path, final_pdf_path)
        
        # Create ZIP file with both Excel and PDF
        zip_filename = base_filename + '_files.zip'
        zip_path = os.path.join(app.config['UPLOAD_FOLDER'], zip_filename)
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            zip_file.write(final_excel_path, final_excel_filename)
            zip_file.write(final_pdf_path, final_pdf_filename)
        
        # Clean up individual files (keep only ZIP)
        os.unlink(final_excel_path)
        os.unlink(final_pdf_path)
        
        # Send ZIP file containing both Excel and PDF
        return send_file(
            zip_path,
            as_attachment=True,
            download_name=zip_filename,
            mimetype='application/zip'
        )
    
    except Exception as e:
        # Clean up files in case of error
        if 'filepath' in locals() and os.path.exists(filepath):
            os.unlink(filepath)
        if 'final_excel_path' in locals() and os.path.exists(final_excel_path):
            os.unlink(final_excel_path)
        if 'final_pdf_path' in locals() and os.path.exists(final_pdf_path):
            os.unlink(final_pdf_path)
        if 'zip_path' in locals() and os.path.exists(zip_path):
            os.unlink(zip_path)
        return str(e), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
