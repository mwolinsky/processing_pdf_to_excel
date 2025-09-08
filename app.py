import os
import re
from flask import Flask, request, render_template, send_file
import pdfplumber
import pandas as pd
from werkzeug.utils import secure_filename
import tempfile

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

def process_pdf(pdf_path):
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
        
        # Extract bonificación from text
        pattern = r'(Subtotal Cotización|Bonificación|Subtotal Neto|IVA|Total Cotización)\s*:\s*([\\d.,]+)'
        matches = re.findall(pattern, text)
        resultados = {label: float(valor.replace(',', '').replace('.', '', valor.count('.')-1)) 
                     for label, valor in matches}
        
        # Calculate bonificación percentage
        if 'Subtotal Cotización' in resultados and 'Bonificación' in resultados:
            pct_bonificacion = resultados['Bonificación'] / resultados['Subtotal Cotización']
        else:
            pct_bonificacion = 0
            
        # Convert numeric columns
        df_final["Cantidad"] = pd.to_numeric(df_final["Cantidad"], errors="coerce")
        df_final["Precio_Unit"] = pd.to_numeric(df_final["Precio Unit"], errors="coerce")
        df_final["% Desc."] = pd.to_numeric(df_final["% Desc."], errors="coerce").fillna(0)
        df_final["% IVA"] = pd.to_numeric(df_final["% IVA"], errors="coerce").fillna(0)
        df_final["Importe"] = pd.to_numeric(df_final["Importe"], errors="coerce").fillna(0)
        
        # Calculate prices
        df_final["Precio_Lista"] = df_final["Cantidad"] * df_final["Precio_Unit"]
        df_final["Precio Lista con Descuento Prod."] = df_final["Precio_Lista"] * (1 - df_final["% Desc."] / 100)
        df_final["Precio Neto"] = df_final["Precio Lista con Descuento Prod."] * (1 - pct_bonificacion)
        df_final["Precio Neto Unitario"] = df_final["Precio Neto"] / df_final["Cantidad"]
        df_final["Precio con Impuestos"] = df_final["Precio Neto"] * (1 + df_final["% IVA"] / 100)
        df_final["Precio"] = df_final["Precio Neto Unitario"]
        
        # Calculate summary
        subtotal = round(pd.to_numeric(df_final["Precio Neto"], errors="coerce").sum(), 2)
        precio_neto_21 = pd.to_numeric(df_final.loc[df_final["% IVA"] == 21, "Precio Neto"].replace("", 0), errors="coerce").fillna(0)
        iva_21 = round((precio_neto_21 * 0.21).sum(), 2)
        precio_neto_105 = pd.to_numeric(df_final.loc[df_final["% IVA"] == 10.5, "Precio Neto"].replace("", 0), errors="coerce").fillna(0)
        iva_105 = round((precio_neto_105 * 0.105).sum(), 2)
        total = round(pd.to_numeric(df_final["Precio con Impuestos"], errors="coerce").sum(), 2)
        
        # Create summary DataFrame
        resumen_df = pd.DataFrame({
            "Concepto": ["Subtotal", "IVA 21%", "IVA 10.5%", "Total"],
            "Importe": [subtotal, iva_21, iva_105, total]
        })
        
        # Create final result DataFrame
        df_result = df_final[["Descripción Artículo", "Desc. Adicional", "Cantidad", "Precio", "% IVA", "Precio Neto"]].copy()
        df_result["Cantidad"] = df_result["Cantidad"].round(2)
        df_result["Precio"] = df_result["Precio"].round(2)
        df_result["% IVA"] = df_result["% IVA"].round(2)
        df_result["Precio Neto"] = df_result["Precio Neto"].round(2)
        
        return df_result, resumen_df

def generate_excel(df_result, resumen_df):
    # Create temporary file
    temp_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    
    # Write to Excel
    with pd.ExcelWriter(temp_file.name, engine="xlsxwriter") as writer:
        df_result.to_excel(writer, index=False, sheet_name="Hoja1", startrow=0)
        n_filas = len(df_result)
        resumen_df.to_excel(writer, index=False, sheet_name="Hoja1", startrow=n_filas + 2)
    
    return temp_file.name

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file uploaded', 400
    
    file = request.files['file']
    if file.filename == '':
        return 'No file selected', 400
    
    if not file.filename.lower().endswith('.pdf'):
        return 'Only PDF files are allowed', 400
    
    # Save uploaded file
    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)
    
    try:
        # Process PDF
        df_result, resumen_df = process_pdf(filepath)
        
        # Generate Excel
        excel_path = generate_excel(df_result, resumen_df)
        
        # Clean up PDF file
        os.unlink(filepath)
        
        # Send Excel file
        # Get the PDF filename without extension and add xlsx
        excel_filename = os.path.splitext(filename)[0] + '.xlsx'
        
        return send_file(
            excel_path,
            as_attachment=True,
            download_name=excel_filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    
    except Exception as e:
        # Clean up files in case of error
        if os.path.exists(filepath):
            os.unlink(filepath)
        return str(e), 500

if __name__ == '__main__':
    app.run(debug=True)
