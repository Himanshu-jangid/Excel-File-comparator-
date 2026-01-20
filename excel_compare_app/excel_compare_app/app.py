from flask import Flask, request, render_template, send_from_directory
import os
import pandas as pd

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'file1' not in request.files or 'file2' not in request.files:
        return 'Please upload two Excel files', 400

    file1 = request.files['file1']
    file2 = request.files['file2']

    path1 = os.path.join(UPLOAD_FOLDER, file1.filename)
    path2 = os.path.join(UPLOAD_FOLDER, file2.filename)

    file1.save(path1)
    file2.save(path2)

    df1 = pd.read_excel(path1)
    df2 = pd.read_excel(path2)

    df1.columns = df1.columns.str.strip()
    df2.columns = df2.columns.str.strip()

    # Use 'Narration' from df1 and 'Description' from df2 if present
    if 'Narration' not in df1.columns or 'Description' not in df2.columns:
        return 'Missing Narration or Description column in uploaded files.', 400

    df1['Narration'] = df1['Narration'].astype(str).str.strip().str.rstrip(',')
    df2['Description'] = df2['Description'].astype(str).str.strip().str.rstrip(',')

    df1.set_index('Narration', inplace=True)
    df2.set_index('Description', inplace=True)

    common_keys = df1.index.intersection(df2.index)
    results = []

    def safe(val):
        return 0 if pd.isna(val) else val

    for key in common_keys:
        row1 = df1.loc[key]
        row2 = df2.loc[key]

        if isinstance(row1, pd.DataFrame) or isinstance(row2, pd.DataFrame):
            continue

        file1_credit = safe(row1.get('Credit', 0))
        file2_credit = safe(row2.get('Credit', 0))
        file1_debit = safe(row1.get('Debit', 0))
        file2_debit = safe(row2.get('Debit', 0))
        file1_balance = safe(row1.get('Balance', 0))
        file2_balance = safe(row2.get('Balance', 0))

        mismatch_fields = []
        if file1_credit != file2_credit:
            mismatch_fields.append("Credit")
        if file1_debit != file2_debit:
            mismatch_fields.append("Debit")
        if file1_balance != file2_balance:
            mismatch_fields.append("Balance")

        if mismatch_fields:
            results.append({
                "Description": key,
                "File1_Credit": file1_credit,
                "File2_Credit": file2_credit,
                "File1_Balance": file1_balance,
                "File2_Balance": file2_balance,
                "Mismatch In": ", ".join(mismatch_fields),
                "File1_Debit": file1_debit,
                "File2_Debit": file2_debit
            })
    if results:
        df_out = pd.DataFrame(results)
        output_file = 'matched_output.xlsx'
        output_path = os.path.join(OUTPUT_FOLDER, output_file)
        df_out.to_excel(output_path, index=False)
        return render_template('index.html', matched_file=output_file)
    else:
        return render_template('index.html', matched_file=None)

@app.route('/download/<filename>')
def download_file(filename):
    
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)