import os
import pandas as pd
from flask import Flask, render_template, request, redirect, send_file

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']

        file.save('upload.xls')
        
        df = pd.read_excel('upload.xls', engine='xlrd')

        states = df['Ship to State'].unique()

        writer = pd.ExcelWriter('report_with_states.xlsx', engine='xlsxwriter')

        df.to_excel(writer, sheet_name='Original Data', index=False)

        for state in states:

            state_df = df[df['Ship to State'] == state]

            state_df.to_excel(writer, sheet_name=state, index=False)

        writer.close()
        
        os.remove('upload.xls')
        
        return redirect('/download')
        
    return render_template('index.html')


@app.route('/download')
def download_file():
    return send_file('report_with_states.xlsx', as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)
