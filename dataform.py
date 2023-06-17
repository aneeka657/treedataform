from flask import Flask, render_template
from flask import Flask, request
import pandas as pd
from datetime import datetime
import os
from flask import send_file

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('index.html')


@app.route('/submit_form', methods=['POST'])
def handle_form():
    data = request.form.to_dict()  # Convert form data to dictionary

    # If a file is sent in the form
    # If a file is sent in the form
    if 'uploadPhoto' in request.files:
        file = request.files['uploadPhoto']
        if file.filename != '':
            file_path = os.path.join('static/images/', file.filename)
            file.save(file_path)
            data['uploadPhoto'] = file_path  # Save path of image file in data dictionary

    # Convert the dictionary to a DataFrame
    df = pd.DataFrame([data])


    filename = 'form_data.xlsx'

    if os.path.isfile(filename):
        df_old = pd.read_excel(filename)
        df_new = pd.concat([df_old, df], ignore_index=True)
    else:
        df_new = df

    df_new.to_excel(filename, index=False)



    return {"message": "Form submitted successfully!"}
@app.route('/download_excel')
def download_excel():
    path = "form_data.xlsx"
    return send_file(path, as_attachment=True)

@app.route('/view_data')
def view_data():
    df = pd.read_excel('form_data.xlsx')
    table = df.to_html(index=False)
    return render_template('view_data.html', table=table)


if __name__ == '__main__':
    app.run(debug=True)
