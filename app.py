#%%
from flask import Flask, render_template, request, send_file
from utils import process_data

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    try:
        if request.method == 'POST':
            text = request.form['data']
            output_file = process_data(text)
            return send_file(output_file, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='general_orders_data_from_message.xlsx')
        return render_template('index.html')
    except Exception as e:
        return str(e)
    
@app.route('/process', methods=['POST'])
def process():
    data = request.form.get('data')
    output = process_data(data)
    
    # Return the in-memory Excel file as a response with the appropriate MIME type
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='general_orders_data_from_message.xlsx')
    
if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True)
