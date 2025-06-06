from flask import Flask, render_template, request
from utils.notebook_functions import process_data  # Assuming this function processes the input data

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Get form data
        option1 = request.form.get('option1')
        option2 = request.form.get('option2')
        output_type = request.form.get('output_type')
        
        # Process the data using the imported function
        results = process_data(option1, option2, output_type)
        
        return render_template('results.html', results=results)
    
    return render_template('index.html')

@app.route('/results')
def results():
    return render_template('results.html')

if __name__ == '__main__':
    app.run(debug=True)