import subprocess
from flask import Flask, render_template, send_file
from flask_cors import CORS
import main

app = Flask(__name__)
CORS(app)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/generate_report', methods=['POST'])
def generate_report():
    try:
        # Display to webpage
        output = "Report Generated!"

        # Run your Python script here
        main.main()

        result_text = f"Output: {output}"

        return result_text
    except Exception as e:
        return f"An error occurred: {str(e)}"


@app.route('/download_report')
def download_report():
    # Provide the file for download
    return send_file('Score Report.pdf', as_attachment=True)


if __name__ == '__main__':
    app.run(debug=False)
