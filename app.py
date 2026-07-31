"""
CholaiSlides - AI Presentation Maker
Built by Cholai Tech
"""

from flask import Flask, render_template, request
app = Flask(__name__)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_slides():
    topic = request.form['topic']
    language = request.form['language']
    return f"""
    <h1 style='text-align:center; color:#2E7D32;'>Generating Your Slides...</h1>
    <p style='text-align:center;'>Topic: <b>{topic}</b></p>
    <p style='text-align:center;'>Language: <b>{language}</b></p>
    <p style='text-align:center;'>AI connection coming next! This is V1 Demo</p>
    """

if __name__ == '__main__':
    app.run(debug=True)
