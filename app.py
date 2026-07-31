"""
CholaiSlides - AI Presentation Maker
Generates slides from 1 sentence in 100+ languages
"""

from flask import Flask, render_template, request
app = Flask(__name__)

@app.route('/')
def home():
    return "CholaiSlides: Type a topic and get slides in 60 seconds"

@app.route('/generate', methods=['POST'])
def generate_slides():
    topic = request.form['topic']
    language = request.form['language']
    # TODO: Connect to AI to generate slides
    return f"Generating slides for: {topic} in {language}"

if __name__ == '__main__':
    app.run(debug=True)
