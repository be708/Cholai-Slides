"""
CholaiSlides - AI Presentation Maker
Built by Cholai Tech
Saves orders to Firebase
"""
import io
from flask import Flask, render_template, request, send_file, redirect
from pptx import Presentation
from pptx.util import Inches

# FIREBASE SETUP
import firebase_admin
from firebase_admin import credentials, firestore
import datetime

app = Flask(__name__)

# Initialize Firebase - REPLACE WITH YOUR SERVICE ACCOUNT JSON
# cred = credentials.Certificate("serviceAccountKey.json")
# firebase_admin.initialize_app(cred)
# db = firestore.client()
db = None # Comment this out after adding your Firebase key

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_slides():
    markdown_text = request.form['topic']
    language = request.form['language']
    email = request.form.get('email', 'no-email') # for order form

    # SAVE ORDER TO FIREBASE CHOLAI HQ
    if db:
        order_data = {
            'topic': markdown_text,
            'language': language,
            'email': email,
            'timestamp': datetime.datetime.now()
        }
        db.collection('orders').add(order_data)

    # CREATE POWERPOINT
    prs = Presentation()
    slides_content = markdown_text.split('---')

    for slide_text in slides_content:
        slide_text = slide_text.strip()
        if not slide_text:
            continue

        # Choose layout
        if '#' in slide_text.split('\n')[0]:
            slide_layout = prs.slide_layouts[0] # Title
        else:
            slide_layout = prs.slide_layouts[1] # Content

        slide = prs.slides.add_slide(slide_layout)
        lines = slide_text.split('\n')

        # Title
        if lines and lines[0].startswith('#'):
            title = lines[0].replace('#', '').strip()
            slide.shapes.title.text = title

        # Content
        if len(slide.shapes) > 1 and len(lines) > 1:
            content = '\n'.join([l.replace('- ', '') for l in lines[1:]])
            if slide.placeholders:
                slide.placeholders[1].text = content

    # Save to memory
    file_stream = io.BytesIO()
    prs.save(file_stream)
    file_stream.seek(0)

    return send_file(
        file_stream,
        as_attachment=True,
        download_name='CholaiSlides_Pitch.pptx',
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )

@app.route('/gold-button')
def gold_button():
    # Redirect to your website
    return redirect("https://be708.github.io")

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=5000)
