"""
CholaiSlides - AI Presentation Maker
Built by Cholai Tech
Saves orders to Firebase
"""
import io
import re
from flask import Flask, render_template, request, send_file, redirect
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.text import PP_ALIGN

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

def smart_split_to_slides(idea_text):
    """
    KING'S AI BRAIN: Turns plain idea into 10 slides automatically
    No markdown needed from user
    """
    slides = []
    
    # Auto-generate 10 slide structure based on idea
    title = idea_text.split('.')[0][:60] # First sentence as title
    
    slides.append({"title": title, "content": ["AI-Powered Presentation by CholaiSlides", "Turn Ideas Into Slides in 60 Seconds"]})
    
    # Simple AI logic - look for keywords and build slides
    if "problem" in idea_text.lower() or "issue" in idea_text.lower():
        slides.append({"title": "The Problem", "content": [s.strip() for s in re.split(r'\.|\n', idea_text) if len(s) > 20][:4]})
    else:
        slides.append({"title": "The Problem", "content": ["Key challenges in the market", "Opportunity for innovation", "Why now is the time"]})
    
    if "solution" in idea_text.lower() or "app" in idea_text.lower():
        slides.append({"title": "Our Solution", "content": [s.strip() for s in re.split(r'\.|\n', idea_text) if "app" in s.lower() or "solution" in s.lower()][:4]})
    else:
        slides.append({"title": "Our Solution", "content": ["Introducing: " + title, "How we solve the problem", "Key features and benefits"]})
    
    slides.append({"title": "Key Features", "content": [s.strip() for s in re.split(r'\.|\n', idea_text) if len(s) > 15][:5]})
    slides.append({"title": "Market Opportunity", "content": ["Target market size", "Growth potential", "Competitive advantage"]})
    slides.append({"title": "Business Model", "content": ["How we make money", "Revenue streams", "Pricing strategy"]})
    slides.append({"title": "Traction & Milestones", "content": ["Current progress", "Key achievements", "What's next"]})
    slides.append({"title": "Our Team", "content": ["Built by Cholai Tech", "Experts in AI + PNG", "Mission driven"]})
    slides.append({"title": "The Ask", "content": ["Investment opportunity", "How funds will be used", "Join us in building the future"]})
    slides.append({"title": "Contact Us", "content": ["Website: be708.github.io", "WhatsApp: 72817573", "Powered by Cholai Tech"]})
    
    return slides[:10] # Max 10 slides


@app.route('/')
def home():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_slides():
    user_idea = request.form['topic'] # User just pastes idea
    language = request.form['language']
    email = request.form.get('email', 'no-email')

    # SAVE ORDER TO FIREBASE CHOLAI HQ
    if db:
        order_data = {
            'topic': user_idea,
            'language': language,
            'email': email,
            'timestamp': datetime.datetime.now()
        }
        db.collection('orders').add(order_data)

    # SMART: Convert plain idea to slides
    slides_data = smart_split_to_slides(user_idea)

    # BUILD POWERPOINT
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    for slide_info in slides_data:
        slide = prs.slides.add_slide(prs.slide_layouts[1]) # Title + Content
        title = slide.shapes.title
        title.text = slide_info["title"]
        
        content = slide.placeholders[1]
        tf = content.text_frame
        tf.clear()
        for point in slide_info["content"]:
            if point.strip():
                p = tf.add_paragraph()
                p.text = point.strip()
                p.level = 0

    # SAVE TO MEMORY
    file_stream = io.BytesIO()
    prs.save(file_stream)
    file_stream.seek(0)
    
    # TODO: Add email sending here in next step

    return send_file(
        file_stream,
        as_attachment=True,
        download_name='CholaiSlides_Pitch.pptx',
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )

if __name__ == '__main__':
    app.run(debug=True)
