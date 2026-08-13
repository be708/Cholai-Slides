"""
CholaiSlides - AI Presentation Maker V4
Built by Cholai Tech
Context-Aware Brain
Saves orders to Firebase
"""
import io
import re
from flask import Flask, render_template, request, send_file, redirect
from pptx import Presentation
from pptx.util import Inches
import datetime

app = Flask(__name__)

# FIREBASE - ADD YOUR KEY LATER
# import firebase_admin
# from firebase_admin import credentials, firestore
# cred = credentials.Certificate("serviceAccountKey.json")
# firebase_admin.initialize_app(cred)
# db = firestore.client()
db = None


def generate_cholaislides_content():
    """Content specifically about CholaiSlides"""
    return {
        "problem": [
            "Founders waste 5+ hours making pitch decks", 
            "Canva is too complex and generic", 
            "No AI tool built specifically for PNG businesses"
        ],
        "solution": [
            "CholaiSlides: Paste idea → Get 10-slide deck in 60 seconds", 
            "Built for PNG founders, NGOs, and students", 
            "No design skills needed"
        ],
        "features": [
            "AI-powered slide generation", 
            "Investor-ready 10 slide structure", 
            "Auto-save to Cholai HQ", 
            "Export to.pptx instantly"
        ],
        "market": [
            "500,000+ SMEs in PNG", 
            "10,000+ students and NGOs", 
            "Global market: $2B presentation software"
        ]
    }


def generate_cholaichat_content():
    """Content specifically about CholaiChat"""
    return {
        "problem": [
            "840+ PNG languages at risk of being lost", 
            "Elders' stories not documented", 
            "Business barriers due to language"
        ],
        "solution": [
            "CholaiChat Hausman: AI that speaks Tokpisin + 840 languages", 
            "PNG Tubunna Archives to preserve culture", 
            "6 month Business management courses"
        ],
        "features": [
            "Speaks and understands 840 PNG languages", 
            "Tubunna Archives: traditions, WW1/WW2 stories", 
            "Business management courses", 
            "Voice chat in local language"
        ],
        "market": [
            "10M+ people in PNG", 
            "840 languages - 12% of world's languages", 
            "Government digitization projects"
        ]
    }


def smart_split_to_slides(idea_text):
    """
    KING'S AI BRAIN V4: Context-Aware
    """
    slides = []
    idea_lower = idea_text.lower()
    sentences = [s.strip() for s in re.split(r'\.|\n', idea_text) if len(s.strip()) > 10]
    
    # DETECT CONTEXT
    if "cholaislides" in idea_lower:
        context = generate_cholaislides_content()
        product_name = "CholaiSlides"
    elif "cholaichat" in idea_lower:
        context = generate_cholaichat_content()
        product_name = "CholaiChat Hausman"
    else:
        # Generic fallback
        context = generate_cholaislides_content()
        product_name = "Your Product"
    
    # SLIDE 1: TITLE
    title = idea_text[:80]
    slides.append({"title": title, "content": ["AI-Powered Presentation by CholaiSlides", "Turn Ideas Into Slides in 60 Seconds"]})
    
    # SLIDE 2: PROBLEM
    slides.append({"title": "The Problem", "content": context["problem"]})
    
    # SLIDE 3: SOLUTION
    slides.append({"title": "Our Solution", "content": context["solution"]})
    
    # SLIDE 4: KEY FEATURES
    slides.append({"title": "Key Features", "content": context["features"]})
    
    # SLIDE 5: MARKET OPPORTUNITY
    slides.append({"title": "Market Opportunity", "content": context["market"]})
    
    # SLIDES 6-10: PRO FILLER
    slides.append({"title": "Business Model", "content": ["Subscription: K15/month", "Government & NGO Partnerships", "Enterprise Licensing", "Freemium Model"]})
    slides.append({"title": "Traction & Milestones", "content": ["MVP Built and Deployed", "First Users Onboarded", "Key Partnerships", "Product Roadmap"]})
    slides.append({"title": "Our Team", "content": ["Built by Cholai Tech", "Experts in AI + PNG", "Mission Driven Founders"]})
    slides.append({"title": "The Ask", "content": ["Investment: K500,000", "18 Month Runway", "Use: Development + Marketing", "Goal: 100,000 Users"]})
    slides.append({"title": "Contact Us", "content": ["Website: be708.github.io", "WhatsApp: 72817573", "Email: bugfreezonepng@gmail.com", "Powered by Cholai Tech"]})
    
    return slides[:10]


@app.route('/')
def home():
    return render_template('index.html')


@app.route('/generate', methods=['POST'])
def generate_slides():
    user_idea = request.form['topic']
    language = request.form['language']
    email = request.form.get('email', 'no-email')

    # SAVE ORDER TO FIREBASE
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

    return send_file(
        file_stream,
        as_attachment=True,
        download_name='CholaiSlides_Pitch.pptx',
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )


if __name__ == '__main__':
    app.run(debug=True)
