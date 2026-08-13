"""
CholaiSlides - AI Presentation Maker V3
Built by Cholai Tech
Saves orders to Firebase
Smart Brain: Turns plain ideas into pro slides
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
    KING'S AI BRAIN V3: No repeats + Extracts Problem from your text
    """
    slides = []
    sentences = [s.strip() for s in re.split(r'\.|\n', idea_text) if len(s.strip()) > 10]
    used_sentences = [] # Track what we already used
    
    # SLIDE 1: TITLE
    title = sentences[0][:80] if sentences else "Your Presentation"
    slides.append({"title": title, "content": ["AI-Powered Presentation by CholaiSlides", "Turn Ideas Into Slides in 60 Seconds"]})
    
    # SLIDE 2: PROBLEM - Force extract disappearing/lost/challenge
    problem_keywords = ['disappear','lost','challenge','problem','dying','struggle','barrier']
    problem_points = [s for s in sentences if any(word in s.lower() for word in problem_keywords)]
    if not problem_points:
        problem_points = ["840+ PNG languages at risk of being lost", "Elders' stories not documented", "Business barriers due to language"]
    used_sentences.extend(problem_points)
    slides.append({"title": "The Problem", "content": problem_points[:3]})
    
    # SLIDE 3: SOLUTION - Remove sentences already used
    solution_keywords = ['app','solution','speaks','understands','chat','ai']
    solution_points = [s for s in sentences if s not in used_sentences and any(word in s.lower() for word in solution_keywords)]
    if not solution_points:
        solution_points = ["CholaiChat Hausman: AI for all of PNG"]
    used_sentences.extend(solution_points)
    slides.append({"title": "Our Solution", "content": solution_points[:3]})
    
    # SLIDE 4: KEY FEATURES - Archives + Courses
    feature_keywords = ['archives','traditions','culture','courses','languages','tokpisin','business','management','stories','ww1','ww2']
    feature_points = [s for s in sentences if s not in used_sentences and any(word in s.lower() for word in feature_keywords)]
    used_sentences.extend(feature_points)
    slides.append({"title": "Key Features", "content": feature_points[:4]})
    
    # SLIDE 5: MARKET OPPORTUNITY - Numbers + PNG
    market_keywords = ['png','people','million','languages','840','population']
    market_points = [s for s in sentences if s not in used_sentences and any(word in s.lower() for word in market_keywords)]
    if not market_points:
        market_points = ["10M+ people in PNG", "840 languages - 12% of world's languages"]
    slides.append({"title": "Market Opportunity", "content": market_points[:3]})
    
    # SLIDES 6-10: PRO FILLER
    slides.append({"title": "Business Model", "content": ["Subscription: K15/month", "Government & NGO Partnerships", "Enterprise Licensing", "Freemium Model"]})
    slides.append({"title": "Traction & Milestones", "content": ["MVP Built and Deployed", "First Users Onboarded", "Key Partnerships", "Product Roadmap"]})
    slides.append({"title": "Our Team", "content": ["Built by Cholai Tech", "Experts in AI + PNG Culture", "Mission Driven Founders"]})
    slides.append({"title": "The Ask", "content": ["Investment: K500,000", "18 Month Runway", "Use: Development + Marketing", "Goal: 100,000 Users"]})
    slides.append({"title": "Contact Us", "content": ["Website: be708.github.io", "WhatsApp: 72817573", "Email: bugfreezonepng@gmail.com", "Powered by Cholai Tech"]})
    
    return slides[:10]


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
    
    # TODO NEXT: Add email sending here

    return send_file(
        file_stream,
        as_attachment=True,
        download_name='CholaiSlides_Pitch.pptx',
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )


if __name__ == '__main__':
    app.run(debug=True)
