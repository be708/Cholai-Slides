"""
CholaiSlides - AI Presentation Maker V2
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
    KING'S AI BRAIN V2: Reads the idea and pulls real content
    """
    slides = []
    # Clean sentences
    sentences = [s.strip() for s in re.split(r'\.|\n', idea_text) if len(s.strip()) > 10]
    
    # SLIDE 1: TITLE
    title = sentences[0][:80] if sentences else "Your Presentation"
    slides.append({"title": title, "content": ["AI-Powered Presentation by CholaiSlides", "Turn Ideas Into Slides in 60 Seconds"]})
    
    # SLIDE 2: PROBLEM
    problem_keywords = ['problem','challenge','disappear','lost','struggle','barrier','issue','dying']
    problem_points = [s for s in sentences if any(word in s.lower() for word in problem_keywords)]
    if not problem_points:
        problem_points = ["Key challenges in the market", "Opportunity for innovation", "Why now is the time"]
    slides.append({"title": "The Problem", "content": problem_points[:4]})
    
    # SLIDE 3: SOLUTION
    solution_keywords = ['app','solution','platform','speaks','understands','chat','ai']
    solution_points = [s for s in sentences if any(word in s.lower() for word in solution_keywords)]
    if not solution_points:
        solution_points = ["Introducing our solution", "How it solves the problem", "Key benefits"]
    slides.append({"title": "Our Solution", "content": solution_points[:4]})
    
    # SLIDE 4: KEY FEATURES
    feature_keywords = ['also','archives','traditions','culture','courses','features','languages','tokpisin','business','management','accounting','stories']
    feature_points = [s for s in sentences if any(word in s.lower() for word in feature_keywords)]
    if not feature_points:
        feature_points = ["Feature 1: Core functionality", "Feature 2: User benefits", "Feature 3: Unique value"]
    slides.append({"title": "Key Features", "content": feature_points[:5]})
    
    # SLIDE 5: MARKET OPPORTUNITY
    market_keywords = ['png','people','million','languages','market','840','population','target','users']
    market_points = [s for s in sentences if any(word in s.lower() for word in market_keywords)]
    if not market_points:
        market_points = ["Large target market", "High growth potential", "Competitive advantage"]
    slides.append({"title": "Market Opportunity", "content": market_points[:4]})
    
    # SLIDE 6: BUSINESS MODEL
    slides.append({"title": "Business Model", "content": ["Subscription: K15/month", "Government & NGO Partnerships", "Enterprise Licensing", "Freemium Model"]})
    
    # SLIDE 7: TRACTION & MILESTONES
    slides.append({"title": "Traction & Milestones", "content": ["MVP Built and Deployed", "First Users Onboarded", "Key Partnerships", "Product Roadmap"]})
    
    # SLIDE 8: OUR TEAM
    slides.append({"title": "Our Team", "content": ["Built by Cholai Tech", "Experts in AI + PNG Culture", "Mission Driven Founders"]})
    
    # SLIDE 9: THE ASK
    slides.append({"title": "The Ask", "content": ["Investment: K500,000", "18 Month Runway", "Use: Development + Marketing", "Goal: 100,000 Users"]})
    
    # SLIDE 10: CONTACT US
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
