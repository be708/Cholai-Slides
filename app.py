"""
CholaiSlides - AI Presentation Maker V5
TRUE AI BRAIN - Thinks and Generates
Built by Cholai Tech
"""
import io
import re
from flask import Flask, render_template, request, send_file, redirect
from pptx import Presentation
from pptx.util import Inches
import datetime

app = Flask(__name__)
db = None

def generate_smart_content(product_name, keywords):
    """AI Brain that generates content based on product + keywords"""
    
    # Detect keywords
    has_languages = any(k in keywords for k in ["language", "tokpisin", "840", "png"])
    has_ai = any(k in keywords for k in ["ai", "siri", "chat", "assistant"])
    has_business = any(k in keywords for k in ["business", "course", "management"])
    
    if has_languages:
        problem = [
            "840+ PNG languages at risk of being lost",
            "No AI assistant that speaks Tokpisin + local languages", 
            "Elders' stories and culture not documented"
        ]
        solution = [
            f"{product_name}: AI that speaks Tokpisin + 840 PNG languages",
            "Voice assistant for everyday PNG people",
            "PNG Tubunna Archives to preserve culture"
        ]
        features = [
            "Speaks and understands 840 PNG languages",
            "Voice chat in local language",
            "Cultural knowledge + WW1/WW2 archives",
            "Business courses in Tokpisin"
        ]
        market = [
            "10M+ people in PNG",
            "840 languages - 12% of world's languages", 
            "Government + Education digitization"
        ]
    else:
        # Generic fallback
        problem = [
            "Manual work takes too much time",
            "No affordable solution for PNG market",
            "Complex tools hard for locals to use"
        ]
        solution = [
            f"{product_name}: Built for PNG",
            "Simple, affordable, and powerful",
            "Made by locals, for locals"
        ]
        features = ["Easy to use", "Affordable pricing", "Built for PNG"]
        market = ["500,000+ SMEs in PNG", "10M+ population"]
    
    return problem, solution, features, market


def smart_split_to_slides(idea_text):
    """
    KING'S AI BRAIN V5: TRUE THINKING
    """
    slides = []
    idea_lower = idea_text.lower()
    
    # STEP 1: EXTRACT PRODUCT NAME
    # Look for "for [PRODUCT]" pattern
    match = re.search(r'for\s+([^-.]+)', idea_text, re.IGNORECASE)
    if match:
        product_name = match.group(1).strip()
    else:
        product_name = idea_text[:40].strip()
    
    # STEP 2: EXTRACT KEYWORDS
    keywords = idea_lower.split()
    
    # STEP 3: GENERATE SMART CONTENT
    problem, solution, features, market = generate_smart_content(product_name, keywords)
    
    # BUILD 10 SLIDES
    slides.append({"title": f"{product_name} - Investor Pitch", "content": ["AI-Powered Presentation by CholaiSlides", "Turn Ideas Into Slides in 60 Seconds"]})
    slides.append({"title": "The Problem", "content": problem})
    slides.append({"title": "Our Solution", "content": solution})
    slides.append({"title": "Key Features", "content": features})
    slides.append({"title": "Market Opportunity", "content": market})
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

    if db:
        order_data = {
            'topic': user_idea,
            'language': language,
            'email': email,
            'timestamp': datetime.datetime.now()
        }
        db.collection('orders').add(order_data)

    slides_data = smart_split_to_slides(user_idea)

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    for slide_info in slides_data:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
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

    file_stream = io.BytesIO()
    prs.save(file_stream)
    file_stream.seek(0)

    return send_file(
        file_stream,
        as_attachment=True,
        download_name=f'{product_name}_Pitch.pptx',
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )

if __name__ == '__main__':
    app.run(debug=True)
