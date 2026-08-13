"""
CholaiSlides - AI Presentation Maker V6.2
SECURE - Uses Environment Variables
Built by Cholai Tech
"""
import io
import os # NEW
import google.generativeai as genai
from flask import Flask, render_template, request, send_file
from pptx import Presentation
from pptx.util import Inches
import datetime

app = Flask(__name__)

# SECURE: READ KEY FROM RENDER, NOT CODE
GOOGLE_API_KEY = os.environ.get("GOOGLE_API_KEY") 
genai.configure(api_key=GOOGLE_API_KEY)
model = genai.GenerativeModel('gemini-1.5-flash')
TEMPLATES = {
    "template02": "template02.pptx"
}

db = None

def generate_with_ai(product_idea):
    """Ask REAL AI to write professional slides"""
    prompt = f"""
    You are a professional pitch deck writer for investors. 
    Create 10 slides for this product: "{product_idea}"
    
    Rules:
    1. Title slide with product name
    2. The Problem - 3 bullet points
    3. Our Solution - 3 bullet points 
    4. Key Features - 4 bullet points
    5. Market Opportunity - 3 bullet points with numbers for PNG
    6. Business Model - 4 bullet points
    7. Traction & Milestones - 4 bullet points
    8. Our Team - 3 bullet points
    9. The Ask - Investment, Use of funds, Goal
    10. Contact Us - Website, WhatsApp, Email
    
    Format as: TITLE: [title]
    CONTENT: [bullet1] | [bullet2] | [bullet3]
    
    Make it professional, specific to PNG market, and investor-ready.
    """
    
    response = model.generate_content(prompt)
    text = response.text
    
    slides = []
    for block in text.split("TITLE:"):
        if "CONTENT:" in block:
            lines = block.strip().split("CONTENT:")
            title = lines[0].strip()
            content = [c.strip() for c in lines[1].split("|") if c.strip()]
            slides.append({"title": title, "content": content})
    
    return slides[:10]


@app.route('/')
def home():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_slides():
    user_idea = request.form['topic']
    email = request.form.get('email', 'no-email')

    slides_data = generate_with_ai(user_idea)
    product_name = user_idea[:30]

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

    safe_name = product_name.replace(" ", "_")[:30]
    return send_file(file_stream, as_attachment=True, download_name=f'{safe_name}_Pitch.pptx')

if __name__ == '__main__':
    app.run(debug=True)
