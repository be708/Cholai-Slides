import os
from flask import Flask, request, render_template, send_file
from pptx import Presentation
from pptx.util import Inches, Pt
import google.generativeai as genai
from io import BytesIO

app = Flask(__name__)

GOOGLE_API_KEY = os.environ.get("GOOGLE_API_KEY")
genai.configure(api_key=GOOGLE_API_KEY)

TEMPLATES = {
    "template02": "template02.pptx"
}
def generate_with_ai(user_idea):
    model = genai.GenerativeModel('gemini-1.0-pro')
    prompt = f"""
    You are an expert presentation designer. Create a 6 slide PowerPoint outline for: {user_idea}
    Format: Use Markdown. Each slide starts with # Title
    Bullet points start with - 
    Separate slides with ---
    Example:
    # Problem
    - Point 1
    - Point 2
    ---
    # Solution
    - Point 1
    """
    response = model.generate_content(prompt)
    return parse_markdown(response.text)

 def parse_markdown(md_text):
    slides = []
    blocks = md_text.split('---')
    for block in blocks:
        lines = [l.strip() for l in block.split('\n') if l.strip()]
        if not lines: 
            continue
        title = lines[0].replace('# ', '')
        content = [l.replace('- ', '') for l in lines[1:]]
        slides.append({"title": title, "content": content})
    return slides


@app.route('/')
def home():
    return render_template('index.html')


@app.route('/generate', methods=['POST'])
def generate_slides():
    user_idea = request.form.get('topic')
    manual_markdown = request.form.get('markdown')
    email = request.form.get('email', 'no-email')
    template_choice = request.form.get('template', 'template02')

    if user_idea and user_idea.strip()!= "":
        slides_data = generate_with_ai(user_idea) # AI MODE
    else:
        slides_data = parse_markdown(manual_markdown) # MANUAL MODE

    product_name = user_idea[:30] if user_idea else "My Presentation"
    template_file = TEMPLATES.get(template_choice, "template02.pptx")
    prs = Presentation(f'static/templates/{template_file}')
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
            p = tf.add_paragraph()
            p.text = point
            p.level = 0
            p.font.size = Pt(18)

    file_stream = BytesIO()
    prs.save(file_stream)
    file_stream.seek(0)

    return send_file(
        file_stream,
        as_attachment=True,
        download_name=f"{product_name}.pptx",
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )

if __name__ == '__main__':
    app.run(debug=True)
