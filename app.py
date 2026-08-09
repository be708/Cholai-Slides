"""
CholaiSlides - AI Presentation Maker
Built by Cholai Tech
"""

from flask import Flask, render_template, request, send_file
import io
from pptx import Presentation
from pptx.util import Inches

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_slides():
    markdown_text = request.form['topic'] # we will use this box for full markdown
    language = request.form['language']

    # Create PowerPoint
    prs = Presentation()

    # Split slides by ---
    slides_content = markdown_text.split('---')

    for slide_text in slides_content:
        slide_text = slide_text.strip()
        if not slide_text:
            continue

        # Choose layout: Title slide vs Content slide
        if '#' in slide_text.split('\n')[0]:
            slide_layout = prs.slide_layouts[0] # Title slide
        else:
            slide_layout = prs.slide_layouts[1] # Content slide

        slide = prs.slides.add_slide(slide_layout)

        lines = slide_text.split('\n')

        # Title
        if lines and lines[0].startswith('#'):
            title = lines[0].replace('#', '').strip()
            slide.shapes.title.text = title

        # Content
        if len(slide.shapes) > 1 and len(lines) > 1:
            content = '\n'.join([l.replace('- ', '• ').strip() for l in lines[1:] if l.strip()])
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

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')
