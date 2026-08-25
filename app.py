import io
import time
import os
from flask import Flask, request, jsonify, send_file
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

app = Flask(__name__)

TITLE_COLOR = RGBColor(0x2C, 0x3E, 0x50)
BODY_COLOR = RGBColor(0x44, 0x62, 0x80)

def build_title_slide(prs, topic):
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    title_box = slide.shapes.title
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = topic
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = TITLE_COLOR

@app.route('/health')
def health():
    return jsonify({"status": "ok"})

@app.route('/generate', methods=['POST'])
def generate():
    time.sleep(3) # Render wake up delay
    
    data = request.get_json(silent=True)
    if not data or 'topic' not in data or not data['topic'].strip():
        return jsonify({"error": "topic required"}), 400
    
    topic = data['topic'].strip()
    
    prs = Presentation()
    build_title_slide(prs, topic)
    
    # Add 26 more slides here - your existing logic
    for i in range(1, 27):
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = f"Slide {i}"
    
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    
    return send_file(
        pptx_io,
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        as_attachment=True,
        download_name=f'{topic}_deck.pptx'
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
