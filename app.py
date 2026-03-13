from flask import Flask, request, send_file, render_template, jsonify
from flask_cors import CORS # <--- 1. เพิ่มบรรทัดนี้
import docx
from docx.shared import RGBColor
import io
import json

app = Flask(__name__)
CORS(app) # <--- 2. เพิ่มบรรทัดนี้ (อนุญาตให้ทุกเว็บส่งข้อมูลมาได้)

# ... (โค้ดส่วนที่เหลือปล่อยไว้เหมือนเดิมเป๊ะเลยครับ) ...

from flask import Flask, request, send_file, render_template, jsonify
import docx
from docx.shared import RGBColor
import io
import json

app = Flask(__name__)

# 1. ฟังก์ชันเช็คสีแดง
def is_reddish(run):
    if run.font.color and run.font.color.rgb:
        hex_color = str(run.font.color.rgb).upper()
        try:
            r = int(hex_color[0:2], 16)
            g = int(hex_color[2:4], 16)
            b = int(hex_color[4:6], 16)
            if r > 130 and g < 100 and b < 100:
                return True
        except:
            pass
    return False

# 2. ฟังก์ชันกาววิเศษ (รวมคำสีแดงที่อยู่ติดกันให้เป็นประโยคเดียว)
def normalize_red_runs(doc):
    def _normalize_paragraph(p):
        first_red_run = None
        for run in p.runs:
            if is_reddish(run):
                if first_red_run is None:
                    first_red_run = run # เจอสีแดงก้อนแรก ให้จำไว้
                else:
                    # เจอสีแดงก้อนถัดมา เอาข้อความมาต่อท้ายก้อนแรก แล้วลบก้อนนี้ทิ้ง
                    first_red_run.text += run.text
                    run.text = ""
            else:
                first_red_run = None # ถ้าเจอสีดำ ให้ตัดจบการรวมร่าง

    # วนลูปติดกาวให้ทุกย่อหน้าและทุกตาราง
    for p in doc.paragraphs:
        _normalize_paragraph(p)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _normalize_paragraph(p)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyze', methods=['POST'])
def analyze():
    file = request.files.get('file')
    if not file:
        return jsonify({'error': 'ไม่พบไฟล์'}), 400
    
    doc = docx.Document(file)
    normalize_red_runs(doc) # <--- สั่งรวมคำก่อนดึงข้อมูลไปแสดงหน้าเว็บ
    
    red_texts = [] 
    seen = set()
    
    def extract_from_paragraphs(paragraphs):
        for p in paragraphs:
            for run in p.runs:
                if is_reddish(run):
                    text = run.text.strip()
                    if text and text not in seen:
                        seen.add(text)
                        red_texts.append(text)
                        
    extract_from_paragraphs(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                extract_from_paragraphs(cell.paragraphs)
                            
    return jsonify({'words': red_texts})

@app.route('/generate', methods=['POST'])
def generate():
    file = request.files.get('file')
    replacements_json = request.form.get('replacements')
    
    if not file or not replacements_json:
        return jsonify({'error': 'ข้อมูลไม่ครบถ้วน'}), 400
        
    replacements = json.loads(replacements_json)
    doc = docx.Document(file)
    
    normalize_red_runs(doc) # <--- สั่งรวมคำก่อนทำการแก้คำด้วย

    def replace_and_recolor(paragraphs, replacements):
        for p in paragraphs:
            for run in p.runs:
                if is_reddish(run):
                    original = run.text.strip()
                    if original in replacements:
                        new_text = replacements[original]
                        # ถ้าช่องกรอกข้อมูลไม่ว่างเปล่า ให้แทนที่คำและเปลี่ยนเป็นสีดำ
                        if new_text != "": 
                            run.text = run.text.replace(original, new_text)
                            run.font.color.rgb = RGBColor(0, 0, 0)
                        # ถ้าเป็นค่าว่าง โค้ดจะข้ามไปและปล่อยให้เป็นสีแดงตามเดิม
                        
    replace_and_recolor(doc.paragraphs, replacements)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_and_recolor(cell.paragraphs, replacements)
                                    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0) 
    
    return send_file(
        bio,
        as_attachment=True,
        download_name='เอกสาร_อัตโนมัติ.docx',
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

if __name__ == '__main__':
    app.run(debug=True, port=5000)