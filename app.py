import os
import base64
from io import BytesIO
import pandas as pd
from flask import Flask, request, redirect, render_template, flash
from werkzeug.utils import secure_filename
from PIL import Image
import pytesseract as tess
from datetime import datetime




# Flask application setup
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # Maximum file size: 16 MB
app.config['ALLOWED_EXTENSIONS'] = {'png', 'jpg', 'jpeg', 'gif'}
app.secret_key = 'supersecretkey'

# Ensure the upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Function to check allowed file extensions
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

import re

# Function to extract only numbers from the OCR output
def extract_text_from_image(image_path):
    try:
        # Set the path to Tesseract
        tess.pytesseract.tesseract_cmd = r'C:\Users\saowalak_t\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'  # Change as necessary
        
        img = Image.open(image_path)
        # custom_config = r'--oem 3 --psm 6 -l tha+eng -c tessedit_char_whitelist=กขค0123456789'
        custom_config = r'tesseract --oem 3 --psm 12 -l tha+osd --dpi 2400 -c tessedit_char_whitelist=กขค0123456789 -c tessedit_char_blacklist=.-'


        # ทำ OCR จากรูปภาพขาวดำ
        text = tess.image_to_string(img, config=custom_config)
        
        # Use regular expressions to extract only numbers (integer and decimal)
        numbers = re.findall(r'\d+(?:\.\d+)?', text)
        print(f"Extracted Numbers: {numbers}")  # Print the extracted numbers
        
        return '\n'.join(numbers)  # Join numbers with newlines for better display
    except Exception as e:
        print(f"Error processing image: {e}")
        return None


# Route for the homepage
@app.route('/')
def index():
    return render_template('index.html')

# Route to handle image uploads and OCR
@app.route('/upload', methods=['POST'])
def upload_image():
    # Check if a captured image is provided
    captured_image_data = request.form.get('captured_image')
    
    if captured_image_data:
        # Extract base64 data and convert it to an image
        image_data = base64.b64decode(captured_image_data.split(',')[1])
        image = Image.open(BytesIO(image_data))
        
        # Save the image
        filename = f"captured_image_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        image.save(file_path)

        # Proceed to extract text from the captured image
        extracted_text = extract_text_from_image(file_path)
        if extracted_text:
            return render_template('show_text.html', extracted_text=extracted_text.split('\n'))
        else:
            flash('Failed to extract text from the image.')
            return redirect('/')

    # Process uploaded image if available
    if 'file' in request.files:
        file = request.files['file']
        if file.filename != '':
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            
            # Extract text from the image
            extracted_text = extract_text_from_image(file_path)
            if extracted_text:
                return render_template('show_text.html', extracted_text=extracted_text.split('\n'))
            else:
                flash('Failed to extract text from the image.')
                return redirect('/')
    
    flash('No image provided')
    return redirect('/')

# Route to handle saving the edited text into an Excel file

@app.route('/save_text', methods=['POST'])
def save_text():
    # ดึงข้อมูลที่แก้ไขจาก form
    edited_text = [
        request.form.get('edited_text_1'),
        request.form.get('edited_text_2'),
        request.form.get('edited_text_3'),
        request.form.get('edited_text_4'),
        request.form.get('edited_text_5')
    ]

    # ตรวจสอบว่ามีข้อมูลที่ถูกส่งมาจริง
    if edited_text:
        try:
            # ชื่อแถวที่จะแสดงใน Excel
            row_names = ['รหัสพาเลท', 'รหัสผลิต', 'รหัสสินค้า', 'วันที่ถอดแผ่น', 'จำนวนแผ่น']
            
            # สร้าง DataFrame สำหรับบันทึกใน Excel
            data = {
                'Line': row_names,
                'Extracted Text': edited_text
            }
            df = pd.DataFrame(data)

            # สร้างชื่อไฟล์ตามวันที่และเวลา
            filename = f"Tag_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)

            # บันทึกลงไฟล์ Excel
            df.to_excel(file_path, index=False, engine='openpyxl')

            # แจ้งเตือนผู้ใช้ว่าไฟล์ถูกบันทึกเรียบร้อยแล้ว
            flash(f'Text saved to Excel file: {filename}')
        except Exception as e:
            flash(f"Error saving text to Excel: {e}")

    return redirect('/')



# Main entry point
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5555, debug=True)
    
