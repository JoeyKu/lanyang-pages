from flask import Flask, request, send_file
from flask_cors import CORS
import os
from docx import Document
from docx.shared import Inches
import io

app = Flask(__name__)
CORS(app)  # 允許跨域請求

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        # 獲取表單數據
        date = request.form.get('date')
        time = request.form.get('time')
        room = request.form.get('room')
        name = request.form.get('name')
        format_type = request.form.get('format', 'docx')

        # 獲取上傳的文件
        file = request.files.get('file')
        if not file:
            return {'error': '沒有上傳文件'}, 400

        # 讀取 Word 文檔
        doc = Document(file)

        # 替換文檔中的佔位符
        # 這裡需要根據實際的 Word 文檔結構來調整替換邏輯
        for paragraph in doc.paragraphs:
            if '日期' in paragraph.text:
                paragraph.text = paragraph.text.replace('日期', f'日期：{date}')
            if '時間' in paragraph.text:
                paragraph.text = paragraph.text.replace('時間', f'時間：{time}')
            if '地點' in paragraph.text:
                paragraph.text = paragraph.text.replace('地點', f'地點：{room}')
            if '分會名稱' in paragraph.text:
                paragraph.text = paragraph.text.replace('分會名稱', f'分會名稱：{name}')

        # 保存到內存
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)

        # 如果需要 PDF，需要額外的庫如 docx2pdf
        # 但為了簡單，這裡只處理 docx

        return send_file(output, as_attachment=True, download_name=f'modified_meeting.{format_type}')

    except Exception as e:
        return {'error': str(e)}, 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8000)