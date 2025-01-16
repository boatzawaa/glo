from flask import Flask, send_file, render_template, jsonify, request
from L6_Master import backendProcess, validateInputData
from openpyxl import Workbook
import io

app = Flask(__name__)
    
@app.route('/', methods=['GET', 'POST'])
def index():    
    if request.method == "POST":
        # ตรวจสอบ Content-Type
        if request.content_type != "application/json":
            return jsonify({"error": "Content-Type must be application/json"}), 400

        # อ่านข้อมูล JSON
        selectpt = request.json.get('selectpt')
        patterns = [request.json.get(f'pattern{i}') for i in range(1, 5)]
        set_value = request.json.get('set')
        charity = request.json.get('charity')
        lot = request.json.get('lot')
        year = request.json.get('year')

        if not request.data:
            return jsonify({"message": 'ไม่มีข้อมูลเข้า' , "status":'Error'}), 400
        
        # validate input        
        validate = validateInputData(selectpt,patterns,set_value,charity,lot,year)
        status = validate[0] 
        result_message = validate[1]
        if status == 'Error':
            return jsonify({"message": result_message , "status":status}), 400
        
        # ประมวลผลข้อมูล
        msg = backendProcess(selectpt,patterns,set_value,charity,lot,year)
        status = msg[0] 
        result_message = msg[1]
        if status == 'Error':
            return jsonify({"message": result_message , "status":status}), 400
        
        # ส่งไฟล์กลับไปยังผู้ใช้        
        wb =Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        for row in msg[1] :
            ws.append(row)
        # เขียน workbook ลงในหน่วยความจำ
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        # กำหนดชื่อไฟล์
        filename = msg[2]
        
        response = send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response.headers['X-Filename'] = filename  # เพิ่มชื่อไฟล์ใน Header
        return response
    
    return render_template("index.html", response_message=None)

# รันเซิร์ฟเวอร์
if __name__ == '__main__':
    #app.run(host="127.0.0.1", port=5000)
    #app.run(debug=True)
    app.run(debug=False)
