from flask import Flask, render_template, request
from L6_Master import backendProcess

app = Flask(__name__)

# Route สำหรับแสดงหน้าเว็บ
@app.route('/', methods=['GET', 'POST'])
def index():
    result_message = None  # ตัวแปรเก็บผลลัพธ์ที่จะส่งกลับไปแสดงผล
    alert = 'alert-success'
    
    if request.method == 'POST':
        # รับค่าจากฟอร์ม
        patterns = [request.form.get(f'pattern{i}') for i in range(1, 5)]
        set_value = request.form.get('set')
        charity = request.form.get('charity')
        lot = request.form.get('lot')
        year = request.form.get('year')

        # ประมวลผลข้อมูล
        msg = None
        msg = backendProcess(patterns,set_value,charity,lot,year)
        print(msg[0])
        if msg[0] != 'success':
            alert = 'alert-danger'   
        result_message = msg[1]
        
    # ส่ง result_message ไปที่ HTML 
    return render_template('index.html', alert=alert, result_message=result_message)

# รันเซิร์ฟเวอร์
if __name__ == '__main__':
    app.run(debug=True)