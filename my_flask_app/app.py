from flask import Flask, render_template, request

app = Flask(__name__)

# Route สำหรับหน้าแรก
@app.route('/')
def index():
    return render_template('form.html')  # แสดงหน้า form.html

# Route สำหรับรับข้อมูลจากฟอร์ม
@app.route('/submit', methods=['POST'])
def submit():
    name = request.form['name']   # รับค่าชื่อจากฟอร์ม
    email = request.form['email'] # รับค่าอีเมลจากฟอร์ม

    # ประมวลผลข้อมูล (เช่น บันทึกลงไฟล์)
    with open("data.txt", "a") as file:
        file.write(f"Name: {name}, Email: {email}\n")

    return f"<h1>Thank you, {name}!</h1><p>Your email ({email}) has been recorded.</p>"

# เริ่มต้นรันเซิร์ฟเวอร์
if __name__ == '__main__':
    app.run(debug=True)