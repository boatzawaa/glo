from flask import Flask, render_template, jsonify, request
from L6_Master import backendProcess, validateInputData

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
            return jsonify({"error": "No input provided"}), 400
        
        # validate input        
        validate = validateInputData(selectpt,patterns,set_value,charity,lot,year)
        status = validate[0] 
        result_message = validate[1]
        if status == 'Error':
            return jsonify({"message": result_message , "status":status})
        
        # ประมวลผลข้อมูล
        msg = backendProcess(selectpt,patterns,set_value,charity,lot,year)
        status = msg[0] 
        result_message = msg[1]    
        return jsonify({"message": result_message , "status":status})
    
    return render_template("index.html", response_message=None)

# รันเซิร์ฟเวอร์
if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000)
    #app.run(debug=True)
    #app.run(debug=False)
