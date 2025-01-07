import mysql.connector

#ทำการเชื่อมต่อกับฐานข้อมูลง่าย ๆ แค่ใส่ Host / User / Password ให้ถูกต้อง
connection = mysql.connector.connect(
 host="127.0.0.1",
 user="root",
 password="root",
 database="l6_db"
)
print(connection)
db_cursor = connection.cursor()

#เราสามารถรันคำสั่ง SQL ในการสร้าง Database  ได้เลย
#db_cursor.execute("CREATE DATABASE L6_db")

#สร้าง Table ลงไป ก็ใช้ Execute ได้เช่นกัน
#db_cursor.execute("CREATE TABLE L6 (Types VARCHAR(255), Year VARCHAR(255),Lotdate_id VARCHAR(255),Sets VARCHAR(255), Book VARCHAR(255),Pattern VARCHAR(255))")

#สร้าง String ไว้รอใส่คำสั่งได้เลย
sql_command = "INSERT INTO L6 (Types, Year , Lotdate_id , Sets , Book , Pattern) VALUES (%s, %s, %s, %s, %s, %s)"

#Value ที่ต้องการใส่ใน Command ทำไว้ในรูปแบบ Tuple ไว้ทำการ map กับคำสั่งด้านบนในตรง VALUES
val = ("02", "88","99", "00","0000", "0")

#สั่งให้คำสั่งทำงานได้เลย
db_cursor.execute(sql_command, val)

connection.commit()

#แสดงว่ามีกี่แถวที่ทำงานสำเร็จ
print(db_cursor.rowcount, "Succeed !")