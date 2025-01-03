import timeit

setCode = """
import pandas as pd
import numpy as np
import configparser
from openpyxl import workbook

#read config
config = configparser.ConfigParser()
config.read("config_test.ini") #config for test
#config.read("config.ini")
set = int(config["input"]["set"])
numPattern = np.array([config["input"]["pattern1"],
                       config["input"]["pattern2"],
                       config["input"]["pattern3"],
                       config["input"]["pattern4"],
                       config["input"]["pattern5"]])
  
#set default #######################################################################
path = config["default"]["path"] #default path
vTypes = str(config["default"]["vTypes"]) #default vTypes
vYear = str(config["default"]["vYear"]) #default vYear
vLotdateId = str(config["default"]["vLotdateId"]) #default vLotdateId
vSet = '01' #default set
vBook = '0000' #default book
vPattern = '1' #default  pattern
df = pd.DataFrame([[vTypes, vYear, vLotdateId, vSet, vBook, vPattern]],
                columns = ['Types', 'Year','Lotdate_id','Set', 'Book' ,'Pattern'])
print(df)

arrData = np.array(['Types', 'Year','Lotdate_id','Set', 'Book' ,'Pattern'])
arraySet = np.array([])
vlpBook = int(config["constant"]["vlpBook"])    # จำนวนเล่มต่อ 1 ชุด
loopBook = set*vlpBook #จำนวนเล่มทั้งหมด(25 x 10000 = 250000 เล่ม)
remainder = set%5 #หาชุดที่เป็นเศษเพื่อหาวิธีการจัดสลาก (25เศษ0/24เศษ4/23เศษ3/22เศษ2/21เศษ1)
remainderRow = remainder*vlpBook #จำนวนเล่มทั้งหมดที่เป็นเศษ (เศษ4 = 4 x 10000 = 40000)
numBook = 0 #ค่าเริ่มต้นจำนวนเล่ม
numSet = 0 #ค่าเริ่มต้นจำนวนสลากที่ให้กับตัวแทน 5 เล่ม เริ่มค่าจาก 0 to 4 = 5 เล่ม
pattern = 0 #ค่าเริ่มต้นจำนวน pattern ที่ใช้ 5 pattern
#default จำนวนเล่มสูงสุด
maxBook1 = int(config["constant"]["maxBook1"])
maxBook2 = int(config["constant"]["maxBook2"])
maxBook3 = int(config["constant"]["maxBook3"])
maxBook4 = int(config["constant"]["maxBook4"])
maxBook5 = int(config["constant"]["maxBook5"])
maxBook6 = int(config["constant"]["maxBook6"])
maxBook7 = int(config["constant"]["maxBook7"])
maxBook8 = int(config["constant"]["maxBook8"])
#default จำนวนเล่มเริ่มต้น
startBook1 = int(config["constant"]["startBook1"])
startBook2 = int(config["constant"]["startBook2"])
startBook3 = int(config["constant"]["startBook3"])
startBook4 = int(config["constant"]["startBook4"])
startBook5 = int(config["constant"]["startBook5"])
startBook6 = int(config["constant"]["startBook6"])
startBook7 = int(config["constant"]["startBook7"])
startBook8 = int(config["constant"]["startBook8"])
startBook9 = int(config["constant"]["startBook9"])
####################################################################################

#เก็บข้อมูล set ลงใน array
for i in range(set):
    arraySet = np.append(arraySet,i+1)
"""

test1 = """
#funtion insert data   
def insert2DataFrame(set,book,pattern,i) :    
    vSet = str(int(arraySet[set])).zfill(2)
    vBook = str(book).zfill(4)
    vPattern = numPattern[pattern]
    df.loc[i] = [vTypes, vYear, vLotdateId, vSet, vBook, vPattern]
    print(df) 
    
#เช็คเงื่อนไขเพื่อหาวิธีการจัดสลาก  
#เศษ 0
if remainder == 0 : #25 set
    for i in range(loopBook) : #250000 เล่ม        
        if (pattern == 4): #เช็คว่าเป็นpatternสุดท้าย
            insert2DataFrame(numSet,numBook,pattern,i)
            numSet = 0
            pattern = 0
            if (numBook == maxBook1) :
                numBook = 0
                arraySet = arraySet[5:] 
            else :
                numBook += 1                    
        else :
            insert2DataFrame(numSet,numBook,pattern,i) 
            pattern += 1
            numSet += 1
df.to_excel("test_df.xlsx",index=False) #สร้างไฟล์ excel
"""
print(timeit.timeit(test1,setup=setCode,number=1))