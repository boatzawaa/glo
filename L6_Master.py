import configparser
from openpyxl import Workbook

'''
section
1.funtion insert data
2.เช็คเงื่อนไขเพื่อหาวิธีการจัดสลาก
'''

############ เปลี่ยน config path ด้วย ############
configPath = 'D:/Boatproject/python-project/'##
###############################################

#read config
config = configparser.ConfigParser()
#config.read(configPath+'config_test.ini') #config for test
config.read(configPath+'config.ini')
set = int(config['input']['set'])
numPattern = [config['input']['pattern1'],
                       config['input']['pattern2'],
                       config['input']['pattern3'],
                       config['input']['pattern4'],
                       config['input']['pattern5']]

#validate config
#เช็ค pattern 1 และ 2 ต้องตรงกัน
if numPattern[0] != numPattern[1]:
    print('Error : กรุณาใส่ Pattern1 และ Pattern2 ให้ตรงกัน')
#เช็คทุก pattern ต้องไม่ซ้ำกัน ยกเว้น1และ2
elif (numPattern[2] == numPattern[3]) or (numPattern[2] == numPattern[4]) or (numPattern[2] == numPattern[1]) or (numPattern[3] == numPattern[1]) or (numPattern[3] == numPattern[4]):
    print('Error : มี Pattern ซ้ำ กรุณาแก้ไขให้ถูกต้อง')
else :  
    #เช็คชุดต่ำสุดที่โปรแกรมจะทำงานได้
    if set < 11 :
        print('Error : จำนวนชุดต้องมีมากกว่า 10 ชุด')
    else :  
        print('start process!!') 
        #set default ##############################################################################
        path = config['default']['excelPath'] #default path
        vTypes = str(config['default']['vTypes']) #default vTypes
        vYear = str(config['default']['vYear']) #default vYear
        vLotdateId = str(config['default']['vLotdateId']) #default vLotdateId
        vSet = '01' #default set
        vBook = '0000' #default book
        vPattern = '1' #default  pattern
        arrData = [['Types', 'Year','Lotdate_id','Set', 'Book' ,'Pattern']]
        arraySet = []
        vlpBook = int(config['constant']['vlpBook'])    # จำนวนเล่มต่อ 1 ชุด
        loopBook = set*vlpBook #จำนวนเล่มทั้งหมด(25 x 10000 = 250000 เล่ม)
        remainder = set%5 #หาชุดที่เป็นเศษเพื่อหาวิธีการจัดสลาก (25เศษ0/24เศษ4/23เศษ3/22เศษ2/21เศษ1)
        remainderRow = remainder*vlpBook #จำนวนเล่มทั้งหมดที่เป็นเศษ (เศษ4 = 4 x 10000 = 40000)
        numBook = 0 #ค่าเริ่มต้นจำนวนเล่ม
        numSet = 0 #ค่าเริ่มต้นจำนวนสลากที่ให้กับตัวแทน 5 เล่ม เริ่มค่าจาก 0 to 4 = 5 เล่ม
        pattern = 0 #ค่าเริ่มต้นจำนวน pattern ที่ใช้ 5 pattern
        #default จำนวนเล่มสูงสุด
        maxBook1 = int(config['constant']['maxBook1'])
        maxBook2 = int(config['constant']['maxBook2'])
        maxBook3 = int(config['constant']['maxBook3'])
        maxBook4 = int(config['constant']['maxBook4'])
        maxBook5 = int(config['constant']['maxBook5'])
        maxBook6 = int(config['constant']['maxBook6'])
        maxBook7 = int(config['constant']['maxBook7'])
        maxBook8 = int(config['constant']['maxBook8'])
        #default จำนวนเล่มเริ่มต้น
        startBook1 = int(config['constant']['startBook1'])
        startBook2 = int(config['constant']['startBook2'])
        startBook3 = int(config['constant']['startBook3'])
        startBook4 = int(config['constant']['startBook4'])
        startBook5 = int(config['constant']['startBook5'])
        startBook6 = int(config['constant']['startBook6'])
        startBook7 = int(config['constant']['startBook7'])
        startBook8 = int(config['constant']['startBook8'])
        startBook9 = int(config['constant']['startBook9'])
        ###########################################################################################
        
        #เก็บข้อมูล set ลงใน array
        for i in range(set):
            arraySet.append(i+1)
            
        #funtion progress bar #####################################################################
        def progress_bar(current, total, bar_length=50):
            progress = int(bar_length * current / total)
            bar = '#' * progress + '_' * (bar_length - progress)
            print(f"\r|{bar}| {current}/{total}", end='', flush=True)
        ###########################################################################################
        
        #funtion insert data ######################################################################  
        def insertData(set,book,pattern,i) :    
            vSet = str(arraySet[set]).zfill(2)
            vBook = str(book).zfill(4)
            vPattern = numPattern[pattern]
            arr = [vTypes, vYear, vLotdateId, vSet, vBook, vPattern]
            return arr
        ###########################################################################################
          
        ####################### เช็คเงื่อนไขเพื่อหาวิธีการจัดสลาก #######################
        #เศษ 0
        if remainder == 0 : #25 set
            for i in range(loopBook) : # 250000 เล่ม        
                if (pattern == 4): #เช็คว่าเป็นpatternสุดท้าย
                    arrData.append(insertData(numSet,numBook,pattern,i))
                    numSet = 0
                    pattern = 0
                    if (numBook == maxBook1) : #เช็คว่าเป็นเล่มสุดท้าย
                        numBook = 0
                        arraySet = arraySet[5:] #จัดสลาก 5 ชุดถัดไป
                    else :
                        numBook += 1                    
                else :
                    arrData.append(insertData(numSet,numBook,pattern,i))
                    pattern += 1
                    numSet += 1
                progress_bar(i+1, loopBook)  
                                               
        #เศษ 1                
        elif remainder == 1 : #21 set
            #print('เศษ 1')
            job2 = loopBook - remainderRow #loop ที่ 2 210000 - 10000 = 200000
            job1 = job2 - (5 * vlpBook) #loop ที่ 1 200000 - 50000 = 150000
            numBook2 = startBook5 #ค่าเริ่มต้นของเล่ม (เล่มที่ 5000)
            numBook3 = startBook3 #ค่าเริ่มต้นของเล่ม (เล่มที่ 7000)
            numBook4 = startBook1 #ค่าเริ่มต้นของเล่ม (เล่มที่ 9000)
            
            #step1 จัดสลาก lot แรกก่อน เช่น ถ้าใส่ชุด 21 ชุด จะจัดสลากที่15ชุดก่อน
            for i in range(job1) : # 150000 เล่ม 
                if (pattern == 4) : #เช็คว่าเป็นpatternสุดท้าย
                    arrData.append(insertData(numSet,numBook,pattern,i))
                    numSet = 0
                    pattern = 0
                    if (numBook == maxBook1) : #เช็คว่าเป็นเล่มสุดท้าย
                        numBook = 0
                        arraySet = arraySet[5:] #จัดสลาก 5 ชุดถัดไป 
                    else :
                        numBook += 1                    
                else :
                    arrData.append(insertData(numSet,numBook,pattern,i))
                    pattern += 1
                    numSet += 1
                progress_bar(i+1, loopBook)
                        
            #step2 จัดสลาก lot 2 เช่น ถ้าใส่ชุด 21 ชุด จะจัดสลากที่15ชุดแรกแล้ว จัดอีก5ชุดหลังต่อ
            numSet = 0 #คุมชุด
            pattern = 0 #คุม pattern
            numFinalSet = 4 #ค่าเริ่มต้นเพื่อเพิ่มชุดสุดท้าย pattern 5  
            numFinalBook = 0 #ค่าเริ่มต้นเล่ม ของชุดสุดท้าย pattern 5   
            for i in range(job1,job2) : # เริ่ม 150000 ถึง 200000 = 50000 เล่ม 
                if (pattern == 4) : #เช็คว่าเป็นpatternสุดท้าย
                    arrData.append(insertData(numFinalSet,numFinalBook,pattern,i))
                    numSet = 0  
                    pattern = 0                  
                    numBook += 1
                    if (numFinalBook == maxBook4) :  #เช็คว่าเป็นเล่มสุดท้าย
                        numFinalBook = 0 #setค่าเริ่มต้น
                        numFinalSet += 1 #เปลี่ยน set เล่มที่ 5
                    else :
                        numFinalBook += 1
                else :
                    arrData.append(insertData(numSet,numBook,pattern,i))
                    pattern += 1
                    numSet += 1
                progress_bar(i+1, loopBook)     
            
            #step3 จัดสลากชุดที่เหลือ
            numSet = 0
            pattern = 0
            arraySet = arraySet[4:]
            numFinalSet = 0 #ค่าเริ่มต้นเพื่อเพิ่มชุดสุดท้าย pattern 5 
            for i in range(job2,loopBook) : # เริ่ม 200000 ถึง 210000 = 10000 เล่ม
                if (pattern == 0) :
                    arrData.append(insertData(numSet,numBook2,pattern,i))
                    numSet += 1
                    pattern += 1
                elif (pattern == 1) :
                    arrData.append(insertData(numSet,numBook2,pattern,i))
                    numSet = 0
                    pattern += 1
                    numBook2 += 1         
                elif (pattern == 2) :
                    arrData.append(insertData(numSet,numBook3,pattern,i))
                    numSet += 1
                    pattern += 1
                elif (pattern == 3) :
                    arrData.append(insertData(numSet,numBook3,pattern,i))
                    numSet = 0
                    pattern += 1
                    numBook3 += 1
                elif (pattern == 4) :
                    arrData.append(insertData(numFinalSet,numBook4,pattern,i))
                    numSet = 0  
                    pattern = 0
                    if (numBook4 == maxBook1) : #เช็คว่าเป็นเล่มสุดท้าย
                        numBook4 = startBook1 #setค่าเริ่มต้น
                        numFinalSet += 1 #เปลี่ยน set เล่มที่ 5
                    else :
                        numBook4 += 1
                progress_bar(i+1, loopBook)   
                     
        #เศษ 2
        elif remainder == 2 : #22 set
            #print('เศษ 2')
            job1 = loopBook - remainderRow #loop ที่ 1 220000 - 20000 = 150000
            numBook3 = startBook6 #setค่าเริ่มต้น 4000
            numBook4 = startBook2 #setค่าเริ่มต้น 8000
            #step1 จัดสลาก lot แรกก่อน เช่น ถ้าใส่ชุด 22 ชุด จะจัดสลากที่ 20 ชุดก่อน
            for i in range(job1) : # 200000 เล่ม         
                if (pattern == 4) : #เช็คว่าเป็นpatternสุดท้าย
                    arrData.append(insertData(numSet,numBook,pattern,i))
                    numSet = 0
                    pattern = 0
                    if (numBook == maxBook1) : #เช็คว่าเป็นเล่มสุดท้าย
                        numBook = 0
                        arraySet = arraySet[5:] #จัดสลาก 5 ชุดถัดไป
                    else :
                        numBook += 1                    
                else :
                    arrData.append(insertData(numSet,numBook,pattern,i))
                    pattern += 1
                    numSet += 1  
                progress_bar(i+1, loopBook)
            
            #step2 จัดสลากชุดที่เหลืออีก 2 ชุด
            numSet = 0
            pattern = 0
            numFinalSet = 0 #ค่าเริ่มต้นเพื่อเพิ่มชุดสุดท้าย pattern 5 
            for i in range(job1,loopBook) : # เริ่ม 200000 ถึง 220000 = 20000 เล่ม
                if (pattern == 0) :
                    arrData.append(insertData(numSet,numBook,pattern,i))
                    numSet += 1
                    pattern += 1
                elif (pattern == 1) :
                    arrData.append(insertData(numSet,numBook,pattern,i))
                    numSet = 0
                    pattern += 1
                    numBook += 1         
                elif (pattern == 2) :
                    arrData.append(insertData(numSet,numBook3,pattern,i))
                    numSet += 1
                    pattern += 1
                elif (pattern == 3) :
                    arrData.append(insertData(numSet,numBook3,pattern,i))
                    numSet = 0
                    pattern += 1
                    numBook3 += 1
                elif (pattern == 4) :
                    arrData.append(insertData(numFinalSet,numBook4,pattern,i))
                    numSet = 0  
                    pattern = 0
                    if (numBook4 == maxBook1) : #เช็คว่าเป็นเล่มสุดท้าย
                        numBook4 = startBook2 #setค่าเริ่มต้นของเล่ม (8000)
                        numFinalSet += 1 #เปลี่ยน set เล่มที่ 5
                    else :
                        numBook4 += 1 
                progress_bar(i+1, loopBook)
                
        #เศษ 3   
        elif remainder == 3 : #23 set
            #print('เศษ 3')
            job2 = loopBook - remainderRow #loop ที่ 2 230000 - 30000 = 200000
            job1 = job2 - (5 * vlpBook) #loop ที่ 1 200000 - 50000 = 150000
            numBook2 = startBook6 #setค่าเริ่มต้นของเล่ม (4000)
            numBook3 = startBook9 #setค่าเริ่มต้นของเล่ม (1000)
            
            #step1 จัดสลาก lot แรกก่อน เช่น ถ้าใส่ชุด 23 ชุด จะจัดสลากที่ 15 ชุดก่อน
            for i in range(job1) : # 150000 เล่ม 
                if (pattern == 4) : #เช็คว่าเป็นpatternสุดท้าย
                    arrData.append(insertData(numSet,numBook,pattern,i))
                    numSet = 0
                    pattern = 0
                    if (numBook == maxBook1) : #เช็คว่าเป็นเล่มสุดท้าย
                        numBook = 0
                        arraySet = arraySet[5:] #จัดสลาก 5 ชุดถัดไป
                    else :
                        numBook += 1                    
                else :
                    arrData.append(insertData(numSet,numBook,pattern,i))
                    pattern += 1
                    numSet += 1
                progress_bar(i+1, loopBook)
                        
            #step2 จัดสลาก lot 2 เช่น ถ้าใส่ชุด 23 ชุด จะจัดสลากที่ 15 ชุดแรกแล้ว จัดอีก 5 ชุดหลังต่อ
            numSet = 0
            pattern = 0
            numFinalSet = 4 #ค่าเริ่มต้นเพื่อเพิ่มชุดสุดท้าย pattern 5  
            numFinalBook = 0 #ค่าเริ่มต้นเล่ม ของชุดสุดท้าย pattern 5   
            for i in range(job1,job2) : # เริ่ม 150000 ถึง 200000 = 50000 เล่ม 
                if (pattern == 4): #เช็คว่าเป็นpatternสุดท้าย
                    arrData.append(insertData(numFinalSet,numFinalBook,pattern,i))
                    numSet = 0  
                    pattern = 0                  
                    numBook += 1
                    if ((numFinalBook == maxBook5) or ((numFinalSet == 6) and (numFinalBook == maxBook8))):  #เช็คว่าเป็นเล่มสุดท้าย
                        numFinalBook = 0 #setค่าเริ่มต้น
                        numFinalSet += 1 #เปลี่ยน set เล่มที่ 5
                    else :
                        numFinalBook += 1
                else :
                    arrData.append(insertData(numSet,numBook,pattern,i))
                    pattern += 1
                    numSet += 1
                progress_bar(i+1, loopBook)            
            
            #step3 จัดสลากชุดที่เหลือ
            numSet = 0
            pattern = 0
            arraySet = arraySet[4:]
            numFinalSet = 2 #ค่าเริ่มต้นเพื่อเพิ่มชุดสุดท้าย pattern 5 
            for i in range(job2,loopBook) : # เริ่ม 200000 ถึง 230000 = 30000 เล่ม 
                if (pattern == 0) :
                    arrData.append(insertData(numSet,numBook2,pattern,i))
                    numSet += 1
                    pattern += 1
                elif (pattern == 1) :
                    arrData.append(insertData(numSet,numBook2,pattern,i))
                    numSet += 1
                    pattern += 1        
                elif (pattern == 2) :
                    arrData.append(insertData(numSet,numBook2,pattern,i))
                    numSet += 1
                    pattern += 1
                elif (pattern == 3) :
                    arrData.append(insertData(numSet,numBook2,pattern,i))
                    numSet = 0
                    pattern += 1
                    numBook2 += 1
                elif (pattern == 4) :
                    arrData.append(insertData(numFinalSet,numBook3,pattern,i))
                    numSet = 0  
                    pattern = 0
                    if (numBook3 == maxBook5) : #เช็คว่าเป็นเล่มสุดท้าย
                        numBook3 = startBook9 #setค่าเริ่มต้น 1000
                        numFinalSet += 1 #เปลี่ยน set เล่มที่ 5
                    else :
                        numBook3 += 1 
                progress_bar(i+1, loopBook)
                
        #เศษ 4                
        elif remainder == 4 : #24 set
            #print('เศษ 4')
            #step1 จัดสลาก lot แรกก่อน เช่น ถ้าใส่ชุด 24 ชุด จะจัดสลากที่ 20 ชุดก่อน
            job1 = loopBook - remainderRow #loop ที่ 1 240000 - 40000 = 200000
            numBook2 = startBook2 #setค่าเริ่มต้น 8000
            for i in range(job1) : # 200000 เล่ม 
                if (pattern == 4) : #เช็คว่าเป็นpatternสุดท้าย
                    arrData.append(insertData(numSet,numBook,pattern,i))
                    numSet = 0
                    pattern = 0
                    if (numBook == maxBook1) : #เช็คว่าเป็นเล่มสุดท้าย
                        numBook = 0
                        arraySet = arraySet[5:] #จัดสลาก 5 ชุดถัดไป 
                    else :
                        numBook += 1                    
                else :
                    arrData.append(insertData(numSet,numBook,pattern,i))
                    pattern += 1
                    numSet += 1 
                progress_bar(i+1, loopBook)                        
            
            #step2 จัดสลากชุดที่เหลืออีก 4 ชุด
            numFinalSet = 0 #ค่าเริ่มต้นเพื่อเพิ่มชุดสุดท้าย pattern 5    
            for i in range(job1,loopBook) : # เริ่ม 200000 ถึง 240000 = 40000 เล่ม 
                if (pattern == 4) : #เช็คว่าเป็นpatternสุดท้าย
                    arrData.append(insertData(numFinalSet,numBook2,pattern,i))
                    numSet = 0  
                    pattern = 0                  
                    numBook += 1
                    if (numBook2 == maxBook1) : #เช็คว่าเป็นเล่มสุดท้าย
                        numBook2 = startBook2 #setค่าเริ่มต้น 8000
                        numFinalSet += 1 #เปลี่ยน set เล่มที่ 5
                    else :
                        numBook2 += 1
                else :
                    arrData.append(insertData(numSet,numBook,pattern,i))
                    pattern += 1
                    numSet += 1
                progress_bar(i+1, loopBook)
                
        #write excel file
        print('\nwrite data . . .')               
        wb =Workbook()
        ws = wb.active
        for row in arrData :
            ws.append(row)
        wb.save(path+'L6_'+str(set)+'k.xlsx')
        print('Done!')