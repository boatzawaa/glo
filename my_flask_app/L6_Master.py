#import configparser
from openpyxl import Workbook

'''
section
1.funtion insert data
2.เช็คเงื่อนไขเพื่อหาวิธีการจัดสลาก
'''

############ เปลี่ยน config path ด้วย ############
#dirExcel = 'D:/Boatproject/python-project/L6/'
#dirExcel = '/home/boatzawa/'
###############################################

def validateInputData(selectpt,patterns,set_value,charity,lot,year) :
    response = []
    sets = int(set_value)
    charitySet = int(charity) #default charity set
    vLotdateId = str(lot) #default vLotdateId
    vYear = str(year) #default vYear
    if sets == 0:
        response = ['Error','จัดสลากไม่ถูกต้อง : กรุณาใส่"จำนวนชุด"ที่จัดสลาก']
    elif charitySet > sets:
        response = ['Error','จัดสลากไม่ถูกต้อง : จำนวนชุดของสลากการกุศล"เกิน"จำนวนชุดของสลากทั้งหมด']
    elif vLotdateId == '0':
        response = ['Error','จัดสลากไม่ถูกต้อง : กรุณาใส่"งวด"ที่จัดสลาก']
    elif vYear == '0':
        response = ['Error','จัดสลากไม่ถูกต้อง : กรุณาใส่"ปี"ที่จัดสลาก']     
    elif selectpt == '2111': #เช็คทุก pattern ต้องไม่ซ้ำกัน
        if (patterns[0] == patterns[1]) or (patterns[0] == patterns[2]) or (patterns[0] == patterns[3]) or (patterns[1] == patterns[2]) or (patterns[1] == patterns[3]) or (patterns[2] == patterns[3]):
            response = ['Error','จัดสลากไม่ถูกต้อง : มีรูปแบบการจัดสลากซ้ำกัน กรุณาแก้ไขให้ถูกต้อง']
    elif selectpt == '221':
        if (patterns[0] == patterns[1]) or (patterns[0] == patterns[2]) or (patterns[1] == patterns[2]):
            response = ['Error','จัดสลากไม่ถูกต้อง : มีรูปแบบการจัดสลากซ้ำกัน กรุณาแก้ไขให้ถูกต้อง']
    
    if response == [] :
        response = ['Success','Success']
    
    return response

def backendProcess(selectpt,patterns,set_value,charity,lot,year) :
    response = []
    try:
        numPattern = []
        if selectpt == "2111":   
            numPattern = [patterns[0],
                            patterns[0],
                            patterns[1],
                            patterns[2],
                            patterns[3]]
        else :
            numPattern = [patterns[0],
                            patterns[0],
                            patterns[1],
                            patterns[1],
                            patterns[2]]
        sets = int(set_value)
        charitySet = int(charity) #default charity set 
        types = str('01') #default vTypes
        charityType = str('02') #default charityType
        vLotdateId = str(lot) #default vLotdateId
        vYear = str(year) #default vYear
                    
        #set default ##############################################################################
        #path = dirExcel #default path
        arrData = [['Types', 'Year','Lotdate_id','Set', 'Book' ,'Pattern']] #header excel
        arraySet = [] 
        arrType = []      
        numBook = 0 #ค่าเริ่มต้นจำนวนเล่ม
        numSet = 0 #ค่าเริ่มต้นจำนวนสลากที่ให้กับตัวแทน 5 เล่ม เริ่มค่าจาก 0 to 4 = 5 เล่ม
        pattern = 0 #ค่าเริ่มต้นจำนวน pattern ที่ใช้ 5 pattern
        #default จำนวนเล่มสูงสุด
        vlpBook = 10000 # จำนวนเล่มต่อ 1 ชุด 
        maxBook9999 = 9999
        maxBook4999 = 4999
        maxBook3999 = 3999
        maxBook999 = 999
        #default จำนวนเล่มเริ่มต้น
        startBook9000 = 9000
        startBook8000 = 8000
        startBook7000 = 7000
        startBook5000 = 5000
        startBook4000 = 4000
        startBook1000 = 1000
        ###########################################################################################
        
        #เก็บข้อมูล set ลงใน array
        for i in range(sets):
            arraySet.append(i+1)
            
        #funtion progress bar #####################################################################
        num_progress_bar = 0
        def progress_bar(current, total, bar_length=50):
            progress = int(bar_length * current / total)
            bar = '#' * progress + '_' * (bar_length - progress)
            print(f"\r|{bar}| {current}/{total}", end='', flush=True)
        ###########################################################################################

        #funtion insert data ######################################################################  
        def insertData(type,set,book,pattern) :    
            vSet = str(arraySet[set]).zfill(2)
            vBook = str(book).zfill(4)
            vPattern = numPattern[pattern]
            arr = [type, vYear, vLotdateId, vSet, vBook, vPattern]
            return arr
        ###########################################################################################    

        print('start process...')     
        ####################### เช็คว่ามีสลากการกุศลหรือไม่ #######################
        set = sets
        if charitySet != 0 :
            arrType = [charityType,types]
            set -= charitySet
        else :
            arrType = [types]
        
        #loop1 สลากการกุศล , loop2 สลากL6   
        for type in arrType : 
            if type == '02' :
                vTypes = charityType #default charity type (02)
                loopBook = charitySet*vlpBook #ลูปของสลากการกุศล
                remainder = charitySet%5 #หาชุดที่เป็นเศษเพื่อหาวิธีการจัดสลาก (25เศษ0/24เศษ4/23เศษ3/22เศษ2/21เศษ1)
                remainderRow = remainder*vlpBook #จำนวนเล่มทั้งหมดที่เป็นเศษ (เศษ4 = 4 x 10000 = 40000)
            else :
                vTypes = types #default normal type (01)
                loopBook = set*vlpBook #จำนวนเล่มทั้งหมด(25 x 10000 = 250000 เล่ม)
                remainder = set%5 #หาชุดที่เป็นเศษเพื่อหาวิธีการจัดสลาก (25เศษ0/24เศษ4/23เศษ3/22เศษ2/21เศษ1)
                remainderRow = remainder*vlpBook #จำนวนเล่มทั้งหมดที่เป็นเศษ (เศษ4 = 4 x 10000 = 40000)
            
            ####################### เช็คเงื่อนไขเพื่อหาวิธีการจัดสลาก #######################  
            #เศษ 0
            if remainder == 0 : #25 set
                for i in range(loopBook) : # 250000 เล่ม            
                    if (pattern == 4): #เช็คว่าเป็นpatternสุดท้าย
                        arrData.append(insertData(vTypes,numSet,numBook,pattern))
                        numSet = 0
                        pattern = 0
                        if (numBook == maxBook9999) : #เช็คว่าเป็นเล่มสุดท้าย
                            numBook = 0
                            arraySet = arraySet[5:] #จัดสลาก 5 ชุดถัดไป
                        else :
                            numBook += 1                    
                    else :
                        arrData.append(insertData(vTypes,numSet,numBook,pattern))
                        pattern += 1
                        numSet += 1
                    num_progress_bar += 1
                    progress_bar(num_progress_bar, sets*vlpBook)                  
                                                    
            #เศษ 1                
            elif remainder == 1 : #21 set
                #print('เศษ 1')
                job2 = loopBook - remainderRow #loop ที่ 2 210000 - 10000 = 200000
                job1 = job2 - (5 * vlpBook) #loop ที่ 1 200000 - 50000 = 150000
                numBook2 = startBook5000 #ค่าเริ่มต้นของเล่ม (เล่มที่ 5000)
                numBook3 = startBook7000 #ค่าเริ่มต้นของเล่ม (เล่มที่ 7000)
                numBook4 = startBook9000 #ค่าเริ่มต้นของเล่ม (เล่มที่ 9000)
                
                #step1 จัดสลาก lot แรกก่อน เช่น ถ้าใส่ชุด 21 ชุด จะจัดสลากที่15ชุดก่อน
                for i in range(job1) : # 150000 เล่ม
                    if (pattern == 4) : #เช็คว่าเป็นpatternสุดท้าย
                        arrData.append(insertData(vTypes,numSet,numBook,pattern))
                        numSet = 0
                        pattern = 0
                        if (numBook == maxBook9999) : #เช็คว่าเป็นเล่มสุดท้าย
                            numBook = 0
                            arraySet = arraySet[5:] #จัดสลาก 5 ชุดถัดไป 
                        else :
                            numBook += 1                    
                    else :
                        arrData.append(insertData(vTypes,numSet,numBook,pattern))
                        pattern += 1
                        numSet += 1
                    num_progress_bar += 1
                    progress_bar(num_progress_bar, sets*vlpBook)
                            
                #step2 จัดสลาก lot 2 เช่น ถ้าใส่ชุด 21 ชุด จะจัดสลากที่15ชุดแรกแล้ว จัดอีก5ชุดหลังต่อ
                numSet = 0 #คุมชุด
                pattern = 0 #คุม pattern
                numFinalSet = 4 #ค่าเริ่มต้นเพื่อเพิ่มชุดสุดท้าย pattern 5  
                numFinalBook = 0 #ค่าเริ่มต้นเล่ม ของชุดสุดท้าย pattern 5   
                for i in range(job1,job2) : # เริ่ม 150000 ถึง 200000 = 50000 เล่ม
                    if (pattern == 4) : #เช็คว่าเป็นpatternสุดท้าย
                        arrData.append(insertData(vTypes,numFinalSet,numFinalBook,pattern))
                        numSet = 0  
                        pattern = 0                  
                        numBook += 1
                        if (numFinalBook == maxBook4999) :  #เช็คว่าเป็นเล่มสุดท้าย
                            numFinalBook = 0 #setค่าเริ่มต้น
                            numFinalSet += 1 #เปลี่ยน set เล่มที่ 5
                        else :
                            numFinalBook += 1
                    else :
                        arrData.append(insertData(vTypes,numSet,numBook,pattern))
                        pattern += 1
                        numSet += 1
                    num_progress_bar += 1
                    progress_bar(num_progress_bar, sets*vlpBook)     
                
                #step3 จัดสลากชุดที่เหลือ
                numSet = 0
                pattern = 0
                arraySet = arraySet[4:]
                numFinalSet = 0 #ค่าเริ่มต้นเพื่อเพิ่มชุดสุดท้าย pattern 5 
                for i in range(job2,loopBook) : # เริ่ม 200000 ถึง 210000 = 10000 เล่ม
                    if (pattern == 0) :
                        arrData.append(insertData(vTypes,numSet,numBook2,pattern))
                        numSet += 1
                        pattern += 1
                    elif (pattern == 1) :
                        arrData.append(insertData(vTypes,numSet,numBook2,pattern))
                        numSet = 0
                        pattern += 1
                        numBook2 += 1         
                    elif (pattern == 2) :
                        arrData.append(insertData(vTypes,numSet,numBook3,pattern))
                        numSet += 1
                        pattern += 1
                    elif (pattern == 3) :
                        arrData.append(insertData(vTypes,numSet,numBook3,pattern))
                        numSet = 0
                        pattern += 1
                        numBook3 += 1
                    elif (pattern == 4) :
                        arrData.append(insertData(vTypes,numFinalSet,numBook4,pattern))
                        numSet = 0  
                        pattern = 0
                        if (numBook4 == maxBook9999) : #เช็คว่าเป็นเล่มสุดท้าย
                            numBook4 = startBook9000 #setค่าเริ่มต้น
                            numFinalSet += 1 #เปลี่ยน set เล่มที่ 5  
                            print(arraySet)                     
                            if (arraySet[numFinalSet-1] == (charitySet)) and (len(arrType) == 2) :
                                numBook = 0
                                arraySet = arraySet[2:] #จัดสลากชุดถัดไป
                        else :
                            numBook4 += 1
                    num_progress_bar += 1
                    progress_bar(num_progress_bar, sets*vlpBook)   
                            
            #เศษ 2
            elif remainder == 2 : #22 set
                #print('เศษ 2')
                job1 = loopBook - remainderRow #loop ที่ 1 220000 - 20000 = 150000
                numBook3 = startBook4000 #setค่าเริ่มต้น 4000
                numBook4 = startBook8000 #setค่าเริ่มต้น 8000
                #step1 จัดสลาก lot แรกก่อน เช่น ถ้าใส่ชุด 22 ชุด จะจัดสลากที่ 20 ชุดก่อน
                for i in range(job1) : # 200000 เล่ม
                    if (pattern == 4) : #เช็คว่าเป็นpatternสุดท้าย
                        arrData.append(insertData(vTypes,numSet,numBook,pattern))
                        numSet = 0
                        pattern = 0
                        if (numBook == maxBook9999) : #เช็คว่าเป็นเล่มสุดท้าย
                            numBook = 0
                            arraySet = arraySet[5:] #จัดสลาก 5 ชุดถัดไป
                        else :
                            numBook += 1                    
                    else :
                        arrData.append(insertData(vTypes,numSet,numBook,pattern))
                        pattern += 1
                        numSet += 1  
                    num_progress_bar += 1
                    progress_bar(num_progress_bar, sets*vlpBook)
                
                #step2 จัดสลากชุดที่เหลืออีก 2 ชุด
                numSet = 0
                pattern = 0
                numFinalSet = 0 #ค่าเริ่มต้นเพื่อเพิ่มชุดสุดท้าย pattern 5 
                for i in range(job1,loopBook) : # เริ่ม 200000 ถึง 220000 = 20000 เล่ม
                    if (pattern == 0) :
                        arrData.append(insertData(vTypes,numSet,numBook,pattern))
                        numSet += 1
                        pattern += 1
                    elif (pattern == 1) :
                        arrData.append(insertData(vTypes,numSet,numBook,pattern))
                        numSet = 0
                        pattern += 1
                        numBook += 1         
                    elif (pattern == 2) :
                        arrData.append(insertData(vTypes,numSet,numBook3,pattern))
                        numSet += 1
                        pattern += 1
                    elif (pattern == 3) :
                        arrData.append(insertData(vTypes,numSet,numBook3,pattern))
                        numSet = 0
                        pattern += 1
                        numBook3 += 1
                    elif (pattern == 4) :
                        arrData.append(insertData(vTypes,numFinalSet,numBook4,pattern))
                        numSet = 0  
                        pattern = 0
                        if (numBook4 == maxBook9999) : #เช็คว่าเป็นเล่มสุดท้าย
                            numBook4 = startBook8000 #setค่าเริ่มต้นของเล่ม (8000)
                            numFinalSet += 1 #เปลี่ยน set เล่มที่ 5
                            #print(arraySet)                     
                            if (arraySet[numFinalSet-1] == (charitySet)) and (len(arrType) == 2) :
                                numBook = 0
                                arraySet = arraySet[2:] #จัดสลากชุดถัดไป
                        else :
                            numBook4 += 1 
                    num_progress_bar += 1
                    progress_bar(num_progress_bar, sets*vlpBook)
                    
            #เศษ 3   
            elif remainder == 3 : #23 set
                #print('เศษ 3')
                job2 = loopBook - remainderRow #loop ที่ 2 230000 - 30000 = 200000
                job1 = job2 - (5 * vlpBook) #loop ที่ 1 200000 - 50000 = 150000
                numBook2 = startBook4000 #setค่าเริ่มต้นของเล่ม (4000)
                numBook3 = startBook1000 #setค่าเริ่มต้นของเล่ม (1000)
                
                #step1 จัดสลาก lot แรกก่อน เช่น ถ้าใส่ชุด 23 ชุด จะจัดสลากที่ 15 ชุดก่อน
                for i in range(job1) : # 150000 เล่ม
                    if (pattern == 4) : #เช็คว่าเป็นpatternสุดท้าย
                        arrData.append(insertData(vTypes,numSet,numBook,pattern))
                        numSet = 0
                        pattern = 0
                        if (numBook == maxBook9999) : #เช็คว่าเป็นเล่มสุดท้าย
                            numBook = 0
                            arraySet = arraySet[5:] #จัดสลาก 5 ชุดถัดไป
                        else :
                            numBook += 1                    
                    else :
                        arrData.append(insertData(vTypes,numSet,numBook,pattern))
                        pattern += 1
                        numSet += 1
                    num_progress_bar += 1
                    progress_bar(num_progress_bar, sets*vlpBook)
                            
                #step2 จัดสลาก lot 2 เช่น ถ้าใส่ชุด 23 ชุด จะจัดสลากที่ 15 ชุดแรกแล้ว จัดอีก 5 ชุดหลังต่อ
                numSet = 0
                pattern = 0
                numFinalSet = 4 #ค่าเริ่มต้นเพื่อเพิ่มชุดสุดท้าย pattern 5  
                numFinalBook = 0 #ค่าเริ่มต้นเล่ม ของชุดสุดท้าย pattern 5   
                for i in range(job1,job2) : # เริ่ม 150000 ถึง 200000 = 50000 เล่ม
                    if (pattern == 4): #เช็คว่าเป็นpatternสุดท้าย
                        arrData.append(insertData(vTypes,numFinalSet,numFinalBook,pattern))
                        numSet = 0  
                        pattern = 0                  
                        numBook += 1
                        if ((numFinalBook == maxBook3999) or ((numFinalSet == 6) and (numFinalBook == maxBook999))):  #เช็คว่าเป็นเล่มสุดท้าย
                            numFinalBook = 0 #setค่าเริ่มต้น
                            numFinalSet += 1 #เปลี่ยน set เล่มที่ 5
                        else :
                            numFinalBook += 1
                    else :
                        arrData.append(insertData(vTypes,numSet,numBook,pattern))
                        pattern += 1
                        numSet += 1
                    num_progress_bar += 1
                    progress_bar(num_progress_bar, sets*vlpBook)            
                
                #step3 จัดสลากชุดที่เหลือ
                numSet = 0
                pattern = 0
                arraySet = arraySet[4:]
                numFinalSet = 2 #ค่าเริ่มต้นเพื่อเพิ่มชุดสุดท้าย pattern 5 
                for i in range(job2,loopBook) : # เริ่ม 200000 ถึง 230000 = 30000 เล่ม
                    if (pattern == 0) :
                        arrData.append(insertData(vTypes,numSet,numBook2,pattern))
                        numSet += 1
                        pattern += 1
                    elif (pattern == 1) :
                        arrData.append(insertData(vTypes,numSet,numBook2,pattern))
                        numSet += 1
                        pattern += 1        
                    elif (pattern == 2) :
                        arrData.append(insertData(vTypes,numSet,numBook2,pattern))
                        numSet += 1
                        pattern += 1
                    elif (pattern == 3) :
                        arrData.append(insertData(vTypes,numSet,numBook2,pattern))
                        numSet = 0
                        pattern += 1
                        numBook2 += 1
                    elif (pattern == 4) :
                        arrData.append(insertData(vTypes,numFinalSet,numBook3,pattern))
                        numSet = 0  
                        pattern = 0
                        if (numBook3 == maxBook3999) : #เช็คว่าเป็นเล่มสุดท้าย
                            numBook3 = startBook1000 #setค่าเริ่มต้น 1000
                            numFinalSet += 1 #เปลี่ยน set เล่มที่ 5
                            print(arraySet)                     
                            if (arraySet[numFinalSet-1] == (charitySet)) and (len(arrType) == 2) :
                                numBook = 0
                                arraySet = arraySet[4:] #จัดสลากชุดถัดไป
                        else :
                            numBook3 += 1 
                    num_progress_bar += 1
                    progress_bar(num_progress_bar, sets*vlpBook)
                    
            #เศษ 4                
            elif remainder == 4 : #24 set
                #print('เศษ 4')
                #step1 จัดสลาก lot แรกก่อน เช่น ถ้าใส่ชุด 24 ชุด จะจัดสลากที่ 20 ชุดก่อน
                job1 = loopBook - remainderRow #loop ที่ 1 240000 - 40000 = 200000
                numBook2 = startBook8000 #setค่าเริ่มต้น 8000
                for i in range(job1) : # 200000 เล่ม
                    if (pattern == 4) : #เช็คว่าเป็นpatternสุดท้าย
                        arrData.append(insertData(vTypes,numSet,numBook,pattern))
                        numSet = 0
                        pattern = 0
                        if (numBook == maxBook9999) : #เช็คว่าเป็นเล่มสุดท้าย
                            numBook = 0
                            arraySet = arraySet[5:] #จัดสลาก 5 ชุดถัดไป 
                        else :
                            numBook += 1                    
                    else :
                        arrData.append(insertData(vTypes,numSet,numBook,pattern))
                        pattern += 1
                        numSet += 1 
                    num_progress_bar += 1
                    progress_bar(num_progress_bar, sets*vlpBook)                        
                
                #step2 จัดสลากชุดที่เหลืออีก 4 ชุด
                numFinalSet = 0 #ค่าเริ่มต้นเพื่อเพิ่มชุดสุดท้าย pattern 5    
                for i in range(job1,loopBook) : # เริ่ม 200000 ถึง 240000 = 40000 เล่ม
                    if (pattern == 4) : #เช็คว่าเป็นpatternสุดท้าย
                        arrData.append(insertData(vTypes,numFinalSet,numBook2,pattern))
                        numSet = 0  
                        pattern = 0                  
                        numBook += 1
                        if (numBook2 == maxBook9999) : #เช็คว่าเป็นเล่มสุดท้าย
                            numBook2 = startBook8000 #setค่าเริ่มต้น 8000
                            numFinalSet += 1 #เปลี่ยน set เล่มที่ 5
                            print(arraySet)                     
                            if (arraySet[numFinalSet-1] == (charitySet)) and (len(arrType) == 2) :
                                numBook = 0
                                arraySet = arraySet[4:] #จัดสลากชุดถัดไป
                        else :
                            numBook2 += 1
                    else :
                        arrData.append(insertData(vTypes,numSet,numBook,pattern))
                        pattern += 1
                        numSet += 1
                    num_progress_bar += 1
                    progress_bar(num_progress_bar, sets*vlpBook)
                
        #generate excel file
        print('\ngenerate file . . .')  
        response = ['Success',arrData, f'L6_{set+charitySet}k.xlsx']
        
        '''
        #สำหรับ write file ลง เครื่อง
        try:        
            wb =Workbook()
            ws = wb.active
            for row in arrData :
                ws.append(row)
            wb.save(f'{path}L6_{set+charitySet}k.xlsx')
            response = ['Success',f'ระบบจัดสลากให้ท่านสำเร็จแล้ว! : ที่อยู่ของไฟล์ : {path}L6_{set+charitySet}k.xlsx']
            print("File written successfully!")

        except FileNotFoundError as e:
            # กรณีที่ไฟล์หรือไดเรกทอรีไม่พบ
            response = ['Error',f"File not found error: {e}"]
            print(f"พบปัญหาในการเขียนไฟล์ : File not found error: {e}")

        except IOError as e:
            # จับข้อผิดพลาด I/O (เช่น การเขียนไฟล์ไม่ได้)
            response = ['Error',f"พบปัญหาในการเขียนไฟล์ : IO error: {e}"]
            print(f"IO error: {e}")

        except Exception as e:
            # ข้อผิดพลาดอื่น ๆ ที่ไม่ได้คาดไว้
            response = ['Error',f"พบปัญหาในการเขียนไฟล์ : Unexpected error: {e}"]
            print(f"Unexpected error: {e}")  
        '''           
            
        return response
    except FileNotFoundError as e:        
        # กรณีที่ไฟล์หรือไดเรกทอรีไม่พบ
        response = ['Error',f"File not found error: {e}"]
        print(f"พบปัญหาในการจัดสลาก : File not found error: {e}")

    except IOError as e:
        # จับข้อผิดพลาด I/O (เช่น การเขียนไฟล์ไม่ได้)
        response = ['Error',f"พบปัญหาในการจัดสลาก : IO error: {e}"]
        print(f"IO error: {e}")

    except Exception as e:
        # ข้อผิดพลาดอื่น ๆ ที่ไม่ได้คาดไว้
        response = ['Error',f"พบปัญหาในการจัดสลาก : Unexpected error: {e}"]
        print(f"Unexpected error: {e}")
    
    return response
     
    