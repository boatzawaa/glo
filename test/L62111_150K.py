import xlsxwriter
import pandas as pd

# จัดสลากรูปแบบ 2 1 1 1 สำหรับสลาก 150000 เล่ม

vlpunit     = 5     # จำนวนสลากที่ให้กับตัวแทน 5 เล่ม เริ่มค่าจาก 0 to 4 = 5 เล่ม
vlpSet      = 3      # รูปแบบ 2 1 1 1 คุมจำนวนรอบ (row) ที่จ่ายสลากให้ตัวแทนในแต่ละเล่มของ 5 เล่ม
#vlpSet =3 (ให้วน 3 รอบ คือ มี 3 pt
vlpBook     = 10000    # จำนวนเล่มทั้งหมดที่มีในชุด
#Loop ของ จำนวนเล่มที่ตัวแทนได้รับสลาก , จำนวนชุดที่ใช้, เล่ม มีสลาก 10000 เล่ม ใน 1 ชุด
# ตัวแปร สำหรับ งวดที่, รูปแบบ, ชุด, เล่ม

vnum2       = '00'  #รูปแบบการแสดงผล mask (ชุด และ เล่ม)
vnum4       = '0000'

vTypes = '01'
vYear  = '66'
vLotdateId  = '49' # เริ่มงวด 1 ตุลาคม 2566 งวดที่ 67 แต่งวดแรก เริ่มงวดที่ 49 งวดวันที่ 3012yyyy
vSet        = '01'
vBook       = '0000'
vPattern    = '4'

df = pd.DataFrame([[vTypes, vYear, vLotdateId, vSet, vBook, vPattern]],
                  columns = ['Types', 'Year','Lotdate_id','Set', 'Book' ,'Pattern' ])

#df.loc[1] = ['67', '1', '02','0000']
#df.loc[2] = ['67', '1', '01','0000']
#df.loc[3] = ['67', '1', '01','0000']
#df.loc[4] = ['67', '1', '01','0000']

#ใช้สลาก pettern ที่ 4 4 3 3 2 (จากเดิม 5 5 2 2 3 และ 1 1 2 2 3)
x = 0 #จำนวนเล่มทั้งหมด
# i คือ ตัวแปรจำนวนชุดสลากที่ใช้
# j คือ ตัวแปรจำนวนเล่มที่ใช้
# k คือ จำนวนเล่มทั้งหมดในชุดสลาก
#vlpunit คือ จำนวนเล่มที่ให้ตัวแทนแต่ละคน


for i in range(vlpSet):  #คุมจำนวนรอบ (row) ที่จ่ายสลากในแต่ละเล่ม ใน 5 เล่ม ที่ให้ตัวแทน คือ ในที่นี้ เล่มที่ 1,2.3.4,5 ใช้ 3 รอบ
    if i==0:
        z = 1  #ชุดเริ่มต้น ของ pettern ที่ 1 และ 2
    else:
        z += 2  # คุมชุดของสลากเล่มที่ 1, 2, 3 และ 4 ที่ตัวแทนจะได้รับ ซึ่งเป็นสลากของ pettern ที่ 1 2 ชุดและ pettern ที่ 2 2 ชุด
    for j in range(vlpBook): #จำนวนเล่มที่ใช้ คุมเล่ม
        for k in range(vlpunit): #จำนวนเล่มที่ให้ตัวแทนต่อ 1 คน k เริ่มจาก 0
            if k==0: # ตัวแทนรับเล่มที่ 1
                vPattern = str(k+4)
                vSet = str(z)
                vSet = vSet.rjust(len(vnum2),'0')
                vBook = str(j)
                vBook = vBook.rjust(len(vnum4),'0')
                df.loc[x] = [vTypes, vYear, vLotdateId, vSet, vBook, vPattern]     #[vLotdateId, vPattern, vSet, vBook]

                x+=1

            elif k==1: # ตัวแทนรับเล่มที่ 2
                vPattern = str(k+3)
                vSet = str(z+1)
                vBook = str(j)
                vBook = vBook.rjust(len(vnum4),'0')
                vSet = vSet.rjust(len(vnum2),'0')
                df.loc[x] = [vTypes, vYear, vLotdateId, vSet, vBook, vPattern] #[vLotdateId, vPattern, vSet, vBook]
                x+=1

            elif k==2: #ตัวแทนรับเล่มที่ 3
                vPattern = str(k+1)
                vSet = str(z+8)
                vSet = vSet.rjust(len(vnum2), '0')
                vBook = str(j)
                vBook = vBook.rjust(len(vnum4), '0')
                df.loc[x] = [vTypes, vYear, vLotdateId, vSet, vBook, vPattern]  # [vLotdateId, vPattern, vSet, vBook]

                x+=1

            elif k==3: # ตัวแทนรับเล่มที่ 4
                vPattern = str(k)
                vSet = str(z+9)
                vSet =  vSet.rjust(len(vnum2),'0')
                vBook = str(9999-j)
                vBook = vBook.rjust(len(vnum4),'0')
                df.loc[x] = [vTypes, vYear, vLotdateId, vSet, vBook, vPattern] #[vLotdateId, vPattern, vSet, vBook]
                x+=1

            elif k==4: #ตัวแทนรับเล่มที่ 5
                vPattern = str(k-2)
                vSet = str(i+17)
                vSet = vSet.rjust(len(vnum2),'0')
                vBook = str(j)
                vBook = vBook.rjust(len(vnum4),'0')
                df.loc[x] = [vTypes, vYear, vLotdateId, vSet, vBook, vPattern] #[vLotdateId, vPattern, vSet, vBook]
                x+=1
            print('Record ที่ '+str(x)+' i='+str(i))
        #if x >= 200000:
            #break

#print(df)
print('จำนวนเล่มทั้งหมด '+str(x)+' เล่ม')

df.to_excel('D:\Mywork/uploadL6/งวด16032567-54/จัดสลาก/l62111_lot16032567_150k.xlsx', sheet_name='l6 lot16032567-1')
