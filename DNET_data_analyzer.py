import time

import xlsxwriter
maxpower = int(input("Enter Max power\n"))
filename = input("Text Filename w/o .txt\n")
dataToFile=[]
with open(filename+'.txt','r') as f:
    for line in f:
        if "5FD" in line:
            data=[]
            temp=line.rstrip().split()
            data.append(float(temp[1]))
            data.append(round((int(temp[6]+temp[5],16)* maxpower)/4095,3))
            # print(data)
            #response line
            resp1=f.readline().rstrip().split()
            if resp1.index("[3FF]") >0:
                data.append(float(resp1[1]))
                data.append(round(int(resp1[7]+resp1[6],16)*maxpower/4095,3))
                data.append(round(int(resp1[9] + resp1[8],16)*maxpower/4095,3))
                # print("next =",resp1 )
                # print(data)
                # print("-----------------------------------------------------")
            resp2 = f.readline().rstrip().split()
            if resp2.index("[3FF]") > 0:
                data.append((float(resp2[1])))
                data.append(int(resp2[6]+resp2[7]))
                print(data)
                dataToFile.append(data)
                # print("-----------------------------------------------------")

#write to excel
excelFile= xlsxwriter.Workbook('output-'+filename+'--'+time.strftime("%Y%m%d-%H%M%S")+'.xlsx')
worksheet=excelFile.add_worksheet()
worksheet.write(0,0,"Setpoint timestamp")
worksheet.write(0,1,"Setpoint")
worksheet.write(0,2,"Forward response timestamp")
worksheet.write(0,3,"Forward Power (W)")
worksheet.write(0,4,"Reflected Power (W)")
worksheet.write(0,5,"status timestamp")
worksheet.write(0,6,"status code")


for row in range(len(dataToFile)):
    for col in range(len(dataToFile[row])):
        worksheet.write(row+1, col, dataToFile[row][col])

excelFile.close()


