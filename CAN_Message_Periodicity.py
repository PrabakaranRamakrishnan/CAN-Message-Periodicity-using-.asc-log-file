
import re
import string
import os
import win32com.client
from win32com.client import Dispatch

# Reading input file path
print ("\nDrag n Drop the Log:\t")
Log_Path = str(input()) #Log path


# Creating file pointer
fp_LogPath=open(Log_Path,"r") #Filepointer for Log data
Log_Data_Lines =  fp_LogPath.readlines()


#variables to store Log data
LogData_list = []	
Timestamp = []
CAN_ID = []
temp= 0
i=0

Message_ID = []
Message_Timestamp = []
Prev_timestamp_flag = 0


#It stores list of CAN IDs
Message_IDs = ['C0320C8x','CF01FC8x','18FEC4C8x','18FEC6C8x','18F020C8x','18E520C8x','18FE5CC8x','374753x'] 
CANID_CycleTime = [0.01,0.01,0.1,0.1,0.1,0.1,0.1]

Prev_timestamp = [0] * len(Message_IDs)
Periodicity = [0] * len(Message_IDs)
row = [0] * len(Message_IDs)


#Opening EXCEL
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = 0
wb = excel.Workbooks.Add()
for id in range(0,len(Message_IDs)):
	globals()[(Message_IDs[id])] = wb.Worksheets.Add()				# adding sheets dynamically
wb.SaveAs('D:\\PRABAKARAN\\CAN_Periodicity.xlsx')
print("Processing ..... wait......\n")

	
# Splitting and storing the Log_Data_Lines
for line in Log_Data_Lines[1:]:
	line = line.rstrip('\n') 							# removes the return from each line
	line = line.strip() 								# removes the front spaces in each line
	words = line.split() 								# split the line in to list
	
	if words != 0:
		Time = words[0]									# stores timestamp from words
		Channel = words[1]								# stores channel number from words
							
		if Time[0:1] >= '0' and Time[0:1] <= '9' and Channel[0:1] >=  '1' and  Channel[0:1]<= '9': 	# checks the first character is a number or not and CAN channel 
			Timestamp.append(words[0]) 					# Stores Timestamps to an Array
			CAN_ID.append(words[2]) 					# Stores CAN IDs to an Array

			for x in range(0,len(Message_IDs)):

				if CAN_ID[i] == Message_IDs[x] : 		# checks current CAN ID with stored CAN ID list
					
					if Prev_timestamp_flag == 1:
						Periodicity[x] = float(Timestamp[i]) - float(Prev_timestamp[x]) # assigning time difference 
		 	
					Prev_timestamp_flag = 1 
					Prev_timestamp[x] = Timestamp[i] # assigning current timestamp to the Previous timestamp
					globals()[str(Message_IDs[x])].Select()
					globals()[str(Message_IDs[x]) ].Cells(row[x]+7,1).Value = Timestamp[i]
					
					if row[x] >=1 : 
						globals()[str(Message_IDs[x])].Cells(row[x]+7,2).Value = Periodicity[x]
						
					row[x] = row[x] + 1
			
			i=i+1


#Storing Verdict in Excel sheets
for x in range(0,len(Message_IDs)):
	globals()[str(Message_IDs[x])].Select()

	if globals()[str(Message_IDs[x]) ].Cells(7,1).Value != None and globals()[str(Message_IDs[x]) ].Cells(8,2).Value != None :
	
		globals()[str(Message_IDs[x]) ].Cells(1,5).Value = Message_IDs[x]
		globals()[str(Message_IDs[x]) ].Cells(1,5).Font.Bold = True
		globals()[str(Message_IDs[x]) ].Cells(1,6).Value = "Observed"
		globals()[str(Message_IDs[x]) ].Cells(1,7).Value = str('Acceptance criteria(+ or - 5%)')
		globals()[str(Message_IDs[x]) ].Cells(1,8).Value = "Verdict"
		globals()[str(Message_IDs[x]) ].Cells(2,5).Value = "Cycle Time in seconds"
		globals()[str(Message_IDs[x]) ].Cells(2,6).Value = CANID_CycleTime[x]
		globals()[str(Message_IDs[x]) ].Cells(2,7).Value = CANID_CycleTime[x]

		globals()[str(Message_IDs[x]) ].Cells(3,5).Value = "Minimum Periodicity time in seconds"
		globals()[str(Message_IDs[x]) ].Cells(3,6).Value = '=MIN(B:B)'
		globals()[str(Message_IDs[x]) ].Cells(3,7).Value = '=((F2/100)*(100-5))'
	
		if globals()[str(Message_IDs[x]) ].Cells(3,6).Value >= globals()[str(Message_IDs[x]) ].Cells(3,7).Value :
			globals()[str(Message_IDs[x]) ].Cells(3,8).Value = "PASS"
		
		else :
			globals()[str(Message_IDs[x]) ].Cells(3,8).Value = "FAIL"

		globals()[str(Message_IDs[x]) ].Cells(4,5).Value = "Maximum Periodicity time in seconds"
		globals()[str(Message_IDs[x]) ].Cells(4,6).Value = '=MAX(B:B)'
		globals()[str(Message_IDs[x]) ].Cells(4,7).Value = '=((F2/100)*(100+5))'
		
		if globals()[str(Message_IDs[x]) ].Cells(4,6).Value <= globals()[str(Message_IDs[x]) ].Cells(4,7).Value :
			globals()[str(Message_IDs[x]) ].Cells(4,8).Value = "PASS"

		else :
			globals()[str(Message_IDs[x]) ].Cells(4,8).Value = "FAIL"

		if globals()[str(Message_IDs[x]) ].Cells(3,8).Value == "PASS" and globals()[str(Message_IDs[x]) ].Cells(4,8).Value == "PASS" :
			globals()[str(Message_IDs[x]) ].Cells(2,8).Value = "PASS"
		
		else :
			globals()[str(Message_IDs[x]) ].Cells(2,8).Value = "FAIL"
		
		globals()[str(Message_IDs[x]) ].Cells(6,1).Value = "Timestamp"
		globals()[str(Message_IDs[x]) ].Cells(6,2).Value = "Periodicity"
	
	else:
		globals()[str(Message_IDs[x]) ].Cells(1,6).Value = "There is no "+Message_IDs[x]+" message found in the log"
			
fp_LogPath.close() 
excel.Application.Quit()

