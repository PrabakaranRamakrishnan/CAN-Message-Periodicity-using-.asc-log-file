# CAN Message Periodicity using .asc log file
 can able to find the periodicity of a CAN message using valid .asc file by configuring the CAN Message Identifiers(N) and its Periodic Time.
 
 the resulted file will be a .xlsx format workbook with N number of sheets.
 
 
 - Line Number 31, 32 and 45 needs to be Cofigured before running the script
 
 Line Number 31 - provide your CAN Message IDs as a string in the list variable Message_IDs Ex: 'C0320C8x', 'CF01FC8x'

 Line Number 32 - provide the Periodic timimg in seconds of CAN Messages in the list variable CANID_CycleTime Ex: 0.01, 0.1
 
 Line Number 45 - provide the PATH to store the resulted file Ex:'D:\\PRABAKARAN\\CAN_Periodicity.xlsx'