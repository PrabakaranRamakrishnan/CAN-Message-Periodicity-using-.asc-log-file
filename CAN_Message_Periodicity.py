"""This Module is used to calculate periodicity of CAN messages
using .asc file by giving Message ids and its periodic time in seconds"""

import win32com.client

# Reading input file path
print("\nDrag n Drop the Log:\t")
LOG_PATH = str(input())  # Log path


# Creating file pointer
with open(LOG_PATH, "r", encoding="utf8") as fp_LogPath:
    Log_Data_Lines = fp_LogPath.readlines()


# variables to store Log data
LogData_list = []
Timestamp = []
CAN_ID = []
inc = 0

Message_ID = []
Message_Timestamp = []
prev_timestamp_flag = 0


# It stores list of CAN IDs
Message_IDs = [
    "C0320C8x",
    "CF01FC8x",
    "18FEC4C8x",
    "18FEC6C8x",
    "18F020C8x",
    "18E520C8x",
    "18FE5CC8x",
    "374753x",
]
CANID_CycleTime = [0.01, 0.01, 0.1, 0.1, 0.1, 0.1, 0.1]

Prev_timestamp = [0] * len(Message_IDs)
Periodicity = [0] * len(Message_IDs)
row = [0] * len(Message_IDs)


# Opening EXCEL
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = 0
wb = excel.Workbooks.Add()
for i, msg_id in enumerate(Message_IDs):
    globals()[msg_id] = wb.Worksheets.Add()  # adding sheets dynamically
    globals()[msg_id].Name = msg_id
wb.SaveAs("D:\\PRABAKARAN\\CAN_Periodicity.xlsx")
print("Processing ..... wait......\n")


# Splitting and storing the Log_Data_Lines
for line in Log_Data_Lines[1:]:
    line = line.rstrip("\n")  # removes the return from each line
    line = line.strip()  # removes the front spaces in each line
    words = line.split()  # split the line in to list

    if words != 0:
        Time = words[0]  # stores timestamp from words
        Channel = words[1]  # stores channel number from words

        if (
            Time[0:1] >= "0"
            and Time[0:1] <= "9"
            and Channel[0:1] >= "1"
            and Channel[0:1] <= "9"
        ):  # checks the first character is a number or not and CAN channel
            Timestamp.append(words[0])  # Stores Timestamps to an Array
            CAN_ID.append(words[2])  # Stores CAN IDs to an Array

            for i, msg_id in range(0, len(Message_IDs)):

                if (
                    CAN_ID[i] == msg_id
                ):  # checks current CAN ID with stored CAN ID list

                    if prev_timestamp_flag == 1:
                        Periodicity[i] = float(Timestamp[i]) - float(
                            Prev_timestamp[i]
                        )  # assigning time difference

                    prev_timestamp_flag = 1
                    Prev_timestamp[i] = Timestamp[
                        i
                    ]  # assigning current timestamp to the Previous timestamp
                    globals()[msg_id].Select()
                    globals()[msg_id].Cells(
                        row[i] + 7, 1
                    ).Value = Timestamp[i]

                    if row[i] >= 1:
                        globals()[msg_id].Cells(
                            row[i] + 7, 2
                        ).Value = Periodicity[i]

                    row[i] = row[i] + 1

            inc +=  1


# Storing Verdict in Excel sheets
for x in range(0, len(Message_IDs)):
    globals()[msg_id].Select()

    if (
        globals()[msg_id].Cells(7, 1).Value is not None
        and globals()[msg_id].Cells(8, 2).Value is not None
    ):

        globals()[msg_id].Cells(1, 5).Value = msg_id
        globals()[msg_id].Cells(1, 5).Font.Bold = True
        globals()[msg_id].Cells(1, 6).Value = "Observed"
        globals()[msg_id].Cells(1, 7).Value = str(
            "Acceptance criteria(+ or - 5%)"
        )
        globals()[msg_id].Cells(1, 8).Value = "Verdict"
        globals()[msg_id].Cells(2, 5).Value = "Cycle Time in seconds"
        globals()[msg_id].Cells(2, 6).Value = CANID_CycleTime[i]
        globals()[msg_id].Cells(2, 7).Value = CANID_CycleTime[i]

        globals()[msg_id].Cells(
            3, 5
        ).Value = "Minimum Periodicity time in seconds"
        globals()[msg_id].Cells(3, 6).Value = "=MIN(B:B)"
        globals()[msg_id].Cells(3, 7).Value = "=((F2/100)*(100-5))"

        if (
            globals()[msg_id].Cells(3, 6).Value
            >= globals()[msg_id].Cells(3, 7).Value
        ):
            globals()[msg_id].Cells(3, 8).Value = "PASS"

        else:
            globals()[msg_id].Cells(3, 8).Value = "FAIL"

        globals()[msg_id].Cells(
            4, 5
        ).Value = "Maximum Periodicity time in seconds"
        globals()[msg_id].Cells(4, 6).Value = "=MAX(B:B)"
        globals()[msg_id].Cells(4, 7).Value = "=((F2/100)*(100+5))"

        if (
            globals()[msg_id].Cells(4, 6).Value
            <= globals()[msg_id].Cells(4, 7).Value
        ):
            globals()[msg_id].Cells(4, 8).Value = "PASS"

        else:
            globals()[msg_id].Cells(4, 8).Value = "FAIL"

        if (
            globals()[msg_id].Cells(3, 8).Value == "PASS"
            and globals()[msg_id].Cells(4, 8).Value == "PASS"
        ):
            globals()[msg_id].Cells(2, 8).Value = "PASS"

        else:
            globals()[msg_id].Cells(2, 8).Value = "FAIL"

        globals()[msg_id].Cells(6, 1).Value = "Timestamp"
        globals()[msg_id].Cells(6, 2).Value = "Periodicity"

    else:
        globals()[msg_id].Cells(1, 6).Value = (
            "There is no " + msg_id + " message found in the log"
        )

excel.Application.Quit()
