import os
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.styles import Alignment, Border, Side
from openpyxl.styles import colors, Color
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

print("hello")

# source path
# filePath = "D://Workspace//Output//20230803_2271_sde//"
# name = "sda"
filePath = "D://Workspace//py_example//src//Samsung_1T_20230814//"

# output path
outFilePath = "D://Workspace//py_example//src//rslt//"
# source files
sourceFileList = [
    "128K_seqR.log",
    "128K_seqW.log",
    "4K_randRT1Q32.log",
    "4K_randWT1Q32.log",
    "4K_mixWRT1Q32.log",
    "8K_randRT1Q32.log",
    "8K_randWT1Q32.log",
    "8K_mixWRT1Q32.log",
    "4K_randRT1Q1.log",
    "4K_randWT1Q1.log",
]

# string location flag
strFlag0 = "Run status group 0 (all jobs):"
strFlag1 = "IOPS="
strFlagList = [
    strFlag0,
    strFlag0,
    strFlag1,
    strFlag1,
    strFlag1,
    strFlag1,
    strFlag1,
    strFlag1,
]

strLanFlag = "clat percentiles (usec):"
strLanFlagQD1 = "clat (nsec):"
strLanFlagQD2 = "clat (usec):"

# rslt
rsltFileList = []
rsltBwList = []
rsltLanList = []
rsltLanListQD1 = []

# handle
for i in range(0, 8):
    # print(sourceFileList[i])
    tempIOPS = 1.1
    with open(filePath + sourceFileList[i]) as f:
        line = f.readline()
        tempIOPS = 0
        while line:
            # print(strFlagList[i])
            if i == 2 or i == 3:
                if strLanFlag in line:
                    line = f.readline()
                    lanIdx = 0
                    while lanIdx < len(line):
                        startIdx = line.find("[", lanIdx, len(line))
                        if startIdx == -1:
                            break
                        endIdx = line.find("]", startIdx, len(line))
                        # print(int(line[startIdx + 1 : endIdx]))
                        rsltLanList.append(int(line[startIdx + 1 : endIdx]))
                        lanIdx = endIdx + 1

            if strFlagList[i] in line:
                if i == 0 or i == 1:
                    line = f.readline()
                    startIdx = line.find("(", 0, len(line)) + 1
                    endIdx = line.find("MB/s", 0, len(line))
                    rsltBwList.append(float(line[startIdx:endIdx]))
                    # print(line[startIdx:endIdx])
                    break
                else:
                    startIdx = line.find("=", 0, len(line)) + 1
                    endIdx = line.find("k", 0, len(line))
                    if endIdx == -1:
                        endIdx = line.find(",", 0, len(line))
                        print(line[startIdx:endIdx])
                        print(float(line[startIdx:endIdx]) / 1000)
                        # tempIOPS += round(float(line[startIdx:endIdx]) / 1000, 2)
                        tempIOPS += float(line[startIdx:endIdx]) / 1000
                    else:
                        tempIOPS += float(line[startIdx:endIdx])

                    # print(line[startIdx:endIdx])
                    # tempIOPS += float(line[startIdx:endIdx])

            line = f.readline()

    if tempIOPS != 0:
        rsltBwList.append(tempIOPS)

for i in range(8, 10):
    with open(filePath + sourceFileList[i]) as f:
        line = f.readline()
        while line:
            # if strLanFlagQD1 in line or strLanFlagQD2 in line:    #时延为us的情况
            if strLanFlagQD1 in line:
                startIdx = line.find("min=", 0, len(line))
                endIdx = line.find(", stdev", startIdx, len(line))
                # print(line[startIdx:endIdx])
                rsltLanListQD1.append(line[startIdx:endIdx])
                lanIdx = endIdx + 1
                break

            line = f.readline()


# output excel

# form title
title0 = "Unit"
titleList0 = ["MB/s", "KIOPS", "usec", "nsec"]
title1 = "Steady Status QD32"
titleList1 = [
    "100%Sequential Reads-128K, QD32, SS",
    "100%Sequential Writes-128K, QD32, SS",
    "100%Random Reads-4K, QD32, SS",
    "100%Random Writes-4K, QD32, SS",
    "70%R 30%W Mixed-4K, QD32, SS",
    "100%Random Reads-8K, QD32, SS",
    "100%Random Writes-8K, QD32, SS",
    "70%R 30%W Mixed-8K, QD32, SS",
]
title2 = "Performace"
title3 = "t3"
titleLan = ["Qos, QD32, SS", "qos", "latency", "Lantency, QD1, SS"]
titleLanList = ["99%", "99.9%", "99.99%"]

# form body
perf2 = [
    {title0: title0, title1: title1, title2: title2},
    {title0: titleList0[0], title1: titleList1[0], title2: rsltBwList[0]},
    {title0: titleList0[0], title1: titleList1[1], title2: rsltBwList[1]},
    {title0: titleList0[1], title1: titleList1[2], title2: rsltBwList[2]},
    {title0: titleList0[1], title1: titleList1[3], title2: rsltBwList[3]},
    {title0: titleList0[1], title1: titleList1[4], title2: rsltBwList[4]},
    {title0: titleList0[1], title1: titleList1[5], title2: rsltBwList[5]},
    {title0: titleList0[1], title1: titleList1[6], title2: rsltBwList[6]},
    {title0: titleList0[1], title1: titleList1[7], title2: rsltBwList[7]},
]

perf3 = [
    {title0: title0, title1: titleLan[0], title2: titleLan[1], title3: titleLan[2]},
    {
        title0: titleList0[2],
        title1: titleList1[2],
        title2: titleLanList[0],
        title3: rsltLanList[0],
    },
    {
        title0: titleList0[2],
        title1: titleList1[2],
        title2: titleLanList[1],
        title3: rsltLanList[1],
    },
    {
        title0: titleList0[2],
        title1: titleList1[2],
        title2: titleLanList[2],
        title3: rsltLanList[2],
    },
    {
        title0: titleList0[2],
        title1: titleList1[3],
        title2: titleLanList[0],
        title3: rsltLanList[3],
    },
    {
        title0: titleList0[2],
        title1: titleList1[3],
        title2: titleLanList[1],
        title3: rsltLanList[4],
    },
    {
        title0: titleList0[2],
        title1: titleList1[3],
        title2: titleLanList[2],
        title3: rsltLanList[5],
    },
]


perf4 = [
    {title0: title0, title1: titleLan[3], title2: titleLan[2]},
    {title0: titleList0[3], title1: titleList1[2], title2: rsltLanListQD1[0]},
    {title0: titleList0[3], title1: titleList1[3], title2: rsltLanListQD1[1]},
]


# output process
workbook = Workbook()
sheet = workbook.active

sheet.title = "性能报告"
# sheet.append([title0, title1, title2])
for data in perf2:
    sheet.append(list(data.values()))

for data in perf3:
    sheet.append(list(data.values()))

for data in perf4:
    sheet.append(list(data.values()))


# format
for i in range(1, 10):
    sheet.merge_cells(range_string="C" + str(i) + ":" + "D" + str(i))

sheet.merge_cells(range_string="A11:A13")
sheet.merge_cells(range_string="A14:A16")
sheet.merge_cells(range_string="B11:B13")
sheet.merge_cells(range_string="B14:B16")

for i in range(17, 20):
    sheet.merge_cells(range_string="C" + str(i) + ":" + "D" + str(i))

alignment = Alignment(
    horizontal="center", vertical="center", text_rotation=0, wrap_text=True
)
font = Font(name="Times New Roman", size=10, bold=False, italic=False)
for m in ["A", "B", "C", "D"]:
    for n in range(1, 20):
        cell = sheet[m + str(n)]
        cell.alignment = alignment
        cell.font = font
        cell.fill = PatternFill("solid", start_color="b5b9b1")
        # bgColor="b5b9b1")


for m in ["A", "B", "C", "D"]:
    for n in range(1, 20):
        cell = sheet[m + str(n)]
        cell.border = Border(
            left=Side(style="thin"),
            bottom=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
        )

for m in ["1", "10", "17"]:
    for i in ["A", "B", "C", "D"]:
        cell = sheet[i + m]
        font = Font(name="Times New Roman", size=12, bold=True, italic=False)
        cell.font = font


height = 20
width = 8

# print("row:", sheet.max_row, "column:", sheet.max_column)
for i in range(1, sheet.max_row + 1):
    sheet.row_dimensions[i].height = height
for i in range(1, sheet.max_column + 1):
    sheet.column_dimensions[get_column_letter(i)].width = width

# sheet.column_dimensions[get_column_letter(1)].width = 10
sheet.column_dimensions[get_column_letter(2)].width = 40
sheet.column_dimensions[get_column_letter(3)].width = 20
sheet.column_dimensions[get_column_letter(4)].width = 20


workbook.save("perf.xlsx")
