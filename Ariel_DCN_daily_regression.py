import io
import sys
import openpyxl


# Change the default encoding of standard output
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf8')

ccl_report_file = "C:\\Users\\bichao\Desktop\\Daily_report\\Ariel-CCL-summary.txt"
status_tracker_file = "C:\\Users\\bichao\\Desktop\\Ariel_DCN2KA_IP_Diagnostics_Status_Tracker.xlsx"
report_result_file = "C:\\Users\\bichao\Desktop\\Daily_report\\report_result.txt"

# read the CCL summary file
source_summary = open(ccl_report_file, "r+", encoding='utf8')
report_result = open(report_result_file, "w+", encoding='utf8')
hang_case_list = []
skip_case_list = []
fail_case_list = []

# read the Status Tracker
wb = openpyxl.load_workbook(status_tracker_file)
all_sheets = wb.sheetnames
print("Get the status tracker all sheet ... ")
print(all_sheets)

# create the new report
ower_hang_case = []
ower_skip_case = []
ower_fail_case = []


def find_false_in_sheet(sheet):
    for column in sheet.iter_cols():
        for cell in column:
            if cell.value in hang_case_list:
#                print(cell.value + sheet.cell(row=cell.row, column=3).value)
                ower_hang_case.append(sheet.cell(row=cell.row, column=3).value + ' \t ' + cell.value)
            if cell.value in skip_case_list:
#                print(cell.value + sheet.cell(row=cell.row, column=3).value)
                ower_skip_case.append(sheet.cell(row=cell.row, column=3).value + ' \t ' + cell.value)
            if cell.value in fail_case_list:
#                print(cell.value + sheet.cell(row=cell.row, column=3).value)
                ower_fail_case.append(sheet.cell(row=cell.row, column=3).value + ' \t ' + cell.value)


for temp_str in source_summary:
    if temp_str.find("Hang") != -1:
        hang_case_id = temp_str.split('.')[0]
        if hang_case_id not in hang_case_list:
            hang_case_list.append(hang_case_id)
    if temp_str.find("Skip") != -1:
        skip_case_id = temp_str.split('.')[0]
        if skip_case_id not in skip_case_list:
            skip_case_list.append(skip_case_id)
    if temp_str.find("Fail") != -1:
        fail_case_id = temp_str.split('.')[0]
        if fail_case_id not in fail_case_list:
            fail_case_list.append(fail_case_id)


print("Get the CCL summary all HANG case ... ")
print(len(hang_case_list))
print(hang_case_list)
print("Get the CCL summary all SKIP case ... ")
print(len(skip_case_list))
print(skip_case_list)
print("Get the CCL summary all FAIL case ... ")
print(len(fail_case_list))
print(fail_case_list)

for i in range(len(all_sheets)):
        sheet = wb[all_sheets[i]]
        find_false_in_sheet(sheet)

print("Report all HANG case with owner ... ")
ower_hang_case.sort()
print(len(ower_hang_case))
print(ower_hang_case)
report_result.write("All HANG case ...\n")
for report in ower_hang_case:
    report_result.write(report+"\n")

print("Report all SKIP case with owner ... ")
ower_skip_case.sort()
print(len(ower_skip_case))
print(ower_skip_case)
report_result.write("All SKIP case ...\n")
for report in ower_skip_case:
    report_result.write(report+"\n")

print("Report all FAIL case with owner ... ")
ower_fail_case.sort()
print(len(ower_fail_case))
print(ower_fail_case)
report_result.write("All FAIL case ...\n")
for report in ower_fail_case:
    report_result.write(report+"\n")

