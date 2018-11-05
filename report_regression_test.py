import io
import sys


# Change the default encoding of standard output
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf8')

ccl_report_file = "C:\\Users\\bichao\Desktop\\Daily_report\\Ariel-CCL-summary.txt"
status_tracker_file = "C:\\Users\\bichao\\Desktop\\Ariel_DCN2KA_IP_Diagnostics_Status_Tracker.xlsx"
report_result_file = "C:\\Users\\bichao\Desktop\\Daily_report\\report_result.txt"

source_summary = open(ccl_report_file, "r+", encoding='utf8')
report_result = open(report_result_file, "w+", encoding='utf8')
skip_case_list = []
fail_case_list = []

temp_str = source_summary.readline()
while temp_str:
    if temp_str.find("Skip") != -1:
        skip_case_id = temp_str.split('.')[0]
        if skip_case_id not in skip_case_list:
            skip_case_list.append(skip_case_id)
    if temp_str.find("Fail") != -1:
        fail_case_id = temp_str.split('.')[0]
        if fail_case_id not in fail_case_list:
            fail_case_list.append(fail_case_id)
    temp_str = source_summary.readline()

print(skip_case_list)
print(fail_case_list)
source_summary.close()
report_result.close()
