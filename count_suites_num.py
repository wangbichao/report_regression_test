import io
import sys
import numpy


# Change the default encoding of standard output
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf8')

ccl_report_file = "Z:\\bin_build\\ariel_build\\ariel_a0-x86_64-linux-dbg\\test_dcn_suite.txt"

source_summary = open(ccl_report_file, "r+", encoding='utf8')
suites_list = []


temp_str = source_summary.readline()
while temp_str:
    suites_id = temp_str.split('.')[0]
    suites_list.append(suites_id)
    temp_str = source_summary.readline()

#print(suites_list)
source_summary.close()

print(dict(zip(*numpy.unique(suites_list, return_counts=True))))
