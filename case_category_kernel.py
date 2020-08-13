# This script is used for categorising cases for SysMgmt
# Author - Nikhil Jain (nikjain@redhat.com)
print("This script use three python libraries: xlrd, xlwt and xluitls. \nPlease make sure these libraries are installed using pip3 before proceeding with the script.\n"
      "You can install them using: pip3 install xlrd xlwt xlutils")
import xlrd
import xlwt
from xlutils.copy import copy
from copy import deepcopy
location = input("Please enter the absolute path (/my/complete/path/data.xlsx) of the workbook/excel with case data: ")
final_location = input("Please enter the absolute path (/my/complete/path) at which the final data with case category needs to be saved: ")
#location = "/home/niks/Downloads/sheet.xlsx"
sheet = input("Please enter the sheet number which has the data in workbook. Please note numbering starts from 0 : ")
sheet_index = int(sheet)
rb = xlrd.open_workbook(location)
wb = copy(rb)
read_sheet = rb.sheet_by_index(sheet_index)
write_sheet = wb.get_sheet(sheet_index)

category_meta_data = { #"Upgrade": [ "upgrade", "install" ],
    "RCA/Hang": {"not responding": [], "hung": [], "unresponsive": [], "freeze": [], "hangs": [], "hard reset": [], "frozen": [], "unreachable": [], "locked": [], "inaccessible": [], "cant access": [], "not accessible" :[], "hangup": [],
                 "hungup": [], "vmem": [], "not able to access": [], "hanging": [], "unexpected": [], "reachable": [], "rcu stall": [], "filesystem frozen": [], "snapshot": [], "halted": [], "sluggish": [], "halt": [], "unavailable": [],
                 "not ready": [], "unable to connect": []},
    "RCA/Panic": {"BUG": [], "soft lockup": [], "panic": [], "crash": [], "RIP": [], "cpu stuck": [], "hard lockup": [], "crashed": [], "core dump": [], "vmcore": [], "crash dump": [], "kernel crash dump": [], "killer": [],
                  "invoked": [], "vmcore-dmesg.txt": [], "crashes": [], "oops": [], "kworker": [], "backport": []},
    "RCA/Reboot": {"bumped": [], "reset": [], "abrupt reboot": [], "auto reboot": [], "reboot": [], "rebooted": [], "rebooting":[], "restarted": [], "power off": [], "power-off": [], "evict": [], "fenced": [], "fencing": [], "cluster": [],
                   "oracle rac": [], "vertias cluster": [], "VCS": [], "uptime" :[]},
    "RCA/Hardware": {"mce": [], "Hardware": [], "ACPI": [], "ghes": [], "pci": [], "edac": [], "firmware": [], "driver": [], "bank": [], "USB": [], "DIMM": [], "Machine check events": [], "sbridge": [], "SMP": [], "processor": [],
                     "cpu": [], "cpus": [], "infiniband": [], "microcode": [], "microcode_ctl": []},
    "RCA/Third-Party": {"fileaccess_mod": [], "mfe_fileaccess": [], "hip_access_authorize": [], "hipsec": [], "custom_file_open": [], "lshook": [], "linuxshield": [], "mcafee": [], "scdrv": [],
                        "schook": [], "splxmod": [], "gsch": [], "redirfs": [], "dsa_filter": [], "vxfs": [], "vxnetd": [], "VRTS": [], "veritas": [], "symev": [], "sysmv": [], "vxio": [],
                        "vxdump": [], "krg_": [], "seos": [], "sisips": [], "twnotify": [], "vsock": [], "mmfs": [], "lin_tape": [], "mvfs" :[], "nvidia": [], "sisfim": [], "lustre": [],
                        "ktap_": [], "cbsensor": [], "cyprotectdrv": [], "onload": [], "sfc": [], "hpdsa": [], "hpvsa": [], "oracleacfs": [], "oracleasm": [], "emc": [], "secfs2": [], "dattobd": [], "falcon": [], "(U)": [], "(P)": []},
    "Tools": {"PCP": [], "collectl": [], "tuned": [], "cgroups": [], "kdump": [], "perf": [], "capture": [], "kexec": [], "kpatch": [], "tuna": [], "numad": [], "ftrace": [], "strace": [], "trace-cmd": [], "kmod": [], "kexec-tools": []},
    "Performance": {"memory": [], "cpu": [], "load average": [], "slowness": [], "perf.tar": [], "unaccount": [], "process": [], "swap": [], "oracle rac": [], "nmon": [], "tuning": [], "high load": [],
                    "slab": [], "huge pages": [], "spikes": [], "spiked": [], "kswapd0": [], "usage": [], "elevated": [], "resources": [], "slow": [], "top": [], "high": [], "performance": [], "load": [],
                    "free": [], "consume": [], "consumption": [], "shoots up": [], "cpufreq": [], "cache": [], "unaccounted": [], "shared": [], "fragmentation": [], "THP": [], "swapping": [], "leak": [],
                    "leakage": [], "defunct": [], "processes": [], "zombie": [], "nice value": [], "iowait": [], "monitoring": []},
    "CantBoot": {"after patching": [], "recently patched": [], "unable to boot": [], "not booting": []},
    "Others": {"other": []}
}

category_meta_data_ignore_words = { #"Upgrade": [ "upgrade", "install" ],
    "RCA/Hang": [],
    "RCA/Panic": [],
    "RCA/Reboot": [],
    "RCA/Hardware": [],
    "RCA/Third-Party": [],
    "Tools": [],
    "Performance": [],
    "CantBoot": [],
    "Others": []
}

#Final dictionory to hold category and cases for eachcategory
# final_category_wise_case_list = {"Upgrade": [], "Performance":[], "Config Management":[], "Candlepin":[], "Pulp":[], "CLI":[], "Backup Restore":[], "Provisioning":[],
#                "Formeman task":[], "katello-agent":[], "Openscap": [], "RHUI":[], "AWS":[], "Remote Execution":[], "Insights": [], "External Authentication": [], "Other":[]
#              }
processed_cases= []
final_category_wise_case_list = deepcopy(category_meta_data)
final_category_wise_case_list_comment_count = deepcopy(category_meta_data)
account_wise_case_details = {}
# start with problem statement first
for i in range(read_sheet.nrows):
# ignore the first row of column headings
    if i == 0:
        continue
    case_number = int(read_sheet.cell_value(i, 1))
    problem_statement = read_sheet.cell_value(i, 4)
    case_comment = int(read_sheet.cell_value(i,6))
    account_number = int(read_sheet.cell_value(i, 7))
    account_name = read_sheet.cell_value(i,8)
    if account_number not in account_wise_case_details.keys():
        #account_wise_case_details[account_number] = deepcopy(category_meta_data)
        account_wise_case_details[account_number] = {}
        account_wise_case_details[account_number][account_name] = deepcopy(category_meta_data)
    # iterate over each category one by one
    for key in category_meta_data.keys():
        # iterate over keyword for each category, till there is match.
        for keyword in category_meta_data[key]:
            if len(final_category_wise_case_list[key][keyword]) ==0:
                final_category_wise_case_list_comment_count[key][keyword] = 0
            dont_add = False
            # look for the keyword in the problem statement
            if keyword.find('/') != -1:
                key1, key2 = keyword.split('/')
                if problem_statement.lower().find(key1) != -1 and problem_statement.lower().find(key2) != -1 and\
                        case_number not in final_category_wise_case_list[key]:
                    for ignore_keyword in category_meta_data_ignore_words[key]:
                        if problem_statement.lower().find(ignore_keyword) != -1:
                            dont_add = True
                    if not dont_add:
                        final_category_wise_case_list[key][keyword].append(case_number)
                        account_wise_case_details[account_number][account_name][key][keyword].append(case_number)
                        processed_cases.append(case_number)
                        final_category_wise_case_list_comment_count[key][keyword] += case_comment
                        write_sheet.write(i, 10, key)
                        break
            elif problem_statement.lower().find(keyword) != -1 and case_number not in final_category_wise_case_list[key]:
                for ignore_keyword in category_meta_data_ignore_words[key]:
                    if problem_statement.lower().find(ignore_keyword) != -1:
                        dont_add = True
                if not dont_add:
                    final_category_wise_case_list[key][keyword].append(case_number)
                    processed_cases.append(case_number)
                    account_wise_case_details[account_number][account_name][key][keyword].append(case_number)
                    final_category_wise_case_list_comment_count[key][keyword] += case_comment
                    write_sheet.write(i, 10, key)
                    break
        else:
            continue
        break

# do the same thing for case description but only for cases which did't get any categoryhmmm
for i in range(read_sheet.nrows):
    if i == 0:
        continue
    case_number = int(read_sheet.cell_value(i,1))
    case_description = str(read_sheet.cell_value(i,5))
    case_comment = int(read_sheet.cell_value(i,6))
    account_number = int(read_sheet.cell_value(i, 7))
    account_name = read_sheet.cell_value(i,8)
    if account_number not in account_wise_case_details.keys():
        #account_wise_case_details[account_number] = deepcopy(category_meta_data)
        account_wise_case_details[account_number] = {}
        account_wise_case_details[account_number][account_name] = deepcopy(category_meta_data)
    #iterate over each category one by one
    for key in category_meta_data.keys():
        #iterate over keyword for each category, till there is match.
        for keyword in category_meta_data[key]:
            if len(final_category_wise_case_list[key][keyword]) ==0:
                final_category_wise_case_list_comment_count[key][keyword] = 0
            #look for the keyword in the case description
            dont_add = False
            if case_number not in processed_cases:
                if keyword.find('/') != -1:
                    key1, key2 = keyword.split('/')
                    if case_description.lower().find(key1) != -1 and case_description.lower().find(key2) != -1 and \
                            case_number not in final_category_wise_case_list[key]:
                        for ignore_keyword in category_meta_data_ignore_words[key]:
                            if case_description.lower().find(ignore_keyword) != -1:
                                dont_add = True
                        if not dont_add:
                            final_category_wise_case_list[key][keyword].append(case_number)
                            processed_cases.append(case_number)
                            account_wise_case_details[account_number][account_name][key][keyword].append(case_number)
                            final_category_wise_case_list_comment_count[key][keyword] += case_comment
                            write_sheet.write(i, 10, key)
                            break
                elif case_description.lower().find(keyword) != -1 and case_number not in final_category_wise_case_list[key]:
                    for ignore_keyword in category_meta_data_ignore_words[key]:
                        if case_description.lower().find(ignore_keyword) != -1:
                            dont_add = True
                    if not dont_add:
                        final_category_wise_case_list[key][keyword].append(case_number)
                        processed_cases.append(case_number)
                        account_wise_case_details[account_number][account_name][key][keyword].append(case_number)
                        final_category_wise_case_list_comment_count[key][keyword] += case_comment
                        write_sheet.write(i,10,key)
                        break
        else:
            continue
        break

             # if no match put it in Other bucket
    if case_number not in processed_cases and case_number not in final_category_wise_case_list["Others"]:
        final_category_wise_case_list["Others"]["other"].append(case_number)
        account_wise_case_details[account_number][account_name]["Others"]["other"].append(case_number)
        final_category_wise_case_list_comment_count["Others"]["other"] += case_comment
        write_sheet.write(i, 10, 'Others')

wb.save(final_location+'/processed_data.xlsx')

#Create the category wise case report
new_wb= xlwt.Workbook()
new_ws = new_wb.add_sheet('Category wise cases')
new_ws.write(0,0,'Main Category')
new_ws.write(0,1,'Sub Category')
new_ws.write(0,2,'Cases')
new_ws.write(0,3,'List of cases')
new_ws.write(0,4,'Total Case comments')
total_cases = 0
i=1
for key in final_category_wise_case_list:
    print(key, ':')
    new_ws.write(i,0, key)
    i = i+1
    for keyword in final_category_wise_case_list[key]:
      print('\r', keyword, ':', len(final_category_wise_case_list[key][keyword]), 'Case(s). List of cases: ', final_category_wise_case_list[key][keyword])
      new_ws.write(i, 1, keyword)
      new_ws.write(i, 2, len(final_category_wise_case_list[key][keyword]))
      new_ws.write(i, 3, str(final_category_wise_case_list[key][keyword]))
      new_ws.write(i, 4, str(final_category_wise_case_list_comment_count[key][keyword]))
      i = i+1
      total_cases = total_cases + len(final_category_wise_case_list[key][keyword])
    print('----------------------------------------------------------------------------------------------')
print("Total cases processed: %s" % total_cases)
new_wb.save(final_location+'/Category_wise_cases.xlsx')

#Create the account wise case category report
new_account_wb = xlwt.Workbook()
new_account_ws = new_account_wb.add_sheet('Account wise cases')
new_account_ws.write(0,0,'Account Number')
new_account_ws.write(0,1,'Account Name')
new_account_ws.write(0,2,'Main Category')
new_account_ws.write(0,3,'Sub Category')
new_account_ws.write(0,4,'List of cases')
new_account_ws.write(0,5,'Number of cases')
i=1
# for key in account_wise_case_details:
#     new_account_ws.write(i,0, key)
#     i = i+1
#     for main_key in account_wise_case_details[key]:
#       new_account_ws.write(i, 1, main_key)
#       i = i+1
#       for sub_key in account_wise_case_details[key][main_key]:
#           if len(account_wise_case_details[key][main_key][sub_key]) > 0:
#             new_account_ws.write(i, 2, sub_key)
#             new_account_ws.write(i, 3, str(account_wise_case_details[key][main_key][sub_key]))
#             i = i+1
for acc_num in account_wise_case_details:
    account_cases = 0
    new_account_ws.write(i, 0, acc_num)
    for acc_name in account_wise_case_details[acc_num]:
        new_account_ws.write(i, 1, acc_name)
        i=i+1
        for main_cat in account_wise_case_details[acc_num][acc_name]:
            new_account_ws.write(i, 2, main_cat)
            i=i+1
            for sub_key in account_wise_case_details[acc_num][acc_name][main_cat]:
                if len(account_wise_case_details[acc_num][acc_name][main_cat][sub_key]) > 0:
                    new_account_ws.write(i, 3, sub_key)
                    new_account_ws.write(i, 4, str(account_wise_case_details[acc_num][acc_name][main_cat][sub_key]))
                    new_account_ws.write(i, 5, len(account_wise_case_details[acc_num][acc_name][main_cat][sub_key]))
                    account_cases = account_cases + len(account_wise_case_details[acc_num][acc_name][main_cat][sub_key])
                    i = i+1
    new_account_ws.write(i, 4, 'Total Account Cases')
    new_account_ws.write(i, 5, account_cases)
    i = i+1
new_account_ws.write(i,4, 'Total Cases processed')
new_account_ws.write(i,5, total_cases)
new_account_wb.save(final_location+'/Account_wise_cases.xlsx')
