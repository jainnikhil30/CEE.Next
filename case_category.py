import xlrd
from xlutils.copy import copy
location = "/home/niks/Downloads/sheet.xlsx"
rb = xlrd.open_workbook(location)
wb = copy(rb)
read_sheet = rb.sheet_by_index(1)
write_sheet = wb.get_sheet(1)
#Cateogries and keywords for each category
category_meta_data = { #"Upgrade": [ "upgrade", "install" ],
                       "Upgrade": ["upgrade"],
                       "Performance": ["performance", "memory"],
                       "Candlepin": ["subscription", "manifest", "virt-who", "register", "bootstrap", "license"],
                       "Config Management": ["puppet", "ansible", "playbook"],
                       "Pulp": ["patch", "repos", "dependency", "pulp", "yum", "content view", "sync", "promote", "capsule", "download packages", "publish"],
                       "CLI": ["hammer", "api"],
                       "Backup Restore": ["migration", "migrate", "backup", "restore"],
                       "Provisioning": ["pxe", "cloud-init"],
                       "Formeman task": ["foreman-task"],
                       "katello-agent": ["katello-agent"],
                       "Openscap": ["openscap"],
                       "RHUI": ["rhui"],
                       "AWS": ["aws"],
                       "Remote Execution": ["remote execution"],
                       "Insights": ["insights"],
                       "External Authentication": ["ldap", "ipa", "authentication", "external authentication"]
            }

category_meta_data_ignore_words = {"Upgrade": ["after"],
                                   "Performance":[],
                                   "Config Management":[],
                                   "Candlepin":[],
                                   "Pulp":[],
                                   "CLI":[],
                                   "Backup Restore":[],
                                   "Provisioning":[],
                                   "Formeman task":[],
                                   "katello-agent":[],
                                   "Openscap": [],
                                   "RHUI":[],
                                   "AWS":[],
                                   "Remote Execution":[],
                                   "Insights": [],
                                   "External Authentication": [],
                                   "Other":[]
            }

#Final dictionory to hold category and cases for eachcategory
final_category_wise_case_list = {"Upgrade": [], "Performance":[], "Config Management":[], "Candlepin":[], "Pulp":[], "CLI":[], "Backup Restore":[], "Provisioning":[],
              "Formeman task":[], "katello-agent":[], "Openscap": [], "RHUI":[], "AWS":[], "Remote Execution":[], "Insights": [], "External Authentication": [], "Other":[]
            }

for i in range(read_sheet.nrows):
    #ignore the first row of column headings
    case_added = False
    if i == 0:
        continue
    case_number = int(read_sheet.cell_value(i,0))
    case_description = read_sheet.cell_value(i,3)
    problem_statement = read_sheet.cell_value(i,2)
    #iterate over each category one by one
    for key in category_meta_data.keys():
        #iterate over keyword for each category, till there is match.
        for keyword in category_meta_data[key]:
            #look for the keyword in the case description
            dont_add = False
            if case_description.lower().find(keyword) != -1 and case_number not in final_category_wise_case_list[key]:
                for ignore_keyword in category_meta_data_ignore_words[key]:
                    if case_description.lower().find(ignore_keyword) != -1:
                        dont_add = True
                if not dont_add:
                    final_category_wise_case_list[key].append(case_number)
                    case_added = True
                    write_sheet.write(i,10,key)
                    break
            # look for the keyword in the problem state if not found in case description
            elif problem_statement.lower().find(keyword) != -1 and case_number not in final_category_wise_case_list[key]:
                for ignore_keyword in category_meta_data_ignore_words[key]:
                    if case_description.lower().find(ignore_keyword) != -1:
                        dont_add = True
                if not dont_add:
                    final_category_wise_case_list[key].append(case_number)
                    case_added = True
                    write_sheet.write(i, 10, key)
                    break
        else:
            continue
        break
             # if no match put it in Other bucket
    if not case_added and case_number not in final_category_wise_case_list["Other"]:
        final_category_wise_case_list["Other"].append(case_number)
        write_sheet.write(i, 10, 'Other')

wb.save('/home/niks/Downloads/sheet.xlsx')
total_cases = 0
for key, value in final_category_wise_case_list.items():
    print(key, ' : ', len(value), 'Case(s). List of cases is: ', value)
    total_cases = total_cases + len(value)
print("Total cases processed: %s" % total_cases)