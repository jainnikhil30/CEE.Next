import xlrd
from xlutils.copy import copy
from copy import deepcopy
location = "/home/niks/Downloads/sheet.xlsx"
rb = xlrd.open_workbook(location)
wb = copy(rb)
read_sheet = rb.sheet_by_index(1)
write_sheet = wb.get_sheet(1)
#Cateogries and keywords for each category
category_meta_data = { #"Upgrade": [ "upgrade", "install" ],
    "Upgrade": {"upgrade": []},
    "Manifest": {"manifest": [], "simple content access":[]},
    "Content Management": {"content view": [], "promote": [], "publish": [], "pulp": [], "sync/capsule": [], "repos/capsule": []},
    "Subscription & Registration": {"virt-who": [], "bootstrap" :[], "license": [], "register": [], "subscription": [], "repos/enable": [], "capsule/enable": [] },
    "System Patching": {"patch": [], "katello-agent": [], "download packages": [], "dependencies": [], "yum": [], "repos": []},
    "Insights": {"inventory": [], "insights": []},
    "Config Management": {"puppet": [], "ansible": [], "playbook": [], "module": []},
    "Performance": {"performance": [], "memory": [], "cpu": [], "swap": [], "mongodb": []},
    "Provisioning": { "pxe": [], "cloud-init": [], "boot disk": [], "provisioning": [], "kickstart": [], "host image": []},
    "Remote Execution": {"remote execution": [], "rex": []},
    "Openscap": {"openscap": []},
    "RHUI & AWS": {"rhui": [], "aws": [], "rhua": []},
    "External Authentication": {"ldap": [], "active directory": [], "ipa": [], "authentication": [], "external authentication": [], "kerberos": [], "sssd": [], "keytab": []},
    "Custom Certificate": {"custom cert": [], "ssl cert": []},
    "CLI": {"hammer": [], "api": []},
    "Backup Restore": {"migration": [], "migrate": [], "backup": [], "restore":[]},
    "Others": {"other": []}
    }

category_meta_data_ignore_words = {
    "Upgrade": ["after", "upgraded", "browser", "package"],
    "Manifest": [],
    "Content Management": [],
    "Subscription & Registration": [],
    "System Patching": [],
    "Insights": [],
    "Config Management":["repo", "subscription"],
    "Performance":[],
    "Provisioning":[],
    "Remote Execution":[],
    "Openscap": [],
    "RHUI & AWS":[],
    "External Authentication": [],
    "Custom Certificate": [],
    "CLI":[],
    "Backup Restore":[],
    "Others":[]
}

#Final dictionory to hold category and cases for eachcategory
# final_category_wise_case_list = {"Upgrade": [], "Performance":[], "Config Management":[], "Candlepin":[], "Pulp":[], "CLI":[], "Backup Restore":[], "Provisioning":[],
#                "Formeman task":[], "katello-agent":[], "Openscap": [], "RHUI":[], "AWS":[], "Remote Execution":[], "Insights": [], "External Authentication": [], "Other":[]
#              }
processed_cases= []
final_category_wise_case_list = deepcopy(category_meta_data)
for i in range(read_sheet.nrows):
# ignore the first row of column headings
    if i == 0:
        continue
    case_number = int(read_sheet.cell_value(i, 0))
    problem_statement = read_sheet.cell_value(i, 2)
    # iterate over each category one by one
    for key in category_meta_data.keys():
        # iterate over keyword for each category, till there is match.
        for keyword in category_meta_data[key]:
            # look for the keyword in the case description
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
                        processed_cases.append(case_number)
                        write_sheet.write(i, 10, key)
                        break
            elif problem_statement.lower().find(keyword) != -1 and case_number not in final_category_wise_case_list[key]:
                for ignore_keyword in category_meta_data_ignore_words[key]:
                    if problem_statement.lower().find(ignore_keyword) != -1:
                        dont_add = True
                if not dont_add:
                    final_category_wise_case_list[key][keyword].append(case_number)
                    processed_cases.append(case_number)
                    write_sheet.write(i, 10, key)
                    break
        else:
            continue
        break

for i in range(read_sheet.nrows):
    #ignore the first row of column headings
    if i == 0:
        continue
    case_number = int(read_sheet.cell_value(i,0))
    case_description = read_sheet.cell_value(i,3)
    #iterate over each category one by one
    for key in category_meta_data.keys():
        #iterate over keyword for each category, till there is match.
        for keyword in category_meta_data[key]:
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
                            write_sheet.write(i, 10, key)
                            break
                elif case_description.lower().find(keyword) != -1 and case_number not in final_category_wise_case_list[key]:
                    for ignore_keyword in category_meta_data_ignore_words[key]:
                        if case_description.lower().find(ignore_keyword) != -1:
                            dont_add = True
                    if not dont_add:
                        final_category_wise_case_list[key][keyword].append(case_number)
                        processed_cases.append(case_number)
                        write_sheet.write(i,10,key)
                        break
        else:
            continue
        break

             # if no match put it in Other bucket
    if case_number not in processed_cases and case_number not in final_category_wise_case_list["Others"]:
        final_category_wise_case_list["Others"]["other"].append(case_number)
        write_sheet.write(i, 10, 'Others')

wb.save('/home/niks/Downloads/sheet.xlsx')
total_cases = 0
for key in final_category_wise_case_list:
    print(key, ':')
    for keyword in final_category_wise_case_list[key]:
      print('\r', keyword, ':', len(final_category_wise_case_list[key][keyword]), 'Case(s). List of cases: ', final_category_wise_case_list[key][keyword])
        #print (keyword, '  :  ', len(value), 'Cases(s). List of cases is: ', value)
# for key, value in final_category_wise_case_list.items():
#     print(key, ' : ', len(value), 'Case(s). List of cases is: ', value)
      total_cases = total_cases + len(final_category_wise_case_list[key][keyword])
    print('----------------------------------------------------------------------------------------------')
print("Total cases processed: %s" % total_cases)