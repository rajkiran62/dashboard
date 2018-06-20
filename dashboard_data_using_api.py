from os import system
import json
import xlsxwriter
import pandas as pd

workbook = xlsxwriter.Workbook('dashboard.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write('A' + '1', 'Project Name')
#worksheet.write('B' + '1', 'Project ID')
worksheet.write('C' + '1', 'Version Name')
#worksheet.write('D' + '1', 'Version ID')
worksheet.write('E' + '1', 'High')
worksheet.write('F' + '1', 'Medium')
worksheet.write('G' + '1', 'Low')

#system("curl -X POST --data 'j_username=testuser1^&j_password=blackduck' -i https://blackduck.philips.com/j_spring_security_check --insecure -c cookie.txt")
#system('curl -b cookie.txt -X GET --header "Accept: application/json" "https://blackduck.philips.com:443/api/projects?limit=999" --insecure > projects.json')
with open('projects.json') as all_projects_name_info:
    data = json.load(all_projects_name_info)
projects_count = data["totalCount"]
print(projects_count)
error_report = open('error_report.txt','w')
projectsinfo = open('projects.txt','w')
for i in range(projects_count):
    projectsinfo.write(data["items"][i]["name"]+", "+ data["items"][i]["_meta"]["href"]+",\n")
all_projects_name_info.close()
projectsinfo.close()
all_projects_name_info=open('projects.txt','r')
rows = 2
for each_project in all_projects_name_info:
    project_name = each_project[:each_project.index(',')]
    each_project_url = each_project[each_project.index(',')+1:-2]
    version_url = each_project_url.strip(" ")+"/versions?limit=999"
    version_command = 'curl -b cookie.txt -X GET --header "Accept: application/json" "'+version_url+'" --insecure > versions.json'
    #system(version_command)
    #print(version_command)
    try:
        with open('versions_onebackend.json', 'r') as all_version_info:
            version_data = json.load(all_version_info)
            num_of_versions = version_data["totalCount"]
            #print(num_of_versions)
            for i in range(num_of_versions):
                #logs -> print i =>
                version_name = version_data["items"][i]["versionName"]
                risk_profile_url = version_data["items"][i]["_meta"]["links"][2]["href"]
                risk_profile_command = 'curl -b cookie.txt -X GET --header "Accept: application/json" "'+risk_profile_url+'" --insecure > risk_profile.json'
                #print(latest_version_name + ": " + risk_profile_url)
                #system(risk_profile_command)
                with open('risk_profile.json','r') as risk_profile_info:
                    risk_data = json.load(risk_profile_info)
                    high = risk_data["categories"]["VULNERABILITY"]["HIGH"]
                    medium = risk_data["categories"]["VULNERABILITY"]["MEDIUM"]
                    low = risk_data["categories"]["VULNERABILITY"]["LOW"]
                    security_vulnerable_components =[high, medium, low]
                    #print(rows)
                    #print(project_name+"\n"+version_name+"\n"+security_vulnerable_components)
                    worksheet.write('A' + str(rows), project_name)
                    #worksheet.write('B' + str(rows), "NA")
                    worksheet.write('C' + str(rows), version_name)
                    #worksheet.write('D' + str(rows), 'NA')
                    worksheet.write('E' + str(rows), high)
                    worksheet.write('F' + str(rows), medium)
                    worksheet.write('G' + str(rows), low)
                rows += 1
            all_version_info.close()
    except IndexError as error:
        #error_report.write(data)
        all_version_info.close()
        workbook.close()
workbook.close()