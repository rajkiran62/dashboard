'''
Gives information about the number of vulnerable components.
Note: Issue count may be different
'''

from os import system
import json
import xlsxwriter
#import pandas as pd

workbook = xlsxwriter.Workbook('dashboard_latest_version.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write('A' + '1', 'Project Name')
worksheet.write('B' + '1', 'Version Name')
worksheet.write('C' + '1', 'High Vuln Components')
worksheet.write('D' + '1', 'Medium Vuln Components')
worksheet.write('E' + '1', 'Low Vuln Components')
worksheet.write('F' + '1', 'High Vuln')
worksheet.write('G' + '1', 'Medium Vuln')
worksheet.write('H' + '1', 'Low Vuln')

#system("curl -X POST --data 'j_username=testuser1&j_password=blackduck' -i https://blackduck.philips.com/j_spring_security_check --insecure -c cookie.txt")
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
    print(version_command)
    try:
        with open('versions.json', 'r') as all_version_info:
            version_data = json.load(all_version_info)
            num_of_versions = version_data["totalCount"]
            check_duplicate_CVE = []
            issues_count = [0, 0, 0]
            #print(num_of_versions)
            #for i in range(num_of_versions):
                #logs -> print i =>
            version_name = version_data["items"][num_of_versions-1]["versionName"]
            risk_profile_url = version_data["items"][num_of_versions-1]["_meta"]["links"][2]["href"]
            vulnerable_bom_components_url = version_data["items"][num_of_versions-1]["_meta"]["links"][4]["href"]
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
                worksheet.write('B' + str(rows), version_name)
                worksheet.write('C' + str(rows), high)
                worksheet.write('D' + str(rows), medium)
                worksheet.write('E' + str(rows), low)
            all_version_info.close()
            if high != 0 and medium != 0 and low != 0:
                pass
            else:
                vulnerable_bom_components_command = 'curl -b cookie.txt -X GET --header "Accept: application/json" "' \
                                                    + vulnerable_bom_components_url +\
                                                    '?limit=9999" --insecure > vulnerable_bom_components.json'
                #system(vulnerable_bom_components_command)
                with open('vulnerable_bom_components.json', 'r') as security_vuln_info:
                    vuln_data = json.load(security_vuln_info)
                    num_of_vuln = vuln_data["totalCount"]
                    for each_vuln in range(num_of_vuln):
                        CVE_ID = vuln_data["items"][each_vuln]["vulnerabilityWithRemediation"]["vulnerabilityName"]
                        remediationStatus = vuln_data["items"][each_vuln]["vulnerabilityWithRemediation"]["remediationStatus"]
                        if CVE_ID in check_duplicate_CVE:
                            pass
                        else:
                            check_duplicate_CVE.append(CVE_ID)
                            if remediationStatus == "NEW":
                                severity = vuln_data["items"][each_vuln]["vulnerabilityWithRemediation"]["severity"]
                                if severity == "HIGH":
                                    issues_count[0] += 1
                                if severity == "MEDIUM":
                                    issues_count[1] += 1
                                if severity == "LOW":
                                    issues_count[2] += 1
                            else:
                                print("New remediation status: %s" % remediationStatus)
                security_vuln_info.close()
                worksheet.write('F' + str(rows), issues_count[0])
                worksheet.write('G' + str(rows), issues_count[1])
                worksheet.write('H' + str(rows), issues_count[2])
            rows += 1
    except Exception:
        #error_report.writelines(rows+"\n"+project_name+"\n"+version_name)
        rows += 1
        #error_report.writelines(Exception)
        print("#####################")
        print(Exception)
        #all_version_info.close()
        security_vuln_info.close()
        #workbook.close()
workbook.close()