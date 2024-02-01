import pandas as pd


def chairCount(sbNameInput, volunteerExcelFile, sbRebateEligibilityFile):
    sbVolunteersDf = pd.read_excel(volunteerExcelFile)
    sbRebateEligibilityDf = pd.read_excel(sbRebateEligibilityFile)

    result = 0
    for index, row in sbVolunteersDf.iterrows():
        if (sbNameInput == row["OU Name"]):
            if (sbVolunteersDf.at[index, "Position"] == "Chair"):
                result +=1
    return result

def counselorCount(sbNameInput, volunteerExcelFile, sbRebateEligibilityFile):
    sbVolunteersDf = pd.read_excel(volunteerExcelFile)
    sbRebateEligibilityDf = pd.read_excel(sbRebateEligibilityFile)

    result = 0
    for index, row in sbVolunteersDf.iterrows():
        if (sbNameInput == row["OU Name"]):
            if (sbVolunteersDf.at[index, "Position"] == "Counselor"):
                result +=1
    return result




sbRebateEligibilityDf = pd.read_excel("SBs rebate eligibility.xlsx")
sbVolunteersDf = pd.read_excel("Volunteer List by OU.xlsx")

sbRebateEligibilityDf["Reported Chair"] = ""
sbRebateEligibilityDf["Chair End of Term Date"] = ""
sbRebateEligibilityDf["Chair Email"] = ""

sbRebateEligibilityDf["Reported Counselor"] = ""
sbRebateEligibilityDf["Counselor Name"] = ""
sbRebateEligibilityDf["Counselor Email"] = ""





for indexElig, rowElig in sbRebateEligibilityDf.iterrows():
    sb = rowElig['Student Branch']
    sb = sb.split(" - ")[1]
    for index, row in sbVolunteersDf.iterrows():
        if(sb == row['OU Name']):
            if (counselorCount(row['OU Name'], "Volunteer List by OU.xlsx", "SBs rebate eligibility.xlsx") ==0):
                sbRebateEligibilityDf.at[indexElig, "Reported Counselor"] = "No"
                break
            else:
                if (sbVolunteersDf.at[index, "Position"] == "Counselor"):
                    for i in range (counselorCount(row['OU Name'], "Volunteer List by OU.xlsx", "SBs rebate eligibility.xlsx")):
                        sbRebateEligibilityDf.at[indexElig, "Reported Counselor"] = "Yes"
                        sbRebateEligibilityDf.at[indexElig, "Counselor Name"] += sbVolunteersDf.at[index+i, "Last Name   "] +" "+sbVolunteersDf.at[index+i, "First Name   "]+", "
                        sbRebateEligibilityDf.at[indexElig, "Counselor Email"] += sbVolunteersDf.at[index+i, "Email Address  "]+", "

                    break

for indexElig, rowElig in sbRebateEligibilityDf.iterrows():
    sb = rowElig['Student Branch']
    sb = sb.split(" - ")[1]
    for index, row in sbVolunteersDf.iterrows():
        if(sb == row['OU Name']):
            if (chairCount(row['OU Name'], "Volunteer List by OU.xlsx", "SBs rebate eligibility.xlsx") ==0):
                sbRebateEligibilityDf.at[indexElig, "Reported Chair"] = "No"
                break
            else:
                if (sbVolunteersDf.at[index, "Position"] == "Chair"):
                    for i in range (chairCount(row['OU Name'], "Volunteer List by OU.xlsx", "SBs rebate eligibility.xlsx")):
                        sbRebateEligibilityDf.at[indexElig, "Reported Chair"] = "Yes"
                        sbRebateEligibilityDf.at[indexElig, "Chair Email"] += sbVolunteersDf.at[index+i, "Email Address  "]+", "
                        sbRebateEligibilityDf.at[indexElig, "Chair End of Term Date"] += sbVolunteersDf.at[index+i, "Position End "]+", "

                        break

sbRebateEligibilityDf.to_excel("your_excel_file_with_Reported_column.xlsx", index=False)




