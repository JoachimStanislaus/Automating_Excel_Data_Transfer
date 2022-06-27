from openpyxl import Workbook, load_workbook


hidden_list=[]

wb = load_workbook(filename= "March 2022 (CDH, DAC, SW and ARP) Utilization Statistics.xlsx")
wb.active = wb['DAC']
ws = wb.active

def get_field_label():
    remove_hidden_headers()
    sample_list = []
    reject_field_list =['Current Enrolment','Reported Utilization','Waitlist Information (Referral Status)','Not Screened (a)','Screened (b)','On Hold (c)','Waiting time for enrolment','Not Screen', 'Screened - Suitable for Admission']
    FieldLabel_list_DAC=['S/N','Service Type','Name of Service Provider (SP)','Region','Primary Service Group','Approved Capacity','SMM Capacity','Reported Vacancy','Full-Time','Part-Time','Total','% of Utilization','No. of PwDs admitted to the service in current month','No. of PwDs discharge from the service in current month','SGE New Referral','Accepted By SP\n(Accepted clients enrolling on a later date)','Pending\nAssessment', 'Pending\nAssessment', 'Pending Final\nAssessment Outcome','Pending Trial Enrollment', 'On Hold by\nPwD/Caregiver', 'On Hold due to \nCentre Constraint', 'On Hold by\nPwD/Caregiver','Total\n(a + b + c)','Estimated Waiting Time\n(SP Estimated Waiting Time for Enrollment)','Average Waiting Time\n(Average Waiting Time based on cases enrolled the past year)',' Median \n(by Month)','Longest Referral on Waitlist\n(Oldest Case on waitlist that does not have finalized outcome)']
    data_field_list=[]
    delete_dict = {}
    for x in range(1,50): # Cycle through all the rows
        row_data = read_row(ws,x) # Get row_data row by row
        #print("Row",x,":",row_data)
        if x<=6:
            for y in range(0,len(row_data)):
                if y in hidden_list:
                    continue
                else:
                    if row_data[y] != None:
                        if  (isinstance(row_data[y], str)) == True :
                            if row_data[y] in reject_field_list:
                                continue
                            else:
                                sample_list.append(row_data[y])
    print(data_field_list)

def read_row(worksheet, row): # Read row cell by cell and return cell values in a list
    return [cell.value for cell in worksheet[row]]

def remove_hidden_headers():
    global hidden_list
    for colLetter,colDimension in ws.column_dimensions.items():
        if colDimension.hidden == True:
            #print(colDimension.min)
            #print('Last Group',colDimension.max)
            for hidden in range(colDimension.min,colDimension.max+1):
                hidden_list.append(hidden-1)


get_field_label()