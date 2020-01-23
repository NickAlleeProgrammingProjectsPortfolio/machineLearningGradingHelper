# -*- coding: utf-8 -*-

"""
Created on Sat Jan 11 

take in list of students names in the source code
take in name of the assignment
take in dict of the question numbers and how many points each question is worth


make a folder to store folders
make a folder inside that folder for each student named correctly
create an excel file for each one of them
	the file should be named correctly and have a table of question numbers
	and how much they are worth. a column should be left for the points 
	earned. correct calculations for the total should be made.

@author: nick allee
"""

import os, openpyxl, sys
p = os.getcwd()
students = ["nick","taylor","bob","jenny","phil","albert"]
assignmentName = sys.argv[1]

points = list(sys.argv[2].strip("[").strip("]").split(","))
pointTotal = 0
for l in points:
    pointTotal+= int(l)

extraCreditPoint = pointTotal*.05
points.append(str(extraCreditPoint))
'''todo check to see if the points are taken in'''



#makes folder to store folders for each student
try:
    os.mkdir(assignmentName)
except FileExistsError:
    print("directory " +assignmentName + " already exists.")

#change directory to the folder above
os.chdir(p+"/"+assignmentName)

for student in students:
    #make the students folder
    try:
        os.mkdir(student+assignmentName)
    except FileExistsError:
        print("directory " + student+assignmentName + " already exists.")
    #change into the students directory
    os.chdir(p+"/"+assignmentName+"/"+student+assignmentName)
    #make students excel document
    wb = openpyxl.Workbook()
    sheet = wb["Sheet"]
    #setup titles
    sheet['A1'] = assignmentName
    sheet['B2'] = "points earned"
    sheet['C2'] = "points possible"
    sheet['D2'] = "reason"
    sheet['E2'] = "percent"
 
    
    """   This needs to be reworked to make the totals work with the extra credit
    """
    
    count = 3
    totalList = ""
    #this needs to include the extra credit point!
    for point in points:
        #for each point i will place it in the correct cell
        totalList = totalList + "+" + "C" + str(count)
        sheet["C"+str(count)] = point
        count+=1
    #save the extra credit
    ecSpot = totalList[len(totalList)-4:]
    #this line removes the plus at the beginning and the last 3 chars which is the extracredit
    totalList = totalList[1:len(totalList)-4]
    #adding the total at the end for the points possible
    sheet["C"+str(count)] = "="+totalList
    #re-add the ecSpot
    totalList+=ecSpot
    
    #now i need to create the total equation for the points earned
    #changing every c in the total list to a b. 
    newTotalList = totalList.replace("C","B")
    sheet["B"+str(count)] = "="+newTotalList
    
    #set the percentage
    sheet["E"+str(count)] = "=" + "B" + str(count) + "/" + "C" + str(count)

    #save, name, and close the excel document
    wb.save(filename=student+assignmentName+"Grade.xlsx")
    wb.close()
    #empty wb
    wb = None
    os.chdir(p+"/"+assignmentName)
    print(student + "'s file and folder made.")