#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Jan 18 12:57:56 2020

Rework of this grading helper because it needs to work with classroom assistants filesetup

take in list of students names in the source code
take in name of the assignment
take in dict of the question numbers and how many points each question is worth including a point for commits


make a folder to store folders
make a folder inside that folder for each student named correctly
create an excel file for each one of them
	the file should be named correctly and have a table of question numbers
	and how much they are worth. a column should be left for the points 
	earned. correct calculations for the total should be made.

@author: nick allee
"""

import os, openpyxl, sys

assignmentName = sys.argv[1]
p = os.getcwd()
#take in points and add extra credit
points = list(sys.argv[2].strip("[").strip("]").split(","))
pointTotal = 0
for l in points:
    pointTotal+= int(l)
extraCreditPoint = pointTotal*.05
points.append(str(extraCreditPoint))



for folder in os.listdir():
    #change into the students directory
    try:
        os.chdir(p+"/"+folder)
        print(folder)
        #make students excel document
        wb = openpyxl.Workbook()
        sheet = wb["Sheet"]
        #setup titles
        sheet['A1'] = assignmentName
        sheet['B2'] = "points earned"
        sheet['C2'] = "points possible"
        sheet['D2'] = "reason"
        sheet['E2'] = "percent"
 
        
        count = 3
        totalList = ""
        #this needs to include the extra credit point!
        for point in points:
            #for each point i will place it in the correct cell
            totalList = totalList + "+" + "C" + str(count)
            sheet["C"+str(count)] = point
            if (count<=len(points)+2):
                sheet["A"+str(count)] ="part"+str(count-2)
            else:
                sheet["A"+str(count)] = "correct amt of commits"
            count+=1
        #save the extra credit
        ecSpot = totalList[len(totalList)-3:]
        #this line removes the plus at the beginning and the last 3 chars which is the extracredit
        totalList = totalList[1:len(totalList)-3]
        #adding the total at the end for the points possible
        sheet["C"+str(count)] = "="+totalList
        #re-add the ecSpot
        totalList+=ecSpot
        
        #now i need to create the total equation for the points earned
        #changing every c in the total list to a b. 
        newTotalList = totalList.replace("C","B")
        #set the total equation
        sheet["B"+str(count)] = "="+newTotalList
        
        #set the percentage
        sheet["E"+str(count)] = "=" + "B" + str(count) + "/" + "C" + str(count)
    
        #save, name, and close the excel document
        wb.save(filename=str(folder)+assignmentName+"Grade.xlsx")
        wb.close()
        #empty wb
        wb = None
        print(str(folder) + "'s Excel file made.")
    except:
        print(folder)
        print("could not change into that directory.")
    
        
print("finnished making Excel files for all students")