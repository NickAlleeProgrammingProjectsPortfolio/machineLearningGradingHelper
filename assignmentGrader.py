#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Jan 14 12:17:48 2020

program2 - the filegrader
open the first students excel file
print the students name
take input for (grade for question 1,2,3, and so on)
if the input isnt the full grade 
	then prompt for a reason
	insert the reason into the correct spot	
print the grades for each question out and print the total
ask if its any changes need to be made
if not
	write the grades into the file
	move on to the next student
		repeat



@author: nick
"""
import os, openpyxl, sys

assignmentName = sys.argv[1]
points = list(sys.argv[2].strip("[").strip("]").split(","))
pointTotal = 0
for l in points:
    pointTotal+= int(l)
extraCreditPoint = pointTotal*.05
points.append(str(extraCreditPoint))
p = os.getcwd()
studentScores = {}

for folder in os.listdir():
    count = 1
    row = 3
    grades = []
    
    #go into that folder and load the excel file for writing
    os.chdir(p + "/" + folder)
    #print out the student name and assignment name
    print("********************************" + folder + "********************************\n\n")
    #load the excel file
    wb = openpyxl.load_workbook(folder + "Grade.xlsx")
    sheet = wb["Sheet"]
    
    
    
    # for each section it will take in a grade.
    for point in points:
        grade = input("please enter the grade for  part" + str(count))
        grades.append(grade)
        #set the grade of the correct section
        sheet["B"+str(row)] = int(grade)
        #print confirmation
        print("part " + str(count) +" set to " + grade)
        #if points are missed then ask for a reason
        if float(grade) != float(point):
            reason = input("what is the reason for deducting points?")
            #insert the reason
            sheet["D"+str(row)] = reason
        #increment row and count
        row+=1
        count+=1
    total = 0
    for num in grades:
        total+=int(num)
    name = str(folder)[0:len(folder)-len(assignmentName)]
    print("the total score for " + str(name)+ " is " + str(total))
    #save the students name and score into a dictionary then print it out at the end
    studentScores[name] = total
    #save and close the excel file
    wb.save(filename = folder + "Grade.xlsx")
    wb.close()
    
"""todo:   print out student scores in a more pleasant way"""
print(studentScores)
""" todo:   print out scores to a text file and email them to me"""
    