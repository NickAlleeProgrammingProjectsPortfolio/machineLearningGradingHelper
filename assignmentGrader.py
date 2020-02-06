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
import os, openpyxl, sys, pyinputplus as pyip

#a dictionary linking the github name to the student name
userNameToRealName = {"ppusap":"Pratyusha Pusapati","Chaitra543":"Chaitra Vemula"
                      ,"bollamharshavardhanreddy":"Harshavardhan Reddy Bollam","chaturkurma":"Chatur Veda Vyas Kurma","dakotagrvtt":"Dakota Gravitt","Druthi7":"Sharadruthi Beerkuri","Echtniet":"Clinton Davelaar"
                      ,"ForeverAnApple":"Dave Chen","halfnote":"Trick Rex","JaswanthiNannuru":"Jaswanthi Nannuru","KHart0012":"Kevin Hart","kiyuzi":"Paige Braymer"
                      ,"nikithamandala":"Nikitha Mandala","rajeshoo7":"Rajesh Kammari","ravikumaratluri":"Ravi Kumar Atluri","reddylavanya":"Lavanya Reddy Uppula"
                      ,"redhug":"Pavan Kumar Reddy Byreddy","rishikareddygaddam":"Rishika Reddy Gaddam","rohithbharadwaj":"Rohith Bharadwaj","RudraPotturi":"Rudra Teja Potturi","Saikiran5669":"Sai Kiran Reddy Baki"
                      ,"saikirandd":"Sai Kiran Doddapaneni","sanjanabaswa":"Sanjana Baswapuram","SravyaKatpally":"Sravya Katpally","sunilmundru":"Sunil Mundru","Sushma4548":"Sushma Rani Reddy Aleti"
                      ,"vamshiredd":"Vamshikrishna Reddy Yedalla","venkateshkunduru123":"Venkatesh Kunduru","vinusha09":"Vinusha Sandadi"
                      }

assignmentName = sys.argv[1]
points = list(sys.argv[2].strip("[").strip("]").split(","))
pointTotal = 0
for l in points:
    pointTotal+= int(l)
extraCreditPoint = pointTotal*.05
points.append(str(extraCreditPoint))
p = os.getcwd()
studentScores = {}


''' use shelf to hold a list of students that have already been graded.

        look to see if the shelf file for the particular assignment is created
        if so then grab it
        if not then make one
        
    each assignment will need a different list
    before asking for grades for each student the program should make sure it hasnt already been graded
    add the name of the student to the graded list after it is graded. 
    save the shelf file'''
    


for folder in sorted(os.listdir()):
    count = 1
    row = 3
    grades = []
    # the flag is so it doesent go into the files that arent folders
    flag = 0
    
    try:
        #go into that folder and load the excel file for writing
        os.chdir(p + "/" + folder)
        #print out the student name and assignment name
        print("********************************" + userNameToRealName[folder] + "********************************\n\n")
        #load the excel file
        wb = openpyxl.load_workbook(folder + assignmentName + "Grade.xlsx")
        sheet = wb["Sheet"]
        flag = 0
    except:
        print("couldnt change into: " + p + "/" + folder)
        flag = 1
    
    
    # for each section it will take in a grade.
    if flag == 0:
        for point in points:
            """check to see if the grade is an actual grade. if not then ask again. do this with regex"""
            
            grade = pyip.inputNum("please enter the grade for  part" + str(count) + ". The max score for this part is :" + str(point) +".", min = 0 , max =point )
            
            grades.append(grade)
            #set the grade of the correct section
            sheet["B"+str(row)] = int(grade)
            #print confirmation
            print("part " + str(count) +" set to " + str(grade))
            #if points are missed then ask for a reason
            if float(grade) != float(point):
                reason = pyip.inputStr("what is the reason for deducting points?")
                #insert the reason
                sheet["D"+str(row)] = reason
            #increment row and count
            row+=1
            count+=1
        total = 0
        for num in grades:
            total+=int(num)
        userName = str(folder)
        print("the total score for " + userName + " is " + str(total))
        #save the students name and score into a dictionary then print it out at the end
        studentScores[userNameToRealName[userName]] = total
        #save and close the excel file
        wb.save(filename = folder +assignmentName+ "Grade.xlsx")
        wb.close()
    
    

scorefile = open("studentScores.txt","w")

for key,value in sorted(studentScores.items()):
    print(key + "---scored---" + value)
    scorefile.write(key + "---scored---" + value + "/n")
scorefile.close()
