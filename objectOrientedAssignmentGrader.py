#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Feb  9 15:33:22 2020

@author: nick
"""


import os, openpyxl, sys, pyinputplus as pyip, pandas as pd

def addExtraCredit(points):
    pointTotal = 0
    for l in points:
        pointTotal+= int(l)
    extraCreditPoint = pointTotal*.05
    points.append(str(extraCreditPoint))
    return points

def writeToScoresList(name,score):
    append_new_line("studentScores.txt", name + "\t\t\t" + str(score))
    scoreDfList.append([name,score])
    return None


def append_new_line(file_name, text_to_append):
        """Append given text as a new line at the end of file"""
        # Open the file in append & read mode ('a+')
        with open(file_name, "a+") as file_object:
        # Move read cursor to the start of file.
            file_object.seek(0)
        # If file is not empty then append '\n'
            data = file_object.read(100)
            if len(data) > 0:
                file_object.write("\n")
        # Append text at the end of file
            file_object.write(text_to_append)

class studentScore:
    def __init__(self, name, githubName,excelFileName):
        self.name = name
        self.githubName = githubName
        self.excelFileName = excelFileName
        self.total = 0
    
    def gradeAssignment(self,points):
        grades = []
        # count is 1 because it marks the parts for each assignment
        count = 1
        # row set to 3 because its the row that we are gonna need first
        row = 3
        wb = openpyxl.load_workbook(self.excelFileName)
        sheet = wb["Sheet"]
        for point in points:
            intPoint = int(float(point))
            intPoint = int(float(point))
            """check to see if the grade is an actual grade. if not then ask again. do this with regex"""
            
            grade = pyip.inputNum("please enter the grade for  part" + str(count) + ". The max score for this part is :" + str(point) +".", min = 0 , max =intPoint)
            
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
        #set the studentscores total to total
        self.total = total
        print("the total score for " + str(self.name) + " is " + str(self.total))
        wb.save(filename = folder +assignmentName+ "Grade.xlsx")
        wb.close()





scoreDfList = []
userNameToRealName = {"ppusap":"Pratyusha Pusapati","Chaitra543":"Chaitra Vemula"
                      ,"bollamharshavardhanreddy":"Harshavardhan Reddy Bollam","chaturkurma":"Chatur Veda Vyas Kurma","dakotagrvtt":"Dakota Gravitt","Druthi7":"Sharadruthi Beerkuri","Echtniet":"Clinton Davelaar"
                      ,"ForeverAnApple":"Dave Chen","halfnote":"Trick Rex","JaswanthiNannuru":"Jaswanthi Nannuru","KHart0012":"Kevin Hart","kiyuzi":"Paige Braymer"
                      ,"nikithamandala":"Nikitha Mandala","rajeshoo7":"Rajesh Kammari","ravikumaratluri":"Ravi Kumar Atluri","reddylavanya":"Lavanya Reddy Uppula"
                      ,"redhug":"Pavan Kumar Reddy Byreddy","rishikareddygaddam":"Rishika Reddy Gaddam","rohithbharadwaj":"Rohith Bharadwaj","RudraPotturi":"Rudra Teja Potturi","Saikiran5669":"Sai Kiran Reddy Baki"
                      ,"saikirandd":"Sai Kiran Doddapaneni","sanjanabaswa":"Sanjana Baswapuram","SravyaKatpally":"Sravya Katpally","sunilmundru":"Sunil Mundru","Sushma4548":"Sushma Rani Reddy Aleti"
                      ,"vamshiredd":"Vamshikrishna Reddy Yedalla","venkateshkunduru123":"Venkatesh Kunduru","vinusha09":"Vinusha Sandadi","SaiNikhilPippara" : "Sai Nikhil Pippara"
}
assignmentName = sys.argv[1]
points = list(sys.argv[2].strip("[").strip("]").split(","))
p = os.getcwd()
studentScores = []
count = 0

# set the extra credit point
points = addExtraCredit(points)

# itterate thru folders
for folder in sorted(os.listdir()):
    ## cd into that directory
    do = 0
    try:
        os.chdir(p + "/" + folder)
        do = 1
    except Exception as e:
        print(e)
        print("didnt do anything this round.")
    if (do == 1):
        ## new studentscore class with init
        studentScoreX = studentScore(userNameToRealName[folder],folder,folder + assignmentName + "Grade.xlsx")
        ## add the student to the studentScores list
        studentScores.append(studentScoreX)
        ##print the studentscores name and gitname
        print(studentScores[count].name + "   *******    "  + studentScores[count].githubName)
        ## grade the students excel file(points)
        studentScores[count].gradeAssignment(points)
        ##cd out of the directory
        os.chdir(p)
        ##write student score to txt file
        writeToScoresList(studentScores[count].name,studentScores[count].total)
        ## increase the count
        count += 1

df = pd.DataFrame(scoreDfList,columns = ['Name','Score'])
print(df.to_markdown())

print("Finished Grading All Assignments... Goodbye")  
#done



