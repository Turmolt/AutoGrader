#author: sam gates
from openpyxl import Workbook
from openpyxl import cell
from openpyxl import load_workbook
import re
import os

class AutoGrader:

    def main(self):

        self.Assignment = input('Enter Assignment Folder: ')
        self.keyFile = input('Enter key file name: ') + '.sag'
        self.aSyntax=[]
        self.readAssignmentKey(self.keyFile)
        for i in range(0, self.aSyntax[0].__len__()):
            print(self.aSyntax[0][i].__str__() + "\n")

        #gF = Graded File, the file we write the output to
        #self.gF = open(self.Assignment+'Graded.txt','w')
        #self.gF.close()


        # for file in os.listdir(self.Assignment):
        #     if file.endswith(".xlsx"):
        #         print(self.Assignment+'/'+file.__str__())
        #         try:
        #             self.gradePaperHARDCODED(self.Assignment+'/'+file.__str__())
        #         except:
        #             print(file.__str__()+' had an error, OOPS')
        #             self.gF.write('error\n\n')

    def gradePaper(self,path):
        fileName = path.replace(self.Assignment+'/A3 Excel_','')
        fileName = fileName.replace('.xlsx','')

        #self.gF.write(fileName+'\n')
        self.awb = load_workbook(path)
        self.asheets = []
        for s in self.awb._sheets:
            self.asheets.append(s)

        if self.asheets.__len__()>0:
            score = 60
            try:
                #for each sheet
                for sheetNum in self.aSyntax:

                    ws=self.asheets[sheetNum]

                    #Go through each question
                    for qNum in range(0,self.aSyntax[sheetNum].__len__()):

                        curQ = self.aSyntax[sheetNum][qNum]
                        #Check to see if it is a single-condition question, check count if so
                        if curQ.__len__()==1:

                            if not self.checkStatement(curQ[0],ws,score):
                                score-=curQ[4]

                        elif curQ.__len__()>1:
                            #check each condition, if any fail then we break out of loop and subtract points
                            for i in range(0,curQ.__len__()):
                                if not self.checkStatement(curQ[i],ws,score):
                                    break
            except AttributeError:
                #self.gF.write('Student left cells blank\n\n\n')
                print('Student left cells blank')

            print('~~~~~~~~~~~~~~')
            print('Score: ' + score)
            print('~~~~~~~~~~~~~~')

    #Check for Statement stmt in worksheet ws
    def checkStatement(self, stmt, ws, score):

        #parse stmt into useful information
        cellToCheck=stmt[0]
        n=stmt[1]
        valToCheck=stmt[2]
        comment=stmt[3]
        pointVal=stmt[4]

        #check the workbook for the desired statement count, fail and subtract score if is less
        if ws[cellToCheck].value.upper().count(valToCheck)<n:
            #self.gF.write(comment)
            print(comment)
            score-=pointVal
            return False
        else:
            return True

    #Read in an assignment key to the self.aSyntax variable
    def readAssignmentKey(self,keyPath):
        aKey=open(keyPath,'r')

        #Break down lines in the key we just opened
        lines = [line.rstrip('\n') for line in aKey]

        curSheet = 0
        state = 0

        newQuestion = []

        multConditions = False

        for i in range(0,lines.__len__()):

            # Get the current line we are parsing
            l = lines[i]

            #State 0 = Start new Sheet
            #############################################################
            if state == 0:
                if l[0] == '#':
                    # Skip Line, continue because this is a comment
                    continue
                elif l[0] == '=':
                    curSheet = int(l[1])
                    self.aSyntax.append([])
                    state = 1

            #State 1 = Check for New Question
            #############################################################
            elif state == 1:
                if l[0] == '*':
                    print('Start Multiple Condition Statement')
                    newQuestion = []
                    state = 3
                elif l[0] == '[':
                    newQuestion = []
                    state = 2
                else:
                    print('Error on line '+i.__str__()+'), stuck in state 1')

            #State 2 = Read single statement in, fall back to S1
            #############################################################
            if state == 2:

                #parse into tokens separated by spaces
                parsedLine = re.split(' ',l)

                newStatement = self.readStatement(parsedLine)
                newQuestion=[newStatement]
                self.aSyntax[curSheet].append(newQuestion)
                state = 1

            #State 3 = Read multiple statements in as a multi-condition, fall back to S1 when Question ends
            #############################################################
            elif state == 3:
                parsedLine = re.split(' ',l)

                #Start of multi-condition
                if parsedLine[0].startswith('*['):
                    if not multConditions:
                        multConditions = True
                        parsedLine[0]=parsedLine[0].replace('*','')
                        newQuestion.append(self.readStatement(parsedLine))

                    #if multConditions, we need to finish the last question and then start a new one
                    #as this is the start of a new multi-condition question
                    elif multConditions:
                        self.aSyntax[curSheet].append(newQuestion)
                        newQuestion = []
                        parsedLine[0]=parsedLine[0].replace('*','')
                        newQuestion.append(self.readStatement(parsedLine))

                #continuation of multi-condition
                elif parsedLine[0].startswith('**[') and multConditions:
                    parsedLine[0]=parsedLine[0].replace('**','')
                    newQuestion.append(self.readStatement(parsedLine))

                #fall out of multi-condition when the line starts with [
                elif parsedLine[0].startswith('[') and multConditions:
                    print('End Multiple Condition Statement')
                    multConditions = False
                    self.aSyntax[curSheet].append(newQuestion)

                    #ensure that the line we are currently on gets read
                    newQuestion = [self.readStatement(parsedLine)]
                    self.aSyntax[curSheet].append(newQuestion)
                    state = 1

            #if a multi-condition is the last line of the assignment file
            if i == lines.__len__()-1 and multConditions:
                self.aSyntax[curSheet].append(newQuestion)

    #Read a statement from the parsed line in readAssignmentKey
    def readStatement(self,parsedLine):

        #This is the structure of our statements
        #[CellToCheck, operator, ValueToCheckFor, Grader Comment, PointValue]

        newStatement = ['', #Cell To Check
                        0,  #Operator
                        '', #Value we want to check is in Cell To Check
                        '', #Grader's comment about getting this question condition wrong
                        0]  #The point value that this question condition is worth

        newStatement[0] = parsedLine[0][1:-1]
        print(newStatement[0])

        #Check operator here
        if 'Check' in parsedLine[1]:
            newStatement[1] = int(parsedLine[1][-1])
            print(newStatement[1])
        else:
            print(parsedLine[1])

        newStatement[2] = parsedLine[2][1:-1]
        print(newStatement[2])

        commentString = ''

        for i in range(3,parsedLine.__len__()-1):
            commentString+=parsedLine[i] + ' '

        newStatement [4] = int(parsedLine[-1][-1])
        print(newStatement[4])

        commentString = commentString[1:-2] + ' (-'+newStatement[4].__str__()+'pts)'
        newStatement[3] = commentString
        print(newStatement[3])

        return newStatement

AG  = AutoGrader()

AG.main()