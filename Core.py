#author: sam gates
from openpyxl import Workbook
from openpyxl import cell
from openpyxl import load_workbook
import re
import os

class AutoGrader:

    def main(self):

        #self.Assignment = input('Enter Assignment Folder: ')

        self.keyFile = input('Enter key file name: ') + '.sag'

        self.assSyntax=[]
        self.readAssignmentKey(self.keyFile)
        print(self.assSyntax)

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



    #State 0 = Start new Sheet
    #State 1 = Check for New Question
    #State 2 = Read single statement in, fall back to S1
    #State 3 = Read multiple statements in, fall back to S1 when Question ends
    #State 4 = The 'read in statement' section

    def readAssignmentKey(self,keyPath):
        assKey=open(keyPath,'r')

        lines = [line.rstrip('\n') for line in assKey]

        curSheet = 0
        state = 0
        statementNumber = 0

        newStatement = ['',0,'','',0]
        newQuestion = []

        multConditions = False

        for i in range(0,lines.__len__()):

            # Get the current character we are parsing
            c = lines[i]
            curChar = 0

            #############################################################
            if state == 0:
                if c[0] == '#':
                    # Skip Line, continue because this is a comment
                    continue
                elif c[0] == '=':
                    curSheet = int(c[1])
                    self.assSyntax.append([])
                    state = 1

            #############################################################
            elif state == 1:
                if c[0] == '*':
                    print('Start Multiple Condition Statement')
                    newQuestion = []
                    state = 3
                elif c[0] == '[':
                    newQuestion = []
                    state = 2
                else:
                    print('Error on line '+i.__str__()+'), stuck in state 1')

            #############################################################
            if state == 2:
                parsedLine = re.split(' ',c)

                newStatement = self.readStatement(parsedLine)
                newQuestion=[newStatement]

                self.assSyntax[curSheet].append(newQuestion)
                state = 1

            #############################################################
            elif state == 3:
                parsedLine = re.split(' ',c)
                if parsedLine[0].startswith('*[') and not multConditions:
                    multConditions = True
                    parsedLine[0]=parsedLine[0].replace('*','')
                    newQuestion.append(self.readStatement(parsedLine))

                elif parsedLine[0].startswith('**[') and multConditions:
                    parsedLine[0]=parsedLine[0].replace('**','')
                    newQuestion.append(self.readStatement(parsedLine))

                else:
                    print('End Multiple Condition Statement')
                    multConditions = False
                    self.assSyntax[curSheet].append(newQuestion)
                    state = 1

            if i == lines.__len__()-1 and multConditions:
                self.assSyntax[curSheet].append(newQuestion)



    def readStatement(self,parsedLine):
        newStatement = ['',0,'','',0]

        newStatement[0] = parsedLine[0][1:-1]
        print(newStatement[0])
        newStatement[1] = int(parsedLine[1][-1])
        print(newStatement[1])
        newStatement[2] = parsedLine[2][1:-1]
        print(newStatement[2])

        commentString = ''

        for i in range(3,parsedLine.__len__()-2):
            commentString+=parsedLine[i] + ' '

        newStatement [4] = int(parsedLine[-1][-1])
        print(newStatement[4])
        commentString = commentString[1:-2] + ' (-'+newStatement[4].__str__()+'pts)'
        newStatement[3] = commentString
        print(newStatement[3])


        return newStatement






    def gradePaperHARDCODED(self,title):
        fileName = title.replace(self.Assignment+'/A3 Excel_','')
        fileName = fileName.replace('.xlsx','')
        self.gF.write(fileName+'\n')
        self.awb = load_workbook(title)
        self.asheets = []
        for s in self.awb._sheets:
            self.asheets.append(s)

        if self.asheets.__len__()>0:
            ws=self.asheets[0]
            score = 60
            try:
                if 'COUNTIF' not in ws['L1'].value.upper():
                    self.gF.write('Did not use COUNTIF function in cell L1 (-3pts)\n')
                    score-=3

                #Check cells in range 2-5
                for i in range(2,6):
                    curCell = 'L'+i.__str__()
                    curK = 'K'+i.__str__()
                    if 'COUNTIF' not in ws[curCell].value.upper():
                        self.gF.write('Did not use COUNTIF function in cell '+curCell+' (-3pts)\n')
                        score-=3
                    elif '$G$9' not in ws[curCell].value.upper():
                        self.gF.write('Did not use absolute value to cells $G$9:$G$104 in cell '+curCell+' (-2pts)\n')
                        score-=2
                    elif curK not in ws[curCell].value.upper():
                        self.gF.write('Did not use relative value to cell '+curK+' in cell '+curCell+' (-3pts)\n')
                        score-=3

                if 'IF' not in ws['I9'].value.upper():
                    self.gF.write('Did not use IF function in cell I9 (-4pts)\n')
                    score -= 4
                elif '$B$5' not in ws['I9'].value.upper():
                    self.gF.write('Did not use relative address $B$5 in cell I9 (-4pts)\n')
                    score -= 4

                if 'D9=TRUE' in ws['I9'].value.upper():
                    self.gF.write('Used D9=TRUE in cell I9 when D9 is a boolean value (-2pts)\n')
                    score-=2
                elif 'D9="TRUE"' in ws['I9'].value.upper():
                    self.gF.write('Used D9="TRUE" when D9 is a boolean value. Also, "TRUE" is a string, not a boolean value (-4pts)\n')
                    score-=4

                doubleSubtract=False
                if 'AND' not in ws['J9'].value.upper():
                    self.gF.write('Did not use AND function in cell J9 (-3pts)\n')
                    score-=3
                elif '$B$6' not in ws['J9'].value.upper():
                    self.gF.write('Did not use absolute reference to cell $B$6 (-3pts)\n')
                    score-=3
                    doubleSubtract = True
                if 'IF' not in ws['J9'].value.upper():
                    self.gF.write('Did not use IF functtion in cell J9 (-3pts)\n')
                    score-=3
                elif not doubleSubtract and '$B$6' not in ws['J9'].value.upper():
                    self.gF.write('Did not use absolute reference to cell $B$6 (-3pts)\n')
                    score-=3

                times = 0
                if 'IF' not in ws['K9'].value.upper():
                    self.gF.write('Did not use IF statement in cell K9 (-3pts)\n')
                    score-=3
                    times +=1
                if ws['K9'].value.upper().count('IF')< 2:
                    self.gF.write('Did not use nested IF statement in cell K9 (-3pts)\n')
                    score-= 3
                    times +=1

                if times<2 and ws['K9'].value.upper().count('$B$')<=2:
                    self.gF.write('Did not use absolute reference to both $B$3 and $B$4 in cell K9 (-3pts)\n')
                    score-=3

                ifStatements = ws['L9'].value.upper().count('IF')
                if ifStatements == 1:
                    self.gF.write('Did not use nested IF statement in cell L9 (-6pts)\n')
                    score -= 6
                elif ifStatements == 2:
                    self.gF.write('Did not use third IF statement in cell L9 (-3pts)\n')
                    score-=3

                if ws['L9'].value.upper().count('$D$') < 3 or ws['L9'].value.upper().count('$E$')<4:
                    self.gF.write('Did not use absolute reference for correct references in cell L9 (-3pts)\n')
                    score-=3

                if 'IF' not in ws['M9'].value.upper():
                    self.gF.write('Did not use IF statement in cell M9 (-3pts)\n')
                    score-=3
                if 'OR' not in ws['M9'].value.upper():
                    self.gF.write('Did not use OR statement in cell M9 (-3pts)\n')
                    score-=3
                self.gF.write('~~~~~~~~~~~~~~~~~~\n')
                self.gF.write('FINAL SCORE: '+score.__str__()+'\n')
                self.gF.write('~~~~~~~~~~~~~~~~~~\n')
                self.gF.write('\n\n')
            except AttributeError:
                self.gF.write('Student left cells blank\n\n\n')


AG  = AutoGrader()

AG.main()