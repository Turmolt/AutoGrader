#author: sam gates
from openpyxl import Workbook
from openpyxl import cell
from openpyxl import load_workbook
import os

class AutoGrader:

    def main(self):

        self.Assignment = input('Enter Assignment Folder: ')

        #gF = Graded File, the file we write the output to
        self.gF = open(self.Assignment+'Graded.txt','w')

        for file in os.listdir(self.Assignment):
            if file.endswith(".xlsx"):
                print(self.Assignment+'/'+file.__str__())
                try:
                    self.gradePaper(self.Assignment+'/'+file.__str__())
                except:
                    print(file.__str__()+' had an error, OOPS')
                    self.gF.write('error\n\n')

        self.gF.close()


    def gradePaper(self,title):
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