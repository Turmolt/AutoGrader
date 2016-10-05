# AutoGrader
Python program which grades Excel homework built by a lazy TA

```
#Auto Grader Assignment Syntax
#Created by Sam Gates

#<-Before the line means this is a comment... like this line

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
#Syntax as follows:
=1
#=1 means the first sheet

*[CellToCheck] CheckN [ValueToCheck] (Comment) -PointValue
#CheckN means to check for count of ValueToCheck >= N in CellToCheck

#* is an optional operator at the start of the line
#* means that the next statements beginning with ** will be hitting the same problem
#This means that we can have multiple queries checking the same problem but not
#deduct points multiple times for the same answer


#EXAMPLE:

=1
*[A1] Check1 [IF] (Did not use IF function in cell A1 -3)
**[A1] Check1 [AND] (Did not use AND function in cell A1 -3)

#Checks if there is 1 or more IF functions in cell A1, -3pts if they did not
```
