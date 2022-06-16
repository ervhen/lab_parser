from curses.ascii import isspace
from gettext import bind_textdomain_codeset
import pylightxl as xl
import re

#can add enumerate to count number of rows
def unique(list1):
    unique_list = []
      
    for x in list1:
        # check if exists in unique_list or not
        if x not in unique_list:
            unique_list.append(x.strip())
    # print list and remove the extras in this one
    unique_list.pop(0)
    for x in unique_list:
        return unique_list

#prep the dictionary with the unique items then attatch the list to it
def dictpreper(uniquelist):
    mydict={}
    for x in uniquelist:
        mydict[x] = ["","","","","",""]
    return mydict

#should add something to read first row to get numbers
def get_shit(uniquelist):
    mydict = dictpreper(uniquelist)
    for x in range(2,256):
        myrow = db.ws(ws='Copy of MW MS LIMITS').row(row=x)
        if re.search(r'BOD', myrow[8].strip()):
            myBOD = str(myrow[9]) + "/" + str(myrow[10])
            if mydict[myrow[5].strip()][0] != myBOD and not(str(myrow[9]).isspace()):
                mydict[myrow[5].strip()][0] = myBOD
        if re.search(r'CBOD', myrow[8].strip()):
            myCBOD = str(myrow[9]) + "/" + str(myrow[10])
            if mydict[myrow[5].strip()][1] != myCBOD and not(str(myrow[9]).isspace()):
                mydict[myrow[5].strip()][1] = myCBOD
        if "TSS" in myrow[8].strip():
            myTSS = str(myrow[9]) + "/" + str(myrow[10])
            if mydict[myrow[5].strip()][2] != myTSS and not(str(myrow[9]).isspace()):
                mydict[myrow[5].strip()][2] = myTSS
        if "NH3" in myrow[8].strip()and not(str(myrow[9]).isspace()):
            myNH3 = str(myrow[9]) + "/" + str(myrow[10])
            if mydict[myrow[5].strip()][3] != myNH3:
                mydict[myrow[5].strip()][3] = myNH3
        if "FC" in myrow[8].strip():
            myFC = str(myrow[9]) + "/" + str(myrow[10])
            if mydict[myrow[5].strip()][4] != myFC and not(str(myrow[9]).isspace()):
                mydict[myrow[5].strip()][4] = myFC
        if re.search(r'\bOG\b', myrow[8].strip()):
            myTSS = str(myrow[9]) + "/" + str(myrow[10])
            if mydict[myrow[5].strip()][5] != myTSS and not(str(myrow[9]).isspace()):
                mydict[myrow[5].strip()][5] = myTSS
    laboutput(mydict)
        
        
        

def laboutput(mydict):
    for x in mydict:
        print(f"{x},{mydict[x][0]},{mydict[x][1]},{mydict[x][2]},{mydict[x][3]},{mydict[x][4]},{mydict[x][5]}", end=",")
        print("\n")

db = xl.readxl("C:\\Users\ErvinHennrich\Downloads\Copy of MW MS LIMITS.xlsx")
#its not 0 based
workinglist = unique(db.ws(ws='Copy of MW MS LIMITS').col(col=6))
get_shit(workinglist)