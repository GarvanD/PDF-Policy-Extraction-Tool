#This file is the control for the entire PDF Automation process
#1. Nuance Power PDF will convert all PDF's to JPEG
#1a. Alt: The Poppler framework must be installed to do this via Python directly
#   as such Nuance is a work around until permission is granted to install Poppler
#2. The Tesseract CLI is called to create hOCR files for each image
#3. A function will accept the hOCR file as an input and create a Row object which
#   will store all the values extracted from the PDF 
#4. A function will place all the Row objects into a CSV file to be handled by the VBA Macro

#Packages for file-system traversal
import os, glob, ntpath, shutil
#hOCR navigation and parse
from bs4 import BeautifulSoup
#For writing to Excel
import xlsxwriter

#For PDF Trim
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import BytesIO


from pathlib import Path
#[SearchTerm,Search Result, yMax, yMin, Validation Flag,[String Validation]]
frontPage = [["Name","",600,300, False, ["Name", "of", "Insured"]],
             ["Mailing","",500,250, False, ["Mailing", "Address"]],
             ["Number","",650,450,False,["Policy","Number"]],
             ["Period","",700,500,False,["Policy","Period","to"]],
             ["POLICY","",400,300,False,["POLICY","DECLARATIONS"]]
             ]
buildingInfo = [["Address","",300,100,False,["Risk","Address"]],
                ["IBC","",300,100,False,["IBC", "Code"]],
                ["Year","",400,150,False,["Year","Built"]],
                ["No.","",350,200,False,["No.","of","Storeys"]],
                ["Walls","",400,200,False,["Walls"]]
                ]
policyInfo = [["Building","",500,250,False,["1","$","51236", "RC"]],
              ["Contents","",500,250,False,["51236","$","RC"]],
              ["Stock","",500,250,False,[""]],
              ["Equipment","",500,250,False,["RC","$","51236"]],
              ["Rental","",500,250,False,["Income","$","13515"]],
              ["Earthquake","",500,250,False,["endorsement","$","%"]],
              ["Flood","",500,250,False,["endorsement","Deductible", "$"]],
              ["Sewer","",500,250,False,["Back","up","endorsement","$"]],
              ["Crime","",500,250,False,["Package","1","$"]],
              ["Profits","",500,250,False,[""]],
              ["Perils","",500,250,False,["Named","Perils","Property","Coverage"]],
              ["Broad","",500,250,False,["Broad","Form","Property","Coverage"]],
              ["Miscellaneous","",500,250,False,["property","$"]]
             ]

participationInfo = [["Property","",5000,0,False,["Aviva","Damage","$","%"]],
                      ["Liability","",5000,0,False,["Aviva","$","%"]],
                       ["Equipment","",5000,0,False,["Aviva","%","$"]]
                       ]
liability = ""
def qpdf(filename,path,page_num):
    os.chdir(scriptDir + r"\qpdf\bin")
    cmd = 'qpdf ' + pdfIn + '\\' + filename + ' --pages . 1-' + str(page_num) + ' -- ' + path + '\\' + filename
    os.system(cmd)
def convertToPng(file,path,outFolder,fileName):
    os.chdir(scriptDir + r"\poppler\bin")
    cmd = 'pdftoppm.exe -png ' + path + '\\3_Page_Policies\\' + file + ' ' + outFolder + '\\'
    os.system(cmd)
    
def createOutputDir(fileName, path):
    newpath = path + '\\' + fileName
    if not os.path.exists(newpath):
       os.makedirs(newpath)
       os.chdir(newpath)
    else:
        os.chdir(newpath)
    return newpath

def convertToHocr(file, path, outFolder, fileName):
    os.chdir(scriptDir + r"\tesseract")
    print("\nDocument: " + outFolder)
    print("Page: " + filename[1])
    cmd = 'tesseract.exe ' + path + '\\' + file + ' ' + outFolder +'\\' + filename + ' hocr'
    os.system(cmd)

#Finds the page which a string occurs at
def find_text_return_page(_path, _string):
    filepath = open(_path, 'rb+')
    i = 0
    for page in PDFPage.get_pages(filepath, check_extractable=True):
        manager_1 = PDFResourceManager()
        retstr_1 = BytesIO()
        layout_1 = LAParams(all_texts=True)
        device_1 = TextConverter(manager_1, retstr_1, laparams=layout_1)
        interpreter_1 = PDFPageInterpreter(manager_1, device_1)
        interpreter_1.process_page(page)
        text_1 = retstr_1.getvalue()
        if str.encode(_string) in text_1:
            print("Document: " + ntpath.basename(_path) + "\nPage: " + str(i + 1) + "\nStatus: Scanning...\n")
            i += 1
        device_1.close()
        retstr_1.close()
    filepath.close()
    if i == 0:
        print("Document: " + ntpath.basename(_path) + "has no useful information")
        return -1
    else:
        print("Document: " + ntpath.basename(_path) + "\nPolicy info on pages 1 to " + str(i) +"\n")
        return i

#Trim PDFs to appropriate length
def trim():
    os.chdir(pdfIn)
    files = glob.glob("*.pdf")
    count = 0
    for file in files:
        count += 1
        #Specific for Chess 
        last_page = find_text_return_page(pdfIn + '\\' +file, "Claims Assist" )
        #if last_page == -1:
            #Handle Error
        print("Document: " + file + "\nTrimmed to range (page 1 - page " + str(last_page) + ")\n")
        if last_page == 3:
            newpath = pdfTrim +'\\3_Page_policies'
        else:
            newpath = pdfTrim + '\\Unable_to_process'
            print("Unable to process multiple location policies")
        if not os.path.exists(newpath):
            os.makedirs(newpath)
            os.chdir(newpath)
        else:
            os.chdir(newpath)
        qpdf(file,newpath,last_page)
    return count
def findPolicyInfo(_min,_max,_list,_soup):
    for x in _list:
        words = _soup.body.findAll("span", {"class": "ocrx_word"})
        for word in words:
            if word.string == x[0]:
                bbox = word['title']
                coords = bbox.split()
                yStart = int(coords[2])
                yEnd = coords[4]
                yEnd = int(yEnd[:-1])
                if (yStart > _min) and (yEnd < _max):
                    output =""
                    for data in word.parent.contents:
                        output += data.string.replace('\n', ' ').replace('\r', '')
                        for test in x[5]:
                            flag = True
                            if not test in output:
                                flag = False
                        if flag == True:
                            x[4] = flag
                            x[1] = output
                
def findLiability(_min,_soup):
    words = _soup.body.findAll("span", {"class": "ocrx_word"})
    for word in words:
        bbox = word['title']
        coords = bbox.split()
        yStart = int(coords[2])
        yEnd = coords[4]
        yEnd = int(yEnd[:-1])
        flag = False
        if (yStart > _min) and (yEnd < 3000):
            output =""
            for data in word.parent.contents:
                output += data.string.replace('\n', ' ').replace('\r', '')
            for test in ["33204"]:
                if (test in output):
                    return output
                else:
                    flag = False
            for test in ["51236"]:
                if (test in output):
                    return output
                else:
                    flag = False
            for not_test in ["33000","Liability","Coverage"]:
                if test in output:
                    flag = False
            if flag == True:
                return output
                       
def findInfo(_dict,_soup):
    words = _soup.body.findAll("span", {"class": "ocrx_word"})
    for word in words:
        for x in _dict:
            if(word.string == x[0]) and (x[4] == False):
                bbox = word['title']
                coords = bbox.split()
                yStart = int(coords[2])
                yEnd = coords[4]
                yEnd = int(yEnd[:-1])
                if (yStart > x[3]) and (yEnd < x[2]):
                    output =""
                    for data in word.parent.contents:
                        output += data.string.replace('\n', ' ').replace('\r', '')
                        for test in x[5]:
                            if (test in output):
                                flag = True
                            else:
                                flag = False
                        if flag == True:
                            x[4] = flag
                            x[1] = output
                        
def verify(_dict):
    for v in _dict:
        if v[4] == True:
            return True
        else:
            return False
        
def findBox(_min,_max,yStart,yEnd,val,validation):
    if val == 0:
        if (yStart > _min) & (yEnd < _max):
            output =""
            for data in word.parent.contents:
                output += data.string.replace('\n', ' ').replace('\r', '')
                for test in validation:
                    flag = True
                    if not test in output:
                        flag = False
                    if flag == True:
                        return yEnd
        else:
            return 0
    else:
        return val

def isBox(_start, _end):
    if (_start != 0) and (_end != 0):
        return True
    else:
        return False
def safeMakeFolder(newpath):
    if not os.path.exists(newpath):
       os.makedirs(newpath)
    else:
        shutil.rmtree(newpath)
        os.makedirs(newpath)
    
if __name__ == "__main__":
    scriptDir = os.getcwd()
    os.chdir('..')
    safeMakeFolder('PDF_TRIM')
    os.chdir('PDF_TRIM')
    pdfTrim = os.getcwd()
    os.chdir('..')
    safeMakeFolder('PNG_CONV')
    os.chdir('PNG_CONV')
    pngOut = os.getcwd()
    os.chdir('..')
    safeMakeFolder('hOCR')
    os.chdir('hOCR')
    hOut = os.getcwd()
    os.chdir('..')
    safeMakeFolder('EXCEL_OUT')
    os.chdir('EXCEL_OUT')
    EXCEL = os.getcwd()
    os.chdir('..')
    os.chdir('PDF_IN')
    pdfIn = os.getcwd()
    os.chdir(scriptDir)
   
    isValid_BI = False
    isValid_FP = False
    isValid_PI = False
    isValid_P = False
    
    page1 = []
    page2 = []
    page3 = []
  
    count_of_docs = trim()
    count_of_trimmed_docs = 0
    print("\nFinished Trim\n")
    if os.path.exists(pdfTrim + '\\3_page_Policies\\'):
        os.chdir(pdfTrim + '\\3_page_Policies\\')
        PDF_FILES = glob.glob("*.pdf")
        
        row_count = 0
        os.chdir(scriptDir)
        importCSV = []
        with open('ter_and_fus.csv','r+') as csvfile:
            for line in csvfile.readlines():
                importCSV.append(line.split(','))
        ter_and_fus_Codes =[]
        for loc in importCSV:
            new_loc = []
            for i in loc:
                i = i.strip()
                i = i.replace("\n","")
                new_loc.append(i)
            ter_and_fus_Codes.append(new_loc)
        
        for pdf in PDF_FILES:
            count_of_trimmed_docs += 1
            filename = Path(pdf).resolve().stem
            PNG_OUT = createOutputDir(filename,pngOut)
            HOCR_OUT = createOutputDir(filename,hOut)
            convertToPng(pdf, pdfTrim,PNG_OUT,filename)
            os.chdir(pngOut)
        for folder in os.listdir(pngOut):
            os.chdir(pngOut + '\\' + folder)
            PNG_FILES = glob.glob("*.png")
            for png in PNG_FILES:
                filename = Path(png).resolve().stem
                convertToHocr(png,PNG_OUT,HOCR_OUT,filename)
            os.chdir('..')
    
        os.chdir(scriptDir)
        os.chdir("..")
        os.chdir(hOut)
        for dir in os.walk(os.getcwd()):
            for folder in dir[1]:
                os.chdir("..")
                os.chdir(hOut + '\\' + folder)
                hOCRfiles = glob.glob("*.hocr") 
                print("\nDocument: " + folder)
                for file in hOCRfiles:
                    boxMax = 0
                    boxMin = 0
                    print("Page: " + file)
                    with open(file, "r") as f:
                        contents = f.read()
                        soup = BeautifulSoup(contents,'lxml')
                        words = soup.body.findAll("span", {"class": "ocrx_word"})
                        for word in words:
                            bbox = word['title']
                            coords = bbox.split()
                            yStart = int(coords[2])
                            yEnd = coords[4]
                            yEnd = int(yEnd[:-1])
                            
                            #Front Page Info
                            if "1" in file:
                                findInfo(frontPage,soup)
                                
                            isValid_FP = verify(frontPage)
                            
                                
                            
                            #Building Info                               
                            if ("2" in file) and (not isValid_BI):
                                findInfo(buildingInfo, soup)
                            isValid_BI = verify(buildingInfo)
                            
                                
                                
                            
                            #Policy Info
                            if ("2" in file) and (isValid_BI == True):
                                #Find Bounding Box
                                if (word.string == "FORM") or (word.string == "NO."):
                                    v = ["FORM","NO,","COVERAGE","RATE","PREMIUM"]
                                    boxMin = findBox(400,550,yStart,yEnd,boxMin,v)
                                if (word.string == "FORM") or (word.string == "NO."):
                                    v = ["FORM","NO","IS","COVERAGE","DEDUCTIBLE","TO"]
                                    boxMax = findBox(600,1200,yStart,yEnd,boxMax,v)
                                #Find Policy Info
                                findPolicyInfo(boxMin,boxMax,policyInfo,soup)
                                liability = findLiability(boxMax,soup)
                                    
                            
                            if ("3" in file):
                                if (word.string == "Aviva"):
                                    findInfo(participationInfo,soup)
                output_row = []
                contents_prem = 0
                contents_lim = 0
                for i in frontPage:
                    tmp = i[1].strip().split()
                    if i[0] == "Name":
                        print([tmp[0:3],tmp[3:]])
                        #output_row.append([0,tmp[3:]])
                    elif i[0] == "Mailing":
                        print([tmp[0:2],tmp[2:]])
                        #output_row.append([10,tmp[2:]])
                    elif i[0] == "Number":
                        print([tmp[0:2],tmp[2:]])
                        #output_row.append([1,tmp[2:]])
                    elif i[0] == "Period":
                        print([tmp[0:2],tmp[2:]])
                        #output_row.append([4,tmp[2:5]])
                        #output_row.append([5,tmp[6:9]])
                    elif i[0] == "POLICY":
                        print(tmp[0:2],tmp[3:])
                        #output_row.append([7,tmp[3:]])
                        
                for i in buildingInfo:
                    tmp = i[1].strip().split()
                    if i[0] == "Address":
                        print([tmp[0:2],tmp[2:-3]])
                        #output_row.append([19,tmp[2:-3]])
                        for city in ter_and_fus_Codes:
                            if "".join(tmp[-7:-6]).replace(",","") in city[0]:
                                print (city)
                                #output_row.append([25,city[2]])
                                #output_row.append([26,city[3]])
                    elif i[0] == "IBC":
                        print([tmp[-3:-1],tmp[-1]])
                        #output_row.append([28,tmp[-2:]])
                    elif i[0] == "Year":
                        print([tmp[0:2],tmp[2:3]])
                        #output_row.append([30,tmp[2:3]])
                    elif i[0] == "No.":
                        print([tmp[0:3],tmp[3:4]])
                        #output_row.append([33,tmp[3:4]])
                    elif i[0] == "Walls":
                        print(tmp[0:1],tmp[1:2])
                        #output_row.append([31,tmp[1:2]])
                for i in policyInfo:
                    if i[4] == True:
                        tmp = i[1].strip().split()
                        #lim = "".join(tmp[5:6])
                       # prem = "".join(tmp[6:7])
                        #lim = lim.replace("$","")
                        #lim = lim.replace(",","")
                        #if "." in lim:
                        #    lim = lim[:lim.index(".")]
                        #prem = prem.replace("$","")
                        #prem = prem.replace(",","")
                        #if "." in prem:
                        #    prem = prem[:prem.index(".")]
                        
                        if i[0] == "Building":
                            print([tmp[1],tmp[5:7]])
                            #output_row.append([34,tmp[5:6]])
                            #output_row.append([35,tmp[6:7]])
                        elif i[0] == "Contents":
                            print([tmp[1],tmp[5:7]])
                            #print(lim,prem)
                            #contents_lim += int(lim)
                            #contents_prem += int(prem)
                            
                        elif i[0] == "Stock":
                            print([tmp[0:2],tmp[5:7]])
                            #output_row.append([40,tmp[5:6]])
                            #output_row.append([41,tmp[6:7]])
                        elif i[0] == "Equipment":
                            print([tmp[1],tmp[5:7]])
                            #output_row.append([38,tmp[5:6]])
                            #output_row.append([39,tmp[6:7]])
                        elif i[0] == "Rental":
                            print([tmp[1],tmp[5:7]])
                            #contents_lim += int(lim)
                            #contents_prem += int(prem)
                        elif i[0] == "Earthquake":
                            print([tmp[1],tmp[5:7]])
                            #output_row.append([38,tmp[5:6]])
                            #output_row.append([39,tmp[6:7]])
                            #output_row.append([39,tmp[6:7]])
                        elif i[0] == "Flood":
                            print([tmp[1],tmp[5:7]])
                        elif i[0] == "Sewer":
                            print([tmp[1:4],tmp[5:8]])
                        elif i[0] == "Crime":
                            print([tmp[1],tmp[5:7]])
                        elif i[0] == "Profits":
                            print([tmp[1],tmp[5:7]])
                            #contents_lim += int(lim)
                            #contents_prem += int(prem)
                        elif i[0] == "Perils":
                            print([tmp[1],tmp[5:7]])
                        elif i[0] == "Broad":
                            print([tmp[1],tmp[5:7]])
                        elif i[0] == "Miscellaneous":
                            print([tmp[1],tmp[5:7]])
                            #contents_lim += int(lim)
                            #contents_prem += int(prem)
                #output_row.append([42,contents_lim])
                #output_row.append([43,contents_prem])       
                for i in participationInfo:
                    tmp = i[1].strip().split()
                    if i[0] == "Property":
                        if tmp[0] == "Aviva":
                            if tmp[1] == "Insurance":
                                if tmp[2] =="Company":
                                    if tmp[3] == "of":
                                        if tmp[4] == "Canada":
                                            if tmp[7] == "Property":
                                                if tmp[8] == "Damage":
                                                    print(tmp[0],tmp[5],tmp[6],tmp[7],tmp[8])
                    elif i[0] == "Liability":
                        tmp = i[1].strip().split()
                        if tmp[0] == "Aviva":
                            if tmp[1] == "Insurance":
                                if tmp[2] =="Company":
                                    if tmp[3] == "of":
                                        if tmp[4] == "Canada":
                                            if tmp[7] == "Liability":
                                                print(tmp[0],tmp[5],tmp[6],tmp[7])
                                                
                                                
                    elif i[0] == "Equipment":
                        tmp = i[1].strip().split()
                        if len(tmp) >= 7:
                            if tmp[0] == "Aviva":
                                if tmp[1] == "Insurance":
                                    if tmp[2] =="Company":
                                        if tmp[3] == "of":
                                            if tmp[4] == "Canada":
                                                if tmp[7] == "Equipment":
                                                    print(tmp[0],tmp[5],tmp[6],tmp[7])
                liability = liability.strip().split()                             
                print ("Liability Coverage", liability[-3:])
                #os.chdir(EXCEL)                
                #workbook = xlsxwriter.Workbook('Output.xlsx')
                #worksheet = workbook.add_worksheet()
                #for info in output_row:
                  #  worksheet.write(row_count,int(info[0])," ".join(info[1]))
                row_count += 1
    else:
        print("Can only process 3 page policies")
            
    #workbook.close()
    
    print("\nExtracted data from " + str(count_of_trimmed_docs) + " out of " + str(count_of_docs) + " PDF's")