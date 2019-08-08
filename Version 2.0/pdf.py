import glob, os, PyPDF2, shutil
class PDF_2_HOCR:
    def qpdf(self,filename,path,pages):
        os.chdir(self.scriptDir + r"\qpdf\bin")
        if len(pages) > 2:
            start = min(pages)
            if start == 0:
                start += 1
            end = max(pages) + 1
            cmd = 'qpdf ' + self.files + '\\' + filename + ' --pages . ' + str(start) +'-' + str(end + self.pageOffset) + ' -- ' + path + '\\' + filename.replace('.pdf','_trim.pdf')
            os.system(cmd)
        

    def convertToHocr(self,file, outFolder, filename):
        os.chdir(self.scriptDir + r"\tesseract")
        print("\nDocument Output Folder: " + outFolder)
        print("Page: " + filename)
        os.chdir(outFolder)
        folderName = filename[:-2]
        if not os.path.isdir(folderName):
                os.mkdir(folderName)
        os.chdir(self.scriptDir + '\\tesseract')
        hout = filename[-1:]
        cmd = 'tesseract.exe ' + self.png_dir + '\\' + file + ' ' + outFolder + '\\' + folderName + '\\' + hout + ' hocr'
        os.system(cmd)

    def convertToPng(self,file,path,outFolder):
        os.chdir(self.scriptDir + r"\poppler\bin")
        cmd = 'pdftoppm.exe -png ' + path + '\\' + file + ' ' + outFolder + '\\' + file.replace('.pdf','')
        os.system(cmd)

    def findPolicyPages(self,file):
        os.chdir(self.scriptDir)
        pdfFileObj = open(file,'rb')    
        try: 
            pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
            self.num_pg = pdfReader.numPages
            policy_pages = []
            for page in range(self.num_pg):
                if page > self.max_page:
                    break
                pageObj = pdfReader.getPage(page)          
                text = pageObj.extractText()
                flag = False
                for word in self.anchorTerm.split():
                    found = text.find(word)
                    if found != -1:
                        flag = True
                    else:
                        flag = False
                if (page) not in policy_pages and flag:
                    policy_pages.append(page)
            return policy_pages
        except Exception:
            print("Cant read PDF yet....wait for version 2.0")

    def __init__(self,anchorTerm,PDF_directory,maxPage,pagesAfterLastAnchor = 0):
        #The anchor term is the term which distinguishes pages as policy pages
        self.anchorTerm = anchorTerm
        #The offset is an optional paramter which is used in the case that a policy page exists after the last anchor page
        self.pageOffset = pagesAfterLastAnchor
        #This is the directory containing all the PDF files
        self.files = PDF_directory
        #Max page is general assumption for where there could not be a policy page past this point
        #The smaller this number is the faster the program will work.
        self.max_page = maxPage
        #Files structure setup
        self.scriptDir = os.getcwd()
        
        folders = ['PNG_OUT','hOCR_OUT','PDF_TRIM']
        for folder in folders:
            if not os.path.isdir(folder):
                os.mkdir(folder)
            else:
                shutil.rmtree(folder)
                os.mkdir(folder)
            os.chdir(self.scriptDir)
        os.chdir(self.scriptDir)
        os.chdir('PNG_OUT')
        self.png_dir = os.getcwd()
        os.chdir(self.scriptDir)
        os.chdir('hOCR_OUT')
        self.hocr_dir = os.getcwd()
        os.chdir(self.scriptDir)
        os.chdir('PDF_TRIM')
        self.pdf_trim = os.getcwd()
        os.chdir(self.files)
        pdf_dir = glob.glob('*.pdf')
        
        
        for pdf in pdf_dir:
            pages = self.findPolicyPages(pdf)
            if pages == None:
                continue
            self.qpdf(pdf,self.pdf_trim,pages)

        os.chdir(self.pdf_trim)
        pdf_trimmed = glob.glob('*.pdf')
        for pdf in pdf_trimmed:
            self.convertToPng(pdf,self.pdf_trim,self.png_dir)  
        os.chdir(self.png_dir)
        png_trimmed = glob.glob('*.png')
        for png in png_trimmed:
            self.convertToHocr(png,self.hocr_dir,png.replace('.png',''))
        
    
        
        

