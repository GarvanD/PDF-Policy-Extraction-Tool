# PDFExtractionTool
Project Structure:

----PDF Extraction Tool 

-------------.py files

-------------poppler

-------------qpdf

-------------tesseract

-------------Binary Classification


Version 1.0

1. This tool identifies and trims only the policy infomation rich components 
   of an insurance policy PDF and converts PDF files to PNG using poppler. 
   This is done by identifying an anchor word for the policy pages.

2. Then uses tesseract OCR to create hOCR files, which include bounding box inofmration. 
   A dictionary in the Automate.py script has specific search parameters and validations which 
   is used to locate meaningful information in the hOCR.
 
Version 2.0

1. This tool identifies and trims only the policy infomation rich components 
   of an insurance policy PDF and converts PDF files to PNG using poppler. 
   This is done by identifying an anchor word for the policy pages.
   
2. Then uses tesseract OCR to create hOCR files, which include bounding box inofmration. 
   Several "centers of mass" are identified within the PDF siginifying areas of dense information rich text.
   This is done by measuring the relative TF - IDF of certain key words. This decreases missed data points.
   
Version 3.0

1. This tool identifies and trims only the policy infomation rich components 
   of an insurance policy PDF and converts PDF files to PNG using poppler. 
   This is done by using a voting system between an SVM model and a Naive Bayes model
   which has been trained on labeled PDF data.

2. Then uses tesseract OCR to create hOCR files, which include bounding box inofmration. 
   Several "centers of mass" are identified within the PDF siginifying areas of dense information rich text.
   This is done by measuring the relative TF - IDF of certain key words. This decreases missed data points.
   
   
