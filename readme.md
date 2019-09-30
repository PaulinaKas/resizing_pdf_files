# The app resizing .pdf files

## Table of contents:
#### * General info
#### * Technologies 
#### * Setup

## **General info**:
The main goal of this app is reducing paper waste during printing a large .pdf file.
The company I help with gets a .pdf file of over 100 pages, where the content on each page covers only about 20-30% of the page. 
Thanks to this app the final file to print contains about 30% of the origin number of pages and the app saves at least 70 print pages.

*Mode of action*:
1. The app sections single pages from given .pdf file and convert them into .jpg image.
2. Then crops image borders according to contours obtained thanks to cv2 module.
3. Afterwards the app rocognizes text on images and these texts are base for folders names.
4. Resized images are placed in proper folders.
5. Then the app merges every image (sets many small images on one page) and creates a .pdf file ready to print.
     
## **Technologies**:
Python 3.7.4

### Libraries and packages which have been used:
 - sys
 - tempfile
 - calendar
 - pytesseract
 - os
 - imghdr
 - fpdf
 - shutil
 - img2pdf
 - PyPDF2
 - pdf2image
 - **cv2**
 - **numpy**
 - **PIL**
 
 ## **Setup**:
 Windows 8.1 Pro
 
 
 

 That's all ;)


