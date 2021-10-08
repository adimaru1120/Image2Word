# Image2Word
Program to convert image with tables( with words in japanese and polish into Word document 

To process image (remove tables) it uses OpenCV, after processing it uses tesseract to convert
image to text and then with docx library, text is inserted to word document.

Program have also feature that checks if line of text alredy exists and uses to this base in temporary
.txt file and uses regexs to check for existing in base. 

Parameters used in image processing for now are hard coded.
