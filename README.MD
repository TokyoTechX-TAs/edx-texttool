
<b> Text Modification Tool for edX Online course</b><br>
A new open edX Content Modification Tool for text replacement, which is developed to automatically replace English text with translated text of Online edX course. Tool extracts course structure from exported edX course and to excel file. Then the English text is replaced by translated text which is noted in the excel file by the user. 

<b> Python package dependenies</b>
- lxml, xlrd, xlwt, beautifulsoup4

<b><font size="5"> Workspace folder  </font></b>
- 2 python scripts 
<br>course_struc_extract.py --> this script generates an excel file ‘conversion table’ 
<br>text_replacement.py --> this script replaces target text components to exported course  
- 4 folders
<br>‘course’ folder --> Unzip an exported edX course and move into this directory 
<br>‘original_course’ --> copy 'course' folder and place it in this directory
<br>‘source’ folder --> Move all translated texts (HTML files) to this directories
<br>'Example' folder --> contain an ‘conversion table’ example files (both before and after filled up contents)


<b><font size="5"> How to run </font></b>
- Run a python script ‘course_struc_extract.py’ 
   --> excel file ‘conversion_table.xlm’ will be created
- Open the conversion_table 
   --> manually add the path location of translated texts (HTML files) contained inside ‘source’ folder (please see example file for detail)
- Run a ‘text_replacement.py’ 
   --> Original text in 'course' folder will be replaced by the translated texts in 'source' folder
- If program is successfully executed, ready-to-upload course (course.tar.gz) is created inside the 'course' folder


<b><font size="5">NOTICE!!!</font></b>


- This tool <u>only replaces</u> HTML(text) component.
- This script is only tested with 1 edX course. Therefore, unexpected error should happen while using the program. 
- There is a bug after uploading course back to edX. It is found that some text components disappeared in edX page. In our test, 1 component is missing.
- Some HTML component listed in ‘error_files.txt’ will not have figures showed properly in edX. This happens when original component (Eng) and translated component (local language) have different number of figures   


