# coding=utf-8

import os,sys, tarfile, shutil, xlrd,xlwt
from lxml import etree
from bs4 import BeautifulSoup
from datetime import datetime


#reload(sys)
#sys.setdefaultencoding('utf8')
"""
	this script will replace HTML component to exported courses according to excel file
	
"""

"""
	initialize index of excel sheet  
	xlsmPath 			-> excel filename 
	HTMLSHEET 			-> sheetname        
	HTMLINDEX 			-> col1
	HTMLSECTION 		-> col2
	HTMLSUBSECTION 		-> col3
	HTMLUNIT 			-> col4
	HTMLCOMPNAME 		-> col5
	HTMLLOC 			-> col6
	HTMLFILE 			-> col7

"""

xlsmPath = "conversion_table.xls"
HTMLSHEET = "html"
HTMLINDEX = 0
HTMLSECTION = 1
HTMLSUBSECTION = 2
HTMLUNIT = 3
HTMLCOMPNAME = 4
HTMLLOC = 5
HTMLFILE = 6

"""
	initialize path  

	path 				-> course folder
	source_path 		-> source folder        
	backup_souce_path 	-> original course folder
	

"""


global path,source_path
path = "course"
source_path = "source"
backup_souce_path = "original_course"

wb = xlrd.open_workbook(xlsmPath)


"""
	create error log file as 'error_files.txt'
"""
file_err = open('error_files.txt', 'a') 
txt = "starttime" + datetime.now().strftime('%Y-%m-%d %H:%M:%S') +"\n"
file_err.write(txt)
file_err.close()



def read_find_html():

	"""
		1. open excel file
		2. begin reading course structure from excel's rows [section_name,subsection_name,unit_name,component_name]
				
	"""

	print("---------------------------------begin reading html metadata from excel file----------------------------")

	global sheetstruc
	sheetstruc = wb.sheet_by_name(HTMLSHEET)
	for row in range(1, sheetstruc.nrows):


		chapter_name 	= sheetstruc.cell_value(row,HTMLSECTION)
		seq_name 		= sheetstruc.cell_value(row,HTMLSUBSECTION)
		ver_name 		= sheetstruc.cell_value(row,HTMLUNIT)
		comp_name 		= sheetstruc.cell_value(row,HTMLCOMPNAME)

		if (len(sheetstruc.cell_value(row,HTMLLOC)) == 0) or (len(sheetstruc.cell_value(row,HTMLFILE)) == 0):

			print("--------------- no translated text file & directory :skipped HTML component  < " + str(comp_name) + " > -------------------")
			continue

		print("-------------------- begin replacing HTML component < " + str(comp_name) + " > --------------------------------")


		map_html_chapter(row,chapter_name)



		print("-------------------------------------------------------------------------------------------------------"+ "\n")
		
		
		
		

def map_html_chapter(_row,_chapter_name):

	"""
		1. seach location of section_name listed from excel at exported course (subfolder -> chapter)
			2.1 if found, search subsection
			2.2 if not found, stop script
		param:
		current_chapter_name -> section_name
		seq_ls 				 ->  object of associated subsection's XML filenames
				
	"""

	current_chapter_name = _chapter_name.rstrip().replace(u'\xa0', ' ')
	chap_path = path + "/chapter"
	chap_ls = os.listdir(chap_path)
	for each_chap in chap_ls:
		tree = etree.parse(chap_path+"/"+each_chap)
		root = tree.getroot()
		if current_chapter_name.lower() == root.get('display_name').lower():
			print("Found Sections: <" + str(current_chapter_name) + "> at filename --> " + str(each_chap))
			seq_ls = root.findall(".sequential")
			map_html_seq(seq_ls,_row)
			return()
	print("could not find Section: <" + (current_chapter_name) + "> Please check Section name on excel again")
	quit()

def map_html_seq(_current_seq_list,_row):

	"""
		1. seach location of associated subsection_name listed from excel at exported course (subfolder -> sequential)
			2.1 if found, search unit
			2.2 if not found, stop script
		
		param:
		current_seq_name -> subsection_name
		ver_ls 			 ->  object of associated unit's XML filenames
				
	"""


	current_seq_name = sheetstruc.cell_value(_row,HTMLSUBSECTION).rstrip().replace(u'\xa0', ' ')
	seq_path = path +"/sequential"
	for each_seq in _current_seq_list:
		seq_filename = each_seq.get('url_name') + ".xml"
		tree = etree.parse(seq_path+"/"+seq_filename)
		root = tree.getroot()
		if current_seq_name.lower() == root.get('display_name').lower():
			print("Found Subsections: <" + str(current_seq_name) + "> at filename --> " + str(each_seq.get('url_name')))
			ver_ls = root.findall(".vertical")
			map_html_ver(ver_ls,_row)
			return()
	print("could not find Section: <" + (current_seq_name) + "> Please check Subsection name on excel again")
	quit()


def map_html_ver(_current_ver_list,_row):

	"""
		1. seach location of associated unit_name listed from excel at exported course (subfolder -> vertical)
			2.1 if found, search html component
			2.2 if not found, stop script
		
		param:
		current_ver_name -> unit_name
		html_ls 			 ->  object of associated html_component's XML filenames
				
	"""

	

	ver_path = path +"/vertical"
	current_ver_name = sheetstruc.cell_value(_row,HTMLUNIT).rstrip().replace(u'\xa0', ' ')
	for each_ver in _current_ver_list:
		ver_filename = each_ver.get('url_name') + ".xml"
		tree = etree.parse(ver_path+"/"+ver_filename)
		root = tree.getroot()
		if current_ver_name.lower() == root.get('display_name').lower():
			print("Found Units: <" + str(current_ver_name) + "> at filename --> " + str(each_ver.get('url_name')))
			html_ls = root.findall(".html")
			map_html_component(html_ls,_row)
			return()

	print("could not find Unit: <" + str(current_ver_name) + "> Please check Unit name on excel again") 
	quit()


def map_html_component(_current_html_list,_row):


	"""
		1. seach location of text_component_name listed from excel at exported course (subfolder -> html)
			2.1 if found, start replacing HTML file
			2.2 if not found, stop script
		
		param:
		current_html_name -> html_name
		html_link 		  -> original HTML file of text
				
	"""


	html_path = path +"/html"
	current_html_name = sheetstruc.cell_value(_row,HTMLCOMPNAME).rstrip().replace(u'\xa0', ' ')

	for each_html in _current_html_list:
		html_link = each_html.get('url_name')
		html_filename = html_link + ".xml"
		tree = etree.parse(html_path+"/"+html_filename)
		root = tree.getroot()
		
		if root.get('display_name') == None:
			each_htmlname = ""
		else:
			each_htmlname	 = root.get('display_name')

		if current_html_name.lower() == each_htmlname.lower():
			print("index number "+ str(sheetstruc.cell_value(_row,HTMLINDEX)) + " Found HTML: <" + str(current_html_name) + "> at filename --> " + str(html_filename))
			replace_content(html_link,_row)
			return()

	print("index number "+ str(sheetstruc.cell_value(_row,HTMLINDEX)) +" could not find HTML Component: <" + str(current_html_name) + "> Please check HTML name on excel again")
	quit()

def replace_content(_filename,_row):


	"""
		1. locate translated HTML file in local directory listed on excel file
		2. replace HTML file in exported course with (1)
		
		param:
		source_loc  -> location of translated text in HTML file 
		source_file -> location of original text in HTML file 
				
	"""


	source_loc = str(sheetstruc.cell_value(_row,HTMLLOC))
	source = str(sheetstruc.cell_value(_row,HTMLFILE))
	source_file = source_path + source_loc+"/"+ source + ".html"


	html_path = path + "/html"
	des_file = html_path+"/"+_filename + ".html"
	shutil.copyfile(source_file, des_file)

	modify_figure_src(des_file,_filename,_row)
	
	





	print("-------------------- Eng-to-Jap text replacing is complete --------------------------------") 



def modify_figure_src(_des_path,_filename,_row):


	"""
		changing figure's source in HTMl file
		param:
		img_tag_eng -> original figure's source from original HTML file
		img_tag_translated　-> translated ver. figure's source from translated text HTML file

		beware:
		this script requires number of figure and order are the same as original HTML file.
		Otherwise, figure's source will not be changed, thus not shown in edx course 
		
				
	"""


	backup_path = path +"/html"
	backup_file = backup_souce_path+"/"+backup_path+"/"+_filename + ".html"


	file_eng = open(backup_file, 'r') 
	eng_text = file_eng.read() 
	tag_eng = BeautifulSoup(eng_text, 'html.parser')
	img_tag_eng = tag_eng.find_all('img')
	file_eng.close()

	file_translated = open(_des_path, 'r') 
	translated_text = file_translated.read() 
	tag_translated = BeautifulSoup(translated_text, 'html.parser')
	img_tag_translated = tag_translated.find_all('img')
	file_translated.close()


	if len(img_tag_eng) != 0:
		if len(img_tag_translated) == len(img_tag_eng):

			for i in range(len(img_tag_translated)):
				img_tag_translated[i].attrs['src'] =img_tag_eng[i].attrs['src']
			translated_text = tag_translated.prettify()
			file_translated_mod = open(_des_path, 'w') 
			file_translated_mod.write(translated_text) 
			file_translated_mod.close()
			print("figure sources are all modified")

		else:
			print("number of figure is not equal to the source")
			file_err = open('error_files.txt', 'a') 
			txt = "row in excel = " + str(sheetstruc.cell_value(_row,HTMLINDEX)) +"\n"
			txt = txt + "--> fig in translated text = "+  str(len(img_tag_translated)) + " ,fig in original text = "  +str(len(img_tag_eng))  +"\n"
			txt = txt + "----------------------------------------------\n"
			file_err.write(txt)
			file_err.close()


	else:
		print("no-figure text content")

	
def make_tarfile():

	"""
		compress exported course into tar.gz	
	"""

	print("file is being compressed as tar.gz ")
	with tarfile.open(path + '/' + path + '.tar.gz', 'w:gz') as tar:
		for f in os.listdir(path):
			tar.add(path + "/" + f, arcname=os.path.basename(f))
		tar.close()
	print("uploadable file is created at " + path + '/' + path + '.tar.gz')




def main():
	
	read_find_html()
	make_tarfile()


if __name__ == '__main__':
	try:
		main()
		closeInput = input("\nPress ENTER to exit")
		print ("Closing...")

	except KeyboardInterrupt:
		logging.warn("\n\nCTRL-C detected, shutting down....")
		sys.exit(ExitCode.OK)

