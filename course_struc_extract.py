# coding=utf-8

import os,sys, tarfile, shutil, xlrd,xlwt, datetime,logging,importlib
from lxml import etree

#importlib.reload(sys)
#sys.setdefaultencoding('utf8')

"""
	this script will extract course structure of exported courses into excel file
	excel　file will be created in this directory
	
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


book = xlwt.Workbook()
sheet = book.add_sheet(HTMLSHEET)






idx = 1 



def read_course():
	"""
		write row title in the first row of excel sheet
	"""

	sheet.col(HTMLINDEX).width 		= 256 * 4				### col size = 4  characters-long
	sheet.col(HTMLSECTION).width 	= 256 * 50				### col size = 50 characters-long
	sheet.col(HTMLSUBSECTION).width = 256 * 50				### col size = 50 characters-long
	sheet.col(HTMLUNIT).width 		= 256 * 50				### col size = 50 characters-long
	sheet.col(HTMLCOMPNAME).width 	= 256 * 50				### col size = 50 characters-long
	sheet.col(HTMLLOC).width 		= 256 * 50				### col size = 50 characters-long
	sheet.col(HTMLFILE).width 		= 256 * 50				### col size = 50 characters-long

	print("----------------generate excel file " + xlsmPath + " , sheetname " + HTMLSHEET + "----------------")
	row_title_idx = 0
	sheet.write(row_title_idx,HTMLINDEX, 		"no")
	sheet.write(row_title_idx,HTMLSECTION, 		"section")
	sheet.write(row_title_idx,HTMLSUBSECTION, 	"subsection")
	sheet.write(row_title_idx,HTMLUNIT, 		"unit")
	sheet.write(row_title_idx,HTMLCOMPNAME, 	"component_name")
	sheet.write(row_title_idx,HTMLLOC, 			"file_loc")
	sheet.write(row_title_idx,HTMLFILE, 		"file_name")
	read_chapter()

def read_chapter():
	"""
		extract section name from exported course (subfolder -> chapter)
		param: 
		- chapter_name -> section name
		- seq_ls -> object of associated subsection's XML filenames
	"""

	chap_path = "course/chapter"
	chap_ls = os.listdir(chap_path)
	for each_chap in chap_ls:
		tree = etree.parse(chap_path+"/"+each_chap)
		root = tree.getroot()
		chapter_name = root.get('display_name')
		seq_ls = root.findall(".sequential")
		read_sequential(chapter_name,seq_ls)
		
	
def read_sequential(_current_chap_name,_seq_ls):

	"""
		extract associated subsection names from exported course (subfolder -> sequential)
		param:
		- seq_name -> associated subsection name
		- ver_ls -> object of associated unit's XML filenames
	"""

	seq_path = "course/sequential"
	for each_seq in _seq_ls:
		seq_filename = each_seq.get('url_name') + ".xml"
		tree = etree.parse(seq_path+"/"+seq_filename)
		root = tree.getroot()
		seq_name = root.get('display_name')
		ver_ls = root.findall(".vertical")
		read_vertical(_current_chap_name,seq_name,ver_ls)

def read_vertical(_current_chap_name,_current_seq_name,_ver_ls):

	"""
		extract associated unit names from exported course (subfolder -> vertical)
		param: 
		- ver_name -> associated unit name
		- html_ls -> object of associated html_component's XML filenames
	"""

	ver_path = "course/vertical"
	for each_ver in _ver_ls:
		ver_filename = each_ver.get('url_name') + ".xml"
		tree = etree.parse(ver_path+"/"+ver_filename)
		root = tree.getroot()
		ver_name = root.get('display_name')
		html_ls = root.findall(".html")
		read_html(_current_chap_name,_current_seq_name,ver_name,html_ls)
		#print(html_ls)
		


def read_html(_current_chap_name,_current_seq_name,_current_ver_name,_html_ls):

	"""
		extract associated html names from exported course (subfolder -> html)
		param: 
		- html_name -> associated html name
		
	"""
	html_path = "course/html"
	for each_html in _html_ls:
		html_filename = each_html.get('url_name') + ".xml"
		tree = etree.parse(html_path+"/"+html_filename)
		root = tree.getroot()
		html_name = root.get('display_name')
	
		#### this condition avoid listing 'licensing' HTML component.  #####
		
		if html_name == "Licensing":         
			continue

		######################################################################

		add_course_struc(_current_chap_name,_current_seq_name,_current_ver_name,html_name)

def add_course_struc(_current_chap_name,_current_seq_name,_current_ver_name,_current_html_name):
	"""
		write course structure in each row of excel sheet 
		param:
		idx 				-> row_index
		section_name 		-> _current_chap_name
		subsection_name		-> _current_seq_name
		unit_name			-> _current_ver_name
		component_name		-> _current_html_name
	"""
	
	global idx 
	
	sheet.write(idx,0, idx)
	sheet.write(idx,1, _current_chap_name)
	sheet.write(idx,2, _current_seq_name)
	sheet.write(idx,3, _current_ver_name)
	sheet.write(idx,4, _current_html_name)
	if _current_html_name is None:
		print("row " + str(idx) +" : text component with 'No title' " + "is extracted")
	else:
		print("row " + str(idx) +" : text component: '" + str(_current_html_name) +"' is extracted")
	idx +=1
	
	


def main():
	
	read_course()                   #### course extraction
	book.save(xlsmPath)				#### save excel file 


if __name__ == '__main__':
	try:
		main()
		closeInput = input("\nPress ENTER to exit")
		print("Closing...")

	except KeyboardInterrupt:
		logging.warn("\n\nCTRL-C detected, shutting down....")
		sys.exit(ExitCode.OK)

