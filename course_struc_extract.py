# coding=utf-8

import os, tarfile, shutil, xlrd,xlwt, datetime
from lxml import etree


"""
	Const declaration
	this consts will help us in case of modifications in the excel sheet
"""

"""
	sheet->Video
"""
HTMLSHEET = "html"
HTMLINDEX = 0
HTMLSECTION = 1
HTMLSUBSECTION = 2
HTMLUNIT = 3
HTMLCOMPNAME = 4
HTMLLOC = 5
HTMLFILE = 6






idx = 0

#export_path = ""
"""
hardcoded xlsmpath must change to a parameter
"""
xlsmPath = "conversion_table.xls"
book = xlwt.Workbook()
sheet = book.add_sheet(HTMLSHEET)

def read_course():

	print "----------------generate excel file " + xlsmPath + " , sheetname " + HTMLSHEET + "----------------"
	idx = 0
	sheet.write(idx,0, "no")
	sheet.write(idx,1, "section")
	sheet.write(idx,2, "subsection")
	sheet.write(idx,3, "unit")
	sheet.write(idx,4, "component_name")
	sheet.write(idx,5, "file_loc")
	sheet.write(idx,6, "file_name")
	read_chapter()

def read_chapter():
	chap_path = "course/chapter"
	chap_ls = os.listdir(chap_path)
	for each_chap in chap_ls:
		tree = etree.parse(chap_path+"/"+each_chap)
		root = tree.getroot()
		chapter_name = root.get('display_name')
		seq_ls = root.findall(".sequential")
		read_sequential(chapter_name,seq_ls,)
		
	
def read_sequential(_current_chap_name,_seq_ls):
	seq_path = "course/sequential"
	#seq_ls = os.listdir(seq_path)
	for each_seq in _seq_ls:
		seq_filename = each_seq.get('url_name') + ".xml"
		tree = etree.parse(seq_path+"/"+seq_filename)
		root = tree.getroot()
		seq_name = root.get('display_name')
		ver_ls = root.findall(".vertical")
		read_vertical(_current_chap_name,seq_name,ver_ls)

def read_vertical(_current_chap_name,_current_seq_name,_ver_ls):
	ver_path = "course/vertical"

	#seq_ls = os.listdir(seq_path)
	for each_ver in _ver_ls:
		ver_filename = each_ver.get('url_name') + ".xml"
		tree = etree.parse(ver_path+"/"+ver_filename)
		root = tree.getroot()
		ver_name = root.get('display_name')
		html_ls = root.findall(".html")
		read_html(_current_chap_name,_current_seq_name,ver_name,html_ls)
		#print(html_ls)
		#read_vertical(_current_chap_name,seq_name,ver_ls)


def read_html(_current_chap_name,_current_seq_name,_current_ver_name,_html_ls):
	html_path = "course/html"

	#seq_ls = os.listdir(seq_path)
	for each_html in _html_ls:
		html_filename = each_html.get('url_name') + ".xml"
		tree = etree.parse(html_path+"/"+html_filename)
		root = tree.getroot()
		html_name = root.get('display_name')
		if html_name == "Licensing":
			continue
		add_course_struc(_current_chap_name,_current_seq_name,_current_ver_name,html_name)

def add_course_struc(_current_chap_name,_current_seq_name,_current_ver_name,_current_html_name):

	#print type(row_count),type(_current_chap_name.encode('ascii','ignore')),type(_current_seq_name.encode('ascii','ignore')),type(_current_ver_name.encode('ascii','ignore')),type(_current_html_name.encode('ascii','ignore'))
	global idx 
	idx +=1
	#print row_count
	sheet.write(idx,0, idx)
	sheet.write(idx,1, _current_chap_name)
	sheet.write(idx,2, _current_seq_name)
	sheet.write(idx,3, _current_ver_name)
	sheet.write(idx,4, _current_html_name)
	if _current_html_name is None:
		print "row " + str(idx) +" : text component with 'No title' " + "is copied"
	else:
		print "row " + str(idx) +" : text component: '" + _current_html_name +"' is copied"
	
	


def generate_Edx():
	

	read_course()
	book.save(xlsmPath)



generate_Edx()