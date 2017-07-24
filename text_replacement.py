# coding=utf-8

import os,sys, tarfile, shutil, xlrd,xlwt
from lxml import etree
from bs4 import BeautifulSoup
from datetime import datetime


reload(sys)
sys.setdefaultencoding('utf8')


HTMLSHEET = "html"
HTMLINDEX = 0
HTMLSECTION = 1
HTMLSUBSECTION = 2
HTMLUNIT = 3
HTMLCOMPNAME = 4
HTMLLOC = 5
HTMLFILE = 6




global path,source_path
path = "course"
source_path = "source"
backup_souce_path = "original_course"
xlsmPath = "conversion_table.xls"
wb = xlrd.open_workbook(xlsmPath)

file_err = open('error_files.txt', 'a') 
txt = "starttime" + datetime.now().strftime('%Y-%m-%d %H:%M:%S') +"\n"
file_err.write(txt)
file_err.close()


def read_find_html():
	print "---------------------------------begin reading html metadat from excel file----------------------------"


	global sheetstruc
	sheetstruc = wb.sheet_by_name(HTMLSHEET)
	for row in range(1, sheetstruc.nrows):


		chapter_name = sheetstruc.cell_value(row,HTMLSECTION)
		seq_name = sheetstruc.cell_value(row,HTMLSUBSECTION)
		ver_name = sheetstruc.cell_value(row,HTMLUNIT)
		comp_name = sheetstruc.cell_value(row,HTMLCOMPNAME)
		print "-------------------- begin replacing HTML component < " + comp_name + " > --------------------------------" 
		#print type(comp_name)
		map_html_chapter(row,chapter_name)
		print "-------------------------------------------------------------------------------------------------------"+ "\n"
		
		
		
		

def map_html_chapter(_row,_chapter_name):

	current_chapter_name = _chapter_name.rstrip().replace(u'\xa0', ' ')
	chap_path = path + "/chapter"
	chap_ls = os.listdir(chap_path)
	for each_chap in chap_ls:
		tree = etree.parse(chap_path+"/"+each_chap)
		root = tree.getroot()
		if current_chapter_name.lower() == root.get('display_name').lower():
			print "Found Sections: <" + str(current_chapter_name) + "> at filename --> " + str(each_chap)
			seq_ls = root.findall(".sequential")
			map_html_seq(seq_ls,_row)
			return()
	print "could not find Section: <" + (current_chapter_name) + "> Please check Section name on excel again" 
	quit()

def map_html_seq(_current_seq_list,_row):

	current_seq_name = sheetstruc.cell_value(_row,HTMLSUBSECTION).rstrip().replace(u'\xa0', ' ')
	seq_path = path +"/sequential"
	for each_seq in _current_seq_list:
		seq_filename = each_seq.get('url_name') + ".xml"
		tree = etree.parse(seq_path+"/"+seq_filename)
		root = tree.getroot()
		if current_seq_name.lower() == root.get('display_name').lower():
			print "Found Subsections: <" + str(current_seq_name) + "> at filename --> " + str(each_seq.get('url_name'))
			ver_ls = root.findall(".vertical")
			map_html_ver(ver_ls,_row)
			return()
	print "could not find Section: <" + (current_seq_name) + "> Please check Subsection name on excel again" 
	quit()


def map_html_ver(_current_ver_list,_row):
	

	ver_path = path +"/vertical"
	current_ver_name = sheetstruc.cell_value(_row,HTMLUNIT).rstrip().replace(u'\xa0', ' ')
	for each_ver in _current_ver_list:
		ver_filename = each_ver.get('url_name') + ".xml"
		tree = etree.parse(ver_path+"/"+ver_filename)
		root = tree.getroot()
		if current_ver_name.lower() == root.get('display_name').lower():
			print "Found Units: <" + str(current_ver_name) + "> at filename --> " + str(each_ver.get('url_name'))
			html_ls = root.findall(".html")
			map_html_component(html_ls,_row)
			return()

	print "could not find Unit: <" + (current_ver_name) + "> Please check Unit name on excel again" 
	quit()


def map_html_component(_current_html_list,_row):
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
			print "index number "+ str(sheetstruc.cell_value(_row,HTMLINDEX)) + " Found HTML: <" + (current_html_name) + "> at filename --> " + html_filename
			replace_content(html_link,_row)
			return()

	print "index number "+ str(sheetstruc.cell_value(_row,HTMLINDEX)) +" could not find HTML Component: <" + (current_html_name) + "> Please check HTML name on excel again" 
	quit()

def replace_content(_filename,_row):

	source_loc = str(sheetstruc.cell_value(_row,HTMLLOC))
	source = str(sheetstruc.cell_value(_row,HTMLFILE))
	source_file = source_path + source_loc+"/"+ source + ".html"


	html_path = path + "/html"
	des_file = html_path+"/"+_filename + ".html"
	shutil.copyfile(source_file, des_file)

	modify_figure_src(des_file,_filename)
	
	





	print "-------------------- Eng-to-Jap text replacing is complete --------------------------------" 



def modify_figure_src(_des_path,_filename):

	backup_path = path +"/html"
	backup_file = backup_souce_path+"/"+backup_path+"/"+_filename + ".html"


	file_eng = open(backup_file, 'r') 
	eng_text = file_eng.read() 
	tag_eng = BeautifulSoup(eng_text, 'html.parser')
	img_tag_eng = tag_eng.find_all('img')
	file_eng.close()

  	file_jap = open(_des_path, 'r') 
	jap_text = file_jap.read() 
	tag_jap = BeautifulSoup(jap_text, 'html.parser')
	img_tag_jap = tag_jap.find_all('img')
	file_jap.close()


	if len(img_tag_eng) != 0:
		if len(img_tag_jap) == len(img_tag_eng):

			for i in range(len(img_tag_jap)):
				img_tag_jap[i].attrs['src'] =img_tag_eng[i].attrs['src']
			jap_text = tag_jap.prettify()
			file_jap_mod = open(_des_path, 'w') 
			file_jap_mod.write(jap_text) 
			file_jap_mod.close()
			print "figure sources are all modified"




		else:
			print "number of figure is not equal to the source"
			file_err = open('error_files.txt', 'a') 
			txt = "fig in Jap text,fig in Eng text, path = " + str(len(img_tag_jap)) + " , " + str(len(img_tag_eng)) +" , " +_des_path + "\n"
			file_err.write(txt)
			file_err.close()


	else:
		print "no-figure text content"

	
def make_tarfile():

	print "file is being compressed as tar.gz "
	with tarfile.open(path + '/' + path + '.tar.gz', 'w:gz') as tar:
		for f in os.listdir(path):
			tar.add(path + "/" + f, arcname=os.path.basename(f))
		tar.close()
	print "uploadable file is created at " + path + '/' + path + '.tar.gz'




def generate_Edx():
	
	read_find_html()
	make_tarfile()



generate_Edx()