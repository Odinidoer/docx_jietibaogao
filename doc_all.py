#!/usr/bin/env python
#encoding:utf-8
#jun.yan@majorbio.com
#20170320

import argparse
import commands
import os
import sys
import re
from docx import Document

parser = argparse.ArgumentParser(description = "iTRAQ prok report automatically generate scripts")
parser.add_argument("i",dest = "MJ",type = str,required=True,help ="just as MJ20160606006!!")
parser.add_argument("--org",dest = "org",type = str,required=True,help ="'ork' or 'prok'")
parser.add_argument("-m","--test_department_m1",dest = "M1",type = str,required=True,help ="test_department report file")
parser.add_argument("-f","--file",dest = "file",type = str,required = True,help = "please input file,normally 'Result_Files'")
args = parser.parse_args()

docx1 = Document(args.M1)
docx2 = Document('/mnt/ilustre/users/jun.yan/scripts/pm_report_module/iTRAQ/%s/iTRAQ2.docx' %args.org)
docx3 = Document('/mnt/ilustre/users/jun.yan/scripts/pm_report_module/iTRAQ/%s/iTRAQ3.docx' %args.org)

text1 = '\n'.join([docx1.paragraphs[i].text for i in range(len(docx1.paragraphs))])

def _first_page(docx):
	name = re.findall(u'�ͻ�������(.*)',text1)[0]
	print name
	MJnum = re.findall(u'��Ŀ��ţ�(.*)',text1)[0]
	print MJnum
	doctime = re.findall(u'ʱ    �䣺(.*)',text1)[0]
	print doctime
	for i in range(len(docx.paragraphs)):
		#print  docx.paragraphs[i].text
		if u'�ͻ�������' in docx.paragraphs[i].text:
			print i
			docx.paragraphs[i].runs[1].text = name
			print '1'
		if u'��Ŀ��ţ�' in docx.paragraphs[i].text:
			docx.paragraphs[i].runs[1].text = MJnum
			print '2'
		if u'ʱ    �䣺' in docx.paragraphs[i].text:
			docx.paragraphs[i].runs[1].text = doctime
			print '3'

_first_page(docx2)
_first_page(docx3)

def _project_info(docx):
	project_info = []
	for i in range(len(docx1.tables[0].rows)):
		for j in range(len(docx1.tables[0].rows[i].cells)):
			project_info.append(docx1.tables[0].rows[i].cells[j].text)
	project_info = '\t'.join(project_info)
	print project_info
	
	for i in range(len(docx.tables)):
		for j in range(len(docx.tables[i].rows)):
			for k in range(len(docx.tables[i].rows[j].cells)):
				if u'��Ŀ����' in docx.tables[i].rows[j].cells[k].text:
					docx.tables[i].rows[j+1].cells[k].text
				
			if u'��Ŀ����' in docx.tables[i].rows[j].cells[0].text:
				docx.tables[i].rows[j+1].cells[0].text = docx.tables[i].rows[j+1].cells[0].text

_project_info(docx2)
_project_info(docx3)


					
def __MJ_table(header1,header2,file,judgment1,judgment2): 
	file = commands.getoutput('''ls %s |head -1 ''' %file)
	if u'%s' %header1 in docx.tables[i].rows[0].cells[0].text and u'%s' %header2 in docx.tables[i].rows[0].cells[1].text:
		with open(file,'r') as table:
			mark = 1
			for line in table.readlines():
				if mark<3:
					items = line.strip().split('\t')
					count = re.subn(r'%s' %judgment1,r'%s' %judgment1,line)[1]
					if count<judgment2:
						for j in range(len(docx.tables[i].rows[mark].cells)):
							docx.tables[i].rows[mark].cells[j].text = items[j]
							docx.tables[i].rows[mark].cells[j].paragraphs[0].paragraph_format.alignment=table_format
						mark += 1
						
def _MJ3_tables(docx):
	for i in range(len(docx.tables)):
		table_format = docx.tables[i].rows[0].cells[0].paragraphs[0].paragraph_format.alignment
		#4.1.1
		##GO.list
		__MJ_table(u'����Accession��',u'��Ӧ��GO���','%s/2.Annotation/2.1.GO/GO.list',r'GO:' %args.file,5)
		##*level2/3/4.xls
		__MJ_table(u'GOע�ͷ���ķ�֧',u'GO����Ķ���','%s/2.Annotation/2.1.GO/GO.list.level2.xls'%args.file,r'GO:',10)
		#4.1.2
		##pathway.txt
		__MJ_table(u'���׵�Accession���',u'��Ӧ��KO��','%s/2.Annotation/2.2.KEGG/pathway.txt'%args.file,r'KO:',10)
		##pathway_table.xls
		__MJ_table(u'ͨ·�ı��',u'ͨ·�Ķ���','%s/2.Annotation/2.2.KEGG/pathways/pathway_table.xls'%args.file,r'KO:',5)
		##kegg_table.xls
		__MJ_table(u'��������',u'KO���','%s/2.Annotation/2.2.KEGG/pathways/kegg_table.xls'%args.file,r'ko:',5)
		#4.1.3
		##COG.list
		__MJ_table(u'����Accession��',u'��Ӧ��COG��','%s/2.Annotation/2.3.COG/COG.list'%args.file,r'COG:',5)
		##COG.annot.xls
		__MJ_table(u'��������',u'COG���','%s/2.Annotation/2.3.COG/COG.annot.xls'%args.file,r'COG:',10)
		##COG.class.catalog.xls
		__MJ_table(u'COG 4������',u'COG���ܷ��࣬��25','%s/2.Annotation/2.3.COG/COG.class.catalog.xls'%args.file,r'\[',10)
		##*_vs_*.diff.exp.xls
		__MJ_table(u'����Accession���',u'����1�иõ�����Ա������ֵ','%s/3.DiffExpAnalysis/3.1.Statistics/Volcano/*_vs_*.diff.exp.xls'%args.file,r'\.',5)
		#4.2.3
		##*.enrichment.detail.xls
		__MJ_table(u'Id',u'Enrichment','%s/3.DiffExpAnalysis/3.2.GO/Enrichment/*_vs_*.diff.exp.xls.DE.list.enrichment.detail.xls'%args.file,r'GO:',5)
		##*.pathway.xls
		__MJ_table(u'KEGG pathway����',u'���ݿ�','%s/3.DiffExpAnalysis/3.3.KEGG/Enrichment/*.pathway.xls'%args.file,r'ko:',5)
		
_MJ3_tables(docx3)

def __MJ_jpg(mark,file):
	file = commands.getoutput('''ls %s |head -1 ''' %file)
	if u'%s' %mark in docx.paragraphs[i]:	
		if 'pdf' in str(file):
			file_pdf = file
			file_jpg = **
			os.system('''convert -density 300  %s %s >/dev/null 2>&1''' %(file_pdf,file_jpg)
		else:
			file_jpg = file
		docx2.paragraphs[i+1].add_run(style = 26)
		docx2.paragraphs[i+1].runs[1].add_picture(file_jpg,width = 5555555)
		docx2.paragraphs[i+1].runs[0].clear()

def _MJ3_jpg(docx):
	if not os.path.isdir('tmp'):
		os.mkdir('tmp')
	for i in range(len(docx.paragraphs)):
		#4.1.1
		##GO��������ͳ������ͼ
		__MJ_jpg(u'level2.go.txt.pdf��GO��������ͳ������ͼ������ͼ',file)
		##GO�������������ļ�����ͳ�ƾű�ͼ
		__MJ_jpg(u'level234.pdf��GO�������������ļ�����ͳ�ƾű�ͼ������ͼ',file)
		#4.1.2
		##����������Ŀ����ǰ20��ͨ·
		__MJ_jpg(u'pathway.top20.pdf������������Ŀ����ǰ20��ͨ·������ͼ',file)
		##KEGGͨ·ͼƬչʾ
		__MJ_jpg(u'pathways�ļ����µ�.png�ļ���KEGGͨ·ͼƬչʾ������ͼ',file)
		##KEGG��лͨ·���з���
		__MJ_jpg(u'kegg_classification.pdf���Ե�����KOע�ͺ󣬿ɸ������ǲ����KEGG��лͨ·���з���',file)
		#4.1.3
		##COG���ܷ���ͳ����ͼ
		__MJ_jpg(u'COG.class.catalog.pdf��COG���ܷ���ͳ����ͼ������ͼ',file)
		##���쵰�׿��ӻ���ɽͼ
		__MJ_jpg(u'*_vs_*.volcano.pdf���������������쵰�׿��ӻ���ɽͼ������ͼ',file)
		##���쵰��Vennͼ
		__MJ_jpg(u'*.Venn.pdf�����쵰��Vennͼ������ͼ',file)
		#4.2.1
		##���µ�����GOע������ͼ		
		__MJ_jpg(u'*gobars.pdf�����µ�����GOע������ͼ������ͼ',file)
		#4.2.2
		##GO���ܸ���������״ͼ
		__MJ_jpg(u'*.enrichment.detail.xls.go.pdf����������쵰��GO���ܸ���������״ͼ������ͼ',file)
		#4.2.3
		##���쵰��KEGGͨ·ͼƬչʾ
		__MJ_jpg(u'*.png�����쵰��KEGGͨ·ͼƬչʾ������ͼ',file)
		##������쵰��KEGG pathway����������״ͼ
		__MJ_jpg(u'*.pathway.pdf����������쵰��KEGG pathway����������״ͼ������ͼ',file)
		#4.2.4
		##���쵰�ױ��ģʽ����ͼ
		__MJ_jpg(u'Heatmap.pdf�����쵰�ױ��ģʽ����ͼ',file)
		##���쵰��ģ�飨clusters�������������ͼ
		__MJ_jpg(u'Heatmap_trendlines_for_*_subclusters.pdf�����쵰��ģ�飨clusters�������������ͼ',file)
		#4.2.5
		##��������֮����쵰�׵ĵ��׻���ͼ��png��ʽ��
		__MJ_jpg(u'*_vs_*.network.png����������֮����쵰�׵ĵ��׻���ͼ��png��ʽ��������ͼ',file)
		#4.2.6
		##Ipath����ͨ·ͼ
		__MJ_jpg(u'*_vs_*.Ipath.png��Ipath����ͨ·ͼ��png��ʽ��������ͼ',file)	
	os.system('rm -rf tmp')	

_MJ3_jpg(docx3)


docx2.save('%s���ݿ��ʿ�.docx' %args.MJ)
docx3.save('%s����.docx' %args.MJ)