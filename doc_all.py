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
	name = re.findall(u'客户姓名：(.*)',text1)[0]
	print name
	MJnum = re.findall(u'项目编号：(.*)',text1)[0]
	print MJnum
	doctime = re.findall(u'时    间：(.*)',text1)[0]
	print doctime
	for i in range(len(docx.paragraphs)):
		#print  docx.paragraphs[i].text
		if u'客户姓名：' in docx.paragraphs[i].text:
			print i
			docx.paragraphs[i].runs[1].text = name
			print '1'
		if u'项目编号：' in docx.paragraphs[i].text:
			docx.paragraphs[i].runs[1].text = MJnum
			print '2'
		if u'时    间：' in docx.paragraphs[i].text:
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
				if u'项目名称' in docx.tables[i].rows[j].cells[k].text:
					docx.tables[i].rows[j+1].cells[k].text
				
			if u'项目名称' in docx.tables[i].rows[j].cells[0].text:
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
		__MJ_table(u'蛋白Accession号',u'对应的GO编号','%s/2.Annotation/2.1.GO/GO.list',r'GO:' %args.file,5)
		##*level2/3/4.xls
		__MJ_table(u'GO注释分类的分支',u'GO分类的定义','%s/2.Annotation/2.1.GO/GO.list.level2.xls'%args.file,r'GO:',10)
		#4.1.2
		##pathway.txt
		__MJ_table(u'蛋白的Accession编号',u'对应的KO号','%s/2.Annotation/2.2.KEGG/pathway.txt'%args.file,r'KO:',10)
		##pathway_table.xls
		__MJ_table(u'通路的编号',u'通路的定义','%s/2.Annotation/2.2.KEGG/pathways/pathway_table.xls'%args.file,r'KO:',5)
		##kegg_table.xls
		__MJ_table(u'蛋白名称',u'KO编号','%s/2.Annotation/2.2.KEGG/pathways/kegg_table.xls'%args.file,r'ko:',5)
		#4.1.3
		##COG.list
		__MJ_table(u'蛋白Accession号',u'对应的COG号','%s/2.Annotation/2.3.COG/COG.list'%args.file,r'COG:',5)
		##COG.annot.xls
		__MJ_table(u'蛋白名称',u'COG编号','%s/2.Annotation/2.3.COG/COG.annot.xls'%args.file,r'COG:',10)
		##COG.class.catalog.xls
		__MJ_table(u'COG 4种类型',u'COG功能分类，共25','%s/2.Annotation/2.3.COG/COG.class.catalog.xls'%args.file,r'\[',10)
		##*_vs_*.diff.exp.xls
		__MJ_table(u'蛋白Accession编号',u'样本1中该蛋白相对表达量均值','%s/3.DiffExpAnalysis/3.1.Statistics/Volcano/*_vs_*.diff.exp.xls'%args.file,r'\.',5)
		#4.2.3
		##*.enrichment.detail.xls
		__MJ_table(u'Id',u'Enrichment','%s/3.DiffExpAnalysis/3.2.GO/Enrichment/*_vs_*.diff.exp.xls.DE.list.enrichment.detail.xls'%args.file,r'GO:',5)
		##*.pathway.xls
		__MJ_table(u'KEGG pathway名称',u'数据库','%s/3.DiffExpAnalysis/3.3.KEGG/Enrichment/*.pathway.xls'%args.file,r'ko:',5)
		
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
		##GO二级分类统计条形图
		__MJ_jpg(u'level2.go.txt.pdf：GO二级分类统计条形图，如下图',file)
		##GO二级、三级、四级分类统计九饼图
		__MJ_jpg(u'level234.pdf：GO二级、三级、四级分类统计九饼图，如下图',file)
		#4.1.2
		##包含蛋白数目最多的前20个通路
		__MJ_jpg(u'pathway.top20.pdf：包含蛋白数目最多的前20个通路，如下图',file)
		##KEGG通路图片展示
		__MJ_jpg(u'pathways文件夹下的.png文件：KEGG通路图片展示，如下图',file)
		##KEGG代谢通路进行分类
		__MJ_jpg(u'kegg_classification.pdf：对蛋白做KO注释后，可根据它们参与的KEGG代谢通路进行分类',file)
		#4.1.3
		##COG功能分类统计柱图
		__MJ_jpg(u'COG.class.catalog.pdf：COG功能分类统计柱图，如下图',file)
		##差异蛋白可视化火山图
		__MJ_jpg(u'*_vs_*.volcano.pdf：各分组样本差异蛋白可视化火山图，见下图',file)
		##差异蛋白Venn图
		__MJ_jpg(u'*.Venn.pdf：差异蛋白Venn图，如下图',file)
		#4.2.1
		##上下调蛋白GO注释柱形图		
		__MJ_jpg(u'*gobars.pdf：上下调蛋白GO注释柱形图，见下图',file)
		#4.2.2
		##GO功能富集分析柱状图
		__MJ_jpg(u'*.enrichment.detail.xls.go.pdf：各分组差异蛋白GO功能富集分析柱状图，见下图',file)
		#4.2.3
		##差异蛋白KEGG通路图片展示
		__MJ_jpg(u'*.png：差异蛋白KEGG通路图片展示，见下图',file)
		##分组差异蛋白KEGG pathway富集分析柱状图
		__MJ_jpg(u'*.pathway.pdf：各分组差异蛋白KEGG pathway富集分析柱状图，见下图',file)
		#4.2.4
		##差异蛋白表达模式聚类图
		__MJ_jpg(u'Heatmap.pdf：差异蛋白表达模式聚类图',file)
		##差异蛋白模块（clusters）表达趋势折线图
		__MJ_jpg(u'Heatmap_trendlines_for_*_subclusters.pdf：差异蛋白模块（clusters）表达趋势折线图',file)
		#4.2.5
		##两组样本之间差异蛋白的蛋白互作图（png格式）
		__MJ_jpg(u'*_vs_*.network.png：两组样本之间差异蛋白的蛋白互作图（png格式），如下图',file)
		#4.2.6
		##Ipath整合通路图
		__MJ_jpg(u'*_vs_*.Ipath.png：Ipath整合通路图（png格式），如下图',file)	
	os.system('rm -rf tmp')	

_MJ3_jpg(docx3)


docx2.save('%s数据科质控.docx' %args.MJ)
docx3.save('%s生信.docx' %args.MJ)