#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Date    : 2015-08-10 17:25:01
# @Author  : MatriQ (Matrixdom@126.com)
# @Link    : https://github.com/MatriQ
# @Version : 0.1

import os
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
#import path from os
import xlrd
import re
import traceback
import argparse


class ColInfo(object):
	isLastField=False
	
	"""docstring for ColInfo"""
	def __init__(self, idx,key):
		super(ColInfo, self).__init__()
		self.idx=idx
		self.key=key


# get need gene data cols
def getNeedCols(table):
	cols=[]
	ncols = table.ncols
	for i in range(ncols ):
		if table.cell(2,i).value >= 2:
			colInfo=ColInfo(i,table.cell(1,i).value)
			colInfo.isLastField=i==ncols-1
			cols.append(colInfo)
	return cols

def getContentHead(tarName):
	return '''local\t%s = %s\tor\t{}\n\n%s={''' % (tarName,tarName,tarName)

def getContentEnd(tarName):
	return '''\n}\n\nreturn\t%s'''%tarName

def geneDataConfig(source,output):
	filePath=os.path.abspath(source)
	content=getContentHead(output)
	try:
		data = xlrd.open_workbook(filePath)
		table = data.sheets()[0]          #通过索引顺序获取
		cols=getNeedCols(table)
		nrows=table.nrows
		for r in range(4,nrows):
			for col in cols:
				v=table.cell(r,col.idx).value
				s=""
				if isinstance(v,float) or isinstance(v,int):
					if v.is_integer():
						s=str(int(v))
					else:
						s=str(v)
				else:
					s="'%s'" % v
				if col.idx==0:
					content=content+"\n\t['"+s+"'] = {\n"
				content=content+"\t\t"+col.key+" = "+s
				if not col.isLastField:
					content=content+",\n"
				else:
					content=content+"\n\t}"
			if r!=nrows-1:
				content=content+","
		content=content+getContentEnd(output)
		return content
	except Exception, e:
		traceback.print_exc()
		print("gene data config by %s failed" % source)
	return None

def main():
	argvs = sys.argv[1:]
	#if len(argvs) == 0:
	#	print ("Args cannot empty")
	#	return

	parser = argparse.ArgumentParser(description="This is a description of %(prog)s",
										epilog="This is a epilog of %(prog)s",
										prefix_chars="-+",
										fromfile_prefix_chars="@",
										formatter_class=argparse.ArgumentDefaultsHelpFormatter)

	parser.add_argument("sources",
						help="xml source dir",
						nargs="+")

	parser.add_argument("-o", "-output",
							required=True,
							nargs="+",
							dest="outputs",
							help="used tmpls")


	parser.add_argument("-d", "-targetdir",
							required=False,
							dest="target",
							help="target dir")


	args = parser.parse_args(argvs)

	if len(args.sources)!= len(args.outputs):
		print("sources count and output count dos not match")
		sys.exit()

	targetPath=os.path.abspath("")
	if args.target !=None:
		targetPath=os.path.join(targetPath,args.target)
	else:
		targetPath=os.path.join(targetPath,"Configs")
		if not os.path.exists(targetPath):
			os.mkdir(targetPath)
	sourceCount=len(args.sources)
	for i in range(sourceCount):
		outFileName=os.path.join(targetPath,args.outputs[i]+".lua")
		content=geneDataConfig(args.sources[i],args.outputs[i])
		if content!= None:
			outFile=open(outFileName,"w")
			outFile.write(content)
			outFile.close()
			print("create config file %s" % os.path.basename(outFileName))
	print("finish to create all config\ntarget dir is %s" % targetPath)

if __name__ == '__main__':
	try:
		main()
	except Exception, e:
		traceback.print_exc()