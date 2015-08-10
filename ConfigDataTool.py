#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Date    : 2015-08-10 17:25:01
# @Author  : MatriQ (Matrixdom@126.com)
# @Link    : https://github.com/MatriQ
# @Version : 0.1

import os
import path from os.path
import xlrd

# get need gene data cols
def getNeedCols(table):
	cols=[]
	ncols = table.ncols
	for i in range(ncols ):
		if table.cell(3,i).value >= 2:
				cols.append(i)
	return cols

def main():
	if len(sys.argv) < 1 :
		raise Exception("argv can not be empty")
	fileName=sys.argv[1]
	filePath=os.path.abspath(fileName)
	try:
		data = xlrd.open_workbook(filePath)
		table = data.sheets()[0]          #通过索引顺序获取
		cols=getNeedCols(table)
		
		
	except Exception, e:
		raise e


if __name__ == '__main__':
	main()