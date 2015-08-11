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
	#print source
	filePath=source.decode("utf-8").encode("gbk")# os.path.abspath(source.decode("gbk").encode("utf-8"))
	content=getContentHead(output)
	try:
		print filePath
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
		print("\ngene data config by %s[%s] failed\nplease make sure all of xls/xlsx is closed and data format is right\n----------------------" %
		 (source.decode("utf-8").encode('gbk'),e.message))#)
	return None

def main():
	parser=None
	args=None
	argvs = sys.argv[1:]

	sources=[]
	outputs=[]

	if len(argvs) ==0:
		configs={
			"充值配置表.xlsx":"RechargeConfig",
			"CH_称号_cs.xlsx":"PlayerTitleConfig",
			"DJ_道具_cs.xls":"PropsConfig",
			"FB_副本_cs.xlsx":"GameCopy",
			"GW_怪物_cs_MonsterConfig.xlsx":"MonsterConfig",
			"JN_技能_cs_SkillsConfig.xls":"SkillsConfig",
			"RY_荣誉声望商店_cs_PropsExchangeConfig.xlsx":"PropsExchangeConfig",
			"SJ_升级经验_cs_UpgradeExpConfig.xls":"UpgradeExpConfig",
			"VIP配置表.xlsx":"VipConfig",
			"ZB_装备_cs.xlsx":"EquipConfig",
			"暗金装备系统.xlsx":"AnjinConfig",
			"帮派仓库兑换表.xlsx":"GangStorageConfig",
			"帮派系统.xlsx":"GangLevelConfig",
			"答题.xls":"AnswerConfig",
			"等级成长礼包.xlsx":"LevelUpGiftConfig",
			"等级开放技能表.xls":"SkillOpenLevelConfig",
			"活动奖励库.xlsx":"ActivityAwardsConfig",
			"活动配置表.xlsx":"ActivityConfig",
			"活动商城配置表.xlsx":"ActivityShopConfig",
			"技能组合效果表.xls":"SkillsCompConfig",
			"角色属性配置.xlsx":"ActorType",
			"经脉丹田.xlsx":"MeridianDanTianConfig",
			"经脉等级.xlsx":"MeridianLevelConfig",
			"经脉品阶.xlsx":"MeridianRankConfig",
			"矿洞排行显示.xlsx":"MineRank",
			"令牌配置表.xlsx":"TokenConfig",
			"龙魂商店.xlsx":"DragonConfig",
			"名门挑战.xlsx":"MmChallengeConfig",
			"名配置表.xls":"ActorLastNameConfig",
			"任务系统.xlsx":"TaskConfig",
			"套装系统.xlsx":"SuitConfig",
			"提示文本配置表.xlsx":"MsgCodeConfig",
			"姓配置表.xls":"ActorFirstNameConfig",
			"游戏功能开放等级表.xlsx":"FeatureConfig",
			"装备强化配置表.xls":"EquipStrengConfig",
			"装备强化消耗表.xls":"EquipStrengConsumptionConfig"
			}
		for item in configs.items():
			sources.append(item[0])
			outputs.append(item[1])
		print "use defaut args"
	else:
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
		sourceCount=len(args.sources)
		for i in range(sourceCount):
			sources.append(args.sources[i])
			outputs.append(args.outputs[i])

	targetPath=os.path.abspath("")
	if args!=None and args.target !=None:
		targetPath=os.path.join(targetPath,args.target)
	else:
		targetPath=os.path.join(targetPath,"Configs")
		if not os.path.exists(targetPath):
			os.mkdir(targetPath)
	sourceCount=len(sources)
	for i in range(sourceCount):
		outFileName=os.path.join(targetPath,outputs[i]+".lua")
		content=geneDataConfig(sources[i],outputs[i])
		if content!= None:
			outFile=open(outFileName,"w")
			outFile.write(content)
			outFile.close()
			print("create config file %s\n" % os.path.basename(outFileName))
	print("finish to create all config\ntarget dir is %s" % targetPath)

if __name__ == '__main__':
	try:
		main()
	except Exception, e:
		traceback.print_exc()