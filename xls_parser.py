#!/usr/bin/python
#coding:utf-8 
import os  
import json
import shutil
import pprint
import copy
import sys 
import xlrd
import sqlite3
from slpp import SLPP
import xml.etree.ElementTree as ET
import xml.dom.minidom

default_encoding = 'utf-8'
# sourcePath = "../Documents/xls"
# if os.path.isfile("./xls_parser.pyc"):
sourcePath = "./xls"
AllData = {}

if sys.getdefaultencoding() != default_encoding:
    reload(sys)
    sys.setdefaultencoding(default_encoding)

def _print(s):
	print(str(s).decode(default_encoding).encode(sys.getfilesystemencoding()))

def checkFolderExist(path):
	if os.path.exists(path):
		pass
	else:
		os.mkdir(path) 
def saveToJson(path, name, contentArray):
	checkFolderExist(path)
	destFile = path +"/" + name + ".json"
	_print("输出：" + destFile)
	output = open(destFile, 'w')
	output.write(json.dumps(contentArray, ensure_ascii=False, indent=4))
	output.close()
	pass

def saveToLua(path, name, contentArray):
	checkFolderExist(path)
	lua = SLPP()
	destFile = path +"/" + name + ".lua"
	_print("输出：" + destFile)
	output = open(destFile, 'w')
	output.write("return " + lua.encode(contentArray))
	output.close()
	pass

def saveToXML(path, name, contentArray, saveType):
	checkFolderExist(path)
	destFile = path +"/" + name + ".xml"
	_print("输出：" + destFile)
	elem = ET.Element('root')
	if saveType == "2":
		child = ET.SubElement(elem, 'element')
		for key, val in contentArray.items():

			if isinstance(val, dict):
				subChild = ET.SubElement(child, key)
				for key2, val2 in val.items():
					subChild.set(str(key2),str(val2))
					pass
				pass
			else:
				child.set(str(key),str(val))
			
	else:
		for key, val in contentArray.items():
			child = ET.SubElement(elem, 'element')
			for key2, val2 in val.items():
				if isinstance(val2, dict):
					subChild = ET.SubElement(child, key2)
					for key3, val3 in val2.items():
						subChild.set(str(key3),str(val3))
						pass
					pass
				else:
					child.set(str(key2),str(val2))
		pass
	with open(destFile,'w') as f:
		f.write(xml.dom.minidom.parseString(ET.tostring(elem,encoding="utf-8")).toprettyxml())
#end saveToXML	

def saveToSqlite(path, name, dataArray):
	checkFolderExist(path)
	destFile = path +"/data"
	conn = sqlite3.connect(destFile)
	conn.execute("DROP TABLE IF EXISTS " + name)
	conn.commit()
	ddl = "CREATE TABLE " + name + "("
	colType = "VARCHAR(64)"
	if name == "i18n":
		colType = "TEXT"
	t = True
	for key, val in dataArray["data"].items():
		for x in dataArray["keys"]:
			if x != "" and x != 0:
				if isinstance(val[x], int):
					ddl = ddl + (x + " INT")
					if t:
						ddl = ddl + " PRIMARY KEY NOT NULL,"
						t = False
					else:
						ddl = ddl + ", "
				elif isinstance(val[x], str):
					ddl = ddl + (x + " " + colType + ", ")
		break
	ddl = ddl[0:-2] + ");"
	# print(ddl)
	conn.execute(ddl)

	for key, val in dataArray["data"].items():
		dml = "INSERT INTO " + name + " VALUES ("
		for x in dataArray["keys"]:
			if x != "" and x != 0:
				if isinstance(val[x], int):
					dml = dml + str(val[x]) + ", "
				elif isinstance(val[x], str):
					dml = dml + "'" + str(val[x]) + "'" + ", "
		dml = dml[0:-2] + ");"
		# print(dml)
		conn.execute(dml)
	conn.commit()
	conn.close()
	# print(dataArray["data"])
	# dataArray["keys"]
#end saveToSqlite	

def tonumber(value):
	try:
		return int(value)
	except ValueError:
		return value
	
def loadXlsFile(filepath):
	_print("加载数据表：" + filepath)
	lineIndex = 0
	lines = []
	
	lineKeyData = []
	lineMetaData = []
	data = xlrd.open_workbook(filepath)
	table = data.sheets()[0]
	nrows = table.nrows
	ncols = table.ncols
	for row in range(nrows):
		
		lineData = {}
		datArray = []
		for j in range(ncols):
			cell_value = table.row(row)[j].value
			if table.row(row)[j].ctype in (2,3) and int(cell_value) == cell_value:
				cell_value = int(cell_value)
				datArray.append(cell_value)
			else:
				datArray.append(str(cell_value))

		if lineIndex == 0:			#key字段
			for i in xrange(0,len(datArray)):
				lineData[str(datArray[i])] = 0
				lineKeyData.append(datArray[i])
		elif lineIndex == 1:		#描述字段
			for i in xrange(0,len(datArray)):
				lineMetaData.append(datArray[i])
		else:						#数据字段
			for i in xrange(0,len(datArray)):
				if lineKeyData[i] != "":
					splitArr = str(datArray[i]).split("|")
					if len(splitArr) > 1:
						lineData[str(lineKeyData[i])] = splitArr
					else:
						lineData[str(lineKeyData[i])] = tonumber(datArray[i])

					#如果有配置数据关联，则递归加载需要的数据表，并且建立对应的关联
					t = str(lineMetaData[i]).split("->")
					if len(t) > 1:
						name = t[1].split(".")[0]
						key = t[1].split(".")[1]
						newKey = t[0].split(":")[1]
						if not name in AllData:
							AllData[name] = loadXlsFile(sourcePath + "/" + name+".xls")
						if str(datArray[i]) != "" and tonumber(datArray[i]) != 0 :
							if str(datArray[i]) in AllData[name]["data"]:
								lineData[newKey] = copy.deepcopy(AllData[name]["data"][str(datArray[i])])
							else:
								_print("警告：关联键" + str(datArray[i]) + "在" + name + "表中对应的数据不存在")
						
					
			lines.append(lineData)	
			
		lineIndex = lineIndex + 1
		
	
	finalObj = {}
	for line in lines:
		try:
			key = str(line[str(lineKeyData[0])])
			if key == "":
				continue
			finalObj[key] = line
		except:
			_print("Some error occured")
	tmp = {}
	for i in xrange(0, len(lineKeyData)):
		tmp[str(lineKeyData[i])] = lineMetaData[i]
	return {"keys":lineKeyData, "metaDesc":tmp, "data":finalObj}
#end loadXlsFile

if len(sys.argv) < 2:
	_print("Usage: python xls_parser.py <TARGET_XLS> [OUTPUT_TYPE] [OUTPUT_FORMATION]")
	_print("OUTPUT_TYPE(默认值为1): \n\t值为1：将指定TARGET_XLS文件的全部内容输出到一个json文件中\n\t值为2：将指定TARGET_XLS文件中的每一行输出成一个Json文件\n\t值为3：将TARGET_XLS的数据输出到SQLITE数据库对应表中")
	_print("OUTPUT_FORMATION：（默认值为json，支持json, lua, xml, sqlite）")
else:
	
	
	outputType = len(sys.argv) > 2 and sys.argv[2] or "1"
	outputFormation = len(sys.argv) > 3 and sys.argv[3] or "json"
	destPath = outputFormation
	

	AllData[sys.argv[1]] = loadXlsFile(sourcePath + "/" + sys.argv[1]+".xls")


	if outputType == "1":
		if outputFormation == "json":
			saveToJson(destPath, sys.argv[1], AllData[sys.argv[1]]["data"])
		elif outputFormation == "lua":
			saveToLua(destPath, sys.argv[1], AllData[sys.argv[1]]["data"])
		elif outputFormation == "xml":
			saveToXML(destPath, sys.argv[1], AllData[sys.argv[1]]["data"], outputType)
		elif outputFormation == "all":
			saveToJson("json", sys.argv[1], AllData[sys.argv[1]]["data"])
			saveToLua("lua", sys.argv[1], AllData[sys.argv[1]]["data"])
			saveToXML("xml", sys.argv[1], AllData[sys.argv[1]]["data"],outputType)
			pass
		
	elif outputType == "2":
		for k,v in AllData[sys.argv[1]]["data"].items():
			if outputFormation == "json":
				saveToJson(destPath, sys.argv[1]+"_"+k, v)
			elif outputFormation == "lua":
				saveToLua(destPath, sys.argv[1]+"_"+k, v)
			elif outputFormation == "xml":
				saveToXML(destPath, sys.argv[1]+"_"+k, v, outputType)
			elif outputFormation == "all":
				saveToJson("json", sys.argv[1]+"_"+k, v)
				saveToLua("lua", sys.argv[1]+"_"+k, v)
				saveToXML("xml", sys.argv[1]+"_"+k, v,outputType)
				pass
	elif outputType == "3":
		if outputFormation == "sqlite":
			saveToSqlite("sqlite", sys.argv[1], AllData[sys.argv[1]])
			pass
