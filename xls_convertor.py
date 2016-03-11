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
import logging
import traceback
from slpp import SLPP
import xml.etree.ElementTree as ET
import xml.dom.minidom

#日志级别：DEBUG INFO WARNING ERROR CRITICAL
logging.basicConfig(level=logging.DEBUG)

default_encoding = 'utf-8'
if sys.getdefaultencoding() != default_encoding:
    reload(sys)
    sys.setdefaultencoding(default_encoding)

class XlsConvertor:

	# @param xlsPath: excel文件路径
	def __init__(self, path, name, recursive=False):
		try:
			logging.info("加载数据：" + path + "/" + name + ".xls")
			self.xlsData = xlrd.open_workbook(path + "/" + name + ".xls")
		except Exception,e:
			logging.error(e)
		self.name = name
		self.path = path
		self.data = {}
		self.data[name] = self.parse(recursive)

	def convertTo(self, formatType="sql", recordPerFile=False, colType=0):
		if formatType == "sql":
			self.saveToSqlite(formatType, self.name, self.data[self.name], "sqlite.db3", colType)
		else:
			if recordPerFile:
				for k,v in self.data[self.name]["data"].items():
					if formatType == "json":
						self.saveToJson(formatType, self.name+"_"+k, v)
					elif formatType == "lua":
						self.saveToLua(formatType, self.name+"_"+k, v)
					elif formatType == "xml":
						self.saveToXML(formatType, self.name+"_"+k, v)
						pass
				pass
			else:
				if formatType == "json":
					self.saveToJson(formatType, self.name, self.data[self.name]["data"])
				elif formatType == "lua":
					self.saveToLua(formatType, self.name, self.data[self.name]["data"])
				elif formatType == "xml":
					self.saveToXML(formatType, self.name, self.data[self.name]["data"])

	def parse(self, recursive=False):
		lineIndex = 0
		lines = []
		lineKeyData = []
		lineMetaData = []
		table = self.xlsData.sheets()[0]
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
						lineData[str(lineKeyData[i])] = datArray[i]

						
						if recursive: #如果有配置数据关联，则递归加载需要的数据表，并且建立对应的关联
							if "->" in str(lineMetaData[i]):
								t = ""
								for line in str(lineMetaData[i]).split("\n"):
									if "->" in str(line):
										t = str(line).split("->")
								if len(t) > 1:
									name = t[1].split(".")[0]
									key = t[1].split(".")[1]
									newKey = t[0]
									if not name in self.data:
										self.data[name] = XlsConvertor(self.path, name, True).data[name]
									if str(datArray[i]) != "" and int(datArray[i]) != 0 :
										if str(datArray[i]) in self.data[name]["data"]:
											lineData[newKey] = copy.deepcopy(self.data[name]["data"][str(datArray[i])])
										else:
											logging.warning("警告：关联键" + str(datArray[i]) + "在" + name + "表中对应的数据不存在")

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
				logging.error("Some Error occured")
		tmp = {}
		for i in xrange(0, len(lineKeyData)):
			# print u'',lineMetaData[i]
			tmp[str(lineKeyData[i])] = lineMetaData[i]

		return {"keys":lineKeyData, "metaDesc":tmp, "data":finalObj}

	def checkFolderExist(self, path):
		if os.path.exists(path):
			pass
		else:
			os.mkdir(path) 

	def saveToJson(self, path, name, contentArray):
		self.checkFolderExist(path)
		destFile = path +"/" + name + ".json"
		logging.info("输出：" + destFile)
		output = open(destFile, 'w')
		output.write(json.dumps(contentArray, ensure_ascii=False, indent=4))
		output.close()
		pass

	def saveToLua(self, path, name, contentArray):
		self.checkFolderExist(path)
		lua = SLPP()
		destFile = path +"/" + name + ".lua"
		logging.info("输出：" + destFile)
		output = open(destFile, 'w')
		output.write("return " + lua.encode(contentArray))
		output.close()
		pass

	def saveToXML(self, path, name, contentArray, saveType=1):
		self.checkFolderExist(path)
		destFile = path +"/" + name + ".xml"
		logging.info("输出：" + destFile)
		elem = ET.Element('root')
		if saveType == 2:
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
				# print(key, val)
				child = ET.SubElement(elem, 'element')
				if not isinstance(val, dict):
					child.set(str(key),str(val))
				else:
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

	#loadType=0时不导出包含DONT_LOAD的字段
	#loadType=1时不导出SERVER_ONLY和DONT_LOAD的字段
	#loadType=2时不导出CLIENT_ONLY和DONT_LOAD的字段
	def saveToSqlite(self, path, name, dataArray, dbfileName="sqlite.db3", loadType=0):
		self.checkFolderExist(path)
		logging.info("开始成生sql文件及sqlite数据文件")
		destFile = path +"/" + dbfileName
		sql = "/* Generate By XlsConvertor*/"

		conn = sqlite3.connect(destFile)
		conn.execute("DROP TABLE IF EXISTS " + name)
		conn.commit()
		sql = sql + "\n" + "DROP TABLE IF EXISTS " + name + ";"
		ddl = "CREATE TABLE " + name + "("
		t = True
		keys = []
		colTypes = []
		for x in dataArray["keys"]:
			if x!="" and x!=0 and dataArray["metaDesc"][x].find("DONT_LOAD")==-1:
				shouldAppend = False
				if loadType == 0:
					shouldAppend = True
				elif loadType == 1 and dataArray["metaDesc"][x].find("SERVER_ONLY")==-1:
					shouldAppend = True
				elif loadType == 2 and dataArray["metaDesc"][x].find("CLIENT_ONLY")==-1:
					shouldAppend = True

				if shouldAppend :
					keys.append(x)
					colType = "INT"
					for l in dataArray["metaDesc"][x].split("\n"):
						if l.startswith("$"):
							colType = l[1:]
					colTypes.append(colType)
				

		# print(keys)
		for key, val in dataArray["data"].items():
			for i in range(len(keys)):				
				ddl = ddl + (keys[i] + " " + colTypes[i])
				if t:
					ddl = ddl + " PRIMARY KEY NOT NULL,"
					t = False
				else:
					ddl = ddl + ", "
				
			break
		ddl = ddl[0:-2] + ");"
		logging.debug("将执行SQL创建表：" + ddl)
		sql = sql + "\n" + ddl
		conn.execute(ddl)

		for key, val in dataArray["data"].items():
			dml = "INSERT INTO " + name + " VALUES ("
			for x in keys:
					# if isinstance(val[x], int):
					# 	dml = dml + str(val[x]) + ", "
					# elif isinstance(val[x], str):
					dml = dml + "'" + str(val[x]) + "'" + ", "
			dml = dml[0:-2] + ");"
			logging.debug("将执行SQL插入数据：" + dml)
			sql = sql + "\n" + dml
			conn.execute(dml)
		conn.commit()
		conn.close()

		file_object = open(path + "/" + name + ".sql", 'w+')
		file_object.write(sql)
		file_object.close( )
		# print(dataArray["data"])
		# dataArray["keys"]
	#end saveToSqlite	

XlsConvertor("xls", "ArkData", False).convertTo("sql", False, 2)

#end