#!/usr/bin/python
#coding:utf-8 

import os  
import sys
import shutil
import py_compile


sourcePath = "./xls"
if os.path.isfile("./build.env"):
	sourcePath = "./xls"

def clean():
	if os.path.exists("json"):
		shutil.rmtree("json")
	if os.path.exists("lua"):
		shutil.rmtree("lua")
	if os.path.exists("xml"):	
		shutil.rmtree("xml")
	if os.path.exists("sqlite"):
		shutil.rmtree("sqlite")


paserFile = "python xls_parser.py"
buildFile = "python build.py"
if not os.path.isfile("./build.env"):
	paserFile = "python xls_parser.pyc"
if not os.path.isfile("./build.env"):
	buildFile = "python build.pyc"

if len(sys.argv) > 1:
	if sys.argv[1] == "clean":
		clean()
	elif sys.argv[1] == "all":
		# clean()
		for filename in os.listdir(sourcePath):
			if filename.endswith(".xls"):
				os.system(paserFile + " "+filename.split(".")[0] + " 3 sqlite")
	elif sys.argv[1].startswith("xls="):
		_name = sys.argv[1].split("=")[1]
		_type = 2
		_format = "json"
		t = _name.split("#")
		if len(t) == 3:
			_name = t[0]
			_type = t[1]
			_format = t[2]
		os.system(paserFile + " "+_name+ " " + _type + " " + _format)
	elif sys.argv[1] == "compile":
		py_compile.compile(r'./build.py')
		py_compile.compile(r'./xls_parser.py')
else:
	print "Usage: ./build.py [all|clean|copyfile|xls=XXXX#1#json]"
