#!/bin/sh
shellpath=`dirname $0`
cd $shellpath
# if [[ -d "xls" ]]; then
# 	echo "xls exist."
# else
# 	mkdir xls
# fi
# cp ../../../Documents/z_Data/*.xls ./xls/

python build.py all
# python build.py xls=PVELevelData#1#json
# python build.py compile