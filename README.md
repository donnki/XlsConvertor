#导表工具用法及配置
* 支持将xls数据导出为json、xml、lua、SQL、sqlite_db格式
* 支持数据表关联，生成出树型结构的json、xml、lua。
* 支持加入标签忽略导出数据列，如CLIENT_ONLY：只导出给客户端， SERVER_ONLY：只导出给服务端， DONT_LOAD：忽略该字段

* 如果使用xls_convertor.py，excel表中描述行需要加上数据库类型定义。

数据表格式规范及数据关联
* 每个xls只包含1个sheet页
* 数据第一行为数据项名（每列只能以英文字母开头，英文、数字、下划线）
* 数据第二行为数据项说明，可以使用中文，其中该字段可以建立数据表之间的关联，例如：
<pre>
TB_Copy.xls里的sectionId字段对应的数据项说明为：<br>
章节编号:sectionData->TB_CopySeries.id<br>
它的意思是这个字段的数据对应于TB_CopySeries.xls的id字段。在生成数据时，会自动将TB_CopySeries.xls里的对应数据作为该条纪录的子项，保存到sectionData里。<br>
</pre>
* 从第三行起为数据值。
build.py
#命令参考：
 python build.pyc xls=Role#2#json
将xls目录下的Role表每条记录生成一个json文件，并保存到当前json目录里

python build.pyc xls=Buff#1#json
将xls目录下的Buff.xls表生成一个总的json文件，并保存到当前json目录里
 
python build.pyc freq
同时运行上面两条命令