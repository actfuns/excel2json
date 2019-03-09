# excel2json

将excel表格转化问json对象或数组, 保存到指定位置, 如果序列化的类型为对象只能序列化数据行的第一行数据.

## 表格格式

json表头(json的key)的下一行必须为数据类型, 之后的行都为数据行, 如果列为数组类型数据用’,‘分割.  
数据类型有: string,int,float,boolen,object,array\<T\>   

## 命令行参数

* -e, –excel Required. 输入的Excel文件路径.
* -j, –json 指定输出的json文件路径.
* -h, –header (默认：第一行)表格第几行是json表头.
* -c, –encoding (默认: utf8-nobom) 指定编码的名称.
* -s, –sheet (默认: 第一个sheet)表格sheet名称.
* -a 序列化成数组
* -d, --date:指定日期格式化字符串，例如：dd / MM / yyy hh: mm:ss

## 简单实例

输入表格文件(in.xls)：  

|  prop1 |  prop2 |  prop3 |  prop4 |      prop5      |
|--------|--------|--------|--------|-----------------|
| string | float  | boolen | object |   array\<int\>  |
|  test  |   1.1  |  true  |{"a": 1}| 1,2,3,4,5,6,7,8 |
|  test2 |   1.2  |  false |{"a": 2}| 9,2,3,4,5,6,7,8 |

执行命令：..\ActFuns.Tools.Excel2Json\bin\Release\netcoreapp2.2\publish\excel2json.exe -e in.xls -j out.json -a  
输出json文件(out.json)：  
<pre>
[
  {
    "prop1": "test",
    "prop2": 1.1,
    "prop3": true,
    "prop4": {
      "a": 1
    },
    "prop5": [
      1,
      2,
      3,
      4,
      5,
      6,
      7,
      8
    ]
  },
  {
    "prop1": "test2",
    "prop2": 1.2,
    "prop3": false,
    "prop4": {
      "a": 2
    },
    "prop5": [
      9,
      2,
      3,
      4,
      5,
      6,
      7,
      8
    ]
  }
]
</pre>
