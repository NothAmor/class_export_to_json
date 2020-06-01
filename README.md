# class_export_to_json
将课表导出为JSON文件, 并导入纯粹课表

#使用方法
将Excel文件放入跟脚本同目录, 然后改名为:class.xls

示例Excel:
![example](https://i.loli.net/2020/06/01/GOJcgakWX6r1ioY.png)

注释已经写得蛮清楚了, 如果课表的排列方式相同, 修改一下就可以用

首先安装依赖:
pip install xlrd

然后:
python class_export_to_json.py

就会在脚本文件目录生成json文件了

有个小问题, 就是会在json文件]前多一个逗号
会导致语法问题, 所以要手动打开json文件删除那个逗号, 就可以直接导入了
这个问题会解决的
