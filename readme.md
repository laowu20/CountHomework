### 实现功能
把学号作为关键字统计收到的学生作业，将统计结果保存在excel文件中。
### 脚本使用方法示例
exap1:	用鼠标将主目录下的“example1”文件夹拖到“get_homework_check_example1.bat”上，项目主目录下生成的“统计结果.xls”为统计结果，统计结果中标“1”的单元格表示在对应文件夹中存在对应学号的文件。

exap2:	将“example2”拖到“get_homework_check_example2.bat”上，项目主目录下生成的“统计结果.xls”为统计结果。

### 测试环境
该脚本在win10系统python3.7(64bit)环境下测试通过。
### 注意事项
1.学号表.xlsl”为学生学号列表，使用时要配置。

2.python需要安装xlrd和xlwt包。