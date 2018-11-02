# VBS-filter-txt
一个过滤文本的玩意儿
Windows环境

新建一个file.txt的文件内容为
123456
1234567
qz123456
qz123z456
qz123z4563
qz12z45 63
q#$12.45 63

放到vbs的同级目录

双击filter.vbs

生成out.txt 内容为qz123z4563

过滤内容是：包含6个数字，只能是6个数字，不能纯数字，能有空格
