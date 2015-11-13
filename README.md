# mytop
实现了两个功能

1.生成pdf工单并导出

2.使附件不直接存入mysql，提高性能

部署方法：

1.安装python2.7和web.py

2.将webpy部署在Apache上

3.将本项目放在 /var/www/ 目录

4.备份好之前的所有附件，然后更改mysql的attachment表，使contents_data字段的属性为varchar

