# excel2pdm
将excel中的数据设计通过脚本导入到PowerDesigner生成sql建立数据库
###excel格式

---表名前有空行，包括第一个表。

###代码参数配置

---列的顺序和excel文件地址需要在代码中配置，见注释。

###PD操作

---打开PowerDesigner。新建PhysicalDataModal，选择数据库。Ctrl + Shift + X 打开脚本编辑器，复制粘贴run，Genernate Database。

另一个脚本用来给所有的表添加通用字段
