Option Explicit

Dim mdl ' the current model
Set mdl = ActiveModel
If (mdl Is Nothing) Then
   MsgBox "There is no Active Model"
End If


addCommonProp  mdl

sub addCommonProp(mdl)
dim table
dim tablecount
'on error Resume Next
for each table in mdl.tables 
 addGeneral table
 tablecount = tablecount + 1
next

MsgBox "已给 " & tablecount &  "表添加通用字段"

Exit Sub
End sub

function addGeneral(table)
   dim prop

   set prop = table.Columns.CreateNew '通用字段1
   prop.Name = "enabled"
   prop.Code = "enabled"
   prop.DataType = "int"

   set prop = table.Columns.CreateNew '通用字段2
   prop.Name = "deleted"
   prop.Code = "deleted"
   prop.DataType = "int"

   set prop = table.Columns.CreateNew '通用字段3
   prop.Name = "creator"
   prop.Code = "creator"
   prop.DataType = "varchar(64)"

end function