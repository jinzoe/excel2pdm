Option Explicit

Dim mdl ' the current model
Set mdl = ActiveModel
If (mdl Is Nothing) Then
   MsgBox "There is no Active Model"
End If

Dim HaveExcel
Dim RQ
Dim excelAddress
excelAddress = "D:/test.xlsx" 'excel文件地址
RQ = vbYes 'MsgBox("Is Excel Installed on your machine ?", vbYesNo + vbInformation, "Confirmation")
If RQ = vbYes Then
   HaveExcel = True
   ' Open & Create Excel Document
   Dim x1  '
   Set x1 = CreateObject("Excel.Application")
   x1.Workbooks.Open excelAddress  
   x1.Workbooks(1).Worksheets("Sheet1").Activate
Else
   HaveExcel = False
End If

excel2pdm x1, mdl

sub excel2pdm(x1, mdl)
dim rwIndex
dim tableName
dim table
dim prop
dim count

dim tableNameCol
dim tableCodeCol
dim propNameCol 
dim propCodeCol
dim propTypeCol
dim propKeyCol
dim propComment

tableNameCol = 1 '表名所在列
tableCodeCol = 1 '表Code所在列
propNameCol = 1  '属性名所在列
propCodeCol = 1  '属性code所在列
propTypeCol = 4  '属性类型所在列
propKeyCol = 4   '属性键值所在列
propComment = 100'其他属性自定


'on error Resume Next
For rwIndex = 1 To 2000 step 1 '默认读到2000行
   With x1.Workbooks(1).Worksheets("Sheet1")
      If .Cells(rwIndex, 1).Value = "" and .Cells(rwIndex+1, 1).Value = "" Then   '两个空行即退出 
         Exit For
      End If
      If .Cells(rwIndex, 1).Value = "" Then	'空行下一行为表名
      	rwIndex = rwIndex+1
         set table = mdl.Tables.CreateNew
         table.Name = .Cells(rwIndex , tableNameCol).Value
         table.Code = .Cells(rwIndex , tableCodeCol).Value
         count = count + 1
      Else
         set prop = table.Columns.CreateNew
         prop.Name = .Cells(rwIndex, propNameCol).Value
         prop.Code = .Cells(rwIndex, propCodeCol).Value
         prop.DataType =  .Cells(rwIndex, propTypeCol).Value
         'prop.Comment = .Cells(rwIndex,propComment).Value
         if .Cells(rwIndex,propKeyCol) = "PK" Then    '判断主键
            prop.Primary =true
         end if
      End If
   End With
Next

MsgBox "生成数据表结构共计 " + CStr(count), vbOK + vbInformation, "表"

Exit Sub
End sub
