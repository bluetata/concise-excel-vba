### 8.6 解决办法：使用SaveAs方法保存.xlsx后，再次打开提示: 文件损坏,后缀名错误（格式错误）

**问题描述：** 旧宏文件里使用`SaveAs`方法，保存为.xls文件，当改成保存为.xlsx文件后，再次打开保存后的文件时，提示文件名后缀错误，无法打开宏文件生成的文件。

**解决办法：** 修改`SaveAs`方法中的参数，将FileFormat的参数设置成如下：

```vb
FileFormat:=xlOpenXMLWorkbook
```
