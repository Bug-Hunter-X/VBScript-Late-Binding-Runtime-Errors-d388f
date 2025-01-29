Late Binding: VBScript allows late binding, meaning you can call methods or access properties of an object without explicitly declaring the object type.  This can lead to runtime errors if the object doesn't have the expected member.  Example:
```vbscript
Set objExcel = CreateObject("Excel.Application")
'Assume the user cancels the workbook opening, then objExcel.Workbooks might be Nothing.
Set wb = objExcel.Workbooks.Open("C:\myFile.xlsx")
MsgBox wb.Name
```
If the workbook fails to open, `objExcel.Workbooks` might be `Nothing`, causing an error when accessing `.Open()` and further attempts to access `wb` would also result in an error.

Another example:
```vbscript
Dim obj As Object
Set obj = CreateObject("Some.Object")
' If Some.Object doesn't have a method named DoSomething, this throws an error.
 obj.DoSomething()
```