# Learning basic excel vba

* console.log
1. Crtl+G 顯示終端機

```
Debug.Print "hello"
```

* 獲得現在的位置

```
Application.ActiveWorkbook.Path
Application.ActiveWorkbook.FullName
```

* 打開活頁簿

```
Workbooks.Open "test2.xlsx"
```

* 選擇打開的活頁簿

1. 裡面的1為第1個開的,從1開始
```
Workbooks(1).Activate
```


* MsgBox 可以將內容呈現出來
    * 跟Debug.Print 差在
```
MsgBox "hello"
```

* 獲得最後一列

```
Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row
```

* 獲得最後一欄

```
Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column
```

* 獲得資料夾內的檔案

```
Sub LoopThroughFiles ()
 
Dim oFSO As Object
Dim oFolder As Object
Dim oFile As Object
Dim i As Integer
Dim path As String

path = Application.ActiveWorkbook.Path
Set oFSO = CreateObject("Scripting.FileSystemObject")
 
Set oFolder = oFSO.GetFolder(path)
 
For Each oFile In oFolder.Files
 
    Cells(i + 1, 1) = oFile.Name
 
    i = i + 1
 
Next oFile
 
End Sub
```