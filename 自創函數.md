# 為Basic.bas的解釋


## 查詢表頭位於第幾列(Y)
* 由於每個excel格式不同,所以為了避免寫死自創一個找尋主要表頭
* 表頭有的字串丟進去,即可傳回位於第幾列

```
Function findHeaderY(Name As String) As Integer
    Dim x As Integer
    Dim y As Integer
    Dim value As String
    
    For y = 1 To 20
        Dim xEnd As Integer
        xEnd = Cells(y, ActiveSheet.Columns.count).End(xlToLeft).Column
        For x = 1 To xEnd
            value = Cells(y, x)
            value = Trim(value)
            If value = Name Then
                findHeaderY = y
                Exit Function
            End If
        Next x
    Next y
    
End Function

```

## 查詢表頭在第幾行(X)
* 跟查詢表頭在第幾列差不多
* 使用上要先查詢在第幾列(Y),在使用此函數,避免重複查詢浪費時間

```
Function findHeaderX(y As Integer, Name As String) As Integer
    Dim x As Integer
    Dim value As String
    xEnd = Cells(y, ActiveSheet.Columns.count).End(xlToLeft).Column
    For x = 1 To xEnd
        value = Cells(y, x)
        value = Trim(value)
        If value = Name Then
            findHeaderX = x
            Exit Function
        End If
    Next x

End Function
```

## 查詢此列的最後一個位於第幾行(X)
* 傳入在第幾列,則會找出那列的最後一行是多少
```
Function getEndX(y As Integer) As Integer
    Dim xEnd As Integer
    xEnd = Cells(y, ActiveSheet.Columns.count).End(xlToLeft).Column
    getEndX = xEnd
End Function
```

## 查詢此列的最後一個位於第幾行(Y)
* 傳入在第幾行,則會找出那列的最後一列是多少
```
Function getEndY(x As Integer) As Integer
    Dim yEnd As Integer
    yEnd = Cells(ActiveSheet.Rows.count, x).End(xlUp).Row
    getEndY = yEnd
End Function
```