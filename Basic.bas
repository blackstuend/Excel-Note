Attribute VB_Name = "Module2"
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

Function getEndX(y As Integer) As Integer
    Dim xEnd As Integer
    xEnd = Cells(y, ActiveSheet.Columns.count).End(xlToLeft).Column
    getEndX = xEnd
End Function
Function getEndY(x As Integer) As Integer
    Dim yEnd As Integer
    yEnd = Cells(ActiveSheet.Rows.count, x).End(xlUp).Row
    getEndY = yEnd
End Function

Function sizecount(headerY As Integer) As Integer
    Dim pos As Integer
    Dim count As Integer
    Dim value As String
    Dim x As Integer
    Dim xEnd As Integer
    count = 0
    xEnd = getEndX(headerY)
    For x = 1 To xEnd
        value = Cells(headerY, x).value
        value = Trim(value)
        pos = InStr(value, "size")
        If pos <> 0 Then
            count = count + 1
        End If
    Next x
    sizecount = count
End Function

