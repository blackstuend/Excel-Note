# Function 傳入excel 來使用

```
Public Function good(rng As Range)
    Dim all As Integer
    all = 0
    For Each cell In rng
    all = cell.value + all
    Next cell
    MsgBox (all)
    good = all
End Function
```