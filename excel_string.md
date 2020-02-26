# Split

```
Sub test()
Dim MyDynArr() As String
MyDynArr = Split("hello world", " ")
For Each element In MyDynArr
    MsgBox (element)
Next element
End Sub

```

# Len

```
MsgBox Len("Hello, world.")
```

# includes

````
Dim pos As Integer
pos = InStr("Hello, world.", "world")
MsgBox "pos = " & pos
````