# 正規表達式

* 匯入RegEx
```
    Dim RegEx As Object, MyString As String
    Set RegEx = CreateObject("VBScript.RegExp")
```

* 新增規則

```
    RegEx.Pattern = ""a\w+e"" '比對字串
    regEx.Global = True '設定全域
    regEx.IgnoreCase = True '不分大小寫
```

* 常用函式

```
    RegEx.test(string) 'return true or false
    RegEx.Replace(MyText, MyReplace)
    RegEx.execute(string)
```

* Execute example

```
Sub RegExec()
    Dim RegEx As Object, MyString As String
    Set RegEx = CreateObject("VBScript.RegExp")
    Dim o As Variant
    Dim allMatches As Object
    RegEx.Pattern = "a(.*)le"
    Set allMatches = RegEx.Execute("applle")
    For Each o In allMatches
        Debug.Print o.subMatches.Item(0) 'lle
    Next
End Sub

```