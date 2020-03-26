# VBA 文字處理

## Split
* 將字串分離成陣列
* 常用於將文字檔用","或是空白來做分割
* 儲存array宣告時必須加入(),來做為陣列

```
Dim Arr() As String
Arr = Split("hello world", " ")
```

* 將內容輸出可以用for each
```
Sub test()
Dim MyDynArr() As String
MyDynArr = Split("hello world", " ")
For Each element In MyDynArr
    MsgBox (element)
Next element
End Sub

```

## 長度
* 獲得字串長度參數傳入String

```
MsgBox Len("Hello, world.")
```

## 包含

* 類似Js的includes可以查找到字串位於第幾個如果沒有查到則傳回0

````
Dim pos As Integer
pos = InStr("Hello, world.", "world")
MsgBox "pos = " & pos
````
* 從右邊找

```
pos =   InStrRev("Hello,world","world")
```

## 去空白
* 將空白去掉跟js的String.prototyp.trim()一樣的函數

    1. 去掉左右兩邊空白
    ```
    Trim("   hello ")
    ```

    2. 指去掉左邊空白

    ```
    LTrime(" hello")
    ```

    3. 去掉右邊空白

    ```
    RTrime(" hello")
    ```
## 字串取代
* 將字串的文字取代
```
Replace(字串, 搜尋文字, 替換文字[, 起始位置[, 替換次數[, 比對方式]]])
```

## 取出字串
* 類似js的substring
    1. 從開頭取出字串
    ```
    MsgBox Left("Hello, world.", 5) 'Hello
    ```
    2. 從尾端取出字串
    ```
    MsgBox Right("Hello, world.", 6) 'world
    ```
    3. 取出任何位置,較為常用
    ```
    MsgBox Mid("This is a message.", 6, 2) 'is
    ```
## 字串比較
* 也可以直接用=來判斷
* 這種是類似c語言的strcmp可以判斷ascii碼來做比大小
* 預設是分大小寫的,傳入參數vbTextCompare可不分
```
MsgBox StrComp("Hello", "Hello") ' 結果為 0
MsgBox StrComp("Hello", "HELLO") ' 結果為 1
MsgBox StrComp("Hello", "hello") ' 結果為 -1
MsgBox StrComp("Hello", "hello", vbTextCompare) ' 結果為 0
```

## 大小寫轉換

```
MsgBox UCase("Hello, world.") ' HELLO, WORLD.
MsgBox LCase("Hello, world.") ' hello, world.
```
