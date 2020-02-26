
# openfile


```
Dim FilePath As String

' 文字檔案位置
FilePath = "C:\ExcelDemo\demo.txt"

' 開啟 FilePath 文字檔，使用編號 #1 檔案代碼
Open FilePath For Input As #1

' 執行迴圈，直到編號 #1 檔案遇到結尾為止
Do Until EOF(1)

  ' 從編號 #1 檔案讀取一行資料
  Line Input #1, LineFromFile

  ' 輸出一行資料
  MsgBox (LineFromFile)

Loop

' 關閉編號 #1 檔案
Close #1
```