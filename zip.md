* Zip Package files

```
Sub ZipFile(strZipFilePath, arrFiles)
 
    Dim intloop         As Long
 
    Dim I               As Integer
 
    Dim objApp          As Object
 
    Dim vFileNameZip


    vFileNameZip = strZipFilePath
 
 
 
 
'-------------------Create new empty Zip File-----------------

    If Len(Dir(vFileNameZip)) > 0 Then Kill vFileNameZip
 
    Open vFileNameZip For Output As #1
 
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
 
    Close #1
 
'=============================================================

    Set objApp = CreateObject("Shell.Application")

    I = 0

    For intloop = LBound(arrFiles) To UBound(arrFiles)

        'Copy file to Zip folder/file created above

        I = I + 1

        objApp.Namespace(vFileNameZip).CopyHere CStr(arrFiles(intloop))



        'Wait until Compressing is complete

        On Error Resume Next

        Do Until objApp.Namespace(vFileNameZip).items.Count = I

            Application.Wait (Now + TimeValue("0:00:01"))

        Loop

        On Error GoTo 0

    Next intloop

ExitH:

    Set objApp = Nothing
 
End Sub
```
* Example

```
Sub main()
    Dim arr(1) As String
    arr(0) = "C:\Users\BlackFloat\Desktop\excel\hello.js"
    arr(1) = "C:\Users\BlackFloat\Desktop\excel\hello.exe"
    ZipFile "C:\Users\BlackFloat\Desktop\556.zip", arr
End Sub
```