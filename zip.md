Sub TestRun()
    Dim arr(2) As String
    arr(1) = "C:\Users\BlackFloat\Desktop\excel\456.txt"
    arr(2) = "C:\Users\BlackFloat\Desktop\excel\zip.md"
    Call ZipFile("C:\Users\BlackFloat\Desktop\excel", "Zipped1", "C:\Users\BlackFloat\Desktop\excel\456.txt", "C:\Users\BlackFloat\Desktop\excel\zip.md", "C:\Users\BlackFloat\Desktop\excel\5566")
    
 
End Sub

Sub ZipFile(strZipFilePath As String, strZipFileName As String, ParamArray arrFiles() As Variant)
 
    Dim intLoop         As Long
 
    Dim I               As Integer
 
    Dim objApp          As Object
 
    Dim vFileNameZip
 
 
 
    If Right(strZipFilePath, 1) <> Application.PathSeparator Then
 
        strZipFilePath = strZipFilePath & Application.PathSeparator
 
    End If
 
    vFileNameZip = strZipFilePath & strZipFileName & ".zip"
 
 
 
    If IsArray(arrFiles) = False Then GoTo ExitH
 
 
 
'-------------------Create new empty Zip File-----------------

    If Len(Dir(vFileNameZip)) > 0 Then Kill vFileNameZip
 
    Open vFileNameZip For Output As #1
 
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
 
    Close #1
 
'=============================================================

    Set objApp = CreateObject("Shell.Application")
 
    I = 0
 
    For intLoop = LBound(arrFiles) To UBound(arrFiles)
 
        'Copy file to Zip folder/file created above

        I = I + 1
 
        objApp.Namespace(vFileNameZip).CopyHere arrFiles(intLoop)
 
 
 
        'Wait until Compressing is complete

        On Error Resume Next
 
        Do Until objApp.Namespace(vFileNameZip).items.Count = I
 
            Application.Wait (Now + TimeValue("0:00:01"))
 
        Loop
 
        On Error GoTo 0
 
    Next intLoop
 
ExitH:
 
    Set objApp = Nothing
 
End Sub
