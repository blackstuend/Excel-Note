# VBA FILESYSTEM FUNCTION

* GetFiles in folder

```
Function GetFiles(FolderPath As String) As String()
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim i As Integer
    Dim f As Variant
    Dim FilesArr() As String
    Dim path As String
    i = 0
    Set oFSO = CreateObject("Scripting.FileSystemObject")
     
    Set oFolder = oFSO.GetFolder(FolderPath)
     
    For Each oFile In oFolder.files
        ReDim Preserve FilesArr(i)
        FilesArr(i) = oFile.Name
        i = i + 1
    Next oFile
    GetFiles = FilesArr
End Function
```

* Get ExtensionName 

```
Function GetExtenFiles(files() As String, ExtensionName As String) As String()
    Dim f As Variant
    Dim fExtensionName As String
    Dim FilesArr() As String
    Dim i As Integer
    Set fs = CreateObject("Scripting.FileSystemObject")
    ExtensionName = LCase(ExtensionName)
    For Each f In files
        fExtensionName = fs.GetExtensionName(f)
        fExtensionName = LCase(fExtensionName)
        If fExtensionName = ExtensionName Then '取得副檔名
            ReDim Preserve FilesArr(i)
            FilesArr(i) = f
            i = i + 1
        End If
    Next f
    GetExtenFiles = FilesArr
End Function
```

* Printf Array 

    * Use to Debug

```
Sub printfArray(files() As String)
    Dim s As Variant
    For Each s In files
        Debug.Print s
    Next s
End Sub
```

* Example

```
Sub main()
    Dim files() As String
    Dim num As Integer
    files = GetFiles("C:\Users\BlackFloat\Desktop\excel")
    files = GetExtenFiles(files, "md")
    printfArray files
End Sub
```