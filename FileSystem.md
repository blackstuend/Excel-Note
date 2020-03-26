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
* Get Same File Name without checkout extensionName
```
Function GetSameName(files() As String, fileName As String) As String()
    Dim f As Variant
    Dim FilesArr() As String
    Dim baseName As String
    Set fs = CreateObject("Scripting.FileSystemObject")
    Dim i As Integer
    fileName = LCase(fileName)
    For Each f In files
        baseName = LCase(fs.GetBaseName(f))
        If baseName = fileName Then '取得副檔名
            ReDim Preserve FilesArr(i)
            FilesArr(i) = f
            i = i + 1
        End If
    Next f
    GetSameName = FilesArr
End Function
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

## FILE operating

* COPY FILE
1. Use normal method
```
    FileCopy source, destination
```
2. Use fs
```
    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.CopyFile source, destination
```

* Move FILE

```
    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.MoveFile  source, destination
```
