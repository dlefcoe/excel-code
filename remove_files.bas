Attribute VB_Name = "remove_files"
Option Explicit



Sub DeleteFiles()
    Call RecursiveFolder("C:\darren\testDel - Copy")
    Debug.Print "job done !"
End Sub
 
Sub RecursiveFolder(MyPath As String)
 
    Dim FileSys As FileSystemObject
    Dim objFolder As Folder
    Dim objSubFolder As Folder
    Dim objFile As File
 
    On Error Resume Next
    
    Set FileSys = CreateObject("Scripting.FileSystemObject")
    Set objFolder = FileSys.GetFolder(MyPath)
 
    For Each objFile In objFolder.Files
        If Left(objFile.Name, 1) <> "~" And objFile.Name <> ThisWorkbook.Name Then
            objFile.Delete
        End If
    Next objFile
 
    For Each objSubFolder In objFolder.SubFolders
        RecursiveFolder MyPath & "\" & objSubFolder.Name
        objSubFolder.Delete 'Added for deleting the empty folder
    Next objSubFolder
 
    Set FileSys = Nothing
    Set objFolder = Nothing
    Set objSubFolder = Nothing
    Set objFile = Nothing
 
    
End Sub

