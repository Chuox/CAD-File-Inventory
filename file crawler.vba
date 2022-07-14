Private Sub CommandButton1_Click()
    Dim MyFSO As FileSystemObject
    Dim MyFile As File
    Dim MyFolder As Folder
    Dim MySubFolder As Folder
    Dim i As Integer
    Path = Cells(3, 10)
    
    'checking if path exist
    Set MyFSO = New FileSystemObject
    If MyFSO.FolderExists(Path) Then
        'MsgBox "The Folder Exists"
    Else
        MsgBox "The Folder Does Not Exist"
        Exit Sub
    End If
    'confirm overwrite of cells content
    answer = MsgBox("This action will overwrite the contents of the cells." & vbNewLine & "Want to Continue?", vbOKCancel + vbExclamation)
    If answer = 2 Then
        Exit Sub
    End If
    
    'check for subfolders
    Set MyFSO = New Scripting.FileSystemObject
    Set MyFolder = MyFSO.GetFolder(Path)
    i = 2
    For Each MySubFolder In MyFolder.SubFolders
        For Each MyFile In MySubFolder.Files
            If MyFSO.GetExtensionName(MyFile.Path) = "dwl" Or MyFSO.GetExtensionName(MyFile.Path) = "dwl2" _
            Or MyFSO.GetExtensionName(MyFile.Path) = "bak" _
            Or MyFSO.GetExtensionName(MyFile.Path) = "adt" _
            Or MyFSO.GetExtensionName(MyFile.Path) = "ds$" _
            Or MyFSO.GetExtensionName(MyFile.Path) = "err" _
            Or MyFSO.GetExtensionName(MyFile.Path) = "log" Then
                'do nothing, those extensions arent inventoried
            Else
                Cells(i, 1) = MySubFolder.Name
                Cells(i, 2) = MyFile.Name
                
                If Left(MySubFolder.Name, 3) = "C3D" And MyFSO.GetExtensionName(MyFile.Path) = "dwg" Then
                    Cells(i, 5) = "DREF"
                    Cells(i, 6) = "1"
                ElseIf Left(MySubFolder.Name, 4) = "XREF" And MyFSO.GetExtensionName(MyFile.Path) = "dwg" Then
                    Cells(i, 5) = "DREF"
                    Cells(i, 6) = "2"
                ElseIf Left(MySubFolder.Name, 1) = "_" And MyFSO.GetExtensionName(MyFile.Path) = "dwg" Then
                    Cells(i, 5) = "PROD"
                    Cells(i, 6) = "3"
                End If
                i = i + 1
            End If
        Next MyFile
    Next MySubFolder
End Sub


Sub DoFolder(Folder)
    Dim SubFolder
    For Each SubFolder In Folder.SubFolders
        DoFolder SubFolder
    Next
    Dim File
    i = 2
    For Each File In Folder.Files
        ' Operate on each file
    Next
End Sub
