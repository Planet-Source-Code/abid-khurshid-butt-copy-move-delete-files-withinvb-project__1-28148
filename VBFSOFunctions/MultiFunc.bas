Attribute VB_Name = "MultiFunc"
Public Function Check_Value() As Boolean
    If Len(Trim(FrmCopyFile.TxtSource.Text)) = 0 Then
        MsgBox "Source has not been defined? ", vbCritical + vbOKOnly, "First Select Source File"
        FrmCopyFile.TxtSource.SetFocus
        Check_Value = False
        Exit Function
    End If
    If Len(Trim(FrmCopyFile.TxtTarget.Text)) = 0 Then
        MsgBox "Target has not been defined? ", vbCritical + vbOKOnly, "First Select Source File"
        FrmCopyFile.TxtSource.SetFocus
        Check_Value = False
        Exit Function
    End If
    Check_Value = True
End Function
Public Sub Copy_File()
    Dim fso As New FileSystemObject, DbFile
    ' Get a handle to the file in root of C:\.
    Set DbFile = fso.GetFile(FrmCopyFile.TxtSource.Text)
    ' Copy the file to \temp.
    DbFile.Copy (FrmCopyFile.TxtTarget.Text)
    MsgBox "File " & FrmCopyFile.TxtSource.Text & " Copied to " & FrmCopyFile.TxtTarget.Text, vbInformation + vbOKOnly
    fso = Nothing
    DbFile = Nothing
End Sub
Public Sub Move_File()
    Dim fso As New FileSystemObject, DbFile
    ' Get a handle to the file in root of C:\.
    Set DbFile = fso.GetFile(FrmCopyFile.TxtSource.Text)
    ' Move the file to your targated directory
    DbFile.Move (FrmCopyFile.TxtTarget.Text)
    MsgBox "File " & FrmCopyFile.TxtSource.Text & " Moved to " & FrmCopyFile.TxtTarget.Text, vbInformation + vbOKOnly
    fso = Nothing
    DbFile = Nothing
End Sub
Public Sub Delete_File()
    Dim fso As New FileSystemObject, DbFile
    ' Get a handle to the file in root of C:\.
    Set DbFile = fso.GetFile(FrmCopyFile.TxtSource.Text)
    ' Delete the file from your targated directory
    DbFile.Move (FrmCopyFile.TxtTarget.Text)
    MsgBox "File " & FrmCopyFile.TxtSource.Text & " Deleted Successfully! ", vbInformation + vbOKOnly
    fso = Nothing
    DbFile = Nothing
End Sub

