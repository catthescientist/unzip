VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "unzip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim FilesInZip() As String
Dim TempFiles As New Collection
Dim sh As New Shell32.Shell
Dim ZipFile As Shell32.Folder
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function ArchiveRead(ByVal ZipName As String) As String()
    If Dir(ZipName) = "" Then
        MsgBox ("File " + ZipName + " is not exists")
        Exit Function
    End If
    
    If (Right(ZipName, 3) <> "zip") And (Right(ZipName, 3) <> "Zip") And (Right(ZipName, 3) <> "ZIP") Then
        MsgBox ("File " + ZipName + " is not ZIP")
        Exit Function
    End If
    
    Dim j As Integer
    j = 0
    
On Error GoTo ErrMessage
    ReDim FilesInZip(j)
    Set ZipFile = sh.NameSpace(ZipName)
    
    For Each jjj In ZipFile.Items
        ReDim Preserve FilesInZip(j)
        FilesInZip(j) = jjj.Path
        'MsgBox FilesInZip(j)
        j = j + 1
    Next
    
    ArchiveRead = FilesInZip
    Exit Function
            
ErrMessage:
    MsgBox "Error with shell32.dll"
    'ArchiveRead = Nothing
    Exit Function
        
End Function

Public Function ExtractFromZip(ByVal ZipName As String, ByVal DestFolder As String, ByVal nofitem As Integer, ByVal movetotrash As Boolean) As String
    If Dir(ZipName) = "" Then
        MsgBox ("File " + ZipName + " is not exists")
        ExtractFromZip = ""
        Exit Function
    End If
    
    If (Right(ZipName, 3) <> "zip") And (Right(ZipName, 3) <> "Zip") And (Right(ZipName, 3) <> "ZIP") Then
        MsgBox ("File " + ZipName + " is not ZIP")
        ExtractFromZip = ""
        Exit Function
    End If
    
On Error GoTo ErrMessage
    Set ZipFile = sh.NameSpace(ZipName)
    Dim OutputFolder As Shell32.Folder
    Set OutputFolder = sh.NameSpace(DestFolder)
    Call OutputFolder.CopyHere(ZipFile.Items.Item(nofitem))
    If movetotrash Then TempFiles.Add (DestFolder + "\" + ZipFile.Items.Item(nofitem).Name)
    ExtractFromZip = (DestFolder + "\" + ZipFile.Items.Item(nofitem).Name)
    Exit Function
    
ErrMessage:
    MsgBox "Error with shell32.dll"
    ExtractFromZip = ""
    Exit Function

End Function

Public Sub DeleteTemp()
On Error GoTo ErrMessage
    For Each f In TempFiles
        Kill f
    Next
    Set TempFiles = New Collection
    Exit Sub
    
ErrMessage:
    MsgBox "Error with killFile procedure"
    
End Sub

Public Sub AppendToZip(ByVal ZipName As String, ByVal FileName As String, ByVal deletefile As Boolean)
    If Dir(FileName) = "" Then
        MsgBox ("File " + FileName + " is not exists")
        Exit Sub
    End If
    
    If (Right(ZipName, 3) <> "zip") And (Right(ZipName, 3) <> "Zip") And (Right(ZipName, 3) <> "ZIP") Then
        MsgBox ("File " + ZipName + " is not ZIP")
        Exit Sub
    End If
    
    If Dir(ZipName) = "" Then
        Open "C:\temp\2.zip" For Output As #1
        Close #1
    End If
    
    Dim zt As Double
    Dim whilestop As Byte
                
On Error GoTo ErrMessage
    Set ZipFile = sh.NameSpace(ZipName)
    zc = ZipFile.Items.Count
    zt = FileSystem.FileDateTime(ZipName)
    
    If deletefile Then
        ZipFile.MoveHere (FileName)
    Else
        ZipFile.CopyHere (FileName)
    End If
    
    whilestop = 0
    While (FileSystem.FileDateTime(ZipName) - zt) = 0 And (whilestop < 100)
        Call Sleep(50)
        whilestop = whilestop + 1
    Wend
    
    Exit Sub
    
ErrMessage:
    MsgBox "Error with shell32.dll"
    Exit Sub
    
End Sub
