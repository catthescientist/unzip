VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   11328
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   11328
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6600
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Choose"
      Height          =   252
      Left            =   10320
      TabIndex        =   14
      Top             =   3240
      Width           =   732
   End
   Begin VB.TextBox Text3 
      Height          =   288
      Left            =   9360
      TabIndex        =   13
      Text            =   "C:\temp\11.txt"
      Top             =   3480
      Width           =   1692
   End
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   9360
      TabIndex        =   11
      Text            =   "C:\temp"
      Top             =   1560
      Width           =   1692
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Choose"
      Height          =   252
      Left            =   10320
      TabIndex        =   9
      Top             =   240
      Width           =   732
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   9360
      TabIndex        =   7
      Text            =   "C:\temp\1.zip"
      Top             =   480
      Width           =   1692
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Try AppendToZip vb6"
      Height          =   732
      Left            =   6600
      TabIndex        =   6
      Top             =   3240
      Width           =   2412
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Try DeleteTemp vb6"
      Height          =   732
      Left            =   6600
      TabIndex        =   5
      Top             =   2280
      Width           =   2412
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Test ExtractFromZip vb6"
      Height          =   732
      Left            =   6600
      TabIndex        =   4
      Top             =   1320
      Width           =   2412
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Test ArchiveRead vb6"
      Height          =   852
      Left            =   6600
      TabIndex        =   3
      Top             =   240
      Width           =   2412
   End
   Begin VB.ListBox List1 
      Height          =   6768
      Left            =   240
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   240
      Width           =   6012
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   612
      Left            =   6600
      TabIndex        =   1
      Top             =   5040
      Width           =   2412
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Read textfile from zip"
      Height          =   612
      Left            =   6600
      TabIndex        =   0
      Top             =   4200
      Width           =   2412
   End
   Begin VB.Label Label3 
      Caption         =   "FileName"
      Height          =   252
      Left            =   9360
      TabIndex        =   12
      Top             =   3240
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "DestFolder"
      Height          =   252
      Left            =   9360
      TabIndex        =   10
      Top             =   1320
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "ZipName"
      Height          =   252
      Left            =   9360
      TabIndex        =   8
      Top             =   240
      Width           =   1092
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim unzip As New unzip.unzip
'Dim uz As New unzip
Dim uz As New unzipdll.unzip

Private Sub Command1_Click()
    CommonDialog1.DialogTitle = "Select ZIP"
    CommonDialog1.ShowOpen
    Text1.Text = CommonDialog1.FileName
End Sub

Private Sub Command2_Click()
    CommonDialog1.DialogTitle = "Select Destination path"
    
End Sub

Private Sub Command3_Click()
    CommonDialog1.DialogTitle = "Select file"
    CommonDialog1.ShowOpen
    Text3.Text = CommonDialog1.FileName
End Sub

Private Sub Command4_Click()
Dim ZipFile, DestFolder, readedfile, tempstr As String
Dim nofitem As Integer
Dim deletefile As Boolean
Dim filenum1 As Integer

'ZipFile = "C:\temp\1.zip"
ZipFile = Text1.Text
'DestFolder = "C:\Users\1\Desktop"
DestFolder = Text2.Text
'nofitem = 9
nofitem = -1
For i = 0 To List1.ListCount - 1 Step 1
    If List1.Selected(i) Then nofitem = i
Next i
addfiletotemp = True

If nofitem = -1 Then
    MsgBox "Choose some file from list"
    Exit Sub
End If

readedfile = uz.ExtractFromZip(ZipFile, DestFolder, nofitem, addfiletotemp)

filenum1 = FreeFile
Open readedfile For Input As filenum1

555:

Line Input #filenum1, tempstr

If ((MsgBox(tempstr, vbRetryCancel)) = 4) Then GoTo 555

Close #filenum1
uz.DeleteTemp

End Sub

Private Sub Command5_Click()
Unload Form1
End Sub

Private Sub Command9_Click()
Dim ZipFile As String
Dim filesarray() As String

'ZipFile = "C:\temp\1.zip"
ZipFile = Text1.Text
filesarray = uz.ArchiveRead(ZipFile)

List1.Clear
For Each fff In filesarray
List1.AddItem (fff)
Next

MsgBox ("Choose some files in list")

End Sub

Private Sub Command10_Click()
Dim ZipFile, DestFolder, readedfile As String
Dim nofitem As Integer
Dim addfiletotemp As Boolean

'ZipFile = "C:\temp\1.zip"
ZipFile = Text1.Text
'DestFolder = "C:\Users\1\Desktop"
DestFolder = Text2.Text
'nofitem = 0
addfiletotemp = True

sf = 0
For i = 0 To List1.ListCount - 1 Step 1
    If List1.Selected(i) Then MsgBox (uz.ExtractFromZip(ZipFile, DestFolder, i, addfiletotemp) + " is extracted to " + DestFolder)
Next i

'readedfile = uz.ExtractFromZip(ZipFile, DestFolder, nofitem, addfiletotemp)
End Sub

Private Sub Command11_Click()
uz.DeleteTemp
End Sub

Private Sub Command12_Click()
Dim ZipFile As String
Dim appendfile As String
Dim deletefile As Boolean

'ZipFile = "C:\temp\1.zip"
ZipFile = Text1.Text
'appendfile = "C:\temp\10.txt"
appendfile = Text3.Text
deletefile = False

uz.AppendToZip ZipFile, appendfile, deletefile
End Sub
