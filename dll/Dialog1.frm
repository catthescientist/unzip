VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   6576
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6576
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   612
      Left            =   3600
      TabIndex        =   2
      Top             =   4200
      Width           =   2652
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   612
      Left            =   240
      TabIndex        =   1
      Top             =   4200
      Width           =   2892
   End
   Begin VB.ListBox List1 
      Height          =   3888
      Left            =   240
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   120
      Width           =   6012
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Me.Caption = True
    Unload Form1
End Sub

Private Sub Command2_Click()
    Me.Caption = False
    Unload Form1
End Sub
