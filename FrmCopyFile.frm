VERSION 5.00
Begin VB.Form FrmCopyFile 
   Caption         =   "FSO Multipul Functions"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton CmdMove 
      Caption         =   "&Move"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox TxtTarget 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox TxtSource 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton CmdCopy 
      Caption         =   "&Copy"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label LblTarget 
      Caption         =   "Target"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.Label LblSource 
      Caption         =   "Source"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FrmCopyFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Check As Boolean
Private Sub CmdCopy_Click()
    Check = Check_Value()
    If Check = True Then
        Call Copy_File
    End If
End Sub

Private Sub CmdDelete_Click()
    Check = Check_Value()
    If Check = True Then
        Call Delete_File
    End If
End Sub

Private Sub CmdMove_Click()
    Check = Check_Value()
    If Check = True Then
        Call Move_File
    End If
End Sub
