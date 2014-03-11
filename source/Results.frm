VERSION 5.00
Begin VB.Form frmResults 
   Caption         =   "Form2"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3585
   LinkTopic       =   "Form2"
   ScaleHeight     =   6045
   ScaleWidth      =   3585
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstFiles 
      Height          =   3960
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblSpeed 
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   5640
      Width           =   2775
   End
   Begin VB.Label Label8 
      Caption         =   "Speed:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label lblTimeSpent 
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   5280
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "Time:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   5280
      Width           =   390
   End
   Begin VB.Label lblCurrentFileNumber 
      Height          =   195
      Left            =   720
      TabIndex        =   6
      Top             =   4920
      Width           =   2640
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "File#:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   4920
      Width           =   390
   End
   Begin VB.Label lblTotalFileSize 
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "KBs:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   285
   End
   Begin VB.Label lblCurrentFile 
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   4200
      Width           =   3015
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lblStatus_Click()

End Sub

Private Sub lstFiles_Click()
End Sub

Private Sub lstFiles_DblClick()
    Call OpenAnyFile(GetPathWithSlash(frmMain.txtPath.Text) & lstFiles.List(lstFiles.ListIndex))
End Sub
