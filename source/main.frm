VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Shital's Text In File Finder"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrSearchSpead 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3000
      Top             =   720
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "Search Now!"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtSearchCriteria 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   3615
   End
   Begin VB.TextBox txtFilter 
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Text            =   "*.*"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Text            =   "C:\Shital\PlanetSource"
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Lable3 
      AutoSize        =   -1  'True
      Caption         =   "Type the search criteria here:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2070
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Filter:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Path:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mlTotalBytesSeached As Long
Dim mlTotSeconds As Long
Dim bIsStart As Boolean

Private Sub btnSearch_Click()

    Dim sNextFile As String
    Dim sFileContent As String
    Dim sPath As String
    Dim oclWordList As New Collection
    Dim lFileCount As Long
    Dim lCurrentFileSize As Long
    
    bIsStart = Not bIsStart
    
    If bIsStart Then
        btnSearch.Caption = "Stop!"
    Else
        btnSearch.Caption = "Next &Search"
    End If
    
    If bIsStart Then
    
        lFileCount = 0
        mlTotalBytesSeached = 0
        mlTotSeconds = 0
        
        sPath = GetPathWithSlash(txtPath.Text)
        sNextFile = Dir$(sPath & txtFilter.Text)
        
        Call GetWordList(txtSearchCriteria.Text, oclWordList)
        
        frmResults.lstFiles.Clear
        tmrSearchSpead.Enabled = True
        frmResults.lblTimeSpent.Caption = "0 Sec"
        frmResults.Show
        Do While (sNextFile <> vbNullString) And bIsStart
            Call LoadAFileAndCheckForWords(sPath & sNextFile, oclWordList, lCurrentFileSize)
            mlTotalBytesSeached = mlTotalBytesSeached + lCurrentFileSize
            lFileCount = lFileCount + 1
            frmResults.lblCurrentFileNumber.Caption = lFileCount
            frmResults.lblTotalFileSize.Caption = mlTotalBytesSeached
            sNextFile = Dir$
        Loop
        tmrSearchSpead.Enabled = False
        
    End If
End Sub

Private Sub GetWordList(ByVal vsInputText As String, ByVal voclOutputCollection As Collection)
    Dim sTextToParse As String
    Dim lWhiteSpaceIndex As Long
    Dim sWord As String
    
    sTextToParse = LTrim$(vsInputText)
    
    Do
        lWhiteSpaceIndex = InStr(1, sTextToParse, " ")
        If lWhiteSpaceIndex <> 0 Then
            sWord = Mid$(sTextToParse, 1, lWhiteSpaceIndex - 1)
            sTextToParse = LTrim$(Mid$(sTextToParse, lWhiteSpaceIndex))
        Else
            sWord = sTextToParse
        End If
        
        Call voclOutputCollection.Add(sWord)
        
    Loop While lWhiteSpaceIndex <> 0
    
End Sub

Private Function LoadAFileAndCheckForWords(ByVal vsFileName As String, ByVal voclWordList As Collection, ByRef rlFileSize As Long) As String
    Dim sFileContent As String
    Dim lFileHandle As Long
    Dim lFileLen As Long
    Dim lWordIndex As Long
    Dim bWordNotFound As Boolean
    
    frmResults.lblCurrentFile.Caption = "Now searching in " & vsFileName
    frmResults.lblCurrentFile.ToolTipText = frmResults.lblCurrentFile.Caption
    
    lFileLen = FileLen(vsFileName)
    rlFileSize = lFileLen
    
    lFileHandle = FreeFile
    
    Open vsFileName For Binary Access Read As #lFileHandle
        sFileContent = Input(lFileLen, #lFileHandle)
        bWordNotFound = False
        For lWordIndex = 1 To voclWordList.Count
            bWordNotFound = (InStr(1, sFileContent, voclWordList(lWordIndex), vbTextCompare) = 0)
            If bWordNotFound Then Exit For
        Next lWordIndex
        If Not bWordNotFound Then
            Call frmResults.lstFiles.AddItem(vsFileName)
        End If
        DoEvents
    Close #lFileHandle
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    bIsStart = False
    DoEvents
    End
End Sub

Private Sub tmrSearchSpead_Timer()
    mlTotSeconds = mlTotSeconds + 1
    With frmResults
        .lblTimeSpent.Caption = mlTotSeconds
        .lblSpeed.Caption = ((mlTotalBytesSeached / mlTotSeconds) \ 1000) & "KB/S"
    End With
End Sub
