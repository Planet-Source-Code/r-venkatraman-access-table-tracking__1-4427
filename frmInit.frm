VERSION 5.00
Begin VB.Form frmInit 
   Caption         =   "Scan for .mdb files"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8895
   Icon            =   "frmInit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   5400
      TabIndex        =   2
      Top             =   3000
      Width           =   2775
   End
   Begin VB.CommandButton cmdScan 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Scan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      Picture         =   "frmInit.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Select Drive and click on the scan button to start scanning for Access Database Files."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   5280
      Width           =   6855
   End
End
Attribute VB_Name = "frmInit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function ShellToBrowser%(frm As Form, ByVal URL$, ByVal WindowStyle%)
         
Dim api%
api% = ShellExecute(frm.hWnd, "open", URL$, "", App.Path, WindowStyle%)
      
'Check return value
If api% < 31 Then
   MsgBox App.Title & " had a problem running your web browser.", 48, "Browser Unavailable"
   ShellToBrowser% = False
ElseIf api% = 32 Then
   'no file association
   MsgBox App.Title & " could not find a file association", 48, "Browser Unavailable"
   ShellToBrowser% = False
Else
   'It worked!
   ShellToBrowser% = True
End If
End Function
Private Sub FindIt(sStartPath As String, sPattern As String)
    Dim fso As FileSystemObject
    Dim fld As Folder
    Dim fldCurrent As Folder
    Dim sFile As String
    
    Screen.MousePointer = vbHourglass
    
    Set fso = New FileSystemObject
        If Right$(sStartPath, 1) <> "\" Then
            sStartPath = sStartPath & "\"
        End If
        
    Set fldCurrent = fso.GetFolder(sStartPath)
    sFile = Dir(sStartPath & sPattern, vbNormal)
    
    Do While Len(sFile) > 0
        List1.AddItem sStartPath & sFile
        sFile = Dir
    Loop
        
    If fldCurrent.SubFolders.Count > 0 Then
        For Each fld In fldCurrent.SubFolders
            FindIt sStartPath & fld.Name, sPattern
        Next
    End If

    Screen.MousePointer = vbNormal
    
End Sub
Private Sub cmdScan_Click()
    cmdScan.Enabled = False
    List1.Clear
    Call FindIt(Drive1.Drive & "\", "*.mdb")
    cmdScan.Enabled = True
End Sub

Private Sub lblEmail_Click()
    Site = "mailto:venkatraman_r@hotmail.com"
    success% = ShellToBrowser(Me, Site, 0)
End Sub

Private Sub List1_DblClick()
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = True Then
            FileName = List1.List(i)
        End If
    Next i
        frmInit.Hide
        frmTrace.Show
End Sub

