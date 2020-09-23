VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTrace 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Object Viewer"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9525
   Icon            =   "frmTrace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   9525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00C0E0FF&
      Height          =   735
      Left            =   6360
      Picture         =   "frmTrace.frx":0742
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6720
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5640
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrace.frx":0B84
            Key             =   "car"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrace.frx":0FD6
            Key             =   "sine"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   4200
      Top             =   3840
   End
   Begin MSComctlLib.ListView lvwFields 
      Height          =   5055
      Left            =   2640
      TabIndex        =   1
      Top             =   1440
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8916
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView tvwTables 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   8916
      _Version        =   393217
      HideSelection   =   0   'False
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Label Label6 
      Caption         =   "Table Contents"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Tables"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Select table from the left hand side to view the corresponding fields and records."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   6960
      Width           =   5535
   End
   Begin VB.Label Label2 
      Caption         =   "Processing...Please Wait"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   7560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   9255
   End
   Begin VB.Label Label1 
      Caption         =   "Access Database Object Viewer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private db As Database
Private rs   As Recordset
Dim lstx As ListItem
Private Sub cmdSearch_Click()
    Unload Me
    frmInit.Show
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub
Private Sub File1_Click()
    cmdOk.Enabled = True
End Sub

Private Sub Command1_Click()
    End
End Sub

Private Sub Form_Load()
    Dim tables As TableDefs
    Dim tvwnode As Node
    
    tvwTables.Indentation = 200
    lvwFields.View = lvwReport

    Label3.Caption = FileName
    tvwTables.Nodes.Clear
    
    'Set tvwnode = tvwTables.Nodes.Add(, , "m", "ACCESS ENGIN")
    'tvwnode.Image = "car"
    lvwFields.ColumnHeaders.Clear
    
    Set db = OpenDatabase(FileName) 'shankar

    For i = 0 To db.TableDefs.Count - 1
        If db.TableDefs(i).Attributes = 0 Then
            Set tvwnode = tvwTables.Nodes.Add(, , "TABLE" & Trim(Str(i + 1)), db.TableDefs(i).Name)
            tvwnode.Image = "car"
        End If
    Next i
    'The first item will always be selected in the tree view
    tvwTables.Nodes(1).Selected = True 'shankar
    'The node click event will be fired so that all records will be displayed automatically
    'in the list view.
    Call tvwTables_NodeClick(tvwTables.Nodes.Item(1))
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = True
        MsgBox "Click Exit Button to Quit", vbOKOnly + vbInformation, "Exit"
    Else
        calcel = False
    End If
End Sub

Private Sub Timer1_Timer()
    If Label2.Visible = False Then
        Label2.Visible = True
    Else
        Label2.Visible = False
    End If
End Sub

Private Sub tvwTables_NodeClick(ByVal Node As MSComctlLib.Node)
    lvwFields.ListItems.Clear
    lvwFields.ColumnHeaders.Clear
    frmTrace.Refresh

    For i = 0 To db.TableDefs(Node.Text).Fields.Count - 1
        lvwFields.ColumnHeaders.Add , , db.TableDefs(Node.Text).Fields(i).Name
    Next i
    
    Set rs = db.OpenRecordset(Node.Text)
 
        Do Until rs.EOF
            Set lstx = lvwFields.ListItems.Add(, , rs.Fields(0), , "sine")
                For i = 1 To db.TableDefs(Node.Text).Fields.Count - 1
                        lstx.SubItems(i) = Trim(rs.Fields(i) & "")
                        Me.MousePointer = vbHourglass
                Next i
             rs.MoveNext
        Loop
        Me.MousePointer = vbNormal
End Sub
