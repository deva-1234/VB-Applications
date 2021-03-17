VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "CSV2MDB"
   ClientHeight    =   2385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   9600
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Convert"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   840
      Width           =   6015
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2760
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   ":::"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   0
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Folder Location"
      BeginProperty Font 
         Name            =   "Noto Kufi Arabic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CSV TO MDB"
      BeginProperty Font 
         Name            =   "Noto Kufi Arabic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   240
      Width           =   1680
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Text1.Text = BrowseForFolder(Me.hWnd, "Select Folder")

    If Text1.Text = "" Then
        MsgBox "Please select the folder location", vbCritical
        Exit Sub
    End If

    List1.Clear
    
    Dim F2 As Files
    
    
    Set F2 = fso.GetFolder(Text1.Text).Files
   
    For Each fls In F2
        If LCase(fso.GetExtensionName(fls.Name)) = "csv" Then
            List1.AddItem fls
        End If
    Next

End Sub


Private Sub Command2_Click()

    Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim mdbpath
    Dim TBLNM
    Dim tdefMDB As TableDef, txtFieldone As Field
    Dim dbDatabase As Database
    Dim sNewDBPathAndName As String
    
    For jj = 0 To List1.ListCount - 1
        
        mdbpath = Text1.Text & "\" & fso.GetBaseName(fso.GetFileName(List1.List(jj))) & ".mdb"
        TBLNM = fso.GetBaseName(fso.GetFileName(List1.List(jj)))
           
        sNewDBPathAndName = mdbpath
        Set dbDatabase = CreateDatabase(sNewDBPathAndName, dbLangGeneral, dbEncrypt)
        
        dbDatabase.Close
        
        If con.State = 1 Then con.Close
        
        With con
            .Open "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & mdbpath & ";Uid=Admin;Pwd=;"
            sql1 = "select * into " & TBLNM & " from [text;HDR=yes;Database=" & Text1.Text & "\].[" & TBLNM & ".csv]"
            con.Execute sql1
            .Close
        End With
        
        
    Next
    
    MsgBox "Runned successfully"
End Sub

Private Sub Form_Load()
    Form1.Show
End Sub
