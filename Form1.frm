VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SQL Server Viewer"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14145
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   14145
   Begin ComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   28
      Top             =   6180
      Width           =   14145
      _ExtentX        =   24950
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   1080
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".sql"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connection"
      Height          =   4215
      Left            =   120
      TabIndex        =   17
      Top             =   240
      Width           =   2415
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   3840
         Width           =   1575
      End
      Begin VB.TextBox txtServerName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Top             =   1080
         Width           =   255
      End
      Begin VB.TextBox txtUserID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "Disconnect"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Server Name or IP Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "SA / No Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "User ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   855
      End
   End
   Begin VB.ComboBox cboDatabases 
      Height          =   315
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   480
      Width           =   2295
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   2760
      TabIndex        =   8
      Top             =   960
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tables/Views/Procedures"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lstProcedures"
      Tab(0).Control(1)=   "TreeView1"
      Tab(0).Control(2)=   "cboProcedures"
      Tab(0).Control(3)=   "cboTables"
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(5)=   "Frame4"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "SQL Query/Results"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "DG"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtSQL"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdExecute"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdClear"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdSave"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdOpenQuery"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.CommandButton cmdOpenQuery 
         Caption         =   "Open Query"
         Height          =   255
         Left            =   9360
         TabIndex        =   29
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save Query"
         Height          =   255
         Left            =   9360
         TabIndex        =   15
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   255
         Left            =   9360
         TabIndex        =   16
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton cmdExecute 
         Caption         =   "Execute"
         Height          =   255
         Left            =   9360
         TabIndex        =   14
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtSQL 
         Height          =   1575
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   480
         Width           =   8895
      End
      Begin MSDataGridLib.DataGrid DG 
         Height          =   2655
         Left            =   240
         TabIndex        =   23
         Top             =   2160
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   4683
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.ListBox lstProcedures 
         Height          =   3180
         Left            =   -69360
         TabIndex        =   12
         Top             =   1320
         Width           =   4455
      End
      Begin ComctlLib.TreeView TreeView1 
         Height          =   3375
         Left            =   -74760
         TabIndex        =   10
         Top             =   1320
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   5953
         _Version        =   327682
         Style           =   7
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cboProcedures 
         Height          =   315
         Left            =   -69360
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   720
         Width           =   4455
      End
      Begin VB.ComboBox cboTables 
         Height          =   315
         Left            =   -74760
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   4455
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tables/Views"
         Height          =   4335
         Left            =   -74880
         TabIndex        =   24
         Top             =   480
         Width           =   4695
         Begin VB.Label Label6 
            Caption         =   "Data Types"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   600
            Width           =   2415
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Procedures"
         Height          =   4335
         Left            =   -69480
         TabIndex        =   25
         Top             =   480
         Width           =   4695
         Begin VB.Label Label5 
            Caption         =   "Procedure Parameters"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   1935
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Databases"
      Height          =   5895
      Left            =   2640
      TabIndex        =   22
      Top             =   240
      Width           =   11415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsrecordset As ADODB.Recordset
Dim rs As ADODB.Recordset

Private Sub cbodatabases_Click()
    ResetData
    strDatabase = cboDatabases.Text
    SSTab1.Enabled = True
    Tables
    Procedures
   Label6.Enabled = True
    Label5.Enabled = True
    Frame4.Enabled = True
    Frame3.Enabled = True
    cboProcedures.Enabled = True
    cboTables.Enabled = True
    SSTab1.Tab = 0
    sb.Panels(2).Text = "Database: " & UCase(cboDatabases.Text)
     sb.Panels(3).Text = "Table/View: N/A"
     sb.Panels(4).Text = "Procedure: N/A"
     sb.Panels(5).Text = "SQL: N/A"
End Sub
Private Sub ResetData()
    On Error Resume Next
    cboTables.Clear
    cboProcedures.Clear
    TreeView1.Nodes.Clear
    lstProcedures.Clear
    txtSQL.Text = ""
    rs.Close
    DG.Refresh
End Sub
Private Sub Procedures()
    Set rsrecordset = cn.OpenSchema(adSchemaProcedures)
    Do Until rsrecordset.EOF
       
            cboProcedures.AddItem rsrecordset!procedure_name
            
    rsrecordset.MoveNext
    Loop
End Sub



Private Sub cboProcedures_Click()
Dim i
lstProcedures.Clear
i = Mid(cboProcedures.Text, 1, InStr(cboProcedures.Text, ";") - 1)
   Set rsrecordset = cn.Execute("sp_procedure_params_rowset" & " " & i)
   Do Until rsrecordset.EOF
        
           lstProcedures.AddItem rsrecordset!parameter_name
     
        rsrecordset.MoveNext
    Loop
  sb.Panels(4).Text = "Procedure: " & UCase(cboProcedures.Text)
End Sub

Private Sub cboTables_Click()
     Dim Start As Node
    Dim middle As Node
    Dim ends As Node
    Dim sqlstmt As String
    Dim i As Integer
    Dim nulls As String
    Dim sqlstmtt As String
    Dim cnt As Integer
   
    
    TreeView1.Nodes.Clear
  '  CboFields.Clear 'Clear the list.
  
    'Get the name of the table selected.
  
    Set rsrecordset = New ADODB.Recordset
    Set rsrecordset = _
        cn.Execute("sp_columns [" & cboTables.Text & "]")
        
    Set Start = TreeView1.Nodes.Add(, 0, rsrecordset.Index, "Table: " & rsrecordset!table_name, 0, 0)
    Do Until rsrecordset.EOF
        cnt = cnt + 1

        Set middle = TreeView1.Nodes.Add(Start, 4, rsrecordset.Index, rsrecordset!column_name, 0, 0)
        middle.EnsureVisible
        If rsrecordset!nullable = "0" Then
            nulls = "Not Null"
        Else
            nulls = "Null"
        End If
        Set ends = TreeView1.Nodes.Add(middle, 4, rsrecordset.Index, rsrecordset!type_name & " (" & rsrecordset!length & "," & " " & nulls & ")", 0, 0)
        rsrecordset.MoveNext
     Loop
     sb.Panels(3).Text = "Table/View: " & UCase(cboTables.Text)
End Sub

Private Sub Check1_Click()
    If Click = False Then
        txtUserID.Enabled = False
        txtPassword.Enabled = False
        Label3.Enabled = False
        Label4.Enabled = False
        Click = True
    Else
        txtUserID.Enabled = True
        txtPassword.Enabled = True
          Label3.Enabled = True
        Label4.Enabled = True
        Click = False
    End If
End Sub

Private Sub cmdClear_Click()
    On Error Resume Next
    txtSQL.Text = ""
    rs.Close
    DG.Refresh
    sb.Panels(5).Text = "SQL: N/A"
End Sub

Private Sub cmdConnect_Click()
sb.Panels(1).Text = "Connecting to " & txtServerName.Text
    If txtServerName.Text = "" Then
        Call MsgBox("Server Name is Needed.", vbOKOnly, "Server Name")
        Exit Sub
    End If
    strServer = txtServerName.Text
    If Click = True Then
        strUID = "sa"
        strPWD = ""
    Else
        strUID = txtUserID.Text
        strPWD = txtPassword.Text
    End If
    Open_cn
    If cnOpen = True Then
        Databases
        cboDatabases.Enabled = True
        Close_cn
        cmdDisconnect.Enabled = True
        cmdConnect.Enabled = False
        txtServerName.Enabled = False
        txtUserID.Enabled = False
        txtPassword.Enabled = False
        Check1.Enabled = False
         sb.Panels(1).Text = "Connected to " & UCase(strServer)
    Else
        Exit Sub
    End If
    
End Sub


Private Sub cmdDisconnect_Click()
    Close_cn
    cmdConnect.Enabled = True
    cboDatabases.Clear
    ResetData
    SSTab1.Enabled = False
    cmdDisconnect.Enabled = False
    cboDatabases.Enabled = False
        txtServerName.Enabled = True
        txtUserID.Enabled = True
        txtPassword.Enabled = True
        Check1.Enabled = True
        SSTab1.Tab = 0
        Label6.Enabled = False
    Label5.Enabled = False
    Frame4.Enabled = False
    Frame3.Enabled = False
    cboProcedures.Enabled = False
    cboTables.Enabled = False
    txtUserID.Text = ""
    txtPassword.Text = ""
    txtServerName.Text = ""
       sb.Panels(1).Text = "Not Connected"
     sb.Panels(2).Text = "Database: N/A"
     sb.Panels(3).Text = "Table/View: N/A"
     sb.Panels(4).Text = "Procedure: N/A"
     sb.Panels(5).Text = "SQL: N/A"
     strServer = ""
     strPWD = ""
     strUID = ""
End Sub

Private Sub cmdExecute_Click()
On Error GoTo errhandler
Dim nulls As String
Dim sqlstmt
sb.Panels(5).Text = "SQL: Querying"
  sqlstmt = txtSQL.Text
   DGError = True
 Set rs = New ADODB.Recordset
    rs.Open sqlstmt, cn, adOpenStatic, adLockOptimistic, _
        adCmdText
    Set DG.DataSource = rs
  sb.Panels(5).Text = "SQL: Results"
Exit Sub
errhandler:
    Call MsgBox("Error in SQL Statment", vbOKOnly, "Error")
    sb.Panels(5).Text = "SQL:Error"
    Exit Sub
End Sub


Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdOpenQuery_Click()
On Error GoTo errhandler
    sb.Panels(5).Text = "SQL:Open"
    Dim filename As String
     CD.Filter = "SQL Files (*.sql)|*.sql"
    CD.FilterIndex = 2

    CD.ShowOpen
  

    filename = CD.filename
    Dim LinesFromFile, NextLine As String
    Dim myarray() As String
    Dim i
Open filename For Input As #1

Do Until EOF(1)
   Line Input #1, NextLine
   LinesFromFile = LinesFromFile + NextLine + Chr(13) + Chr(10)
Loop
Close #1
    txtSQL.Text = LinesFromFile
   filename = ""
   
    rs.Close
    DG.Refresh
    sb.Panels(5).Text = "SQL: N/A"
    Exit Sub
errhandler:
Exit Sub
End Sub


Private Sub cmdSave_Click()
On Error GoTo errhandler
sb.Panels(5).Text = "SQL:Saving"
If txtSQL.Text <> "" Then
CD.Filter = "SQL Files (*.sql)|*.sql"
    CD.FilterIndex = 2
    CD.ShowSave
    Open CD.filename For Output As #1    ' Open/Create file
    Print #1, txtSQL.Text
    Close #1
Else
    Call MsgBox("No Data To Save.", vbOKOnly, "No Save")
    Exit Sub
End If
sb.Panels(5).Text = "SQL: N/A"
Exit Sub
errhandler:
Exit Sub
End Sub

Private Sub Form_Load()
Click = False
    cmdDisconnect.Enabled = False
    strDatabase = ""
    SSTab1.Enabled = False
    cboDatabases.Enabled = False
    SSTab1.Tab = 0
    Label6.Enabled = False
    Label5.Enabled = False
    Frame4.Enabled = False
    Frame3.Enabled = False
    cboProcedures.Enabled = False
    cboTables.Enabled = False
    sb.Panels(1).Text = "Not Connected"
     sb.Panels(2).Text = "Database: N/A"
     sb.Panels(3).Text = "Table/View: N/A"
     sb.Panels(4).Text = "Procedure: N/A"
     sb.Panels(5).Text = "SQL: N/A"
End Sub
Private Sub Databases()
    Set rsrecordset = New ADODB.Recordset
    Set rsrecordset = cn.Execute("sp_databases")
     Do Until rsrecordset.EOF
        cboDatabases.AddItem (rsrecordset.Fields("Database_Name"))
        rsrecordset.MoveNext
    Loop
End Sub
Private Sub Tables()
Open_cn
    Set rsrecordset = cn.OpenSchema(adSchemaTables)
    Do Until rsrecordset.EOF
        If UCase(Left(rsrecordset!table_name, 4)) <> "MSYS" Then
            cboTables.AddItem rsrecordset!table_name
        End If
        rsrecordset.MoveNext
    Loop

End Sub

