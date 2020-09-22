VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ODBC Database Data Scripter"
   ClientHeight    =   8265
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10275
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   10275
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSQL 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   6345
      Width           =   10170
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   7170
      Top             =   330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0994
            Key             =   "SaveBlue"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0EE6
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0FF8
            Key             =   "Database"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1592
            Key             =   "Table"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18E4
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A3E
            Key             =   "Clear"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B98
            Key             =   "Settings"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2132
            Key             =   "ClassEvent"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbData 
      Height          =   360
      Left            =   3465
      TabIndex        =   5
      Top             =   375
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   635
      ButtonWidth     =   1746
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Select"
            Key             =   "SelectAll"
            Object.ToolTipText     =   "Select All"
            ImageKey        =   "SelectAll"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clear"
            Key             =   "Clear"
            Object.ToolTipText     =   "Clear Selected"
            ImageKey        =   "Clear"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Settings"
            Key             =   "Settings"
            Object.ToolTipText     =   "Script Settings"
            ImageKey        =   "Settings"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Execute"
            Key             =   "Execute"
            Object.ToolTipText     =   "Execute SQL or fill entire list from table"
            ImageKey        =   "ClassEvent"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComDlg.CommonDialog cdSave 
      Left            =   4485
      Top             =   1395
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "sql"
      DialogTitle     =   "Save Insert SQL Script"
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   285
      Left            =   3525
      TabIndex        =   2
      Top             =   6015
      Visible         =   0   'False
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   4
      Top             =   7920
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Key             =   "pnlMain"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwDSN 
      Height          =   5610
      Left            =   0
      TabIndex        =   1
      Top             =   375
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   9895
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   226
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "imlToolbarIcons"
      Appearance      =   1
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   635
      ButtonWidth     =   1349
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageKey        =   "Exit"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
            Object.ToolTipText     =   "Save Script"
            ImageKey        =   "Save"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ListView lvwColumns 
      Height          =   5220
      Left            =   3450
      TabIndex        =   3
      Top             =   765
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   9208
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "imlToolbarIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "SQL Statement"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   7
      Top             =   6075
      Width           =   2955
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnuSettingsCommit 
         Caption         =   "Set Commit"
      End
      Begin VB.Menu mnuSettingsDelete 
         Caption         =   "Include DELETE ALL on first line"
      End
      Begin VB.Menu mnuSettingsOnePerLine 
         Caption         =   "One Item Per Line"
      End
      Begin VB.Menu mnuSettingsSpace01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSettingsDateFormat 
         Caption         =   "Database Target Output"
         Begin VB.Menu mnuSettingsDateFormatODBC 
            Caption         =   "&ODBC"
         End
         Begin VB.Menu mnuSettingsDateFormatOracle 
            Caption         =   "O&racle"
         End
         Begin VB.Menu mnuSettingsDateFormatSQLServer 
            Caption         =   "&SQL Server"
         End
      End
   End
   Begin VB.Menu mnuSelect 
      Caption         =   "Select"
      Visible         =   0   'False
      Begin VB.Menu mnuSelectColumns 
         Caption         =   "Columns"
      End
      Begin VB.Menu mnuSelectRows 
         Caption         =   "Rows"
      End
   End
   Begin VB.Menu mnuClear 
      Caption         =   "Clear"
      Visible         =   0   'False
      Begin VB.Menu mnuClearColumns 
         Caption         =   "Columns"
      End
      Begin VB.Menu mnuClearRows 
         Caption         =   "Rows"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer
Private Declare Function SQLAllocEnv% Lib "ODBC32.DLL" (env&)

Const SQL_SUCCESS As Long = 0
Const SQL_FETCH_NEXT As Long = 1

Dim m_oADOConnect As adodb.Connection

Private Sub GetDSNs()

    On Error Resume Next
      Dim i As Integer
      Dim sDSNItem As String * 1024
      Dim sDRVItem As String * 1024
      Dim sDSN As String
      Dim sDRV As String
      Dim iDSNLen As Integer
      Dim iDRVLen As Integer
      Dim lHenv As Long       'handle to the environment
      Dim iCurrent As Integer 'Index to currentItem in CBO
      Dim itmX As Node

        'Initialize
        tvwDSN.Nodes.Clear

        iCurrent = 0
        i = SQL_SUCCESS

        'get the DSNs
        If SQLAllocEnv(lHenv) <> -1 Then
            i = SQL_SUCCESS
            Do Until i <> SQL_SUCCESS
                sDSNItem = Space(1024)
                sDRVItem = Space(1024)
                i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
                sDSN = VBA.Left(sDSNItem, iDSNLen)

                If sDSN <> Space(iDSNLen) Then

                    Set itmX = tvwDSN.Nodes.Add()
                    itmX.Text = sDSN
                    itmX.Key = sDSN
                    itmX.Image = "Database"
                    Set itmX = Nothing
                End If
            Loop
        End If
    
    Exit Sub

End Sub

Private Sub Form_Load()
    
    GetDSNs
    
    sbStatusBar.Panels(2).Picture = imlToolbarIcons.ListImages("Database").ExtractIcon
    sbStatusBar.Panels(3).Picture = imlToolbarIcons.ListImages("Table").ExtractIcon
    
    LoadSettings
    
    Me.Caption = Me.Caption & " " & App.Major & "." & App.Minor & "." & App.Revision
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_oADOConnect = Nothing
End Sub

Private Sub lvwColumns_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ColumnHeader.Index <> 1 Then
        If ColumnHeader.Icon = "SelectAll" Then
            ColumnHeader.Icon = "Clear"
        Else
            ColumnHeader.Icon = "SelectAll"
        End If
    ElseIf ColumnHeader.Index = 1 Then
        If lvwColumns.Tag = "" Or lvwColumns.Tag = "NONE" Then
            RowSelect True
            lvwColumns.Tag = "ALL"
        Else
            RowSelect False
            lvwColumns.Tag = "NONE"
        End If
    End If
    
    
    
End Sub

Private Sub mnuClearColumns_Click()
    ColSelect False
End Sub

Private Sub mnuClearRows_Click()
    RowSelect False
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileSave_Click()
    SaveScript
End Sub

Private Sub mnuSelectColumns_Click()
    ColSelect True
End Sub

Private Sub mnuSelectRows_Click()
    RowSelect True
End Sub

Private Sub mnuSettingsOnePerLine_Click()
    mnuSettingsOnePerLine.Checked = Not mnuSettingsOnePerLine.Checked
    SaveSetting App.ProductName, "Settings", "OneItemPerLine", mnuSettingsOnePerLine.Checked

End Sub

Private Sub mnuSettingsCommit_Click()

Dim sCommit As String
Dim sReg As String

    sReg = GetSetting(App.ProductName, "Settings", "Commit", "5")

    sCommit = InputBox("Place commit statement after how many records?", "Commit Statement", sReg)
    
    If Len(sCommit) <> 0 Then
        SaveSetting App.ProductName, "Settings", "Commit", Trim$(Str$(Val(sCommit)))
    End If
    
End Sub

Private Sub mnuSettingsDateFormatODBC_Click()
    
    mnuSettingsDateFormatOracle.Checked = False
    mnuSettingsDateFormatSQLServer.Checked = False
    SaveSetting App.ProductName, "Settings", "DateORACLE", False
    SaveSetting App.ProductName, "Settings", "DateSQLSERVER", False
    
    mnuSettingsDateFormatODBC.Checked = Not mnuSettingsDateFormatODBC.Checked
    SaveSetting App.ProductName, "Settings", "DateODBC", mnuSettingsDateFormatODBC.Checked
    
    If lvwColumns.ListItems.Count > 0 Then
        FillTable sbStatusBar.Panels(3).Text, Trim$(txtSQL.Text)
    End If

End Sub

Private Sub mnuSettingsDateFormatOracle_Click()
    mnuSettingsDateFormatODBC.Checked = False
    mnuSettingsDateFormatSQLServer.Checked = False
    SaveSetting App.ProductName, "Settings", "DateODBC", False
    SaveSetting App.ProductName, "Settings", "DateSQLSERVER", False
    
    mnuSettingsDateFormatOracle.Checked = Not mnuSettingsDateFormatOracle.Checked
    SaveSetting App.ProductName, "Settings", "DateORACLE", mnuSettingsDateFormatOracle.Checked
    
    If lvwColumns.ListItems.Count > 0 Then
        FillTable sbStatusBar.Panels(3).Text, Trim$(txtSQL.Text)
    End If

End Sub

Private Sub mnuSettingsDateFormatSQLServer_Click()
     mnuSettingsDateFormatOracle.Checked = False
    mnuSettingsDateFormatODBC.Checked = False
    SaveSetting App.ProductName, "Settings", "DateORACLE", False
    SaveSetting App.ProductName, "Settings", "DateODBC", False
    
    mnuSettingsDateFormatSQLServer.Checked = Not mnuSettingsDateFormatSQLServer.Checked
    SaveSetting App.ProductName, "Settings", "DateSQLSERVER", mnuSettingsDateFormatSQLServer.Checked
        
    If lvwColumns.ListItems.Count > 0 Then
        FillTable sbStatusBar.Panels(3).Text, Trim$(txtSQL.Text)
    End If
    
    
End Sub

Private Sub mnuSettingsDelete_Click()
    mnuSettingsDelete.Checked = Not mnuSettingsDelete.Checked
    SaveSetting App.ProductName, "Settings", "DeleteLine", mnuSettingsDelete.Checked
End Sub

Private Sub tlbData_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case "SelectAll"
            PopupMenu mnuSelect, , tlbData.Left + tlbData.Buttons("SelectAll").Left, tlbData.Top + tlbData.Height
        Case "Clear"
            PopupMenu mnuClear, , tlbData.Left + tlbData.Buttons("Clear").Left, tlbData.Top + tlbData.Height
        Case "Settings"
            PopupMenu mnuSettings, , tlbData.Left + tlbData.Buttons("Settings").Left, tlbData.Top + tlbData.Height
        Case "Execute"
            FillTable sbStatusBar.Panels(3).Text, txtSQL.Text
    End Select
    
End Sub
Private Sub ColSelect(pbValue As Boolean)

Dim colX As ColumnHeader

    For Each colX In lvwColumns.ColumnHeaders
        If colX.Index <> 1 Then
            colX.Icon = IIf(pbValue = True, "SelectAll", "Clear")
        End If
    Next
    
    Set colX = Nothing
    
End Sub
Private Sub RowSelect(pbValue As Boolean)

Dim i As Long

    If lvwColumns.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    ProgressBar.Visible = True
    ProgressBar.Max = lvwColumns.ListItems.Count
    For i = 1 To lvwColumns.ListItems.Count
        ProgressBar.Value = i
        lvwColumns.ListItems(i).Checked = pbValue
    Next i
    ProgressBar.Visible = False
    
End Sub
Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Exit"
           mnuFileExit_Click
        Case "Save"
            mnuFileSave_Click
    End Select
End Sub


Private Sub tvwDSN_DblClick()
    FillTable sbStatusBar.Panels(3).Text, txtSQL.Text
End Sub

Private Sub tvwDSN_NodeClick(ByVal Node As MSComctlLib.Node)

On Error GoTo tvwDSN_NodeClick_Error

Dim oLogin As frmODBCLogon
Dim itmX As Node
Dim oTable As adodb.Recordset
Dim i As Long
Dim sConnectString As String

    If Node.Parent Is Nothing Then
        sbStatusBar.Panels(1).Text = "ODBC Login"
         Set oLogin = New frmODBCLogon
         oLogin.Initialize Node.Text
         oLogin.Show vbModal, Me
         sConnectString = oLogin.ConnectString
         Set oLogin = Nothing
         
         If Len(sConnectString) = 0 Then
            sbStatusBar.Panels(1).Text = "No Connection String"
             Exit Sub
         End If
         
         If m_oADOConnect Is Nothing Then
             Set m_oADOConnect = New adodb.Connection
         End If
         
         
        If m_oADOConnect.State = adStateOpen Then
            m_oADOConnect.Close
        End If
        
        sbStatusBar.Panels(1).Text = "Executing Connection String"
        m_oADOConnect.Open sConnectString
         
         Set oTable = m_oADOConnect.OpenSchema(adSchemaTables)
         If oTable Is Nothing Then
             Exit Sub
         End If
         
        ProgressBar.Max = oTable.RecordCount
        
        sbStatusBar.Panels(1).Text = "Loading Tables"
        Do Until oTable.EOF
             i = i + 1
             ProgressBar.Value = i
             Set itmX = tvwDSN.Nodes.Add(Node.Key, tvwChild)
             itmX.Image = "Table"
             itmX.Text = oTable.Fields("TABLE_NAME")
             itmX.Key = Trim$(Str$(i)) & "-" & Node.Key & "." & itmX.Text
             oTable.MoveNext
        Loop
        ProgressBar.Visible = False
        Node.Expanded = True
        Node.EnsureVisible
        sbStatusBar.Panels(2).Text = Node.Text
        sbStatusBar.Panels(1).Text = ""
    Else
        sbStatusBar.Panels(3).Text = Node.Text
        lvwColumns.Tag = "ALL"
        txtSQL.Text = "SELECT * FROM " & Node.Text
    End If

    sbStatusBar.Panels(1).Text = ""
Exit Sub
tvwDSN_NodeClick_Error:
    MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbCritical
    ProgressBar.Visible = False
    lvwColumns.ListItems.Clear
    lvwColumns.ColumnHeaders.Clear
    sbStatusBar.Panels(1).Text = ""
End Sub

Private Sub FillTable(psTable As String, psSQL As String)

On Error GoTo FillTable_Error

Dim oRS As adodb.Recordset
Dim colX As ColumnHeader
Dim itmX As ListItem
Dim i As Long
Dim j As Long
Dim lRecordCount As Long
Dim sFrom As String
Dim iFrom As Integer
Dim sType As String

    sType = UCase$(Mid$(psSQL, 1, 6))
    
    If sType <> "SELECT" Then
        If MsgBox("Are you sure you want to execute the " & sType & " action query!", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
        End If
    End If
    
    Set oRS = m_oADOConnect.Execute(psSQL)
    iFrom = InStr(1, UCase$(psSQL), "FROM")
    
    sFrom = Mid$(psSQL, iFrom + 4)
    
    If oRS Is Nothing Then
        Exit Sub
    End If
    
    lvwColumns.ListItems.Clear
    lvwColumns.ColumnHeaders.Clear
    
    sbStatusBar.Panels(1).Text = "Loading Columns"
    Set colX = lvwColumns.ColumnHeaders.Add
    colX.Text = "Select"
    colX.Width = "1000"
    
    ProgressBar.Visible = True
    ProgressBar.Max = oRS.Fields.Count
    For i = 0 To oRS.Fields.Count - 1
        ProgressBar.Value = i
        Set colX = lvwColumns.ColumnHeaders.Add
        colX.Text = oRS.Fields(i).Name
        colX.Icon = "SelectAll"
    Next i
    
    sbStatusBar.Panels(1).Text = "Loading Data..."
    ProgressBar.Value = 0
    lRecordCount = m_oADOConnect.Execute("Select Count(*) as RecCount from " & sFrom).Fields("RecCount").Value
    
    If lRecordCount <> 0 Then
        ProgressBar.Max = lRecordCount
    End If
    sbStatusBar.Panels(4).Text = "Table Count = " & Format$(Trim$(Str$(lRecordCount)), "#,#")
    
    Do Until oRS.EOF
        j = j + 1
        ProgressBar.Value = j
        
        Set itmX = lvwColumns.ListItems.Add
        itmX.Text = j
        itmX.Checked = True
        For i = 0 To oRS.Fields.Count - 1
            itmX.SubItems(i + 1) = Trim$("" & SqlValue(GetValue(oRS.Fields(i).Value), oRS.Fields(i).Type))
            'DoEvents
        Next i
        oRS.MoveNext
    Loop
    
    sbStatusBar.Panels(4).Text = "Record Count = " & Format$(Trim$(Str$(lvwColumns.ListItems.Count)), "#,#")
    
    
    
    ProgressBar.Visible = False
    sbStatusBar.Panels(1).Text = ""
Exit Sub
FillTable_Error:
    MsgBox Trim$(Str$(Err.Number)) & " - " & Err.Description & vbCrLf & "On line: " & Erl, vbCritical, "Fill Table Error"
    Set oRS = Nothing
End Sub
 
Private Sub SaveScript()

Dim sFileName As String
Dim oStream As Scripting.TextStream
Dim oFSO As Scripting.FileSystemObject
Dim i As Long
Dim j As Long
Dim lSelected As Long
Dim lCommit As Long

Dim sSQL As String
Dim sFields As String
Dim sValues As String
Dim bIdentityInsert As Boolean
Dim bCommit As Boolean

    ProgressBar.Max = lvwColumns.ListItems.Count
    ProgressBar.Visible = True
    
    For i = 1 To lvwColumns.ListItems.Count
        ProgressBar.Value = i
        If lvwColumns.ListItems(i).Checked = True Then
            lSelected = lSelected + 1
        End If
    Next i
    
    ProgressBar.Visible = False
    
    If lSelected = 0 Then
        If MsgBox("No rows selected, Select All and continue?", vbQuestion + vbYesNo) = vbYes Then
            mnuSelectRows_Click
        Else
            Exit Sub
        End If
    End If
    
    cdSave.CancelError = True
    cdSave.Filter = "SQL Script (*.sql)|*.sql"
    cdSave.FilterIndex = 1
    cdSave.FileName = "DATA_" & sbStatusBar.Panels(3).Text & "_" & Format$(Now, "mmddyy")
    cdSave.ShowSave
    sFileName = cdSave.FileName
        
    If Len(sFileName) = 0 Then
        Exit Sub
    End If
        
    If LCase(Right$(sFileName, 4)) <> ".sql" Then
        sFileName = sFileName & ".sql"
    End If
        
    Set oFSO = New Scripting.FileSystemObject
    Set oStream = oFSO.OpenTextFile(sFileName, ForAppending, True)
    
    'Add Delete From ... on first line
    If mnuSettingsDelete.Checked Then
        oStream.WriteLine "DELETE FROM " & sbStatusBar.Panels(3).Text & IIf(Me.mnuSettingsDateFormatOracle.Checked, ";", "")
    End If
    
    lCommit = CLng(GetSetting(App.ProductName, "Settings", "Commit", "5"))
    
    'SQL Server
    'Begin Transaction
    If Me.mnuSettingsDateFormatSQLServer.Checked Then
        bIdentityInsert = MsgBox("Turn IDENTITY INSERT ON", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes
        
        If bIdentityInsert Then
            oStream.WriteLine "SET IDENTITY_INSERT " & sbStatusBar.Panels(3).Text & " ON"
            oStream.WriteLine "GO"
        End If

        oStream.WriteLine "BEGIN TRANSACTION"
    End If
    
    ProgressBar.Max = lvwColumns.ListItems.Count
    ProgressBar.Visible = True
    For j = 1 To lvwColumns.ListItems.Count
        ProgressBar.Value = j
        sFields = ""
        sValues = ""
        bCommit = False
        If lvwColumns.ListItems(j).Checked = True Then
            For i = 2 To lvwColumns.ColumnHeaders.Count
                If lvwColumns.ColumnHeaders.Item(i).Icon = "SelectAll" Then
                    'Don't create if no data for field
                    If Len(GetValue(lvwColumns.ListItems(j).SubItems(i - 1))) <> 0 Then
                        'output one item per line (for debugging and many columns)
                        If mnuSettingsOnePerLine.Checked = True Then
                            If Len(sFields) <> 0 Then sFields = sFields & vbCrLf
                            If Len(sValues) <> 0 Then sValues = sValues & vbCrLf
                        End If
                        
                        If Len(sFields) <> 0 Then sFields = sFields & ","
                        sFields = sFields & lvwColumns.ColumnHeaders(i).Text
                    
                        '& not allowed in Oracle script
                        If mnuSettingsDateFormatOracle.Checked Then
                            sValues = Replace(sValues, "&", "AND")
                        End If
                    
                        If Len(sValues) <> 0 Then sValues = sValues & ","
                        sValues = sValues & lvwColumns.ListItems(j).SubItems(i - 1)
                        
                    End If
                End If
            Next i
            
            If mnuSettingsOnePerLine.Checked = True Then
                sSQL = "INSERT INTO " & sbStatusBar.Panels(3).Text & " (" & vbCrLf & sFields & vbCrLf & ") VALUES (" & vbCrLf & sValues & vbCrLf & ")"
            Else
                sSQL = "INSERT INTO " & sbStatusBar.Panels(3).Text & " (" & sFields & ") VALUES (" & sValues & ")"
            End If
            
            'Oracle add ; at end
            If Me.mnuSettingsDateFormatOracle.Checked Then
                sSQL = sSQL & ";"
            End If
            
            oStream.WriteLine sSQL
            
            If j Mod lCommit = 0 Then
                oStream.WriteLine "Commit" & IIf(Me.mnuSettingsDateFormatOracle.Checked, ";", "")
                bCommit = True
                'SQL Server
                'Begin Transaction
                If Me.mnuSettingsDateFormatSQLServer.Checked Then
                    oStream.WriteLine "BEGIN TRANSACTION"
                End If
    
            End If
        End If
    Next j
    
    If Me.mnuSettingsDateFormatSQLServer.Checked Then
        If (j Mod lCommit <> 0) Or Not bCommit Then
            oStream.WriteLine "Commit" & IIf(Me.mnuSettingsDateFormatOracle.Checked, ";", "")
        End If
        
        If bIdentityInsert Then
            oStream.WriteLine "SET IDENTITY_INSERT " & sbStatusBar.Panels(3).Text & " OFF"
            oStream.WriteLine "GO"
        End If
    Else
        oStream.WriteLine "Commit" & IIf(Me.mnuSettingsDateFormatOracle.Checked, ";", "")
    End If
    
    oStream.Close
    Set oStream = Nothing
    Set oFSO = Nothing
    ProgressBar.Visible = False
    
End Sub

