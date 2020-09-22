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

12  On Error Resume Next
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
25  tvwDSN.Nodes.Clear

27  iCurrent = 0
28  i = SQL_SUCCESS

'get the DSNs
31  If SQLAllocEnv(lHenv) <> -1 Then
32     i = SQL_SUCCESS
33     Do Until i <> SQL_SUCCESS
34        sDSNItem = Space(1024)
35        sDRVItem = Space(1024)
36        i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
37        sDSN = VBA.Left(sDSNItem, iDSNLen)

39        If sDSN <> Space(iDSNLen) Then

41           Set itmX = tvwDSN.Nodes.Add()
42           itmX.Text = sDSN
43           itmX.Key = sDSN
44           itmX.Image = "Database"
45           Set itmX = Nothing
46        End If
47     Loop
48  End If

50  Exit Sub

End Sub

Private Sub Form_Load()

56  GetDSNs

58  sbStatusBar.Panels(2).Picture = imlToolbarIcons.ListImages("Database").ExtractIcon
59  sbStatusBar.Panels(3).Picture = imlToolbarIcons.ListImages("Table").ExtractIcon

61  LoadSettings

63  Me.Caption = Me.Caption & " " & App.Major & "." & App.Minor & "." & App.Revision


End Sub

Private Sub Form_Unload(Cancel As Integer)
69  Set m_oADOConnect = Nothing
End Sub

Private Sub lvwColumns_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
73  If ColumnHeader.Index <> 1 Then
74     If ColumnHeader.Icon = "SelectAll" Then
75        ColumnHeader.Icon = "Clear"
76     Else
77        ColumnHeader.Icon = "SelectAll"
78     End If
79  ElseIf ColumnHeader.Index = 1 Then
80     If lvwColumns.Tag = "" Or lvwColumns.Tag = "NONE" Then
81        RowSelect True
82        lvwColumns.Tag = "ALL"
83     Else
84        RowSelect False
85        lvwColumns.Tag = "NONE"
86     End If
87  End If



End Sub

Private Sub mnuClearColumns_Click()
94  ColSelect False
End Sub

Private Sub mnuClearRows_Click()
98  RowSelect False
End Sub

Private Sub mnuFileExit_Click()
102     Unload Me
End Sub

Private Sub mnuFileSave_Click()
106     SaveScript
End Sub

Private Sub mnuSelectColumns_Click()
110     ColSelect True
End Sub

Private Sub mnuSelectRows_Click()
114     RowSelect True
End Sub

Private Sub mnuSettingsOnePerLine_Click()
118     mnuSettingsOnePerLine.Checked = Not mnuSettingsOnePerLine.Checked
119     SaveSetting App.ProductName, "Settings", "OneItemPerLine", mnuSettingsOnePerLine.Checked

End Sub

Private Sub mnuSettingsCommit_Click()

Dim sCommit As String
Dim sReg As String

128     sReg = GetSetting(App.ProductName, "Settings", "Commit", "5")

130     sCommit = InputBox("Place commit statement after how many records?", "Commit Statement", sReg)

132     If Len(sCommit) <> 0 Then
133        SaveSetting App.ProductName, "Settings", "Commit", Trim$(Str$(Val(sCommit)))
134     End If

End Sub

Private Sub mnuSettingsDateFormatODBC_Click()

140     mnuSettingsDateFormatOracle.Checked = False
141     mnuSettingsDateFormatSQLServer.Checked = False
142     SaveSetting App.ProductName, "Settings", "DateORACLE", False
143     SaveSetting App.ProductName, "Settings", "DateSQLSERVER", False

145     mnuSettingsDateFormatODBC.Checked = Not mnuSettingsDateFormatODBC.Checked
146     SaveSetting App.ProductName, "Settings", "DateODBC", mnuSettingsDateFormatODBC.Checked

148     If lvwColumns.ListItems.Count > 0 Then
149        FillTable sbStatusBar.Panels(3).Text, Trim$(txtSQL.Text)
150     End If

End Sub

Private Sub mnuSettingsDateFormatOracle_Click()
155     mnuSettingsDateFormatODBC.Checked = False
156     mnuSettingsDateFormatSQLServer.Checked = False
157     SaveSetting App.ProductName, "Settings", "DateODBC", False
158     SaveSetting App.ProductName, "Settings", "DateSQLSERVER", False

160     mnuSettingsDateFormatOracle.Checked = Not mnuSettingsDateFormatOracle.Checked
161     SaveSetting App.ProductName, "Settings", "DateORACLE", mnuSettingsDateFormatOracle.Checked

163     If lvwColumns.ListItems.Count > 0 Then
164        FillTable sbStatusBar.Panels(3).Text, Trim$(txtSQL.Text)
165     End If

End Sub

Private Sub mnuSettingsDateFormatSQLServer_Click()
170     mnuSettingsDateFormatOracle.Checked = False
171     mnuSettingsDateFormatODBC.Checked = False
172     SaveSetting App.ProductName, "Settings", "DateORACLE", False
173     SaveSetting App.ProductName, "Settings", "DateODBC", False

175     mnuSettingsDateFormatSQLServer.Checked = Not mnuSettingsDateFormatSQLServer.Checked
176     SaveSetting App.ProductName, "Settings", "DateSQLSERVER", mnuSettingsDateFormatSQLServer.Checked

178     If lvwColumns.ListItems.Count > 0 Then
179        FillTable sbStatusBar.Panels(3).Text, Trim$(txtSQL.Text)
180     End If


End Sub

Private Sub mnuSettingsDelete_Click()
186     mnuSettingsDelete.Checked = Not mnuSettingsDelete.Checked
187     SaveSetting App.ProductName, "Settings", "DeleteLine", mnuSettingsDelete.Checked
End Sub

Private Sub tlbData_ButtonClick(ByVal Button As MSComctlLib.Button)

192     Select Case Button.Key
           Case "SelectAll"
194           PopupMenu mnuSelect, , tlbData.Left + tlbData.Buttons("SelectAll").Left, tlbData.Top + tlbData.Height
           Case "Clear"
196           PopupMenu mnuClear, , tlbData.Left + tlbData.Buttons("Clear").Left, tlbData.Top + tlbData.Height
           Case "Settings"
198           PopupMenu mnuSettings, , tlbData.Left + tlbData.Buttons("Settings").Left, tlbData.Top + tlbData.Height
           Case "Execute"
200           FillTable sbStatusBar.Panels(3).Text, txtSQL.Text
201     End Select

End Sub
Private Sub ColSelect(pbValue As Boolean)

Dim colX As ColumnHeader

208     For Each colX In lvwColumns.ColumnHeaders
209        If colX.Index <> 1 Then
210           colX.Icon = IIf(pbValue = True, "SelectAll", "Clear")
211        End If
212     Next

214     Set colX = Nothing

End Sub
Private Sub RowSelect(pbValue As Boolean)

Dim i As Long

221     If lvwColumns.ListItems.Count = 0 Then
222        Exit Sub
223     End If

225     ProgressBar.Visible = True
226     ProgressBar.Max = lvwColumns.ListItems.Count
227     For i = 1 To lvwColumns.ListItems.Count
228        ProgressBar.Value = i
229        lvwColumns.ListItems(i).Checked = pbValue
230     Next i
231     ProgressBar.Visible = False

End Sub
Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
235     On Error Resume Next
236     Select Case Button.Key
           Case "Exit"
238           mnuFileExit_Click
           Case "Save"
240           mnuFileSave_Click
241     End Select
End Sub


Private Sub tvwDSN_DblClick()
246     FillTable sbStatusBar.Panels(3).Text, txtSQL.Text
End Sub

Private Sub tvwDSN_NodeClick(ByVal Node As MSComctlLib.Node)

251     On Error GoTo tvwDSN_NodeClick_Error

Dim oLogin As frmODBCLogon
Dim itmX As Node
Dim oTable As adodb.Recordset
Dim i As Long
Dim sConnectString As String

259     If Node.Parent Is Nothing Then
260        sbStatusBar.Panels(1).Text = "ODBC Login"
261        Set oLogin = New frmODBCLogon
262        oLogin.Initialize Node.Text
263        oLogin.Show vbModal, Me
264        sConnectString = oLogin.ConnectString
265        Set oLogin = Nothing

267        If Len(sConnectString) = 0 Then
268           sbStatusBar.Panels(1).Text = "No Connection String"
269           Exit Sub
270        End If

272        If m_oADOConnect Is Nothing Then
273           Set m_oADOConnect = New adodb.Connection
274        End If


277        If m_oADOConnect.State = adStateOpen Then
278           m_oADOConnect.Close
279        End If

281        sbStatusBar.Panels(1).Text = "Executing Connection String"
282        m_oADOConnect.Open sConnectString

284        Set oTable = m_oADOConnect.OpenSchema(adSchemaTables)
285        If oTable Is Nothing Then
286           Exit Sub
287        End If

289        ProgressBar.Max = oTable.RecordCount

291        sbStatusBar.Panels(1).Text = "Loading Tables"
292        Do Until oTable.EOF
293           i = i + 1
294           ProgressBar.Value = i
295           Set itmX = tvwDSN.Nodes.Add(Node.Key, tvwChild)
296           itmX.Image = "Table"
297           itmX.Text = oTable.Fields("TABLE_NAME")
298           itmX.Key = Trim$(Str$(i)) & "-" & Node.Key & "." & itmX.Text
299           oTable.MoveNext
300        Loop
301        ProgressBar.Visible = False
302        Node.Expanded = True
303        Node.EnsureVisible
304        sbStatusBar.Panels(2).Text = Node.Text
305        sbStatusBar.Panels(1).Text = ""
306     Else
307        sbStatusBar.Panels(3).Text = Node.Text
308        lvwColumns.Tag = "ALL"
309        txtSQL.Text = "SELECT * FROM " & Node.Text
310     End If

312     sbStatusBar.Panels(1).Text = ""
313     Exit Sub
314 tvwDSN_NodeClick_Error:
315     MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbCritical
316     ProgressBar.Visible = False
317     lvwColumns.ListItems.Clear
318     lvwColumns.ColumnHeaders.Clear
319     sbStatusBar.Panels(1).Text = ""
End Sub

Private Sub FillTable(psTable As String, psSQL As String)

324     On Error GoTo FillTable_Error

Dim oRS As adodb.Recordset
Dim colX As ColumnHeader
Dim itmX As ListItem
Dim i As Long
Dim j As Long
Dim lRecordCount As Long
Dim sFrom As String
Dim iFrom As Integer
Dim sType As String

336     sType = UCase$(Mid$(psSQL, 1, 6))

338     If sType <> "SELECT" Then
339        If MsgBox("Are you sure you want to execute the " & sType & " action query!", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
340           Exit Sub
341        End If
342     End If

344     Set oRS = m_oADOConnect.Execute(psSQL)
345     iFrom = InStr(1, UCase$(psSQL), "FROM")

347     sFrom = Mid$(psSQL, iFrom + 4)

349     If oRS Is Nothing Then
350        Exit Sub
351     End If

353     lvwColumns.ListItems.Clear
354     lvwColumns.ColumnHeaders.Clear

356     sbStatusBar.Panels(1).Text = "Loading Columns"
357     Set colX = lvwColumns.ColumnHeaders.Add
358     colX.Text = "Select"
359     colX.Width = "1000"

361     ProgressBar.Visible = True
362     ProgressBar.Max = oRS.Fields.Count
363     For i = 0 To oRS.Fields.Count - 1
364        ProgressBar.Value = i
365        Set colX = lvwColumns.ColumnHeaders.Add
366        colX.Text = oRS.Fields(i).Name
367        colX.Icon = "SelectAll"
368     Next i

370     sbStatusBar.Panels(1).Text = "Loading Data..."
371     ProgressBar.Value = 0
372     lRecordCount = m_oADOConnect.Execute("Select Count(*) as RecCount from " & sFrom).Fields("RecCount").Value

374     If lRecordCount <> 0 Then
375        ProgressBar.Max = lRecordCount
376     End If
377     sbStatusBar.Panels(4).Text = "Table Count = " & Format$(Trim$(Str$(lRecordCount)), "#,#")

379     Do Until oRS.EOF
380        j = j + 1
381        ProgressBar.Value = j

383        Set itmX = lvwColumns.ListItems.Add
384        itmX.Text = j
385        itmX.Checked = True
386        For i = 0 To oRS.Fields.Count - 1
387           itmX.SubItems(i + 1) = Trim$("" & SqlValue(GetValue(oRS.Fields(i).Value), oRS.Fields(i).Type))
'DoEvents
389        Next i
390        oRS.MoveNext
391     Loop

393     sbStatusBar.Panels(4).Text = "Record Count = " & Format$(Trim$(Str$(lvwColumns.ListItems.Count)), "#,#")



397     ProgressBar.Visible = False
398     sbStatusBar.Panels(1).Text = ""
399     Exit Sub
400 FillTable_Error:
401     MsgBox Trim$(Str$(Err.Number)) & " - " & Err.Description & vbCrLf & "On line: " & Erl, vbCritical, "Fill Table Error"
402     Set oRS = Nothing
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

420     ProgressBar.Max = lvwColumns.ListItems.Count
421     ProgressBar.Visible = True

423     For i = 1 To lvwColumns.ListItems.Count
424        ProgressBar.Value = i
425        If lvwColumns.ListItems(i).Checked = True Then
426           lSelected = lSelected + 1
427        End If
428     Next i

430     ProgressBar.Visible = False

432     If lSelected = 0 Then
433        If MsgBox("No rows selected, Select All and continue?", vbQuestion + vbYesNo) = vbYes Then
434           mnuSelectRows_Click
435        Else
436           Exit Sub
437        End If
438     End If

440     cdSave.CancelError = True
441     cdSave.Filter = "SQL Script (*.sql)|*.sql"
442     cdSave.FilterIndex = 1
443     cdSave.FileName = "DATA_" & sbStatusBar.Panels(3).Text & "_" & Format$(Now, "mmddyy")
444     cdSave.ShowSave
445     sFileName = cdSave.FileName

447     If Len(sFileName) = 0 Then
448        Exit Sub
449     End If

451     If LCase(Right$(sFileName, 4)) <> ".sql" Then
452        sFileName = sFileName & ".sql"
453     End If

455     Set oFSO = New Scripting.FileSystemObject
456     Set oStream = oFSO.OpenTextFile(sFileName, ForAppending, True)

'Add Delete From ... on first line
459     If mnuSettingsDelete.Checked Then
460        oStream.WriteLine "DELETE FROM " & sbStatusBar.Panels(3).Text & IIf(Me.mnuSettingsDateFormatOracle.Checked, ";", "")
461     End If

463     lCommit = CLng(GetSetting(App.ProductName, "Settings", "Commit", "5"))

'SQL Server
'Begin Transaction
467     If Me.mnuSettingsDateFormatSQLServer.Checked Then
468        bIdentityInsert = MsgBox("Turn IDENTITY INSERT ON", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes

470        If bIdentityInsert Then
471           oStream.WriteLine "SET IDENTITY_INSERT " & sbStatusBar.Panels(3).Text & " ON"
472           oStream.WriteLine "GO"
473        End If

475        oStream.WriteLine "BEGIN TRANSACTION"
476     End If

478     ProgressBar.Max = lvwColumns.ListItems.Count
479     ProgressBar.Visible = True
480     For j = 1 To lvwColumns.ListItems.Count
481        ProgressBar.Value = j
482        sFields = ""
483        sValues = ""
484        If lvwColumns.ListItems(j).Checked = True Then
485           For i = 2 To lvwColumns.ColumnHeaders.Count
486              If lvwColumns.ColumnHeaders.Item(i).Icon = "SelectAll" Then
'Don't create if no data for field
488                 If Len(GetValue(lvwColumns.ListItems(j).SubItems(i - 1))) <> 0 Then
'output one item per line (for debugging and many columns)
490                    If mnuSettingsOnePerLine.Checked = True Then
491                       If Len(sFields) <> 0 Then sFields = sFields & vbCrLf
492                       If Len(sValues) <> 0 Then sValues = sValues & vbCrLf
493                    End If

495                    If Len(sFields) <> 0 Then sFields = sFields & ","
496                    sFields = sFields & lvwColumns.ColumnHeaders(i).Text

'& not allowed in Oracle script
499                    If mnuSettingsDateFormatOracle.Checked Then
500                       sValues = Replace(sValues, "&", "AND")
501                    End If

503                    If Len(sValues) <> 0 Then sValues = sValues & ","
504                    sValues = sValues & lvwColumns.ListItems(j).SubItems(i - 1)

506                 End If
507              End If
508           Next i

510           If mnuSettingsOnePerLine.Checked = True Then
511              sSQL = "INSERT INTO " & sbStatusBar.Panels(3).Text & " (" & vbCrLf & sFields & vbCrLf & ") VALUES (" & vbCrLf & sValues & vbCrLf & ")"
512           Else
513              sSQL = "INSERT INTO " & sbStatusBar.Panels(3).Text & " (" & sFields & ") VALUES (" & sValues & ")"
514           End If

'Oracle add ; at end
517           If Me.mnuSettingsDateFormatOracle.Checked Then
518              sSQL = sSQL & ";"
519           End If

521           oStream.WriteLine sSQL

523           If j Mod lCommit = 0 Then
524              oStream.WriteLine "Commit" & IIf(Me.mnuSettingsDateFormatOracle.Checked, ";", "")

'SQL Server
'Begin Transaction
528              If Me.mnuSettingsDateFormatSQLServer.Checked Then
529                 oStream.WriteLine "BEGIN TRANSACTION"
530              End If

532           End If
533        End If
534     Next j

536     If Me.mnuSettingsDateFormatSQLServer.Checked Then
537        If j Mod lCommit <> 0 Then
538           oStream.WriteLine "Commit" & IIf(Me.mnuSettingsDateFormatOracle.Checked, ";", "")
539        End If

541        If bIdentityInsert Then
542           oStream.WriteLine "SET IDENTITY_INSERT " & sbStatusBar.Panels(3).Text & " OFF"
543           oStream.WriteLine "GO"
544        End If
545     Else
546        oStream.WriteLine "Commit" & IIf(Me.mnuSettingsDateFormatOracle.Checked, ";", "")
547     End If

549     oStream.Close
550     Set oStream = Nothing
551     Set oFSO = Nothing
552     ProgressBar.Visible = False

End Sub

