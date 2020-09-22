VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ODBC Database Data Scripter"
   ClientHeight    =   6390
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10275
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   10275
   StartUpPosition =   2  'CenterScreen
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
         NumListImages   =   8
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
         NumButtons      =   3
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
      Left            =   7440
      TabIndex        =   2
      Top             =   6090
      Visible         =   0   'False
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   4
      Top             =   6045
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
      Begin VB.Menu mnuSettingsSpace01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSettingsDateFormat 
         Caption         =   "Date Format Output"
         Begin VB.Menu mnuSettingsDateFormatODBC 
            Caption         =   "&ODBC"
         End
         Begin VB.Menu mnuSettingsDateFormatOracle 
            Caption         =   "O&racle"
         End
         Begin VB.Menu mnuSettingsDateFormatSQLServer 
            Caption         =   "&SQL Server/Access"
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

Private Sub mnuSettingsCommit_Click()

Dim sCommit As String
Dim sReg As String

122     sReg = GetSetting(App.ProductName, "Settings", "Commit", "5")

124     sCommit = InputBox("Place commit statement after how many records?", "Commit Statement", sReg)

126     If Len(sCommit) <> 0 Then
127        SaveSetting App.ProductName, "Settings", "Commit", Trim$(Str$(Val(sCommit)))
128     End If

End Sub

Private Sub mnuSettingsDateFormatODBC_Click()

134     mnuSettingsDateFormatOracle.Checked = False
135     mnuSettingsDateFormatSQLServer.Checked = False
136     SaveSetting App.ProductName, "Settings", "DateORACLE", False
137     SaveSetting App.ProductName, "Settings", "DateSQLSERVER", False

139     mnuSettingsDateFormatODBC.Checked = Not mnuSettingsDateFormatODBC.Checked
140     SaveSetting App.ProductName, "Settings", "DateODBC", mnuSettingsDateFormatODBC.Checked

142     If lvwColumns.ListItems.Count > 0 Then
143        FillTable sbStatusBar.Panels(3).Text
144     End If

End Sub

Private Sub mnuSettingsDateFormatOracle_Click()
149     mnuSettingsDateFormatODBC.Checked = False
150     mnuSettingsDateFormatSQLServer.Checked = False
151     SaveSetting App.ProductName, "Settings", "DateODBC", False
152     SaveSetting App.ProductName, "Settings", "DateSQLSERVER", False

154     mnuSettingsDateFormatOracle.Checked = Not mnuSettingsDateFormatOracle.Checked
155     SaveSetting App.ProductName, "Settings", "DateORACLE", mnuSettingsDateFormatOracle.Checked

157     If lvwColumns.ListItems.Count > 0 Then
158        FillTable sbStatusBar.Panels(3).Text
159     End If

End Sub

Private Sub mnuSettingsDateFormatSQLServer_Click()
164     mnuSettingsDateFormatOracle.Checked = False
165     mnuSettingsDateFormatODBC.Checked = False
166     SaveSetting App.ProductName, "Settings", "DateORACLE", False
167     SaveSetting App.ProductName, "Settings", "DateODBC", False

169     mnuSettingsDateFormatSQLServer.Checked = Not mnuSettingsDateFormatSQLServer.Checked
170     SaveSetting App.ProductName, "Settings", "DateSQLSERVER", mnuSettingsDateFormatSQLServer.Checked

172     If lvwColumns.ListItems.Count > 0 Then
173        FillTable sbStatusBar.Panels(3).Text
174     End If


End Sub

Private Sub mnuSettingsDelete_Click()
180     mnuSettingsDelete.Checked = Not mnuSettingsDelete.Checked
181     SaveSetting App.ProductName, "Settings", "DeleteLine", mnuSettingsDelete.Checked
End Sub

Private Sub tlbData_ButtonClick(ByVal Button As MSComctlLib.Button)

186     Select Case Button.Key
           Case "SelectAll"
188           PopupMenu mnuSelect, , tlbData.Left + tlbData.Buttons("SelectAll").Left, tlbData.Top + tlbData.Height
           Case "Clear"
190           PopupMenu mnuClear, , tlbData.Left + tlbData.Buttons("Clear").Left, tlbData.Top + tlbData.Height
           Case "Settings"
192           PopupMenu mnuSettings, , tlbData.Left + tlbData.Buttons("Settings").Left, tlbData.Top + tlbData.Height

194     End Select

End Sub
Private Sub ColSelect(pbValue As Boolean)

Dim colX As ColumnHeader

201     For Each colX In lvwColumns.ColumnHeaders
202        If colX.Index <> 1 Then
203           colX.Icon = IIf(pbValue = True, "SelectAll", "Clear")
204        End If
205     Next

207     Set colX = Nothing

End Sub
Private Sub RowSelect(pbValue As Boolean)

Dim i As Long

214     If lvwColumns.ListItems.Count = 0 Then
215        Exit Sub
216     End If

218     ProgressBar.Visible = True
219     ProgressBar.Max = lvwColumns.ListItems.Count
220     For i = 1 To lvwColumns.ListItems.Count
221        ProgressBar.Value = i
222        lvwColumns.ListItems(i).Checked = pbValue
223     Next i
224     ProgressBar.Visible = False

End Sub
Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
228     On Error Resume Next
229     Select Case Button.Key
           Case "Exit"
231           mnuFileExit_Click
           Case "Save"
233           mnuFileSave_Click
234     End Select
End Sub


Private Sub tvwDSN_NodeClick(ByVal Node As MSComctlLib.Node)

240     On Error GoTo tvwDSN_NodeClick_Error

Dim oLogin As frmODBCLogon
Dim itmX As Node
Dim oTable As adodb.Recordset
Dim i As Long
Dim sConnectString As String

248     If Node.Parent Is Nothing Then
249        Set oLogin = New frmODBCLogon
250        oLogin.Initialize Node.Text
251        oLogin.Show vbModal, Me
252        sConnectString = oLogin.ConnectString
253        Set oLogin = Nothing

255        If Len(sConnectString) = 0 Then
256           Exit Sub
257        End If

259        If m_oADOConnect Is Nothing Then
260           Set m_oADOConnect = New adodb.Connection
261        End If


264        If m_oADOConnect.State = adStateOpen Then
265           m_oADOConnect.Close
266        End If

268        m_oADOConnect.Open sConnectString

270        Set oTable = m_oADOConnect.OpenSchema(adSchemaTables)
271        If oTable Is Nothing Then
272           Exit Sub
273        End If

275        ProgressBar.Max = oTable.RecordCount

277        Do Until oTable.EOF
278           i = i + 1
279           ProgressBar.Value = i
280           Set itmX = tvwDSN.Nodes.Add(Node.Key, tvwChild)
281           itmX.Image = "Table"
282           itmX.Text = oTable.Fields("TABLE_NAME")
283           itmX.Key = Trim$(Str$(i)) & "-" & Node.Key & "." & itmX.Text
284           oTable.MoveNext
285        Loop
286        ProgressBar.Visible = False
287        Node.Expanded = True
288        Node.EnsureVisible
289        sbStatusBar.Panels(2).Text = Node.Text
290     Else
291        sbStatusBar.Panels(3).Text = Node.Text
292        lvwColumns.Tag = "ALL"
293        FillTable Node.Text
294     End If

296     Exit Sub
297 tvwDSN_NodeClick_Error:
298     MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbCritical
299     ProgressBar.Visible = False
300     lvwColumns.ListItems.Clear
301     lvwColumns.ColumnHeaders.Clear

End Sub

Private Sub FillTable(psTable As String)

Dim oRS As adodb.Recordset
Dim colX As ColumnHeader
Dim itmX As ListItem
Dim i As Long
Dim j As Long
Dim lRecordCount As Long

314     Set oRS = m_oADOConnect.Execute("Select * from " & psTable)

316     If oRS Is Nothing Then
317        Exit Sub
318     End If

320     lvwColumns.ListItems.Clear
321     lvwColumns.ColumnHeaders.Clear

323     sbStatusBar.Panels(1).Text = "Loading Columns"
324     Set colX = lvwColumns.ColumnHeaders.Add
325     colX.Text = "Select"
326     colX.Width = "1000"

328     ProgressBar.Visible = True
329     ProgressBar.Max = oRS.Fields.Count
330     For i = 0 To oRS.Fields.Count - 1
331        ProgressBar.Value = i
332        Set colX = lvwColumns.ColumnHeaders.Add
333        colX.Text = oRS.Fields(i).Name
334        colX.Icon = "SelectAll"
335     Next i

337     sbStatusBar.Panels(1).Text = "Loading Data..."
338     ProgressBar.Value = 0
339     lRecordCount = m_oADOConnect.Execute("Select Count(*) as RecCount from " & psTable).Fields("RecCount").Value

341     If lRecordCount <> 0 Then
342        ProgressBar.Max = lRecordCount
343     End If
344     sbStatusBar.Panels(4).Text = lRecordCount

346     Do Until oRS.EOF
347        j = j + 1
348        ProgressBar.Value = j

350        Set itmX = lvwColumns.ListItems.Add
351        itmX.Text = j
352        itmX.Checked = True
353        For i = 0 To oRS.Fields.Count - 1
354           itmX.SubItems(i + 1) = Trim$("" & SqlValue(GetValue(oRS.Fields(i).Value), oRS.Fields(i).Type))
'DoEvents
356        Next i
357        oRS.MoveNext
358     Loop

360     ProgressBar.Visible = False
361     sbStatusBar.Panels(1).Text = ""
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

378     ProgressBar.Max = lvwColumns.ListItems.Count
379     ProgressBar.Visible = True

381     For i = 1 To lvwColumns.ListItems.Count
382        ProgressBar.Value = i
383        If lvwColumns.ListItems(i).Checked = True Then
384           lSelected = lSelected + 1
385        End If
386     Next i

388     ProgressBar.Visible = False

390     If lSelected = 0 Then
391        If MsgBox("No rows selected, Select All and continue?", vbQuestion + vbYesNo) = vbYes Then
392           mnuSelectRows_Click
393        Else
394           Exit Sub
395        End If
396     End If

398     cdSave.CancelError = True
399     cdSave.Filter = "SQL Script (*.sql)|*.sql"
400     cdSave.FilterIndex = 1
401     cdSave.FileName = "DATA_" & sbStatusBar.Panels(3).Text & "_" & Format$(Now, "mmddyy")
402     cdSave.ShowSave
403     sFileName = cdSave.FileName

405     If Len(sFileName) = 0 Then
406        Exit Sub
407     End If

409     If LCase(Right$(sFileName, 4)) <> ".sql" Then
410        sFileName = sFileName & ".sql"
411     End If

413     Set oFSO = New Scripting.FileSystemObject
414     Set oStream = oFSO.OpenTextFile(sFileName, ForAppending, True)

'Add Delete From ... on first line
417     If mnuSettingsDelete.Checked Then
418        oStream.WriteLine "DELETE FROM " & sbStatusBar.Panels(3).Text & ";"
419     End If

421     lCommit = CLng(GetSetting(App.ProductName, "Settings", "Commit", "5"))

423     ProgressBar.Max = lvwColumns.ListItems.Count
424     ProgressBar.Visible = True
425     For j = 1 To lvwColumns.ListItems.Count
426        ProgressBar.Value = j
427        sFields = ""
428        sValues = ""
429        If lvwColumns.ListItems(j).Checked = True Then
430           For i = 2 To lvwColumns.ColumnHeaders.Count
431              If lvwColumns.ColumnHeaders.Item(i).Icon = "SelectAll" Then
432                 If Len(sFields) <> 0 Then sFields = sFields & ","
433                 sFields = sFields & lvwColumns.ColumnHeaders(i).Text


436                 If Len(sValues) <> 0 Then sValues = sValues & ","
437                 sValues = sValues & lvwColumns.ListItems(j).SubItems(i - 1)
438              End If
439           Next i

441           sSQL = "INSERT INTO " & sbStatusBar.Panels(3).Text & " (" & sFields & ") VALUES (" & sValues & ");"
442           oStream.WriteLine sSQL

444           If j Mod lCommit = 0 Then
445              oStream.WriteLine "Commit;"
446           End If
447        End If
448     Next j

450     If j Mod lCommit <> 0 Then
451        oStream.WriteLine "Commit;"
452     End If

454     oStream.Close
455     Set oStream = Nothing
456     Set oFSO = Nothing
457     ProgressBar.Visible = False

End Sub

