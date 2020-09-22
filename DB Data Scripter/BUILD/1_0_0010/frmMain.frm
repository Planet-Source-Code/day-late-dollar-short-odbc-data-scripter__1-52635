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
   MinButton       =   0   'False
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
      Begin VB.Menu mnuSettingsDelete 
         Caption         =   "Include DELETE ALL on first line"
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

Private Sub mnuSettingsDateFormatODBC_Click()

119     mnuSettingsDateFormatOracle.Checked = False
120     mnuSettingsDateFormatSQLServer.Checked = False
121     SaveSetting App.ProductName, "Settings", "DateORACLE", False
122     SaveSetting App.ProductName, "Settings", "DateSQLSERVER", False

124     mnuSettingsDateFormatODBC.Checked = Not mnuSettingsDateFormatODBC.Checked
125     SaveSetting App.ProductName, "Settings", "DateODBC", mnuSettingsDateFormatODBC.Checked

127     If lvwColumns.ListItems.Count > 0 Then
128        FillTable sbStatusBar.Panels(3).Text
129     End If

End Sub

Private Sub mnuSettingsDateFormatOracle_Click()
134     mnuSettingsDateFormatODBC.Checked = False
135     mnuSettingsDateFormatSQLServer.Checked = False
136     SaveSetting App.ProductName, "Settings", "DateODBC", False
137     SaveSetting App.ProductName, "Settings", "DateSQLSERVER", False

139     mnuSettingsDateFormatOracle.Checked = Not mnuSettingsDateFormatOracle.Checked
140     SaveSetting App.ProductName, "Settings", "DateORACLE", mnuSettingsDateFormatOracle.Checked

142     If lvwColumns.ListItems.Count > 0 Then
143        FillTable sbStatusBar.Panels(3).Text
144     End If

End Sub

Private Sub mnuSettingsDateFormatSQLServer_Click()
149     mnuSettingsDateFormatOracle.Checked = False
150     mnuSettingsDateFormatODBC.Checked = False
151     SaveSetting App.ProductName, "Settings", "DateORACLE", False
152     SaveSetting App.ProductName, "Settings", "DateODBC", False

154     mnuSettingsDateFormatSQLServer.Checked = Not mnuSettingsDateFormatSQLServer.Checked
155     SaveSetting App.ProductName, "Settings", "DateSQLSERVER", mnuSettingsDateFormatSQLServer.Checked

157     If lvwColumns.ListItems.Count > 0 Then
158        FillTable sbStatusBar.Panels(3).Text
159     End If


End Sub

Private Sub mnuSettingsDelete_Click()
165     mnuSettingsDelete.Checked = Not mnuSettingsDelete.Checked
166     SaveSetting App.ProductName, "Settings", "DeleteLine", mnuSettingsDelete.Checked
End Sub

Private Sub tlbData_ButtonClick(ByVal Button As MSComctlLib.Button)

171     Select Case Button.Key
           Case "SelectAll"
173           PopupMenu mnuSelect, , tlbData.Left + tlbData.Buttons("SelectAll").Left, tlbData.Top + tlbData.Height
           Case "Clear"
175           PopupMenu mnuClear, , tlbData.Left + tlbData.Buttons("Clear").Left, tlbData.Top + tlbData.Height
           Case "Settings"
177           PopupMenu mnuSettings, , tlbData.Left + tlbData.Buttons("Settings").Left, tlbData.Top + tlbData.Height

179     End Select

End Sub
Private Sub ColSelect(pbValue As Boolean)

Dim colX As ColumnHeader

186     For Each colX In lvwColumns.ColumnHeaders
187        If colX.Index <> 1 Then
188           colX.Icon = IIf(pbValue = True, "SelectAll", "Clear")
189        End If
190     Next

192     Set colX = Nothing

End Sub
Private Sub RowSelect(pbValue As Boolean)

Dim i As Long

199     If lvwColumns.ListItems.Count = 0 Then
200        Exit Sub
201     End If

203     ProgressBar.Visible = True
204     ProgressBar.Max = lvwColumns.ListItems.Count
205     For i = 1 To lvwColumns.ListItems.Count
206        ProgressBar.Value = i
207        lvwColumns.ListItems(i).Checked = pbValue
208     Next i
209     ProgressBar.Visible = False

End Sub
Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
213     On Error Resume Next
214     Select Case Button.Key
           Case "Exit"
216           mnuFileExit_Click
           Case "Save"
218           mnuFileSave_Click
219     End Select
End Sub


Private Sub tvwDSN_NodeClick(ByVal Node As MSComctlLib.Node)

225     On Error GoTo tvwDSN_NodeClick_Error

Dim oLogin As frmODBCLogon
Dim itmX As Node
Dim oTable As adodb.Recordset
Dim i As Long
Dim sConnectString As String

233     If Node.Parent Is Nothing Then
234        Set oLogin = New frmODBCLogon
235        oLogin.Initialize Node.Text
236        oLogin.Show vbModal, Me
237        sConnectString = oLogin.ConnectString
238        Set oLogin = Nothing

240        If Len(sConnectString) = 0 Then
241           Exit Sub
242        End If

244        If m_oADOConnect Is Nothing Then
245           Set m_oADOConnect = New adodb.Connection
246        End If


249        If m_oADOConnect.State = adStateOpen Then
250           m_oADOConnect.Close
251        End If

253        m_oADOConnect.Open sConnectString

255        Set oTable = m_oADOConnect.OpenSchema(adSchemaTables)
256        If oTable Is Nothing Then
257           Exit Sub
258        End If

260        ProgressBar.Max = oTable.RecordCount

262        Do Until oTable.EOF
263           i = i + 1
264           ProgressBar.Value = i
265           Set itmX = tvwDSN.Nodes.Add(Node.Key, tvwChild)
266           itmX.Image = "Table"
267           itmX.Text = oTable.Fields("TABLE_NAME")
268           itmX.Key = Trim$(Str$(i)) & "-" & Node.Key & "." & itmX.Text
269           oTable.MoveNext
270        Loop
271        ProgressBar.Visible = False
272        Node.Expanded = True
273        Node.EnsureVisible
274        sbStatusBar.Panels(2).Text = Node.Text
275     Else
276        sbStatusBar.Panels(3).Text = Node.Text
277        FillTable Node.Text
278     End If

280     Exit Sub
281 tvwDSN_NodeClick_Error:
282     MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbCritical
283     ProgressBar.Visible = False
284     lvwColumns.ListItems.Clear
285     lvwColumns.ColumnHeaders.Clear

End Sub

Private Sub FillTable(psTable As String)

Dim oRS As adodb.Recordset
Dim colX As ColumnHeader
Dim itmX As ListItem
Dim i As Long
Dim j As Long
Dim lRecordCount As Long

298     Set oRS = m_oADOConnect.Execute("Select * from " & psTable)

300     If oRS Is Nothing Then
301        Exit Sub
302     End If

304     lvwColumns.ListItems.Clear
305     lvwColumns.ColumnHeaders.Clear

307     sbStatusBar.Panels(1).Text = "Loading Columns"
308     Set colX = lvwColumns.ColumnHeaders.Add
309     colX.Text = "Select"
310     colX.Width = "500"

312     ProgressBar.Visible = True
313     ProgressBar.Max = oRS.Fields.Count
314     For i = 0 To oRS.Fields.Count - 1
315        ProgressBar.Value = i
316        Set colX = lvwColumns.ColumnHeaders.Add
317        colX.Text = oRS.Fields(i).Name
318        colX.Icon = "SelectAll"
319     Next i

321     sbStatusBar.Panels(1).Text = "Loading Data..."
322     ProgressBar.Value = 0
323     lRecordCount = m_oADOConnect.Execute("Select Count(*) as RecCount from " & psTable).Fields("RecCount").Value

325     If lRecordCount <> 0 Then
326        ProgressBar.Max = lRecordCount
327     End If
328     sbStatusBar.Panels(4).Text = lRecordCount

330     Do Until oRS.EOF
331        j = j + 1
332        ProgressBar.Value = j

334        Set itmX = lvwColumns.ListItems.Add
335        itmX.Text = j
336        For i = 0 To oRS.Fields.Count - 1
337           itmX.SubItems(i + 1) = Trim$("" & SqlValue(GetValue(oRS.Fields(i).Value), oRS.Fields(i).Type))
'DoEvents
339        Next i
340        oRS.MoveNext
341     Loop

343     ProgressBar.Visible = False
344     sbStatusBar.Panels(1).Text = ""
End Sub

Private Sub SaveScript()

Dim sFileName As String
Dim oStream As Scripting.TextStream
Dim oFSO As Scripting.FileSystemObject
Dim i As Long
Dim j As Long
Dim lSelected As Long

Dim sSQL As String
Dim sFields As String
Dim sValues As String

360     ProgressBar.Max = lvwColumns.ListItems.Count
361     ProgressBar.Visible = True

363     For i = 1 To lvwColumns.ListItems.Count
364        ProgressBar.Value = i
365        If lvwColumns.ListItems(i).Checked = True Then
366           lSelected = lSelected + 1
367        End If
368     Next i

370     ProgressBar.Visible = False

372     If lSelected = 0 Then
373        If MsgBox("No rows selected, Select All and continue?", vbQuestion + vbYesNo) = vbYes Then
374           mnuSelectRows_Click
375        Else
376           Exit Sub
377        End If
378     End If

380     cdSave.CancelError = True
381     cdSave.Filter = "SQL Script (*.sql)|*.sql"
382     cdSave.FilterIndex = 1
383     cdSave.FileName = "DATA_" & sbStatusBar.Panels(3).Text & "_" & Format$(Now, "mmddyy")
384     cdSave.ShowSave
385     sFileName = cdSave.FileName

387     If Len(sFileName) = 0 Then
388        Exit Sub
389     End If

391     If LCase(Right$(sFileName, 4)) <> ".sql" Then
392        sFileName = sFileName & ".sql"
393     End If

395     Set oFSO = New Scripting.FileSystemObject
396     Set oStream = oFSO.OpenTextFile(sFileName, ForAppending, True)

'Add Delete From ... on first line
399     If mnuSettingsDelete.Checked Then
400        oStream.WriteLine "DELETE FROM " & sbStatusBar.Panels(3).Text & ";"
401     End If

403     ProgressBar.Max = lvwColumns.ListItems.Count
404     ProgressBar.Visible = True
405     For j = 1 To lvwColumns.ListItems.Count
406        ProgressBar.Value = j
407        sFields = ""
408        sValues = ""
409        If lvwColumns.ListItems(j).Checked = True Then
410           For i = 2 To lvwColumns.ColumnHeaders.Count
411              If lvwColumns.ColumnHeaders.Item(i).Icon = "SelectAll" Then
412                 If Len(sFields) <> 0 Then sFields = sFields & ","
413                 sFields = sFields & lvwColumns.ColumnHeaders(i).Text


416                 If Len(sValues) <> 0 Then sValues = sValues & ","
417                 sValues = sValues & lvwColumns.ListItems(j).SubItems(i - 1)
418              End If
419           Next i

421           sSQL = "INSERT INTO " & sbStatusBar.Panels(3).Text & " (" & sFields & ") VALUES (" & sValues & ");"
422           oStream.WriteLine sSQL

424           If j Mod 5 = 0 Then
425              oStream.WriteLine "Commit;"
426           End If
427        End If
428     Next j

430     oStream.WriteLine "Commit;"

432     oStream.Close
433     Set oStream = Nothing
434     Set oFSO = Nothing
435     ProgressBar.Visible = False

End Sub

