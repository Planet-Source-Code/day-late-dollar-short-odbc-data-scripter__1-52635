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
149        FillTable sbStatusBar.Panels(3).Text
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
164        FillTable sbStatusBar.Panels(3).Text
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
179        FillTable sbStatusBar.Panels(3).Text
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

200     End Select

End Sub
Private Sub ColSelect(pbValue As Boolean)

Dim colX As ColumnHeader

207     For Each colX In lvwColumns.ColumnHeaders
208        If colX.Index <> 1 Then
209           colX.Icon = IIf(pbValue = True, "SelectAll", "Clear")
210        End If
211     Next

213     Set colX = Nothing

End Sub
Private Sub RowSelect(pbValue As Boolean)

Dim i As Long

220     If lvwColumns.ListItems.Count = 0 Then
221        Exit Sub
222     End If

224     ProgressBar.Visible = True
225     ProgressBar.Max = lvwColumns.ListItems.Count
226     For i = 1 To lvwColumns.ListItems.Count
227        ProgressBar.Value = i
228        lvwColumns.ListItems(i).Checked = pbValue
229     Next i
230     ProgressBar.Visible = False

End Sub
Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
234     On Error Resume Next
235     Select Case Button.Key
           Case "Exit"
237           mnuFileExit_Click
           Case "Save"
239           mnuFileSave_Click
240     End Select
End Sub


Private Sub tvwDSN_NodeClick(ByVal Node As MSComctlLib.Node)

246     On Error GoTo tvwDSN_NodeClick_Error

Dim oLogin As frmODBCLogon
Dim itmX As Node
Dim oTable As adodb.Recordset
Dim i As Long
Dim sConnectString As String

254     If Node.Parent Is Nothing Then
255        Set oLogin = New frmODBCLogon
256        oLogin.Initialize Node.Text
257        oLogin.Show vbModal, Me
258        sConnectString = oLogin.ConnectString
259        Set oLogin = Nothing

261        If Len(sConnectString) = 0 Then
262           Exit Sub
263        End If

265        If m_oADOConnect Is Nothing Then
266           Set m_oADOConnect = New adodb.Connection
267        End If


270        If m_oADOConnect.State = adStateOpen Then
271           m_oADOConnect.Close
272        End If

274        m_oADOConnect.Open sConnectString

276        Set oTable = m_oADOConnect.OpenSchema(adSchemaTables)
277        If oTable Is Nothing Then
278           Exit Sub
279        End If

281        ProgressBar.Max = oTable.RecordCount

283        Do Until oTable.EOF
284           i = i + 1
285           ProgressBar.Value = i
286           Set itmX = tvwDSN.Nodes.Add(Node.Key, tvwChild)
287           itmX.Image = "Table"
288           itmX.Text = oTable.Fields("TABLE_NAME")
289           itmX.Key = Trim$(Str$(i)) & "-" & Node.Key & "." & itmX.Text
290           oTable.MoveNext
291        Loop
292        ProgressBar.Visible = False
293        Node.Expanded = True
294        Node.EnsureVisible
295        sbStatusBar.Panels(2).Text = Node.Text
296     Else
297        sbStatusBar.Panels(3).Text = Node.Text
298        lvwColumns.Tag = "ALL"
299        FillTable Node.Text
300     End If

302     Exit Sub
303 tvwDSN_NodeClick_Error:
304     MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbCritical
305     ProgressBar.Visible = False
306     lvwColumns.ListItems.Clear
307     lvwColumns.ColumnHeaders.Clear

End Sub

Private Sub FillTable(psTable As String)

Dim oRS As adodb.Recordset
Dim colX As ColumnHeader
Dim itmX As ListItem
Dim i As Long
Dim j As Long
Dim lRecordCount As Long

320     Set oRS = m_oADOConnect.Execute("Select * from " & psTable)

322     If oRS Is Nothing Then
323        Exit Sub
324     End If

326     lvwColumns.ListItems.Clear
327     lvwColumns.ColumnHeaders.Clear

329     sbStatusBar.Panels(1).Text = "Loading Columns"
330     Set colX = lvwColumns.ColumnHeaders.Add
331     colX.Text = "Select"
332     colX.Width = "1000"

334     ProgressBar.Visible = True
335     ProgressBar.Max = oRS.Fields.Count
336     For i = 0 To oRS.Fields.Count - 1
337        ProgressBar.Value = i
338        Set colX = lvwColumns.ColumnHeaders.Add
339        colX.Text = oRS.Fields(i).Name
340        colX.Icon = "SelectAll"
341     Next i

343     sbStatusBar.Panels(1).Text = "Loading Data..."
344     ProgressBar.Value = 0
345     lRecordCount = m_oADOConnect.Execute("Select Count(*) as RecCount from " & psTable).Fields("RecCount").Value

347     If lRecordCount <> 0 Then
348        ProgressBar.Max = lRecordCount
349     End If
350     sbStatusBar.Panels(4).Text = lRecordCount

352     Do Until oRS.EOF
353        j = j + 1
354        ProgressBar.Value = j

356        Set itmX = lvwColumns.ListItems.Add
357        itmX.Text = j
358        itmX.Checked = True
359        For i = 0 To oRS.Fields.Count - 1
360           itmX.SubItems(i + 1) = Trim$("" & SqlValue(GetValue(oRS.Fields(i).Value), oRS.Fields(i).Type))
'DoEvents
362        Next i
363        oRS.MoveNext
364     Loop

366     ProgressBar.Visible = False
367     sbStatusBar.Panels(1).Text = ""
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

384     ProgressBar.Max = lvwColumns.ListItems.Count
385     ProgressBar.Visible = True

387     For i = 1 To lvwColumns.ListItems.Count
388        ProgressBar.Value = i
389        If lvwColumns.ListItems(i).Checked = True Then
390           lSelected = lSelected + 1
391        End If
392     Next i

394     ProgressBar.Visible = False

396     If lSelected = 0 Then
397        If MsgBox("No rows selected, Select All and continue?", vbQuestion + vbYesNo) = vbYes Then
398           mnuSelectRows_Click
399        Else
400           Exit Sub
401        End If
402     End If

404     cdSave.CancelError = True
405     cdSave.Filter = "SQL Script (*.sql)|*.sql"
406     cdSave.FilterIndex = 1
407     cdSave.FileName = "DATA_" & sbStatusBar.Panels(3).Text & "_" & Format$(Now, "mmddyy")
408     cdSave.ShowSave
409     sFileName = cdSave.FileName

411     If Len(sFileName) = 0 Then
412        Exit Sub
413     End If

415     If LCase(Right$(sFileName, 4)) <> ".sql" Then
416        sFileName = sFileName & ".sql"
417     End If

419     Set oFSO = New Scripting.FileSystemObject
420     Set oStream = oFSO.OpenTextFile(sFileName, ForAppending, True)

'Add Delete From ... on first line
423     If mnuSettingsDelete.Checked Then
424        oStream.WriteLine "DELETE FROM " & sbStatusBar.Panels(3).Text & ";"
425     End If

427     lCommit = CLng(GetSetting(App.ProductName, "Settings", "Commit", "5"))

429     ProgressBar.Max = lvwColumns.ListItems.Count
430     ProgressBar.Visible = True
431     For j = 1 To lvwColumns.ListItems.Count
432        ProgressBar.Value = j
433        sFields = ""
434        sValues = ""
435        If lvwColumns.ListItems(j).Checked = True Then
436           For i = 2 To lvwColumns.ColumnHeaders.Count
437              If lvwColumns.ColumnHeaders.Item(i).Icon = "SelectAll" Then
'Don't create if no data for field
439                 If Len(GetValue(lvwColumns.ListItems(j).SubItems(i - 1))) <> 0 Then
'output one item per line (for debugging, many columns)
441                    If mnuSettingsOnePerLine.Checked = True Then
442                       If Len(sFields) <> 0 Then sFields = sFields & vbCrLf
443                       If Len(sValues) <> 0 Then sValues = sValues & vbCrLf
444                    End If

446                    If Len(sFields) <> 0 Then sFields = sFields & ","
447                    sFields = sFields & lvwColumns.ColumnHeaders(i).Text

'& not allowed in Oracle script
450                    If mnuSettingsDateFormatOracle.Checked Then
451                       sValues = Replace(sValues, "&", "AND")
452                    End If

454                    If Len(sValues) <> 0 Then sValues = sValues & ","
455                    sValues = sValues & lvwColumns.ListItems(j).SubItems(i - 1)

457                 End If
458              End If
459           Next i

461           sSQL = "INSERT INTO " & sbStatusBar.Panels(3).Text & " (" & vbCrLf & sFields & vbCrLf & ") VALUES (" & vbCrLf & sValues & vbCrLf & ");"
462           oStream.WriteLine sSQL

464           If j Mod lCommit = 0 Then
465              oStream.WriteLine "Commit;"
466           End If
467        End If
468     Next j

470     If j Mod lCommit <> 0 Then
471        oStream.WriteLine "Commit;"
472     End If

474     oStream.Close
475     Set oStream = Nothing
476     Set oFSO = Nothing
477     ProgressBar.Visible = False

End Sub

