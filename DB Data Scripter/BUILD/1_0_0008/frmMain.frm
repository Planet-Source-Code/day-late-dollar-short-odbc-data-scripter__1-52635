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
79  End If

End Sub

Private Sub mnuClearColumns_Click()
84  ColSelect False
End Sub

Private Sub mnuClearRows_Click()
88  RowSelect False
End Sub

Private Sub mnuFileExit_Click()
92  Unload Me
End Sub

Private Sub mnuFileSave_Click()
96  SaveScript
End Sub

Private Sub mnuSelectColumns_Click()
100     ColSelect True
End Sub

Private Sub mnuSelectRows_Click()
104     RowSelect True
End Sub

Private Sub mnuSettingsDateFormatODBC_Click()

109     mnuSettingsDateFormatOracle.Checked = False
110     mnuSettingsDateFormatSQLServer.Checked = False
111     SaveSetting App.ProductName, "Settings", "DateORACLE", False
112     SaveSetting App.ProductName, "Settings", "DateSQLSERVER", False

114     mnuSettingsDateFormatODBC.Checked = Not mnuSettingsDateFormatODBC.Checked
115     SaveSetting App.ProductName, "Settings", "DateODBC", mnuSettingsDateFormatODBC.Checked

117     If lvwColumns.ListItems.Count > 0 Then
118        FillTable sbStatusBar.Panels(3).Text
119     End If

End Sub

Private Sub mnuSettingsDateFormatOracle_Click()
124     mnuSettingsDateFormatODBC.Checked = False
125     mnuSettingsDateFormatSQLServer.Checked = False
126     SaveSetting App.ProductName, "Settings", "DateODBC", False
127     SaveSetting App.ProductName, "Settings", "DateSQLSERVER", False

129     mnuSettingsDateFormatOracle.Checked = Not mnuSettingsDateFormatOracle.Checked
130     SaveSetting App.ProductName, "Settings", "DateORACLE", mnuSettingsDateFormatOracle.Checked

132     If lvwColumns.ListItems.Count > 0 Then
133        FillTable sbStatusBar.Panels(3).Text
134     End If

End Sub

Private Sub mnuSettingsDateFormatSQLServer_Click()
139     mnuSettingsDateFormatOracle.Checked = False
140     mnuSettingsDateFormatODBC.Checked = False
141     SaveSetting App.ProductName, "Settings", "DateORACLE", False
142     SaveSetting App.ProductName, "Settings", "DateODBC", False

144     mnuSettingsDateFormatSQLServer.Checked = Not mnuSettingsDateFormatSQLServer.Checked
145     SaveSetting App.ProductName, "Settings", "DateSQLSERVER", mnuSettingsDateFormatSQLServer.Checked

147     If lvwColumns.ListItems.Count > 0 Then
148        FillTable sbStatusBar.Panels(3).Text
149     End If


End Sub

Private Sub mnuSettingsDelete_Click()
155     mnuSettingsDelete.Checked = Not mnuSettingsDelete.Checked
156     SaveSetting App.ProductName, "Settings", "DeleteLine", mnuSettingsDelete.Checked
End Sub

Private Sub tlbData_ButtonClick(ByVal Button As MSComctlLib.Button)

161     Select Case Button.Key
           Case "SelectAll"
163           PopupMenu mnuSelect, , tlbData.Left + tlbData.Buttons("SelectAll").Left, tlbData.Top + tlbData.Height
           Case "Clear"
165           PopupMenu mnuClear, , tlbData.Left + tlbData.Buttons("Clear").Left, tlbData.Top + tlbData.Height
           Case "Settings"
167           PopupMenu mnuSettings, , tlbData.Left + tlbData.Buttons("Settings").Left, tlbData.Top + tlbData.Height

169     End Select

End Sub
Private Sub ColSelect(pbValue As Boolean)

Dim colX As ColumnHeader

176     For Each colX In lvwColumns.ColumnHeaders
177        If colX.Index <> 1 Then
178           colX.Icon = IIf(pbValue = True, "SelectAll", "Clear")
179        End If
180     Next

182     Set colX = Nothing

End Sub
Private Sub RowSelect(pbValue As Boolean)

Dim i As Long

189     If lvwColumns.ListItems.Count = 0 Then
190        Exit Sub
191     End If

193     ProgressBar.Visible = True
194     ProgressBar.Max = lvwColumns.ListItems.Count
195     For i = 1 To lvwColumns.ListItems.Count
196        ProgressBar.Value = i
197        lvwColumns.ListItems(i).Checked = pbValue
198     Next i
199     ProgressBar.Visible = False

End Sub
Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
203     On Error Resume Next
204     Select Case Button.Key
           Case "Exit"
206           mnuFileExit_Click
           Case "Save"
208           mnuFileSave_Click
209     End Select
End Sub


Private Sub tvwDSN_NodeClick(ByVal Node As MSComctlLib.Node)

215     On Error GoTo tvwDSN_NodeClick_Error

Dim oLogin As frmODBCLogon
Dim itmX As Node
Dim oTable As adodb.Recordset
Dim i As Long
Dim sConnectString As String

223     If Node.Parent Is Nothing Then
224        Set oLogin = New frmODBCLogon
225        oLogin.Initialize Node.Text
226        oLogin.Show vbModal, Me
227        sConnectString = oLogin.ConnectString
228        Set oLogin = Nothing

230        If Len(sConnectString) = 0 Then
231           Exit Sub
232        End If

234        If m_oADOConnect Is Nothing Then
235           Set m_oADOConnect = New adodb.Connection
236        End If


239        If m_oADOConnect.State = adStateOpen Then
240           m_oADOConnect.Close
241        End If

243        m_oADOConnect.Open sConnectString

245        Set oTable = m_oADOConnect.OpenSchema(adSchemaTables)
246        If oTable Is Nothing Then
247           Exit Sub
248        End If

250        ProgressBar.Max = oTable.RecordCount

252        Do Until oTable.EOF
253           i = i + 1
254           ProgressBar.Value = i
255           Set itmX = tvwDSN.Nodes.Add(Node.Key, tvwChild)
256           itmX.Image = "Table"
257           itmX.Text = oTable.Fields("TABLE_NAME")
258           itmX.Key = Trim$(Str$(i)) & "-" & Node.Key & "." & itmX.Text
259           oTable.MoveNext
260        Loop
261        ProgressBar.Visible = False
262        Node.Expanded = True
263        Node.EnsureVisible
264        sbStatusBar.Panels(2).Text = Node.Text
265     Else
266        sbStatusBar.Panels(3).Text = Node.Text
267        FillTable Node.Text
268     End If

270     Exit Sub
271 tvwDSN_NodeClick_Error:
272     MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbCritical
273     ProgressBar.Visible = False
274     lvwColumns.ListItems.Clear
275     lvwColumns.ColumnHeaders.Clear

End Sub

Private Sub FillTable(psTable As String)

Dim oRS As adodb.Recordset
Dim colX As ColumnHeader
Dim itmX As ListItem
Dim i As Long
Dim j As Long
Dim lRecordCount As Long

288     Set oRS = m_oADOConnect.Execute("Select * from " & psTable)

290     If oRS Is Nothing Then
291        Exit Sub
292     End If

294     lvwColumns.ListItems.Clear
295     lvwColumns.ColumnHeaders.Clear

297     sbStatusBar.Panels(1).Text = "Loading Columns"
298     Set colX = lvwColumns.ColumnHeaders.Add
299     colX.Text = "Select"
300     colX.Width = "500"

302     ProgressBar.Visible = True
303     ProgressBar.Max = oRS.Fields.Count
304     For i = 0 To oRS.Fields.Count - 1
305        ProgressBar.Value = i
306        Set colX = lvwColumns.ColumnHeaders.Add
307        colX.Text = oRS.Fields(i).Name
308        colX.Icon = "SelectAll"
309     Next i

311     sbStatusBar.Panels(1).Text = "Loading Data..."
312     ProgressBar.Value = 0
313     lRecordCount = m_oADOConnect.Execute("Select Count(*) as RecCount from " & psTable).Fields("RecCount").Value

315     If lRecordCount <> 0 Then
316        ProgressBar.Max = lRecordCount
317     End If
318     sbStatusBar.Panels(4).Text = lRecordCount

320     Do Until oRS.EOF
321        j = j + 1
322        ProgressBar.Value = j

324        Set itmX = lvwColumns.ListItems.Add
325        itmX.Text = j
326        For i = 0 To oRS.Fields.Count - 1
327           itmX.SubItems(i + 1) = Trim$("" & SqlValue(GetValue(oRS.Fields(i).Value), oRS.Fields(i).Type))
'DoEvents
329        Next i
330        oRS.MoveNext
331     Loop

333     ProgressBar.Visible = False
334     sbStatusBar.Panels(1).Text = ""
End Sub

Private Sub SaveScript()

Dim sFileName As String
Dim oStream As Scripting.TextStream
Dim oFSO As Scripting.FileSystemObject
Dim i As Long
Dim j As Long

Dim sSQL As String
Dim sFields As String
Dim sValues As String

349     cdSave.CancelError = True
350     cdSave.Filter = "SQL Script (*.sql)|*.sql"
351     cdSave.FilterIndex = 1
352     cdSave.FileName = "DATA_" & sbStatusBar.Panels(3).Text & "_" & Format$(Now, "mmddyy")
353     cdSave.ShowSave
354     sFileName = cdSave.FileName

356     If Len(sFileName) = 0 Then
357        Exit Sub
358     End If

360     If LCase(Right$(sFileName, 4)) <> ".sql" Then
361        sFileName = sFileName & ".sql"
362     End If

364     Set oFSO = New Scripting.FileSystemObject
365     Set oStream = oFSO.OpenTextFile(sFileName, ForAppending, True)

'Add Delete From ... on first line
368     If mnuSettingsDelete.Checked Then
369        oStream.WriteLine "DELETE FROM " & sbStatusBar.Panels(3).Text & ";"
370     End If

372     For j = 1 To lvwColumns.ListItems.Count
373        sFields = ""
374        sValues = ""
375        If lvwColumns.ListItems(j).Checked = True Then
376           For i = 2 To lvwColumns.ColumnHeaders.Count
377              If lvwColumns.ColumnHeaders.Item(i).Icon = "SelectAll" Then
378                 If Len(sFields) <> 0 Then sFields = sFields & ","
379                 sFields = sFields & lvwColumns.ColumnHeaders(i).Text


382                 If Len(sValues) <> 0 Then sValues = sValues & ","
383                 sValues = sValues & lvwColumns.ListItems(j).SubItems(i - 1)
384              End If
385           Next i

387           sSQL = "INSERT INTO " & sbStatusBar.Panels(3).Text & " (" & sFields & ") VALUES (" & sValues & ");"
388           oStream.WriteLine sSQL
389        End If
390     Next j

392     oStream.Close
393     Set oStream = Nothing
394     Set oFSO = Nothing
End Sub

