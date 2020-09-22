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
         NumListImages   =   7
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
      ButtonWidth     =   1482
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
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
   Begin MSComctlLib.Toolbar Toolbar1 
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
Public Function SqlValue(ByVal psValue As String, piType As adodb.DataTypeEnum) As String


12  On Error GoTo SqlValue_Error
Dim sValue As String, iLoop As Integer
Dim sTemp As String

'Replace Single Quotes with Double Quotes ??? Better String Handling
17  sValue = ""
18  For iLoop = 1 To Len(psValue)
19     sTemp = Mid$(psValue, iLoop, 1)
20     If sTemp = "'" Then
21        sTemp = "''"
22     End If
23     sValue = sValue & sTemp
24  Next iLoop
25  psValue = sValue

'Build query expression
28  Select Case piType
           Case adNumeric, adDecimal, adInteger, adSmallInt, adDouble, adBigInt, adTinyInt, adVarNumeric, adSingle
'Just Use sValue
           Case adDate
32        If IsDate(psValue) Then
33           If Len(psValue) <> 0 Then
34              psValue = Format$(CVDate(psValue), "yyyy-mm-dd")
35           End If
36        Else
37           psValue = "Null"
38        End If
39        psValue = "{d '" & psValue & "'}"
           Case adDBTime
41        If Len(psValue) <> 0 Then
42           psValue = "{t '" & psValue & "'}"
43        End If
           Case adDBTimeStamp
45        If IsDate(psValue) Then
46           psValue = Format$(CVDate(psValue), "yyyy-mm-dd hh:nn:ss")
47        Else
48           psValue = ""
49        End If

51        If Len(psValue) <> 0 Then
52           psValue = "{ts '" & psValue & "'}"
53        Else
54           psValue = ""
55        End If

           Case adVarChar, adChar, adLongVarChar, adWChar, adLongVarChar, adVarWChar, adLongVarWChar
58        psValue = "'" & psValue & "'"
           Case adBinary, adVarBinary, adLongVarBinary
60        psValue = "''" '???
           Case adBoolean
'Just User CBool Value
63        If CBool(psValue) = True Then
64           psValue = 1
65        Else
66           psValue = 0
67        End If
68  End Select

'Return Value
71  SqlValue = psValue

73  Exit Function

75 SqlValue_Error:


End Function


Sub GetDSNs()

83  On Error Resume Next
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
96  tvwDSN.Nodes.Clear

98  iCurrent = 0
99  i = SQL_SUCCESS

'get the DSNs
102     If SQLAllocEnv(lHenv) <> -1 Then
103        i = SQL_SUCCESS
104        Do Until i <> SQL_SUCCESS
105           sDSNItem = Space(1024)
106           sDRVItem = Space(1024)
107           i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
108           sDSN = VBA.Left(sDSNItem, iDSNLen)

110           If sDSN <> Space(iDSNLen) Then

112              Set itmX = tvwDSN.Nodes.Add()
113              itmX.Text = sDSN
114              itmX.Key = sDSN
115              itmX.Image = "Database"
116              Set itmX = Nothing
117           End If
118        Loop
119     End If

121     Exit Sub

End Sub

Private Sub Form_Load()
126     GetDSNs

128     sbStatusBar.Panels(2).Picture = imlToolbarIcons.ListImages("Database").ExtractIcon
129     sbStatusBar.Panels(3).Picture = imlToolbarIcons.ListImages("Table").ExtractIcon

131     Me.Caption = Me.Caption & " " & App.Major & "." & App.Minor & "." & App.Revision

End Sub

Private Sub Form_Unload(Cancel As Integer)
136     Set m_oADOConnect = Nothing
End Sub

Private Sub lvwColumns_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
140     If ColumnHeader.Index <> 1 Then
141        If ColumnHeader.Icon = "SelectAll" Then
142           ColumnHeader.Icon = "Clear"
143        Else
144           ColumnHeader.Icon = "SelectAll"
145        End If
146     End If

End Sub

Private Sub mnuClearColumns_Click()
151     ColSelect False
End Sub

Private Sub mnuClearRows_Click()
155     RowSelect False
End Sub

Private Sub mnuFileExit_Click()
159     Unload Me
End Sub

Private Sub mnuFileSave_Click()
163     SaveScript
End Sub

Private Sub mnuSelectColumns_Click()
167     ColSelect True
End Sub

Private Sub mnuSelectRows_Click()
171     RowSelect True
End Sub

Private Sub tlbData_ButtonClick(ByVal Button As MSComctlLib.Button)

176     Select Case Button.Key
           Case "SelectAll"
178           PopupMenu mnuSelect, , tlbData.Left + tlbData.Buttons("SelectAll").Left, tlbData.Top + tlbData.Height
           Case "Clear"
180           PopupMenu mnuClear, , tlbData.Left + tlbData.Buttons("Clear").Left, tlbData.Top + tlbData.Height
181     End Select

End Sub
Private Sub ColSelect(pbValue As Boolean)

Dim colX As ColumnHeader

188     For Each colX In lvwColumns.ColumnHeaders
189        If colX.Index <> 1 Then
190           colX.Icon = IIf(pbValue = True, "SelectAll", "Clear")
191        End If
192     Next

194     Set colX = Nothing

End Sub
Private Sub RowSelect(pbValue As Boolean)

Dim i As Long

201     If lvwColumns.ListItems.Count = 0 Then
202        Exit Sub
203     End If

205     ProgressBar.Visible = True
206     ProgressBar.Max = lvwColumns.ListItems.Count
207     For i = 1 To lvwColumns.ListItems.Count
208        ProgressBar.Value = i
209        lvwColumns.ListItems(i).Checked = pbValue
210     Next i
211     ProgressBar.Visible = False

End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
215     On Error Resume Next
216     Select Case Button.Key
           Case "Exit"
218           mnuFileExit_Click
           Case "Save"
220           mnuFileSave_Click
221     End Select
End Sub

Private Sub tvwDSN_NodeClick(ByVal Node As MSComctlLib.Node)

226     On Error GoTo tvwDSN_NodeClick_Error

Dim oLogin As frmODBCLogon
Dim itmX As Node
Dim oTable As adodb.Recordset
Dim i As Long
Dim sConnectString As String

234     If Node.Parent Is Nothing Then
235        Set oLogin = New frmODBCLogon
236        oLogin.Initialize Node.Text
237        oLogin.Show vbModal, Me
238        sConnectString = oLogin.ConnectString
239        Set oLogin = Nothing

241        If Len(sConnectString) = 0 Then
242           Exit Sub
243        End If

245        If m_oADOConnect Is Nothing Then
246           Set m_oADOConnect = New adodb.Connection
247        End If


250        If m_oADOConnect.State = adStateOpen Then
251           m_oADOConnect.Close
252        End If

254        m_oADOConnect.Open sConnectString

256        Set oTable = m_oADOConnect.OpenSchema(adSchemaTables)
257        If oTable Is Nothing Then
258           Exit Sub
259        End If

261        ProgressBar.Max = oTable.RecordCount

263        Do Until oTable.EOF
264           i = i + 1
265           ProgressBar.Value = i
266           Set itmX = tvwDSN.Nodes.Add(Node.Key, tvwChild)
267           itmX.Image = "Table"
268           itmX.Text = oTable.Fields("TABLE_NAME")
269           itmX.Key = Trim$(Str$(i)) & "-" & Node.Key & "." & itmX.Text
270           oTable.MoveNext
271        Loop
272        ProgressBar.Visible = False
273        Node.Expanded = True
274        Node.EnsureVisible
275        sbStatusBar.Panels(2).Text = Node.Text
276     Else
277        sbStatusBar.Panels(3).Text = Node.Text
278        FillTable Node.Text
279     End If

281     Exit Sub
282 tvwDSN_NodeClick_Error:
283     MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbCritical
284     ProgressBar.Visible = False
285     lvwColumns.ListItems.Clear
286     lvwColumns.ColumnHeaders.Clear

End Sub

Private Sub FillTable(psTable As String)

Dim oRS As adodb.Recordset
Dim colX As ColumnHeader
Dim itmX As ListItem
Dim i As Long
Dim j As Long
Dim lRecordCount As Long

299     Set oRS = m_oADOConnect.Execute("Select * from " & psTable)

301     If oRS Is Nothing Then
302        Exit Sub
303     End If

305     lvwColumns.ListItems.Clear
306     lvwColumns.ColumnHeaders.Clear

308     sbStatusBar.Panels(1).Text = "Loading Columns"
309     Set colX = lvwColumns.ColumnHeaders.Add
310     colX.Text = "Select"
311     colX.Width = "500"

313     ProgressBar.Visible = True
314     ProgressBar.Max = oRS.Fields.Count
315     For i = 0 To oRS.Fields.Count - 1
316        ProgressBar.Value = i
317        Set colX = lvwColumns.ColumnHeaders.Add
318        colX.Text = oRS.Fields(i).Name
319        colX.Icon = "SelectAll"
320     Next i

322     sbStatusBar.Panels(1).Text = "Loading Data..."
323     ProgressBar.Value = 0
324     lRecordCount = m_oADOConnect.Execute("Select Count(*) as RecCount from " & psTable).Fields("RecCount").Value

326     If lRecordCount <> 0 Then
327        ProgressBar.Max = lRecordCount
328     End If
329     sbStatusBar.Panels(4).Text = lRecordCount

331     Do Until oRS.EOF
332        j = j + 1
333        ProgressBar.Value = j

335        Set itmX = lvwColumns.ListItems.Add
336        itmX.Text = j
337        For i = 0 To oRS.Fields.Count - 1
338           itmX.SubItems(i + 1) = Trim$("" & SqlValue(GetValue(oRS.Fields(i).Value), oRS.Fields(i).Type))
'DoEvents
340        Next i
341        oRS.MoveNext
342     Loop

344     ProgressBar.Visible = False
345     sbStatusBar.Panels(1).Text = ""
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

360     cdSave.CancelError = True
361     cdSave.Filter = "SQL Script (*.sql)|*.sql"
362     cdSave.FilterIndex = 1
363     cdSave.FileName = "DATA_" & sbStatusBar.Panels(3).Text & "_" & Format$(Now, "mmddyy")
364     cdSave.ShowSave
365     sFileName = cdSave.FileName

367     If Len(sFileName) = 0 Then
368        Exit Sub
369     End If



373     If LCase(Right$(sFileName, 4)) <> ".sql" Then
374        sFileName = sFileName & ".sql"
375     End If

377     Set oFSO = New Scripting.FileSystemObject
378     Set oStream = oFSO.OpenTextFile(sFileName, ForAppending, True)

380     For j = 1 To lvwColumns.ListItems.Count
381        sFields = ""
382        sValues = ""
383        If lvwColumns.ListItems(j).Checked = True Then
384           For i = 2 To lvwColumns.ColumnHeaders.Count
385              If lvwColumns.ColumnHeaders.Item(i).Icon = "SelectAll" Then
386                 If Len(sFields) <> 0 Then sFields = sFields & ","
387                 sFields = sFields & lvwColumns.ColumnHeaders(i).Text


390                 If Len(sValues) <> 0 Then sValues = sValues & ","
391                 sValues = sValues & lvwColumns.ListItems(j).SubItems(i - 1)
392              End If
393           Next i

395           sSQL = "INSERT INTO " & sbStatusBar.Panels(3).Text & " (" & sFields & ") VALUES (" & sValues & ");"
396           oStream.WriteLine sSQL

398        End If

400     Next j

402     oStream.Close
403     Set oStream = Nothing
404     Set oFSO = Nothing
End Sub

Public Function GetValue(pvValue As Variant) As String


410     On Error GoTo GetValue_Error
411     If IsNull(pvValue) Then
412        GetValue = ""
413     Else
414        GetValue = pvValue
415     End If

417     Exit Function

419 GetValue_Error:


End Function

