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
   Begin MSComDlg.CommonDialog cdSave 
      Left            =   4890
      Top             =   2955
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
         NumListImages   =   5
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
      EndProperty
   End
   Begin MSComctlLib.ListView lvwColumns 
      Height          =   5610
      Left            =   3450
      TabIndex        =   3
      Top             =   375
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   9895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
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

Private Sub mnuFileExit_Click()
140     Unload Me
End Sub

Private Sub mnuFileSave_Click()
144     SaveScript
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
148     On Error Resume Next
149     Select Case Button.Key
           Case "Exit"
151           mnuFileExit_Click
           Case "Save"
153           mnuFileSave_Click
154     End Select
End Sub

Private Sub tvwDSN_NodeClick(ByVal Node As MSComctlLib.Node)

159     On Error GoTo tvwDSN_NodeClick_Error

Dim oLogin As frmODBCLogon
Dim itmX As Node
Dim oTable As adodb.Recordset
Dim i As Long
Dim sConnectString As String

167     If Node.Parent Is Nothing Then
168        Set oLogin = New frmODBCLogon
169        oLogin.Initialize Node.Text
170        oLogin.Show vbModal, Me
171        sConnectString = oLogin.ConnectString
172        Set oLogin = Nothing

174        If Len(sConnectString) = 0 Then
175           Exit Sub
176        End If

178        If m_oADOConnect Is Nothing Then
179           Set m_oADOConnect = New adodb.Connection
180        End If


183        If m_oADOConnect.State = adStateOpen Then
184           m_oADOConnect.Close
185        End If

187        m_oADOConnect.Open sConnectString

189        Set oTable = m_oADOConnect.OpenSchema(adSchemaTables)
190        If oTable Is Nothing Then
191           Exit Sub
192        End If

194        ProgressBar.Max = oTable.RecordCount

196        Do Until oTable.EOF
197           i = i + 1
198           ProgressBar.Value = i
199           Set itmX = tvwDSN.Nodes.Add(Node.Key, tvwChild)
200           itmX.Image = "Table"
201           itmX.Text = oTable.Fields("TABLE_NAME")
202           itmX.Key = Trim$(Str$(i)) & "-" & Node.Key & "." & itmX.Text
203           oTable.MoveNext
204        Loop
205        ProgressBar.Visible = False
206        Node.Expanded = True
207        Node.EnsureVisible
208        sbStatusBar.Panels(2).Text = Node.Text
209     Else
210        sbStatusBar.Panels(3).Text = Node.Text
211        FillTable Node.Text
212     End If

214     Exit Sub
215 tvwDSN_NodeClick_Error:
216     MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbCritical
217     ProgressBar.Visible = False
218     lvwColumns.ListItems.Clear
219     lvwColumns.ColumnHeaders.Clear

End Sub

Private Sub FillTable(psTable As String)

Dim oRS As adodb.Recordset
Dim colX As ColumnHeader
Dim itmX As ListItem
Dim i As Long
Dim j As Long
Dim lRecordCount As Long

232     Set oRS = m_oADOConnect.Execute("Select * from " & psTable)

234     If oRS Is Nothing Then
235        Exit Sub
236     End If

238     lvwColumns.ListItems.Clear
239     lvwColumns.ColumnHeaders.Clear

241     sbStatusBar.Panels(1).Text = "Loading Columns"
242     Set colX = lvwColumns.ColumnHeaders.Add
243     colX.Text = "Select"
244     colX.Width = "500"

246     ProgressBar.Visible = True
247     ProgressBar.Max = oRS.Fields.Count
248     For i = 0 To oRS.Fields.Count - 1
249        ProgressBar.Value = i
250        Set colX = lvwColumns.ColumnHeaders.Add
251        colX.Text = oRS.Fields(i).Name
252     Next i

254     sbStatusBar.Panels(1).Text = "Loading Data..."
255     ProgressBar.Value = 0
256     lRecordCount = m_oADOConnect.Execute("Select Count(*) as RecCount from " & psTable).Fields("RecCount").Value

258     If lRecordCount <> 0 Then
259        ProgressBar.Max = lRecordCount
260     End If
261     sbStatusBar.Panels(4).Text = lRecordCount

263     Do Until oRS.EOF
264        j = j + 1
265        ProgressBar.Value = j

267        Set itmX = lvwColumns.ListItems.Add

269        For i = 0 To oRS.Fields.Count - 1
270           itmX.SubItems(i + 1) = Trim$("" & SqlValue(GetValue(oRS.Fields(i).Value), oRS.Fields(i).Type))
'DoEvents
272        Next i
273        oRS.MoveNext
274     Loop

276     ProgressBar.Visible = False
277     sbStatusBar.Panels(1).Text = ""
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

292     cdSave.Filter = "SQL Script (*.sql)|*.sql"
293     cdSave.FilterIndex = 1
294     cdSave.ShowSave
295     sFileName = cdSave.FileName

297     If Len(sFileName) = 0 Then
298        Exit Sub
299     End If



303     If LCase(Right$(sFileName, 4)) <> ".sql" Then
304        sFileName = sFileName & ".sql"
305     End If

307     Set oFSO = New Scripting.FileSystemObject
308     Set oStream = oFSO.OpenTextFile(sFileName, ForAppending, True)

310     For j = 1 To lvwColumns.ListItems.Count
311        sFields = ""
312        sValues = ""
313        If lvwColumns.ListItems(j).Checked = True Then
314           For i = 2 To lvwColumns.ColumnHeaders.Count
315              If Len(sFields) <> 0 Then sFields = sFields & ","
316              sFields = sFields & lvwColumns.ColumnHeaders(i).Text


319              If Len(sValues) <> 0 Then sValues = sValues & ","
320              sValues = sValues & lvwColumns.ListItems(j).SubItems(i - 1)
321           Next i

323           sSQL = "Insert INTO " & sbStatusBar.Panels(3).Text & " (" & sFields & ") VALUES (" & sValues & ");"
324           oStream.WriteLine sSQL

326        End If

328     Next j

330     oStream.Close
331     Set oStream = Nothing
332     Set oFSO = Nothing
End Sub

Public Function GetValue(pvValue As Variant) As String


338     On Error GoTo GetValue_Error
339     If IsNull(pvValue) Then
340        GetValue = ""
341     Else
342        GetValue = pvValue
343     End If

345     Exit Function

347 GetValue_Error:


End Function

