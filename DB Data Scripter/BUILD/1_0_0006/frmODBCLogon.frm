VERSION 5.00
Begin VB.Form frmODBCLogon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ODBC Logon"
   ClientHeight    =   3180
   ClientLeft      =   2850
   ClientTop       =   1755
   ClientWidth     =   4470
   ControlBox      =   0   'False
   Icon            =   "frmODBCLogon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   450
      Left            =   2520
      TabIndex        =   13
      Top             =   2655
      Width           =   1440
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   450
      Left            =   915
      TabIndex        =   12
      Top             =   2655
      Width           =   1440
   End
   Begin VB.Frame fraStep3 
      Caption         =   "Connection Values"
      Height          =   2415
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   4230
      Begin VB.TextBox txtUID 
         Height          =   300
         Left            =   1125
         TabIndex        =   3
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtPWD 
         Height          =   300
         Left            =   1125
         TabIndex        =   5
         Top             =   930
         Width           =   3015
      End
      Begin VB.TextBox txtDatabase 
         Height          =   300
         Left            =   1125
         TabIndex        =   7
         Top             =   1260
         Width           =   3015
      End
      Begin VB.ComboBox cboDSNList 
         Height          =   315
         ItemData        =   "frmODBCLogon.frx":000C
         Left            =   1125
         List            =   "frmODBCLogon.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3000
      End
      Begin VB.TextBox txtServer 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1125
         TabIndex        =   11
         Top             =   1935
         Width           =   3015
      End
      Begin VB.ComboBox cboDrivers 
         Height          =   315
         Left            =   1125
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1590
         Width           =   3015
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "&DSN:"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   0
         Top             =   285
         Width           =   390
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "&UID:"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   2
         Top             =   630
         Width           =   330
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "&Password:"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   4
         Top             =   975
         Width           =   735
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "Data&base:"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   6
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "Dri&ver:"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   8
         Top             =   1665
         Width           =   465
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "&Server:"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   10
         Top             =   2010
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmODBCLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer
Private Declare Function SQLAllocEnv% Lib "ODBC32.DLL" (env&)
Const SQL_SUCCESS As Long = 0
Const SQL_FETCH_NEXT As Long = 1

Public ConnectString As String

Public Sub Initialize(psDSN As String)

Dim i As Long
12  GetDSNsAndDrivers psDSN

' For i = 0 To cboDSNList.ListCount
'     If UCase(cboDSNList.List(i)) = UCase(psDSN) Then
'         cboDSNList.ListIndex = i
'         Exit For
'     End If
' Next i



End Sub


Private Sub cmdCancel_Click()
27  Unload Me
End Sub

Private Sub cmdOK_Click()
Dim sConnect    As String
Dim sADOConnect As String
Dim sDAOConnect As String
Dim sDSN        As String

36  If cboDSNList.ListIndex > 0 Then
37     sDSN = "DSN=" & cboDSNList.Text & ";"
38  Else
39     sConnect = sConnect & "Driver=" & cboDrivers.Text & ";"
40     sConnect = sConnect & "Server=" & txtServer.Text & ";"
41  End If

43  sConnect = sConnect & "UID=" & txtUID.Text & ";"
44  sConnect = sConnect & "PWD=" & txtPWD.Text & ";"

46  If Len(txtDatabase.Text) > 0 Then
47     sConnect = sConnect & "Database=" & txtDatabase.Text & ";"
48  End If

50  sADOConnect = "PROVIDER=MSDASQL;" & sDSN & sConnect
51  sDAOConnect = "ODBC;" & sDSN & sConnect

'MsgBox _
 54 "To open an ADO Connection, use:" & vbCrLf & _
   "Set gConnection = New Connection" & vbCrLf & _
   "gConnection.Open """ & sADOConnect & """" & vbCrLf & vbCrLf & _
   "To open a DAO database object, use:" & vbCrLf & _
   "Set gDatabase = OpenDatabase(vbNullString, 0, 0, sDAOConnect)" & vbCrLf & vbCrLf & _
   "Or to open an RDO Connection, use:" & vbCrLf & _
   "Set gRDOConnection = rdoEnvironments(0).OpenConnection(sDSN, rdDriverNoPrompt, 0, sConnect)"

'ADO:
'Set m_oAdoConn = New Connection
'm_oAdoConn.Open sADOConnect
65  Me.ConnectString = sADOConnect
66  Unload Me
'DAO:
'Set gDatabase = OpenDatabase(vbNullString, 0, 0, sDAOConnect)
'RDO:
'Set gRDOConnection = rdoEnvironments(0).OpenConnection(sDSN, rdDriverNoPrompt, 0, sConnect)
End Sub

Private Sub cboDSNList_Click()
74  On Error Resume Next
75  If cboDSNList.Text = "(None)" Then
76     txtServer.Enabled = False
77     cboDrivers.Enabled = False
78  Else
79     txtServer.Enabled = True
80     cboDrivers.Enabled = True
81  End If
End Sub

Sub GetDSNsAndDrivers(Optional psDSN As String)
Dim i As Integer
Dim sDSNItem As String * 1024
Dim sDRVItem As String * 1024
Dim sDSN As String
Dim sDRV As String
Dim iDSNLen As Integer
Dim iDRVLen As Integer
Dim lHenv As Long         'handle to the environment

Dim sDsnDriver As String

96  On Error Resume Next
97  cboDSNList.AddItem "(None)"

'get the DSNs
100     If SQLAllocEnv(lHenv) <> -1 Then
101        Do Until i <> SQL_SUCCESS
102           sDSNItem = Space$(1024)
103           sDRVItem = Space$(1024)
104           i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
105           sDSN = Left$(sDSNItem, iDSNLen)
106           sDRV = Left$(sDRVItem, iDRVLen)

108           If sDSN <> Space(iDSNLen) Then
109              cboDSNList.AddItem sDSN
110              cboDrivers.AddItem sDRV

112              If UCase$(psDSN) = UCase$(sDSN) Then
113                 sDsnDriver = sDRV
114              End If
115           End If
116        Loop
117     End If

'remove the dupes
120     If cboDSNList.ListCount > 0 Then
121        With cboDrivers
122           If .ListCount > 1 Then
123              i = 0
124              While i < .ListCount
125                 If .List(i) = .List(i + 1) Then
126                    .RemoveItem (i)
127                 Else
128                    i = i + 1
129                 End If
130              Wend
131           End If
132        End With
133     End If
134     cboDSNList.ListIndex = 0

136     If Len(psDSN) <> 0 Then
137        For i = 1 To cboDSNList.ListCount
138           If UCase$(cboDSNList.List(i)) = UCase$(psDSN) Then
139              cboDSNList.ListIndex = i
140              cboDSNList.Locked = True
141              Exit For
142           End If
143        Next i
144     End If

146     If Len(sDsnDriver) <> 0 Then
147        For i = 1 To cboDrivers.ListCount
148           If UCase$(cboDrivers.List(i)) = UCase$(sDsnDriver) Then
149              cboDrivers.ListIndex = i
150              cboDrivers.Locked = True
151              Exit For
152           End If
153        Next i
154     End If

End Sub

