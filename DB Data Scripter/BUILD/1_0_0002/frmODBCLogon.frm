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
12  GetDSNsAndDrivers

14  For i = 0 To cboDSNList.ListCount
15     If UCase(cboDSNList.List(i)) = UCase(psDSN) Then
16        cboDSNList.ListIndex = i
17        Exit For
18     End If
19  Next i



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
76     txtServer.Enabled = True
77     cboDrivers.Enabled = True
78  Else
79     txtServer.Enabled = False
80     cboDrivers.Enabled = False
81  End If
End Sub

Sub GetDSNsAndDrivers()
Dim i As Integer
Dim sDSNItem As String * 1024
Dim sDRVItem As String * 1024
Dim sDSN As String
Dim sDRV As String
Dim iDSNLen As Integer
Dim iDRVLen As Integer
Dim lHenv As Long         'handle to the environment

94  On Error Resume Next
95  cboDSNList.AddItem "(None)"

'get the DSNs
98  If SQLAllocEnv(lHenv) <> -1 Then
99     Do Until i <> SQL_SUCCESS
100           sDSNItem = Space$(1024)
101           sDRVItem = Space$(1024)
102           i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
103           sDSN = Left$(sDSNItem, iDSNLen)
104           sDRV = Left$(sDRVItem, iDRVLen)

106           If sDSN <> Space(iDSNLen) Then
107              cboDSNList.AddItem sDSN
108              cboDrivers.AddItem sDRV
109           End If
110        Loop
111     End If
'remove the dupes
113     If cboDSNList.ListCount > 0 Then
114        With cboDrivers
115           If .ListCount > 1 Then
116              i = 0
117              While i < .ListCount
118                 If .List(i) = .List(i + 1) Then
119                    .RemoveItem (i)
120                 Else
121                    i = i + 1
122                 End If
123              Wend
124           End If
125        End With
126     End If
127     cboDSNList.ListIndex = 0
End Sub

