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
      Default         =   -1  'True
      Height          =   450
      Left            =   915
      TabIndex        =   5
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
         TabIndex        =   0
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtPWD 
         Height          =   300
         Left            =   1125
         TabIndex        =   1
         Top             =   930
         Width           =   3015
      End
      Begin VB.TextBox txtDatabase 
         Height          =   300
         Left            =   1125
         TabIndex        =   2
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
         TabIndex        =   7
         Top             =   240
         Width           =   3000
      End
      Begin VB.TextBox txtServer 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1125
         TabIndex        =   4
         Top             =   1935
         Width           =   3015
      End
      Begin VB.ComboBox cboDrivers 
         Height          =   315
         Left            =   1125
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1590
         Width           =   3015
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "&DSN:"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   6
         Top             =   285
         Width           =   390
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "&UID:"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   8
         Top             =   630
         Width           =   330
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "&Password:"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   9
         Top             =   975
         Width           =   735
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "Data&base:"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   10
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "Dri&ver:"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   11
         Top             =   1665
         Width           =   465
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "&Server:"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   12
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

11  GetDSNsAndDrivers psDSN

End Sub


Private Sub cmdCancel_Click()
17  Unload Me
End Sub

Private Sub cmdOK_Click()
Dim sConnect    As String
Dim sADOConnect As String
Dim sDAOConnect As String
Dim sDSN        As String

26  If cboDSNList.ListIndex > 0 Then
27     sDSN = "DSN=" & cboDSNList.Text & ";"
28  Else
29     sConnect = sConnect & "Driver=" & cboDrivers.Text & ";"
30     sConnect = sConnect & "Server=" & txtServer.Text & ";"
31  End If

33  sConnect = sConnect & "UID=" & txtUID.Text & ";"
34  sConnect = sConnect & "PWD=" & txtPWD.Text & ";"

36  If Len(txtDatabase.Text) > 0 Then
37     sConnect = sConnect & "Database=" & txtDatabase.Text & ";"
38  End If

40  sADOConnect = "PROVIDER=MSDASQL;" & sDSN & sConnect
41  sDAOConnect = "ODBC;" & sDSN & sConnect

43  Me.ConnectString = sADOConnect
44  Unload Me

End Sub

Private Sub cboDSNList_Click()
49  On Error Resume Next
50  If cboDSNList.Text = "(None)" Then
51     txtServer.Enabled = False
52     cboDrivers.Enabled = False
53  Else
54     txtServer.Enabled = True
55     cboDrivers.Enabled = True
56  End If
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

71  On Error Resume Next
72  cboDSNList.AddItem "(None)"

'get the DSNs
75  If SQLAllocEnv(lHenv) <> -1 Then
76     Do Until i <> SQL_SUCCESS
77        sDSNItem = Space$(1024)
78        sDRVItem = Space$(1024)
79        i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
80        sDSN = Left$(sDSNItem, iDSNLen)
81        sDRV = Left$(sDRVItem, iDRVLen)

83        If sDSN <> Space(iDSNLen) Then
84           cboDSNList.AddItem sDSN
85           cboDrivers.AddItem sDRV

87           If UCase$(psDSN) = UCase$(sDSN) Then
88              sDsnDriver = sDRV
89           End If
90        End If
91     Loop
92  End If

'remove the dupes
95  If cboDSNList.ListCount > 0 Then
96     With cboDrivers
97        If .ListCount > 1 Then
98           i = 0
99           While i < .ListCount
100                 If .List(i) = .List(i + 1) Then
101                    .RemoveItem (i)
102                 Else
103                    i = i + 1
104                 End If
105              Wend
106           End If
107        End With
108     End If
109     cboDSNList.ListIndex = 0

111     If Len(psDSN) <> 0 Then
112        For i = 1 To cboDSNList.ListCount
113           If UCase$(cboDSNList.List(i)) = UCase$(psDSN) Then
114              cboDSNList.ListIndex = i
115              cboDSNList.Locked = True
116              Exit For
117           End If
118        Next i
119     End If

121     If Len(sDsnDriver) <> 0 Then
122        For i = 1 To cboDrivers.ListCount
123           If UCase$(cboDrivers.List(i)) = UCase$(sDsnDriver) Then
124              cboDrivers.ListIndex = i
125              cboDrivers.Locked = True
126              Exit For
127           End If
128        Next i
129     End If

End Sub

