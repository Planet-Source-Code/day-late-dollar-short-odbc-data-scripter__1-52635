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
         IMEMode         =   3  'DISABLE
         Left            =   1125
         PasswordChar    =   "*"
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
    
    GetDSNsAndDrivers psDSN
    
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim sConnect    As String
    Dim sADOConnect As String
    Dim sDAOConnect As String
    Dim sDSN        As String
    
    If cboDSNList.ListIndex > 0 Then
        sDSN = "DSN=" & cboDSNList.Text & ";"
    Else
        sConnect = sConnect & "Driver=" & cboDrivers.Text & ";"
        sConnect = sConnect & "Server=" & txtServer.Text & ";"
    End If
    
    sConnect = sConnect & "UID=" & txtUID.Text & ";"
    sConnect = sConnect & "PWD=" & txtPWD.Text & ";"
    
    If Len(txtDatabase.Text) > 0 Then
        sConnect = sConnect & "Database=" & txtDatabase.Text & ";"
    End If
    
    sADOConnect = "PROVIDER=MSDASQL;" & sDSN & sConnect
    sDAOConnect = "ODBC;" & sDSN & sConnect
    
    Me.ConnectString = sADOConnect
    Unload Me
    
End Sub

Private Sub cboDSNList_Click()
    On Error Resume Next
    If cboDSNList.Text = "(None)" Then
        txtServer.Enabled = False
        cboDrivers.Enabled = False
    Else
        txtServer.Enabled = True
        cboDrivers.Enabled = True
    End If
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
    
    On Error Resume Next
    cboDSNList.AddItem "(None)"

    'get the DSNs
    If SQLAllocEnv(lHenv) <> -1 Then
        Do Until i <> SQL_SUCCESS
            sDSNItem = Space$(1024)
            sDRVItem = Space$(1024)
            i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
            sDSN = Left$(sDSNItem, iDSNLen)
            sDRV = Left$(sDRVItem, iDRVLen)
            
            If sDSN <> Space(iDSNLen) Then
                cboDSNList.AddItem sDSN
                cboDrivers.AddItem sDRV
                
                If UCase$(psDSN) = UCase$(sDSN) Then
                    sDsnDriver = sDRV
                End If
            End If
        Loop
    End If
    
    'remove the dupes
    If cboDSNList.ListCount > 0 Then
        With cboDrivers
            If .ListCount > 1 Then
                i = 0
                While i < .ListCount
                    If .List(i) = .List(i + 1) Then
                        .RemoveItem (i)
                    Else
                        i = i + 1
                    End If
                Wend
            End If
        End With
    End If
    cboDSNList.ListIndex = 0
    
    If Len(psDSN) <> 0 Then
        For i = 1 To cboDSNList.ListCount
            If UCase$(cboDSNList.List(i)) = UCase$(psDSN) Then
                cboDSNList.ListIndex = i
                cboDSNList.Locked = True
                Exit For
            End If
        Next i
    End If
    
    If Len(sDsnDriver) <> 0 Then
        For i = 1 To cboDrivers.ListCount
            If UCase$(cboDrivers.List(i)) = UCase$(sDsnDriver) Then
                cboDrivers.ListIndex = i
                cboDrivers.Locked = True
                Exit For
            End If
        Next i
    End If
    
End Sub

