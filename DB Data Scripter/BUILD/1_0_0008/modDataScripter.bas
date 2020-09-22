Attribute VB_Name = "modDataScripter"
Option Explicit

Public Sub LoadSettings()

Dim sReg As String

7   sReg = GetSetting(App.ProductName, "Settings", "DeleteLine", "False")
8   If Len(sReg) <> 0 Then
9      frmMain.mnuSettingsDelete.Checked = CBool(sReg)
10  End If

12  sReg = GetSetting(App.ProductName, "Settings", "DateODBC", "True")
13  If Len(sReg) <> 0 Then
14     frmMain.mnuSettingsDateFormatODBC.Checked = CBool(sReg)
15  End If

17  sReg = GetSetting(App.ProductName, "Settings", "DateORACLE", "False")
18  If Len(sReg) <> 0 Then
19     frmMain.mnuSettingsDateFormatOracle.Checked = CBool(sReg)
20  End If

22  sReg = GetSetting(App.ProductName, "Settings", "DateSQLSERVER", "False")
23  If Len(sReg) <> 0 Then
24     frmMain.mnuSettingsDateFormatSQLServer.Checked = CBool(sReg)
25  End If

End Sub

Public Function GetValue(pvValue As Variant) As String


32  On Error GoTo GetValue_Error
33  If IsNull(pvValue) Then
34     GetValue = ""
35  Else
36     GetValue = pvValue
37  End If

39  Exit Function

41 GetValue_Error:


End Function

Public Function SqlValue(ByVal psValue As String, piType As adodb.DataTypeEnum) As String


49  On Error GoTo SqlValue_Error
Dim sValue As String, iLoop As Integer
Dim sTemp As String

'Replace Single Quotes with Double Quotes ??? Better String Handling
54  sValue = ""
55  For iLoop = 1 To Len(psValue)
56     sTemp = Mid$(psValue, iLoop, 1)
57     If sTemp = "'" Then
58        sTemp = "''"
59     End If
60     sValue = sValue & sTemp
61  Next iLoop
62  psValue = sValue

'Build query expression
65  Select Case piType
           Case adNumeric, adDecimal, adInteger, adSmallInt, adDouble, adBigInt, adTinyInt, adVarNumeric, adSingle
'Just Use sValue
           Case adDate
69        If IsDate(psValue) Then
70           If Len(psValue) <> 0 Then
71              psValue = Format$(CVDate(psValue), "yyyy-mm-dd")
72           End If
73        Else
74           psValue = "Null"
75        End If

'ODBC
78        If frmMain.mnuSettingsDateFormatODBC.Checked Then
79           psValue = "{d '" & psValue & "'}"
'Oracle
81        ElseIf frmMain.mnuSettingsDateFormatOracle.Checked Then
82           psValue = "to_date('" & psValue & "', 'yyyy-mm-dd')"
'SQL Server/Access
84        ElseIf frmMain.mnuSettingsDateFormatSQLServer.Checked Then
85           psValue = "'" & psValue & "'"
86        End If

           Case adDBTime
89        If Len(psValue) <> 0 Then
'ODBC
91           If frmMain.mnuSettingsDateFormatODBC.Checked Then
92              psValue = "{t '" & psValue & "'}"
'Oracle
94           ElseIf frmMain.mnuSettingsDateFormatOracle.Checked Then
95              psValue = "to_date('" & psValue & "', 'hh:mi:ss')"
'SQL Server/Access
97           ElseIf frmMain.mnuSettingsDateFormatSQLServer.Checked Then
98              psValue = "'" & psValue & "'"
99           End If
100           End If
           Case adDBTimeStamp
102           If IsDate(psValue) Then
103              psValue = Format$(CVDate(psValue), "yyyy-mm-dd hh:nn:ss")
104           Else
105              psValue = ""
106           End If

108           If Len(psValue) <> 0 Then
'ODBC
110              If frmMain.mnuSettingsDateFormatODBC.Checked Then
111                 psValue = "{ts '" & psValue & "'}"

'Oracle
114              ElseIf frmMain.mnuSettingsDateFormatOracle.Checked Then
115                 psValue = "to_date('" & psValue & "', 'yyyy-mm-dd hh:mi:ss')"
'SQL Server/Access
117              ElseIf frmMain.mnuSettingsDateFormatSQLServer.Checked Then
118                 psValue = "'" & psValue & "'"
119              End If

121           Else
122              psValue = ""
123           End If

           Case adVarChar, adChar, adLongVarChar, adWChar, adLongVarChar, adVarWChar, adLongVarWChar
126           psValue = "'" & psValue & "'"
           Case adBinary, adVarBinary, adLongVarBinary
128           psValue = "''" '???
           Case adBoolean
'Just User CBool Value
131           If CBool(psValue) = True Then
132              psValue = -1
133           Else
134              psValue = 0
135           End If
136     End Select

'Return Value
139     SqlValue = psValue

141     Exit Function

143 SqlValue_Error:


End Function



