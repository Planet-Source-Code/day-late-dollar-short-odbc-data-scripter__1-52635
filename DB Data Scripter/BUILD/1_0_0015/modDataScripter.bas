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
68        If Len(psValue) = 0 Then
69           psValue = "0"
70        End If

           Case adDate
73        If IsDate(psValue) Then
74           If Len(psValue) <> 0 Then
75              psValue = Format$(CVDate(psValue), "yyyy-mm-dd")
76           End If
77        Else
78           psValue = ""
79        End If

81        If Len(psValue) <> 0 Then
'ODBC
83           If frmMain.mnuSettingsDateFormatODBC.Checked Then
84              psValue = "{d '" & psValue & "'}"
'Oracle
86           ElseIf frmMain.mnuSettingsDateFormatOracle.Checked Then
87              psValue = "TO_DATE('" & psValue & "', 'YYYY-MM-DD')"
'SQL Server/Access
89           ElseIf frmMain.mnuSettingsDateFormatSQLServer.Checked Then
90              psValue = "'" & psValue & "'"
91           End If
92        End If

           Case adDBTime
95        If Len(psValue) <> 0 Then
'ODBC
97           If frmMain.mnuSettingsDateFormatODBC.Checked Then
98              psValue = "{t '" & psValue & "'}"
'Oracle
100              ElseIf frmMain.mnuSettingsDateFormatOracle.Checked Then
101                 psValue = "TO_DATE('" & psValue & "', 'HH24:MI:SS')"

'SQL Server/Access
104              ElseIf frmMain.mnuSettingsDateFormatSQLServer.Checked Then
105                 psValue = "'" & psValue & "'"
106              End If
107           End If

           Case adDBTimeStamp
110           If IsDate(psValue) Then
111              psValue = Format$(CVDate(psValue), "yyyy-mm-dd hh:nn:ss")
112           Else
113              psValue = ""
114           End If

116           If Len(psValue) <> 0 Then
'ODBC
118              If frmMain.mnuSettingsDateFormatODBC.Checked Then
119                 psValue = "{ts '" & psValue & "'}"

'Oracle
122              ElseIf frmMain.mnuSettingsDateFormatOracle.Checked Then
'yyyy-mm-dd hh24:mi:ss
124                 psValue = "TO_TIMESTAMP('" & psValue & "', 'YYYY-MM-DD HH24:MI:SS')"

'SQL Server/Access
127              ElseIf frmMain.mnuSettingsDateFormatSQLServer.Checked Then
128                 psValue = "'" & psValue & "'"
129              End If

131           Else
132              psValue = ""
133           End If

           Case adVarChar, adChar, adLongVarChar, adWChar, adLongVarChar, adVarWChar, adLongVarWChar
136           psValue = "'" & psValue & "'"

           Case adBinary, adVarBinary, adLongVarBinary
139           psValue = "" '???

           Case adBoolean
'Just User CBool Value
143           If Len(psValue) <> 0 Then
144              If CBool(psValue) = True Then
145                 psValue = -1
146              Else
147                 psValue = 0
148              End If
149           Else
150              psValue = 0
151           End If

153     End Select

'Return Value
156     SqlValue = psValue

158     Exit Function

160 SqlValue_Error:


End Function



