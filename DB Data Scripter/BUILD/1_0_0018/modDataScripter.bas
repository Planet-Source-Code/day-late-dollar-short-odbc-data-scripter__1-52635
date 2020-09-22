Attribute VB_Name = "modDataScripter"
Option Explicit

Public Sub LoadSettings()

Dim sReg As String

7   sReg = GetSetting(App.ProductName, "Settings", "DeleteLine", "False")
8   If Len(sReg) <> 0 Then
9      frmMain.mnuSettingsDelete.Checked = CBool(sReg)
10  End If

12  sReg = GetSetting(App.ProductName, "Settings", "OneItemPerLine", "False")
13  If Len(sReg) <> 0 Then
14     frmMain.mnuSettingsOnePerLine.Checked = CBool(sReg)
15  End If

17  sReg = GetSetting(App.ProductName, "Settings", "DateODBC", "True")
18  If Len(sReg) <> 0 Then
19     frmMain.mnuSettingsDateFormatODBC.Checked = CBool(sReg)
20  End If

22  sReg = GetSetting(App.ProductName, "Settings", "DateORACLE", "False")
23  If Len(sReg) <> 0 Then
24     frmMain.mnuSettingsDateFormatOracle.Checked = CBool(sReg)
25  End If

27  sReg = GetSetting(App.ProductName, "Settings", "DateSQLSERVER", "False")
28  If Len(sReg) <> 0 Then
29     frmMain.mnuSettingsDateFormatSQLServer.Checked = CBool(sReg)
30  End If

End Sub

Public Function GetValue(pvValue As Variant) As String


37  On Error GoTo GetValue_Error
38  If IsNull(pvValue) Then
39     GetValue = ""
40  Else
41     GetValue = pvValue
42  End If

44  Exit Function

46 GetValue_Error:


End Function

Public Function SqlValue(ByVal psValue As String, piType As adodb.DataTypeEnum) As String


54  On Error GoTo SqlValue_Error
Dim sValue As String, iLoop As Integer
Dim sTemp As String

'Replace Single Quotes with Double Quotes ??? Better String Handling
59  sValue = ""
60  For iLoop = 1 To Len(psValue)
61     sTemp = Mid$(psValue, iLoop, 1)
62     If sTemp = "'" Then
63        sTemp = "''"
64     End If
65     sValue = sValue & sTemp
66  Next iLoop
67  psValue = sValue

'Build query expression
70  Select Case piType
           Case adNumeric, adDecimal, adInteger, adSmallInt, adDouble, adBigInt, adTinyInt, adVarNumeric, adSingle
'Just Use sValue
73        If Len(psValue) = 0 Then
74           psValue = "0"
75        End If

           Case adDate
78        If IsDate(psValue) Then
79           If Len(psValue) <> 0 Then
80              psValue = Format$(CVDate(psValue), "yyyy-mm-dd")
81           End If
82        Else
83           psValue = ""
84        End If

86        If Len(psValue) <> 0 Then
'ODBC
88           If frmMain.mnuSettingsDateFormatODBC.Checked Then
89              psValue = "{d '" & psValue & "'}"
'Oracle
91           ElseIf frmMain.mnuSettingsDateFormatOracle.Checked Then
92              psValue = "TO_DATE('" & psValue & "', 'YYYY-MM-DD')"
'SQL Server/Access
94           ElseIf frmMain.mnuSettingsDateFormatSQLServer.Checked Then
95              psValue = "'" & psValue & "'"
96           End If
97        End If

           Case adDBTime
100           If Len(psValue) <> 0 Then
'ODBC
102              If frmMain.mnuSettingsDateFormatODBC.Checked Then
103                 psValue = "{t '" & psValue & "'}"
'Oracle
105              ElseIf frmMain.mnuSettingsDateFormatOracle.Checked Then
106                 psValue = "TO_DATE('" & psValue & "', 'HH24:MI:SS')"

'SQL Server/Access
109              ElseIf frmMain.mnuSettingsDateFormatSQLServer.Checked Then
110                 psValue = "'" & psValue & "'"
111              End If
112           End If

           Case adDBTimeStamp
115           If IsDate(psValue) Then
116              psValue = Format$(CVDate(psValue), "yyyy-mm-dd hh:nn:ss")
117           Else
118              psValue = ""
119           End If

121           If Len(psValue) <> 0 Then
'ODBC
123              If frmMain.mnuSettingsDateFormatODBC.Checked Then
124                 psValue = "{ts '" & psValue & "'}"

'Oracle
127              ElseIf frmMain.mnuSettingsDateFormatOracle.Checked Then
'yyyy-mm-dd hh24:mi:ss
129                 psValue = "TO_DATE('" & psValue & "', 'YYYY-MM-DD HH24:MI:SS')"

'SQL Server/Access
132              ElseIf frmMain.mnuSettingsDateFormatSQLServer.Checked Then
133                 psValue = "'" & psValue & "'"
134              End If

136           Else
137              psValue = ""
138           End If

           Case adVarChar, adChar, adLongVarChar, adWChar, adLongVarChar, adVarWChar, adLongVarWChar
141           psValue = "'" & psValue & "'"

           Case adBinary, adVarBinary, adLongVarBinary
144           psValue = "" '???

           Case adBoolean
'Just User CBool Value
148           If Len(psValue) <> 0 Then
149              If CBool(psValue) = True Then
150                 psValue = -1
151              Else
152                 psValue = 0
153              End If
154           Else
155              psValue = 0
156           End If

158     End Select

'Return Value
161     SqlValue = psValue

163     Exit Function

165 SqlValue_Error:


End Function



