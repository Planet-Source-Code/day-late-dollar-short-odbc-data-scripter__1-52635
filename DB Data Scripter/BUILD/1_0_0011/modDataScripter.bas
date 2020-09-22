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
72        If IsDate(psValue) Then
73           If Len(psValue) <> 0 Then
74              psValue = Format$(CVDate(psValue), "yyyy-mm-dd")
75           End If
76        Else
77           psValue = "Null"
78        End If

'ODBC
81        If frmMain.mnuSettingsDateFormatODBC.Checked Then
82           psValue = "{d '" & psValue & "'}"
'Oracle
84        ElseIf frmMain.mnuSettingsDateFormatOracle.Checked Then
85           psValue = "to_date('" & psValue & "', 'yyyy-mm-dd')"
'SQL Server/Access
87        ElseIf frmMain.mnuSettingsDateFormatSQLServer.Checked Then
88           psValue = "'" & psValue & "'"
89        End If

           Case adDBTime
92        If Len(psValue) <> 0 Then
'ODBC
94           If frmMain.mnuSettingsDateFormatODBC.Checked Then
95              psValue = "{t '" & psValue & "'}"
'Oracle
97           ElseIf frmMain.mnuSettingsDateFormatOracle.Checked Then
98              psValue = Format$(psValue, "hh:nn:ss AMPM")
99              psValue = "to_date('" & psValue & "', 'hh:mi:ss')"

101                 psValue = Replace(psValue, "AM", "")
102                 psValue = Replace(psValue, "PM", "")
'SQL Server/Access
104              ElseIf frmMain.mnuSettingsDateFormatSQLServer.Checked Then
105                 psValue = "'" & psValue & "'"
106              End If
107           End If
           Case adDBTimeStamp
109           If IsDate(psValue) Then
110              psValue = Format$(CVDate(psValue), "yyyy-mm-dd hh:nn:ss AMPM")
111           Else
112              psValue = ""
113           End If

115           If Len(psValue) <> 0 Then
'ODBC
117              If frmMain.mnuSettingsDateFormatODBC.Checked Then
118                 psValue = "{ts '" & psValue & "'}"

'Oracle
121              ElseIf frmMain.mnuSettingsDateFormatOracle.Checked Then
122                 psValue = "to_date('" & psValue & "', 'yyyy-mm-dd hh:mi:ss')"
123                 psValue = Replace(psValue, "AM", "")
124                 psValue = Replace(psValue, "PM", "")
'SQL Server/Access
126              ElseIf frmMain.mnuSettingsDateFormatSQLServer.Checked Then
127                 psValue = "'" & psValue & "'"
128              End If

130           Else
131              psValue = ""
132           End If

           Case adVarChar, adChar, adLongVarChar, adWChar, adLongVarChar, adVarWChar, adLongVarWChar
135           psValue = "'" & psValue & "'"
           Case adBinary, adVarBinary, adLongVarBinary
137           psValue = "''" '???
           Case adBoolean
'Just User CBool Value
140           If CBool(psValue) = True Then
141              psValue = -1
142           Else
143              psValue = 0
144           End If
145     End Select

'Return Value
148     SqlValue = psValue

150     Exit Function

152 SqlValue_Error:


End Function



