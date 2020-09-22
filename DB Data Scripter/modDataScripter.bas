Attribute VB_Name = "modDataScripter"
Option Explicit

Public Sub LoadSettings()

Dim sReg As String

    sReg = GetSetting(App.ProductName, "Settings", "DeleteLine", "False")
    If Len(sReg) <> 0 Then
        frmMain.mnuSettingsDelete.Checked = CBool(sReg)
    End If
    
    sReg = GetSetting(App.ProductName, "Settings", "OneItemPerLine", "False")
    If Len(sReg) <> 0 Then
        frmMain.mnuSettingsOnePerLine.Checked = CBool(sReg)
    End If
    
    sReg = GetSetting(App.ProductName, "Settings", "DateODBC", "True")
    If Len(sReg) <> 0 Then
        frmMain.mnuSettingsDateFormatODBC.Checked = CBool(sReg)
    End If
    
    sReg = GetSetting(App.ProductName, "Settings", "DateORACLE", "False")
    If Len(sReg) <> 0 Then
        frmMain.mnuSettingsDateFormatOracle.Checked = CBool(sReg)
    End If
    
    sReg = GetSetting(App.ProductName, "Settings", "DateSQLSERVER", "False")
    If Len(sReg) <> 0 Then
        frmMain.mnuSettingsDateFormatSQLServer.Checked = CBool(sReg)
    End If
    
End Sub

Public Function GetValue(pvValue As Variant) As String

  
    On Error GoTo GetValue_Error
    If IsNull(pvValue) Then
        GetValue = ""
      Else
        GetValue = pvValue
    End If

Exit Function

GetValue_Error:
  

End Function

Public Function SqlValue(ByVal psValue As String, piType As adodb.DataTypeEnum) As String


    On Error GoTo SqlValue_Error
  Dim sValue As String, iLoop As Integer
  Dim sTemp As String

    'Replace Single Quotes with Double Quotes ??? Better String Handling
    sValue = ""
    For iLoop = 1 To Len(psValue)
        sTemp = Mid$(psValue, iLoop, 1)
        If sTemp = "'" Then
            sTemp = "''"
        End If
        sValue = sValue & sTemp
    Next iLoop
    psValue = sValue

    'Build query expression
    Select Case piType
      Case adNumeric, adDecimal, adInteger, adSmallInt, adDouble, adBigInt, adTinyInt, adVarNumeric, adSingle
        'Just Use sValue
        If Len(psValue) = 0 Then
            psValue = "0"
        End If
      
      Case adDate
        If IsDate(psValue) Then
            If Len(psValue) <> 0 Then
                psValue = Format$(CVDate(psValue), "yyyy-mm-dd")
            End If
        Else
            psValue = ""
        End If
        
        If Len(psValue) <> 0 Then
            'ODBC
            If frmMain.mnuSettingsDateFormatODBC.Checked Then
                psValue = "{d '" & psValue & "'}"
            'Oracle
            ElseIf frmMain.mnuSettingsDateFormatOracle.Checked Then
                psValue = "TO_DATE('" & psValue & "', 'YYYY-MM-DD')"
            'SQL Server/Access
            ElseIf frmMain.mnuSettingsDateFormatSQLServer.Checked Then
                psValue = "'" & psValue & "'"
            End If
        End If
        
      Case adDBTime
        If Len(psValue) <> 0 Then
            'ODBC
            If frmMain.mnuSettingsDateFormatODBC.Checked Then
                psValue = "{t '" & psValue & "'}"
            'Oracle
            ElseIf frmMain.mnuSettingsDateFormatOracle.Checked Then
                psValue = "TO_DATE('" & psValue & "', 'HH24:MI:SS')"
                
            'SQL Server/Access
            ElseIf frmMain.mnuSettingsDateFormatSQLServer.Checked Then
                psValue = "'" & psValue & "'"
            End If
        End If
      
      Case adDBTimeStamp
        If IsDate(psValue) Then
            psValue = Format$(CVDate(psValue), "yyyy-mm-dd hh:nn:ss")
        Else
            psValue = ""
        End If
        
        If Len(psValue) <> 0 Then
            'ODBC
            If frmMain.mnuSettingsDateFormatODBC.Checked Then
                psValue = "{ts '" & psValue & "'}"
             
            'Oracle
            ElseIf frmMain.mnuSettingsDateFormatOracle.Checked Then
                'yyyy-mm-dd hh24:mi:ss
                psValue = "TO_DATE('" & psValue & "', 'YYYY-MM-DD HH24:MI:SS')"
                
            'SQL Server/Access
            ElseIf frmMain.mnuSettingsDateFormatSQLServer.Checked Then
                psValue = "'" & psValue & "'"
            End If
            
        Else
            psValue = ""
        End If
        
      Case adVarChar, adChar, adLongVarChar, adWChar, adLongVarChar, adVarWChar, adLongVarWChar
        psValue = "'" & psValue & "'"
      
      Case adBinary, adVarBinary, adLongVarBinary
        psValue = "" '???
      
      Case adBoolean
        'Just User CBool Value
        If Len(psValue) <> 0 Then
            If CBool(psValue) = True Then
                psValue = -1
            Else
                psValue = 0
            End If
        Else
            psValue = 0
        End If
        
    End Select

    'Return Value
    SqlValue = psValue

Exit Function

SqlValue_Error:
    

End Function



