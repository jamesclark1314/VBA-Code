Attribute VB_Name = "Module1"
Sub GetAsset()

Dim cusip As String
Dim cycle_code As String
Dim issuer_id As String
Dim series As String
Dim issue_year As Integer

If connectDB Then

    strsql = "WITH cte_prior AS(SELECT DISTINCT b.cusip, b.name " _
        & " FROM TRPRef.position.v_position_master a, TRPRef.security.v_security_master b, TRPRef.pricing.pricing_master c, TRPRef.account.account_master d, TRPRef.security.v_analytic_fixed_income_current e " _
        & " WHERE a.security_id = b.security_id AND a.security_id = c.security_id AND a.security_id = e.security_id " _
        & " AND b.instrument_type LIKE 'CMBS' " _
        & " AND a.effective_date = '" & Sheets("CMBS").Cells(3, 21) & "' AND a.effective_date = c.effective_date AND a.effective_date = e.effective_date " _
        & " AND d.is_active = 'true' AND d.portfolio_type <> 'EQ' AND d.account_type_trp <> 'INDEX' AND d.account_type_trp <> 'MODEL' AND acnominor <> '' " _
        & " AND a.account_id = d.account_id " _
        & " AND e.provider_code = '" & Sheets("CMBS").Cells(4, 3) & "'), " _
        & " cte_current AS(SELECT DISTINCT b.cusip, b.name " _
        & " FROM TRPRef.position.v_position_master a, TRPRef.security.v_security_master b, TRPRef.pricing.pricing_master c, TRPRef.account.account_master d, TRPRef.security.v_analytic_fixed_income_current e " _
        & " WHERE a.security_id = b.security_id AND a.security_id = c.security_id AND a.security_id = e.security_id " _
        & " AND b.instrument_type LIKE 'CMBS' " _
        & " AND a.effective_date = '" & Sheets("CMBS").Cells(4, 21) & "' AND a.effective_date = c.effective_date AND a.effective_date = e.effective_date " _
        & " AND d.is_active = 'true' AND d.portfolio_type <> 'EQ' AND d.account_type_trp <> 'INDEX' AND d.account_type_trp <> 'MODEL' AND acnominor <> '' " _
        & " AND a.account_id = d.account_id " _
        & " AND e.provider_code = '" & Sheets("CMBS").Cells(4, 3) & "') " _
        & " SELECT * FROM cte_current WHERE NOT EXISTS(SELECT * FROM cte_prior WHERE cte_prior.cusip = cte_current.cusip);"
        
    DataConnect.dbCmd.CommandType = adCmdText
    DataConnect.dbCmd.CommandText = strsql

    DataConnect.dbRS.CursorLocation = adUseClient
    DataConnect.dbRS.CacheSize = 1000
    DataConnect.dbRS.Open DataConnect.dbCmd
    

If dbRS.RecordCount = 0 Then
    MsgBox "No records retrieved, please check your Input and Factor date"
Else
    If dbRS.State = 1 Then    ' -- If Recordset had valied values and is opened
            ' -- Print Data
                Call DataConnect.printData(DataConnect.dbRS, "B7")
    End If
End If

' --------------------------------
'  CLOSE OUT ALL DATABASE OBJECTS
    
    DataConnect.dbRS.Close
    Set DataConnect.dbRS = Nothing
' --------------------------------
End If


End Sub

Sub getassetrpt()
Sheets("CMBS").Calculate
Worksheets("CMBS").Range("B6:L10000").ClearContents
Call GetAsset


End Sub

