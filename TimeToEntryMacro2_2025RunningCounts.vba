The VBA runs fine. but the numbers are still off in every cell so it's either Brooke is counting differently or the cut off date she used for the 2025 counts is different from what I used (which is all 2025 in the current Sheet 1). My manual of all 2025 dates matches the 'Entered' columns for both judgment and qualifications. Prepare an email for Brooke to inquire about this and ask the following questions. Write the questions illustratively so she is could understand what I am trying to convey.

Questions
1. Up to what date was your count based on?
2. Are you using the filed or entered date count?
3. I also noticed there are multiple entries under the same docket, for example: #136364, ATL, Row #3532). It looks like two entries indicating same judgment dates and same qualifications dates (different from each other). Are they being counted as 1 or 2?
4. Other questions for clarification and data validation.



Option Explicit

'=========================================================================
' Time to Entry Macro #2: 2025 Running Counts for Outcomes Numeric
'=========================================================================
' Project: GMS Automation & Reporting (PRD-104081)
' Epic: EPIC-288383
' Author: Rigel, BI Data Specialist, DARSRS
' Created: December 2024
' Purpose: Automate collection and calculation of 2025 running counts for
'          judgments, qualification documents, and six report types by county
'
' Requirements: See Business Requirements Document v2.0
' User Stories: US-001 through US-007
'=========================================================================

Sub Generate_2025_RunningCounts()
    '=====================================================================
    ' MAIN PROCEDURE: Generate 2025 Running Counts Report
    '=====================================================================
    ' This macro processes the source data (Sheet1) and generates a summary
    ' report on the "Outcomes - numeric" sheet with 2025 date counts by
    ' county for all document types.
    '
    ' US-001: Count 2025 Judgments by County
    ' US-002: Count 2025 Qualification Documents by County
    ' US-003: Calculate Qualification to Judgment Percentage
    ' US-004: Count Individual Document Report Types
    ' US-005: Calculate Total Reviewable Reports
    ' US-006: Generate Summary Totals Row
    '=====================================================================
    
    Dim wsSource As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRow As Long
    Dim outputRow As Long
    Dim startTime As Double
    Dim targetYear As Long
    
    ' Validation and error handling
    On Error GoTo ErrorHandler
    
    ' Start timer for performance tracking
    startTime = Timer
    
    ' Initialize target year
    targetYear = 2025
    
    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    '=================================================================
    ' STEP 1: Pre-Processing Validation
    '=================================================================
    Debug.Print "Starting Time to Entry Macro #2 - 2025 Running Counts"
    Debug.Print "Timestamp: " & Now()
    Debug.Print String(80, "=")
    
    ' Validate source worksheet exists
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets("Sheet1")
    On Error GoTo ErrorHandler
    
    If wsSource Is Nothing Then
        MsgBox "ERROR: Source worksheet 'Sheet1' not found." & vbCrLf & vbCrLf & _
               "Please ensure the source data worksheet is named 'Sheet1'.", _
               vbCritical, "Source Data Missing"
        GoTo CleanExit
    End If
    
    ' Validate output worksheet exists
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Worksheets("Outcomes - numeric")
    On Error GoTo ErrorHandler
    
    If wsOutput Is Nothing Then
        MsgBox "ERROR: Output worksheet 'Outcomes - numeric' not found." & vbCrLf & vbCrLf & _
               "Please ensure the output worksheet exists.", _
               vbCritical, "Output Sheet Missing"
        GoTo CleanExit
    End If
    
    ' Get last row of source data
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 3 Then
        MsgBox "ERROR: Source data appears to be empty." & vbCrLf & vbCrLf & _
               "Last row: " & lastRow, _
               vbCritical, "No Data Found"
        GoTo CleanExit
    End If
    
    Debug.Print "Validation passed:"
    Debug.Print "  - Source sheet: Sheet1 (Rows: " & lastRow & ")"
    Debug.Print "  - Output sheet: Outcomes - numeric"
    Debug.Print "  - Target year: " & targetYear
    Debug.Print String(80, "=")
    
    '=================================================================
    ' STEP 2: Clear Previous Output (Preserve Headers)
    '=================================================================
    Debug.Print "Clearing previous output data..."
    
    ' Clear data rows (starting from row 3, preserve rows 1-2 for headers)
    Dim lastOutputRow As Long
    lastOutputRow = wsOutput.Cells(wsOutput.Rows.Count, "A").End(xlUp).Row
    
    If lastOutputRow >= 3 Then
        wsOutput.Range("A3:T" & lastOutputRow).ClearContents
        Debug.Print "  - Cleared rows 3 to " & lastOutputRow
    End If
    
    '=================================================================
    ' STEP 3: Process Data and Generate County Summaries
    '=================================================================
    Debug.Print "Processing data by county..."
    
    Dim counties As Object  ' Dictionary to store unique counties
    Dim countyData As Object  ' Dictionary to store count data per county
    Set counties = CreateObject("Scripting.Dictionary")
    Set countyData = CreateObject("Scripting.Dictionary")
    
    ' Get unique counties and initialize data structure
    Call GetUniqueCounties(wsSource, lastRow, counties)
    Call InitializeCountyData(counties, countyData)
    
    Debug.Print "  - Found " & counties.Count & " unique counties"
    
    ' Count 2025 dates for each county
    Call CountDatesForAllCounties(wsSource, lastRow, countyData, targetYear)
    
    '=================================================================
    ' STEP 4: Write Output to Outcomes - numeric Sheet
    '=================================================================
    Debug.Print "Writing output to Outcomes - numeric sheet..."
    
    outputRow = 3  ' Start at row 3 (after headers)
    
    ' Write county data in alphabetical order
    Dim countyList As Object
    Set countyList = CreateObject("System.Collections.ArrayList")
    
    Dim key As Variant
    Dim county As Variant
    Dim countyStr As String
    
    For Each key In counties.Keys
        countyList.Add key
    Next key
    
    countyList.Sort
    
    ' Write each county's data
    For Each county In countyList
        countyStr = CStr(county)  ' Convert Variant to String for type safety
        Call WriteCountyRow(wsOutput, outputRow, countyStr, countyData(county))
        outputRow = outputRow + 1
    Next county
    
    Debug.Print "  - Wrote " & (outputRow - 3) & " county rows"
    
    '=================================================================
    ' STEP 5: Generate TOTALS Row
    '=================================================================
    Debug.Print "Generating TOTALS row..."
    
    Call WriteTotalsRow(wsOutput, outputRow, countyData)
    
    Debug.Print "  - TOTALS row written at row " & outputRow
    
    '=================================================================
    ' STEP 6: Apply Formatting
    '=================================================================
    Debug.Print "Applying formatting..."
    
    Call FormatOutputSheet(wsOutput, outputRow)
    
    '=================================================================
    ' STEP 7: Post-Processing Validation
    '=================================================================
    Debug.Print String(80, "=")
    Debug.Print "Post-processing validation:"
    
    Dim totalJudgments As Long
    Dim totalQualDocs As Long
    Dim totalReports As Long
    Dim overallPct As Double
    
    ' Calculate totals for validation
    totalJudgments = wsOutput.Cells(outputRow, 3).Value  ' Column C
    totalQualDocs = wsOutput.Cells(outputRow, 4).Value   ' Column D
    totalReports = wsOutput.Cells(outputRow, 19).Value   ' Column S
    
    If totalJudgments > 0 Then
        overallPct = (totalQualDocs / totalJudgments)
    Else
        overallPct = 0
    End If
    
    Debug.Print "  - Total 2025 Judgments: " & totalJudgments
    Debug.Print "  - Total 2025 Qualification Documents: " & totalQualDocs
    Debug.Print "  - Overall Completion Rate: " & Format(overallPct, "0%")
    Debug.Print "  - Total Reviewable Reports: " & totalReports
    Debug.Print "  - Counties Processed: " & (outputRow - 3)
    
    '=================================================================
    ' SUCCESS: Display Summary
    '=================================================================
    Dim executionTime As Double
    executionTime = Timer - startTime
    
    Debug.Print String(80, "=")
    Debug.Print "MACRO COMPLETED SUCCESSFULLY"
    Debug.Print "Execution Time: " & Format(executionTime, "0.00") & " seconds"
    Debug.Print String(80, "=")
    
    MsgBox "2025 Running Counts Generated Successfully!" & vbCrLf & vbCrLf & _
           "Counties Processed: " & (outputRow - 3) & vbCrLf & _
           "Total 2025 Judgments: " & Format(totalJudgments, "#,##0") & vbCrLf & _
           "Total 2025 Qualification Documents: " & Format(totalQualDocs, "#,##0") & vbCrLf & _
           "Overall Completion Rate: " & Format(overallPct, "0%") & vbCrLf & _
           "Total Reviewable Reports: " & Format(totalReports, "#,##0") & vbCrLf & vbCrLf & _
           "Execution Time: " & Format(executionTime, "0.0") & " seconds", _
           vbInformation, "Success - Time to Entry Macro #2"
    
CleanExit:
    ' Re-enable Excel features
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "ERROR: An unexpected error occurred." & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description & vbCrLf & vbCrLf & _
           "Please contact DARSRS for assistance.", _
           vbCritical, "Macro Error"
    
    Debug.Print String(80, "=")
    Debug.Print "ERROR OCCURRED:"
    Debug.Print "  Number: " & Err.Number
    Debug.Print "  Description: " & Err.Description
    Debug.Print String(80, "=")
End Sub

'=========================================================================
' HELPER PROCEDURES
'=========================================================================

Private Sub GetUniqueCounties(ByRef ws As Worksheet, ByVal lastRow As Long, ByRef counties As Object)
    '=====================================================================
    ' Extract unique county codes from source data
    ' US-001, US-002: County aggregation requirement
    '=====================================================================
    Dim i As Long
    Dim county As String
    
    For i = 3 To lastRow  ' Start at row 3 (after headers)
        county = Trim(UCase(ws.Cells(i, 1).Value))  ' Column A
        
        If county <> "" And county <> "COUNTY" Then
            If Not counties.Exists(county) Then
                counties.Add county, True
            End If
        End If
    Next i
End Sub

Private Sub InitializeCountyData(ByRef counties As Object, ByRef countyData As Object)
    '=====================================================================
    ' Initialize data structure for each county
    '=====================================================================
    Dim county As Variant
    Dim dataDict As Object
    
    For Each county In counties.Keys
        Set dataDict = CreateObject("Scripting.Dictionary")
        
        ' Initialize all count fields to zero
        dataDict.Add "Judgments", 0              ' Column C (G+H)
        dataDict.Add "QualDocs", 0               ' Column D (I+J)
        dataDict.Add "Inventory", 0              ' Column G (K+L)
        dataDict.Add "WellBeing", 0              ' Column I (M+N)
        dataDict.Add "SocSecRepPayee", 0         ' Column K (O+P)
        dataDict.Add "EZAccounting", 0           ' Column M (Q+R)
        dataDict.Add "AnnualReport", 0           ' Column O (U+V)
        dataDict.Add "ComprehensiveAcct", 0      ' Column Q (W+X)
        
        countyData.Add county, dataDict
    Next county
End Sub

Private Sub CountDatesForAllCounties(ByRef ws As Worksheet, ByVal lastRow As Long, _
                                     ByRef countyData As Object, ByVal targetYear As Long)
    '=====================================================================
    ' Count 2025 dates for all document types for each county
    ' US-001 through US-004
    '=====================================================================
    Dim i As Long
    Dim county As String
    Dim dataDict As Object
    
    For i = 3 To lastRow
        county = Trim(UCase(ws.Cells(i, 1).Value))  ' Column A
        
        If county <> "" And county <> "COUNTY" And countyData.Exists(county) Then
            Set dataDict = countyData(county)
            
            ' US-001: Count Judgments (Columns G, H)
            ' Count case once if EITHER Filed OR Entered date is in target year
            If IsDate2025(ws.Cells(i, 7).Value, targetYear) Or _
               IsDate2025(ws.Cells(i, 8).Value, targetYear) Then
                dataDict("Judgments") = dataDict("Judgments") + 1
            End If
            
            ' US-002: Count Qualification Documents (Columns I, J)
            ' Count case once if EITHER Filed OR Entered date is in target year
            If IsDate2025(ws.Cells(i, 9).Value, targetYear) Or _
               IsDate2025(ws.Cells(i, 10).Value, targetYear) Then
                dataDict("QualDocs") = dataDict("QualDocs") + 1
            End If
            
            ' US-004: Count Individual Report Types
            ' INVENTORY RPT (Columns K, L)
            If IsDate2025(ws.Cells(i, 11).Value, targetYear) Or _
               IsDate2025(ws.Cells(i, 12).Value, targetYear) Then
                dataDict("Inventory") = dataDict("Inventory") + 1
            End If
            
            ' WELL BEING RPT (Columns M, N)
            If IsDate2025(ws.Cells(i, 13).Value, targetYear) Or _
               IsDate2025(ws.Cells(i, 14).Value, targetYear) Then
                dataDict("WellBeing") = dataDict("WellBeing") + 1
            End If
            
            ' SOCIAL SECURITY REP PAYEE RPT (Columns O, P)
            If IsDate2025(ws.Cells(i, 15).Value, targetYear) Or _
               IsDate2025(ws.Cells(i, 16).Value, targetYear) Then
                dataDict("SocSecRepPayee") = dataDict("SocSecRepPayee") + 1
            End If
            
            ' EZ ACCOUNTING RPT (Columns Q, R)
            If IsDate2025(ws.Cells(i, 17).Value, targetYear) Or _
               IsDate2025(ws.Cells(i, 18).Value, targetYear) Then
                dataDict("EZAccounting") = dataDict("EZAccounting") + 1
            End If
            
            ' ANNUAL REPORT (Columns U, V)
            If IsDate2025(ws.Cells(i, 21).Value, targetYear) Or _
               IsDate2025(ws.Cells(i, 22).Value, targetYear) Then
                dataDict("AnnualReport") = dataDict("AnnualReport") + 1
            End If
            
            ' COMPREHENSIVE ACCOUNTING RPT (Columns W, X)
            If IsDate2025(ws.Cells(i, 23).Value, targetYear) Or _
               IsDate2025(ws.Cells(i, 24).Value, targetYear) Then
                dataDict("ComprehensiveAcct") = dataDict("ComprehensiveAcct") + 1
            End If
        End If
    Next i
End Sub

Private Function IsDate2025(ByVal cellValue As Variant, ByVal targetYear As Long) As Boolean
    '=====================================================================
    ' Check if a cell value is a date in the target year (2025)
    ' Handles Date objects and empty cells
    '=====================================================================
    IsDate2025 = False
    
    If IsEmpty(cellValue) Then Exit Function
    If Not IsDate(cellValue) Then Exit Function
    
    If Year(cellValue) = targetYear Then
        IsDate2025 = True
    End If
End Function

Private Sub WriteCountyRow(ByRef ws As Worksheet, ByVal rowNum As Long, ByVal county As String, _
                          ByRef dataDict As Object)
    '=====================================================================
    ' Write a single county's data to the output sheet
    ' US-001 through US-006
    '=====================================================================
    
    ' Column A: County name
    ws.Cells(rowNum, 1).Value = county
    
    ' Column B: Leave blank (label column)
    
    ' Column C: Judgments count (US-001)
    ws.Cells(rowNum, 3).Value = dataDict("Judgments")
    
    ' Column D: Qualification Documents count (US-002)
    ws.Cells(rowNum, 4).Value = dataDict("QualDocs")
    
    ' Column E: Percentage calculation (US-003)
    If dataDict("Judgments") > 0 Then
        ws.Cells(rowNum, 5).Formula = "=IF(D" & rowNum & ",D" & rowNum & "/C" & rowNum & ",0)"
        ws.Cells(rowNum, 5).NumberFormat = "0%"
    Else
        ws.Cells(rowNum, 5).Value = "--"
    End If
    
    ' Column F: Leave blank (buffer)
    
    ' Column G: INVENTORY RPT count (US-004)
    ' Note: Output Column G is blank, H has the count
    ws.Cells(rowNum, 8).Value = dataDict("Inventory")
    
    ' Column I: WELL BEING RPT count (US-004)
    ' Note: Output Column I is blank, J has the count
    ws.Cells(rowNum, 10).Value = dataDict("WellBeing")
    
    ' Column K: SOCIAL SECURITY REP PAYEE RPT count (US-004)
    ' Note: Output Column K is blank, L has the count
    ws.Cells(rowNum, 12).Value = dataDict("SocSecRepPayee")
    
    ' Column M: EZ ACCOUNTING RPT count (US-004)
    ' Note: Output Column M is blank, N has the count
    ws.Cells(rowNum, 14).Value = dataDict("EZAccounting")
    
    ' Column O: ANNUAL REPORT count (US-004)
    ' Note: Output Column O is blank, P has the count
    ws.Cells(rowNum, 16).Value = dataDict("AnnualReport")
    
    ' Column Q: COMPREHENSIVE ACCOUNTING RPT count (US-004)
    ' Note: Output Column Q is blank, R has the count
    ws.Cells(rowNum, 18).Value = dataDict("ComprehensiveAcct")
    
    ' Column S: Total reports (US-005)
    ' Note: Output Column S is blank, T has the formula
    ws.Cells(rowNum, 20).Formula = "=SUM(H" & rowNum & ",N" & rowNum & ",P" & rowNum & ",R" & rowNum & ")"
End Sub

Private Sub WriteTotalsRow(ByRef ws As Worksheet, ByVal rowNum As Long, ByRef countyData As Object)
    '=====================================================================
    ' Write the TOTALS row at the bottom
    ' US-006: Generate Summary Totals Row
    '=====================================================================
    
    ' Column A: "TOTALS" label
    ws.Cells(rowNum, 1).Value = "TOTALS"
    ws.Cells(rowNum, 1).Font.Bold = True
    
    ' Calculate totals from county data
    Dim totalJudgments As Long
    Dim totalQualDocs As Long
    Dim totalInventory As Long
    Dim totalWellBeing As Long
    Dim totalSocSec As Long
    Dim totalEZAcct As Long
    Dim totalAnnual As Long
    Dim totalComprehensive As Long
    
    Dim county As Variant
    Dim dataDict As Object
    
    For Each county In countyData.Keys
        Set dataDict = countyData(county)
        
        totalJudgments = totalJudgments + dataDict("Judgments")
        totalQualDocs = totalQualDocs + dataDict("QualDocs")
        totalInventory = totalInventory + dataDict("Inventory")
        totalWellBeing = totalWellBeing + dataDict("WellBeing")
        totalSocSec = totalSocSec + dataDict("SocSecRepPayee")
        totalEZAcct = totalEZAcct + dataDict("EZAccounting")
        totalAnnual = totalAnnual + dataDict("AnnualReport")
        totalComprehensive = totalComprehensive + dataDict("ComprehensiveAcct")
    Next county
    
    ' Column C: Total Judgments
    ws.Cells(rowNum, 3).Value = totalJudgments
    ws.Cells(rowNum, 3).Font.Bold = True
    
    ' Column D: Total Qualification Documents
    ws.Cells(rowNum, 4).Value = totalQualDocs
    ws.Cells(rowNum, 4).Font.Bold = True
    
    ' Column E: Overall Percentage (recalculated, not summed)
    If totalJudgments > 0 Then
        ws.Cells(rowNum, 5).Formula = "=IF(C" & rowNum & ",D" & rowNum & "/C" & rowNum & ",0)"
        ws.Cells(rowNum, 5).NumberFormat = "0%"
    Else
        ws.Cells(rowNum, 5).Value = "--"
    End If
    ws.Cells(rowNum, 5).Font.Bold = True
    
    ' Report type totals
    ws.Cells(rowNum, 8).Value = totalInventory
    ws.Cells(rowNum, 8).Font.Bold = True
    
    ws.Cells(rowNum, 10).Value = totalWellBeing
    ws.Cells(rowNum, 10).Font.Bold = True
    
    ws.Cells(rowNum, 12).Value = totalSocSec
    ws.Cells(rowNum, 12).Font.Bold = True
    
    ws.Cells(rowNum, 14).Value = totalEZAcct
    ws.Cells(rowNum, 14).Font.Bold = True
    
    ws.Cells(rowNum, 16).Value = totalAnnual
    ws.Cells(rowNum, 16).Font.Bold = True
    
    ws.Cells(rowNum, 18).Value = totalComprehensive
    ws.Cells(rowNum, 18).Font.Bold = True
    
    ' Column T: Total reports formula
    ws.Cells(rowNum, 20).Formula = "=SUM(H" & rowNum & ",N" & rowNum & ",P" & rowNum & ",R" & rowNum & ")"
    ws.Cells(rowNum, 20).Font.Bold = True
    
    ' Add top border to TOTALS row
    ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 20)).Borders(xlEdgeTop).Weight = xlMedium
End Sub

Private Sub FormatOutputSheet(ByRef ws As Worksheet, ByVal lastRow As Long)
    '=====================================================================
    ' Apply formatting to the output sheet
    '=====================================================================
    
    ' Format percentage column
    ws.Range("E3:E" & lastRow).NumberFormat = "0%"
    
    ' Format number columns with thousand separator
    Dim numCols As Variant
    numCols = Array(3, 4, 8, 10, 12, 14, 16, 18, 20)  ' C, D, H, J, L, N, P, R, T
    
    Dim i As Integer
    For i = LBound(numCols) To UBound(numCols)
        ws.Columns(numCols(i)).NumberFormat = "#,##0"
    Next i
    
    ' Auto-fit columns
    ws.Columns("A:T").AutoFit
    
    ' Center align county names
    ws.Range("A3:A" & lastRow).HorizontalAlignment = xlCenter
    
    ' Right align numeric data
    For i = LBound(numCols) To UBound(numCols)
        ws.Columns(numCols(i)).HorizontalAlignment = xlRight
    Next i
End Sub

'=========================================================================
' END OF MACRO
'=========================================================================
