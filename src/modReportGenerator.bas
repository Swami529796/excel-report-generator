Attribute VB_Name = "modReportGenerator"
'==============================================================================
' Automated Report Generator
' Author : Swaminathan Narayanan
' Repo   : https://github.com/Swami529796/excel-report-generator
'
' Turns raw transaction data into one formatted monthly report per region:
'  - KPI summary with month-on-month change
'  - Product breakdown (units, revenue, share, change vs last month)
'  - 6-month revenue trend chart
'  - Exported to PDF, optionally saved as an Outlook draft with the PDF attached
'
' Run: GenerateReports
'==============================================================================
Option Explicit

'---- Sheet names -------------------------------------------------------------
Private Const DATA_SHEET As String = "Data"
Private Const SETTINGS_SHEET As String = "Settings"
Private Const REPORT_SHEET As String = "Report"
Private Const LOG_SHEET As String = "Log"

'---- Data sheet columns ------------------------------------------------------
Private Const COL_DATE As Long = 1
Private Const COL_REGION As Long = 2
Private Const COL_PRODUCT As Long = 3
Private Const COL_UNITS As Long = 5
Private Const COL_REVENUE As Long = 7

'---- Layout ------------------------------------------------------------------
Private Const DIST_FIRST_ROW As Long = 10   ' first row of the distribution list on Settings
Private Const TREND_MONTHS As Long = 6
Private Const TABLE_ROW As Long = 10        ' header row of the product table on Report

'---- Colours -----------------------------------------------------------------
Private Const CLR_BRAND As Long = 7949855   ' RGB(31, 78, 121)
Private Const CLR_TILE As Long = 15921906   ' RGB(242, 242, 242)
Private Const CLR_UP As Long = 32768        ' RGB(0, 128, 0)
Private Const CLR_DOWN As Long = 192        ' RGB(192, 0, 0)

'---- Run state ---------------------------------------------------------------
Private mLogRow As Long
Private mCalcMode As XlCalculation

'---- Figures for one region --------------------------------------------------
Private Type RegionStats
    CurRevenue As Double
    CurUnits As Double
    CurOrders As Long
    PrevRevenue As Double
    PrevUnits As Double
    PrevOrders As Long
    Trend(1 To TREND_MONTHS) As Double
End Type

'==============================================================================
' Entry point
'==============================================================================
Public Sub GenerateReports()
    Dim wsSet As Worksheet
    Set wsSet = ThisWorkbook.Worksheets(SETTINGS_SHEET)

    '---- read and check settings ----
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Please save this workbook first.", vbExclamation
        Exit Sub
    End If

    If Not IsDate(wsSet.Range("ReportMonth").Value) Then
        MsgBox "Please enter a valid date in Report Month (Settings sheet).", vbExclamation
        Exit Sub
    End If

    Dim reportMonth As Date
    reportMonth = DateSerial(Year(wsSet.Range("ReportMonth").Value), Month(wsSet.Range("ReportMonth").Value), 1)

    Dim companyName As String
    companyName = Trim$(CStr(wsSet.Range("CompanyName").Value))

    Dim outFolder As String
    outFolder = Trim$(CStr(wsSet.Range("OutputFolder").Value))
    If Len(outFolder) = 0 Then outFolder = ThisWorkbook.Path & Application.PathSeparator & "Reports"
    If Right$(outFolder, 1) <> Application.PathSeparator Then outFolder = outFolder & Application.PathSeparator
    If Not EnsureFolder(outFolder) Then
        MsgBox "Could not create the output folder:" & vbLf & outFolder, vbCritical
        Exit Sub
    End If

    Dim createDrafts As Boolean
    createDrafts = (UCase$(Trim$(CStr(wsSet.Range("CreateDrafts").Value))) = "YES")

    '---- load all data into memory once ----
    Dim data As Variant
    data = ThisWorkbook.Worksheets(DATA_SHEET).Range("A1").CurrentRegion.Value
    If Not IsArray(data) Then
        MsgBox "The Data sheet is empty.", vbExclamation
        Exit Sub
    End If
    If UBound(data, 1) < 2 Then
        MsgBox "The Data sheet has no data rows.", vbExclamation
        Exit Sub
    End If

    Dim startTime As Double
    startTime = Timer

    Dim wsReport As Worksheet, wsLog As Worksheet
    Set wsReport = GetOrCreateSheet(REPORT_SHEET)
    Set wsLog = GetOrCreateSheet(LOG_SHEET)

    On Error GoTo CleanFail
    SpeedUp True

    wsLog.Cells.Clear
    wsLog.Range("A1:F1").Value = Array("Region", "Manager", "Revenue", "PDF", "Email", "Status")
    mLogRow = 2

    '---- Outlook (late bound: no library reference needed) ----
    Dim olApp As Object
    If createDrafts Then
        On Error Resume Next
        Set olApp = CreateObject("Outlook.Application")
        On Error GoTo CleanFail
        If olApp Is Nothing Then
            createDrafts = False
            WriteLog wsLog, "-", "-", 0, "", "Outlook not available", "Drafts skipped"
        End If
    End If

    '---- one report per region in the distribution list ----
    Dim r As Long, done As Long, failed As Long
    r = DIST_FIRST_ROW
    Do While Len(Trim$(CStr(wsSet.Cells(r, 1).Value))) > 0
        If UCase$(Trim$(CStr(wsSet.Cells(r, 4).Value))) <> "NO" Then
            Application.StatusBar = "Building report: " & wsSet.Cells(r, 1).Value
            If ProcessRegion(data, Trim$(CStr(wsSet.Cells(r, 1).Value)), _
                             Trim$(CStr(wsSet.Cells(r, 2).Value)), _
                             Trim$(CStr(wsSet.Cells(r, 3).Value)), _
                             reportMonth, companyName, outFolder, _
                             createDrafts, olApp, wsReport, wsLog) Then
                done = done + 1
            Else
                failed = failed + 1
            End If
        End If
        r = r + 1
    Loop

    With wsLog
        .Rows(1).Font.Bold = True
        .UsedRange.Columns.AutoFit
    End With

    SpeedUp False
    Application.StatusBar = False
    wsReport.Activate

    MsgBox "Reports complete for " & Format(reportMonth, "mmmm yyyy") & "." & vbLf & vbLf & _
           "Reports created: " & done & vbLf & _
           "Skipped / errors: " & failed & vbLf & _
           "Email drafts: " & IIf(createDrafts, "saved to Outlook Drafts", "off") & vbLf & _
           "Time taken: " & Format(Timer - startTime, "0.0") & " sec" & vbLf & vbLf & _
           "PDFs saved to:" & vbLf & outFolder, vbInformation
    Exit Sub

CleanFail:
    SpeedUp False
    Application.StatusBar = False
    MsgBox "Report generation stopped:" & vbLf & Err.Description, vbCritical
End Sub

'==============================================================================
' Build, export and (optionally) email one region's report
'==============================================================================
Private Function ProcessRegion(ByVal data As Variant, ByVal region As String, ByVal manager As String, _
                               ByVal email As String, ByVal reportMonth As Date, ByVal companyName As String, _
                               ByVal outFolder As String, ByVal createDrafts As Boolean, ByVal olApp As Object, _
                               ByVal wsReport As Worksheet, ByVal wsLog As Worksheet) As Boolean
    On Error GoTo Fail

    Dim stats As RegionStats
    Dim prodUnits As Object, prodRev As Object, prodPrev As Object
    Set prodUnits = CreateObject("Scripting.Dictionary")
    Set prodRev = CreateObject("Scripting.Dictionary")
    Set prodPrev = CreateObject("Scripting.Dictionary")
    prodUnits.CompareMode = vbTextCompare
    prodRev.CompareMode = vbTextCompare
    prodPrev.CompareMode = vbTextCompare

    CollectStats data, region, reportMonth, stats, prodUnits, prodRev, prodPrev

    If stats.CurOrders = 0 Then
        WriteLog wsLog, region, manager, 0, "", "", "Skipped - no data for " & Format(reportMonth, "mmm yyyy")
        Exit Function
    End If

    BuildReport wsReport, region, manager, reportMonth, companyName, stats, prodUnits, prodRev, prodPrev

    Dim pdfPath As String
    pdfPath = outFolder & "Sales Report - " & SafeFileName(region) & " - " & Format(reportMonth, "yyyy-mm") & ".pdf"
    wsReport.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, _
                                 Quality:=xlQualityStandard, IgnorePrintAreas:=False, _
                                 OpenAfterPublish:=False

    Dim emailStatus As String
    If createDrafts Then
        If Len(email) = 0 Then
            emailStatus = "No email address"
        Else
            SaveDraft olApp, email, manager, region, reportMonth, stats, pdfPath
            emailStatus = "Draft saved"
        End If
    Else
        emailStatus = "Off"
    End If

    WriteLog wsLog, region, manager, stats.CurRevenue, pdfPath, emailStatus, "Done"
    ProcessRegion = True
    Exit Function

Fail:
    WriteLog wsLog, region, manager, 0, "", "", "Error - " & Err.Description
End Function

'==============================================================================
' Aggregate the data for one region in a single pass
'==============================================================================
Private Sub CollectStats(ByVal data As Variant, ByVal region As String, ByVal reportMonth As Date, _
                         ByRef stats As RegionStats, ByVal prodUnits As Object, _
                         ByVal prodRev As Object, ByVal prodPrev As Object)
    Dim prevMonth As Date, trendStart As Date, rowMonth As Date
    prevMonth = DateAdd("m", -1, reportMonth)
    trendStart = DateAdd("m", -(TREND_MONTHS - 1), reportMonth)

    Dim i As Long, product As String, units As Double, revenue As Double
    For i = 2 To UBound(data, 1)
        If StrComp(Trim$(CStr(data(i, COL_REGION))), region, vbTextCompare) = 0 Then
            If IsDate(data(i, COL_DATE)) Then
                rowMonth = DateSerial(Year(data(i, COL_DATE)), Month(data(i, COL_DATE)), 1)
                product = Trim$(CStr(data(i, COL_PRODUCT)))
                units = ToNumber(data(i, COL_UNITS))
                revenue = ToNumber(data(i, COL_REVENUE))

                If rowMonth = reportMonth Then
                    stats.CurRevenue = stats.CurRevenue + revenue
                    stats.CurUnits = stats.CurUnits + units
                    stats.CurOrders = stats.CurOrders + 1
                    prodUnits(product) = prodUnits(product) + units
                    prodRev(product) = prodRev(product) + revenue
                ElseIf rowMonth = prevMonth Then
                    stats.PrevRevenue = stats.PrevRevenue + revenue
                    stats.PrevUnits = stats.PrevUnits + units
                    stats.PrevOrders = stats.PrevOrders + 1
                    prodPrev(product) = prodPrev(product) + revenue
                End If

                If rowMonth >= trendStart And rowMonth <= reportMonth Then
                    stats.Trend(DateDiff("m", trendStart, rowMonth) + 1) = _
                        stats.Trend(DateDiff("m", trendStart, rowMonth) + 1) + revenue
                End If
            End If
        End If
    Next i
End Sub

'==============================================================================
' Lay out the report sheet
'==============================================================================
Private Sub BuildReport(ByVal ws As Worksheet, ByVal region As String, ByVal manager As String, _
                        ByVal reportMonth As Date, ByVal companyName As String, ByRef stats As RegionStats, _
                        ByVal prodUnits As Object, ByVal prodRev As Object, ByVal prodPrev As Object)
    '---- reset ----
    Do While ws.ChartObjects.Count > 0
        ws.ChartObjects(1).Delete
    Loop
    ws.Cells.Clear
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C:F").ColumnWidth = 15
    ws.Columns("G").ColumnWidth = 3
    ws.Columns("H").ColumnWidth = 10
    ws.Columns("I").ColumnWidth = 14

    '---- header ----
    With ws.Range("B1")
        .Value = IIf(Len(companyName) > 0, companyName, "Sales Report")
        .Font.Size = 9
        .Font.Color = RGB(118, 118, 118)
    End With
    With ws.Range("B2")
        .Value = "Monthly Sales Report - " & region
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = CLR_BRAND
    End With
    ws.Range("B3").Value = Format(reportMonth, "mmmm yyyy") & _
                           IIf(Len(manager) > 0, "   |   Prepared for " & manager, "")
    ws.Range("B3").Font.Color = RGB(89, 89, 89)
    With ws.Range("B4:I4").Borders(xlEdgeBottom)
        .Color = CLR_BRAND
        .Weight = xlMedium
    End With

    '---- KPI tiles ----
    Dim aovCur As Double, aovPrev As Double
    If stats.CurOrders > 0 Then aovCur = stats.CurRevenue / stats.CurOrders
    If stats.PrevOrders > 0 Then aovPrev = stats.PrevRevenue / stats.PrevOrders

    WriteKpi ws, "B", "Revenue", stats.CurRevenue, stats.PrevRevenue
    WriteKpi ws, "C", "Units Sold", stats.CurUnits, stats.PrevUnits
    WriteKpi ws, "D", "Orders", stats.CurOrders, stats.PrevOrders
    WriteKpi ws, "E", "Avg Order Value", aovCur, aovPrev
    With ws.Range("B6:E8")
        .Interior.Color = CLR_TILE
        .HorizontalAlignment = xlCenter
    End With

    '---- product table (sorted by revenue, highest first) ----
    ws.Range("B" & TABLE_ROW & ":F" & TABLE_ROW).Value = _
        Array("Product", "Units", "Revenue", "Share", "vs Last Month")
    StyleHeader ws.Range("B" & TABLE_ROW & ":F" & TABLE_ROW)

    Dim keys As Variant, k As Long, r As Long
    keys = SortedKeys(prodRev)
    r = TABLE_ROW + 1
    For k = LBound(keys) To UBound(keys)
        ws.Cells(r, 2).Value = keys(k)
        ws.Cells(r, 3).Value = prodUnits(keys(k))
        ws.Cells(r, 4).Value = prodRev(keys(k))
        If stats.CurRevenue <> 0 Then ws.Cells(r, 5).Value = prodRev(keys(k)) / stats.CurRevenue
        WriteChange ws.Cells(r, 6), prodRev(keys(k)), prodPrev(keys(k))
        If (r - TABLE_ROW) Mod 2 = 0 Then ws.Range(ws.Cells(r, 2), ws.Cells(r, 6)).Interior.Color = CLR_TILE
        r = r + 1
    Next k

    ' total row
    ws.Cells(r, 2).Value = "Total"
    ws.Cells(r, 3).Value = stats.CurUnits
    ws.Cells(r, 4).Value = stats.CurRevenue
    ws.Cells(r, 5).Value = 1
    WriteChange ws.Cells(r, 6), stats.CurRevenue, stats.PrevRevenue
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 6))
        .Font.Bold = True
        .Borders(xlEdgeTop).Color = CLR_BRAND
    End With

    ws.Range(ws.Cells(TABLE_ROW + 1, 3), ws.Cells(r, 4)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(TABLE_ROW + 1, 5), ws.Cells(r, 5)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(TABLE_ROW, 3), ws.Cells(r, 6)).HorizontalAlignment = xlRight

    '---- trend table ----
    ws.Range("H" & TABLE_ROW & ":I" & TABLE_ROW).Value = Array("Month", "Revenue")
    StyleHeader ws.Range("H" & TABLE_ROW & ":I" & TABLE_ROW)
    ws.Range("H" & TABLE_ROW + 1 & ":H" & TABLE_ROW + TREND_MONTHS).NumberFormat = "@"  ' keep labels as text
    Dim m As Long
    For m = 1 To TREND_MONTHS
        ws.Cells(TABLE_ROW + m, 8).Value = Format(DateAdd("m", m - TREND_MONTHS, reportMonth), "mmm yy")
        ws.Cells(TABLE_ROW + m, 9).Value = stats.Trend(m)
    Next m
    ws.Range("I" & TABLE_ROW + 1 & ":I" & TABLE_ROW + TREND_MONTHS).NumberFormat = "#,##0"

    '---- trend chart ----
    Dim chartRow As Long
    chartRow = Application.WorksheetFunction.Max(r, TABLE_ROW + TREND_MONTHS) + 2

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(Left:=ws.Range("B" & chartRow).Left, _
                                 Top:=ws.Range("B" & chartRow).Top, _
                                 Width:=ws.Range("B1:I1").Width, Height:=230)
    With co.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=ws.Range("H" & TABLE_ROW & ":I" & TABLE_ROW + TREND_MONTHS)
        .HasTitle = True
        .ChartTitle.Text = "Revenue - Last " & TREND_MONTHS & " Months"
        .ChartTitle.Font.Size = 11
        .HasLegend = False
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = CLR_BRAND
        .SeriesCollection(1).HasDataLabels = True
        .SeriesCollection(1).DataLabels.NumberFormat = "#,##0"
        .SeriesCollection(1).DataLabels.Font.Size = 8
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(217, 217, 217)
    End With

    '---- footer + page setup ----
    Dim lastRow As Long
    lastRow = co.BottomRightCell.Row + 1
    With ws.Cells(lastRow, 2)
        .Value = "Generated " & Format(Now, "dd-mmm-yyyy hh:nn") & " by the Automated Report Generator"
        .Font.Size = 8
        .Font.Italic = True
        .Font.Color = RGB(150, 150, 150)
    End With

    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = "$A$1:$I$" & lastRow
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
    End With
    Application.PrintCommunication = True
End Sub

'==============================================================================
' Report building helpers
'==============================================================================
Private Sub WriteKpi(ByVal ws As Worksheet, ByVal col As String, ByVal label As String, _
                     ByVal current As Double, ByVal previous As Double)
    With ws.Range(col & "6")
        .Value = label
        .Font.Size = 9
        .Font.Color = RGB(89, 89, 89)
    End With
    With ws.Range(col & "7")
        .Value = current
        .NumberFormat = "#,##0"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = CLR_BRAND
    End With
    WriteChange ws.Range(col & "8"), current, previous
    ws.Range(col & "8").Font.Size = 9
End Sub

' Writes the % change vs previous period, coloured green / red
Private Sub WriteChange(ByVal cell As Range, ByVal current As Double, ByVal previous As Double)
    If previous = 0 Then
        cell.Value = "n/a"
        cell.Font.Color = RGB(118, 118, 118)
    Else
        cell.Value = (current - previous) / previous
        cell.NumberFormat = "+0.0%;-0.0%;0.0%"
        cell.Font.Color = IIf(current >= previous, CLR_UP, CLR_DOWN)
    End If
End Sub

Private Sub StyleHeader(ByVal rng As Range)
    With rng
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = CLR_BRAND
    End With
End Sub

' Dictionary keys sorted by their values, highest first (insertion sort - lists are short)
Private Function SortedKeys(ByVal dict As Object) As Variant
    Dim keys As Variant, i As Long, j As Long, tmp As Variant
    keys = dict.keys
    For i = LBound(keys) + 1 To UBound(keys)
        tmp = keys(i)
        j = i - 1
        Do While j >= LBound(keys)
            If dict(keys(j)) >= dict(tmp) Then Exit Do
            keys(j + 1) = keys(j)
            j = j - 1
        Loop
        keys(j + 1) = tmp
    Next i
    SortedKeys = keys
End Function

'==============================================================================
' Outlook
'==============================================================================
Private Sub SaveDraft(ByVal olApp As Object, ByVal email As String, ByVal manager As String, _
                      ByVal region As String, ByVal reportMonth As Date, ByRef stats As RegionStats, _
                      ByVal pdfPath As String)
    Dim firstName As String
    firstName = Split(manager & " ", " ")(0)
    If Len(firstName) = 0 Then firstName = "Team"

    Dim body As String
    body = "<p>Hi " & firstName & ",</p>" & _
           "<p>Please find attached the <b>" & region & "</b> sales report for <b>" & _
           Format(reportMonth, "mmmm yyyy") & "</b>.</p>" & _
           "<ul>" & _
           "<li>Revenue: <b>" & Format(stats.CurRevenue, "#,##0") & "</b> (" & ChangeText(stats.CurRevenue, stats.PrevRevenue) & " vs last month)</li>" & _
           "<li>Units sold: <b>" & Format(stats.CurUnits, "#,##0") & "</b> (" & ChangeText(stats.CurUnits, stats.PrevUnits) & ")</li>" & _
           "<li>Orders: <b>" & Format(stats.CurOrders, "#,##0") & "</b> (" & ChangeText(stats.CurOrders, stats.PrevOrders) & ")</li>" & _
           "</ul>" & _
           "<p>Regards</p>"

    Dim mail As Object
    Set mail = olApp.CreateItem(0)   ' 0 = olMailItem
    With mail
        .To = email
        .Subject = region & " Sales Report - " & Format(reportMonth, "mmmm yyyy")
        .HTMLBody = body
        .Attachments.Add pdfPath
        .Save                        ' saved to Drafts for review - never sent automatically
    End With
End Sub

Private Function ChangeText(ByVal current As Double, ByVal previous As Double) As String
    If previous = 0 Then
        ChangeText = "n/a"
    Else
        ChangeText = Format((current - previous) / previous, "+0.0%;-0.0%;0.0%")
    End If
End Function

'==============================================================================
' General helpers
'==============================================================================
Private Sub WriteLog(ByVal wsLog As Worksheet, ByVal region As String, ByVal manager As String, _
                     ByVal revenue As Double, ByVal pdfPath As String, ByVal emailStatus As String, _
                     ByVal status As String)
    wsLog.Cells(mLogRow, 1).Resize(1, 6).Value = Array(region, manager, revenue, pdfPath, emailStatus, status)
    wsLog.Cells(mLogRow, 3).NumberFormat = "#,##0"
    mLogRow = mLogRow + 1
End Sub

Private Function ToNumber(ByVal v As Variant) As Double
    If IsNumeric(v) And Not IsEmpty(v) Then ToNumber = CDbl(v)
End Function

Private Function SafeFileName(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, ch, "-")
    Next ch
    SafeFileName = s
End Function

Private Function EnsureFolder(ByVal folderPath As String) As Boolean
    On Error Resume Next
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath
    EnsureFolder = (Len(Dir(folderPath, vbDirectory)) > 0)
End Function

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

Private Sub SpeedUp(ByVal turnOn As Boolean)
    With Application
        If turnOn Then
            mCalcMode = .Calculation
            .Calculation = xlCalculationManual
        ElseIf mCalcMode <> 0 Then
            .Calculation = mCalcMode
        End If
        .ScreenUpdating = Not turnOn
        .EnableEvents = Not turnOn
        .DisplayAlerts = Not turnOn
    End With
End Sub
