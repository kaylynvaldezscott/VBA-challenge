Date: 10-SEP-2023
Project Title:
Stock Total Volume Solution 2 of 2 - Kaylyn Valdez-Scott

Project Description:
This is the more difficult solution - building on Solution 1.  
The purpose of this VBA code is to read all rows of stock data and summarize the change from open and
close for each year.  After the data is processed, the VBA uses conditional formatting to show positive
and negative movements for each ticker symbol - Green for a positive, and Red for a negative.  The code
then shows the greatest increase in percentage; and the greatest decrease in percentage.  Finally the
highest total volume is displayed.

Required setup:
STEP 1)
Open up default "Multiple_year_stock_data" worksheet

STEP 2)
In "ThisWorkbook", cut/paste the following code:

Private Sub Workbook_Open()

    If MsgBox("Press OK to begin process, or CANCEL to exit.", vbQuestion + vbOKCancel, "Kaylyn Valdez-Scott - Solution 2") = vbOK Then
        Call GatherData
        MsgBox "Calculations complete", vbInformation, "Kaylyn Valdez-Scott - Solution 2"
    End If

End Sub


STEP 3)
'Insert a module, cut/paste the following code:

Dim gl_ROW As Long

Sub GatherData()
    
    'Solution 2 add functionality to show greatest increase
    'greatest decrease and greatest total volume
    
    'Define Variables
    Dim sTICKER As String       'ticker symbol
    Dim sTempTICKER As String   'to monitor if ticker changes
    Dim WorksheetName As String 'ws name
    Dim LastRow As Long         'hold # of rows
        
    Dim dblHIGH As Double
    Dim dblHIGHest As Double
    
    Dim dblLOW As Double
    Dim dblLOWest As Double
    
    Dim dblOPEN As Double       'capture initial open for year
    Dim dblCLOSE As Double      'capture final close for year
    Dim lVOL As Double
    Dim lVOLtotal As Double
    Dim lCUR_ROW As Long
    Dim lINNERLOOP As Long
    
    For Each ws In Worksheets
        ' Create a Variable to Hold File Name, Last Row, and Year
        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ' Grabbed the WorksheetName
        WorksheetName = ws.Name
        ws.Activate
        
        'initialize variables
        gl_ROW = 2
        lVOLtotal = 0
        ws.Range("i1", "p99999").Clear
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        ws.Cells(1, 15) = "Ticker"
        ws.Cells(1, 16) = "Value"
        ws.Cells(2, 14) = "Greatest % Increase"
        ws.Cells(3, 14) = "Greatest % Decrease"
        ws.Cells(4, 14) = "Greatest Total Vol"
        ws.Columns("N:P").Select
        Selection.ColumnWidth = 18

        sTempTICKER = ws.Cells(2, 1)
        dblLOWest = ws.Cells(2, 5)
        dblHIGHest = ws.Cells(2, 4)
        dblOPEN = ws.Cells(2, 3)
        
        For lINNERLOOP = 2 To LastRow
        
            sTICKER = ws.Cells(lINNERLOOP, 1)
            
            If sTICKER <> sTempTICKER Then
                'calc all ticker values
                 dblCLOSE = ws.Cells(lINNERLOOP - 1, 6)
                'write values to sheet
                
                ws.Cells(gl_ROW, 9) = sTempTICKER                          'ticker out
                ws.Cells(gl_ROW, 10) = dblCLOSE - dblOPEN                  'yearly change out
                ws.Cells(gl_ROW, 11) = (dblCLOSE - dblOPEN) / Abs(dblOPEN) 'perc% change out
                ws.Cells(gl_ROW, 12) = lVOLtotal                           'total volume out
                gl_ROW = gl_ROW + 1
                sTempTICKER = sTICKER
                dblOPEN = ws.Cells(lINNERLOOP, 3)
                lVOLtotal = 0
                'Sheets(WorksheetName).Activate
                DoEvents
            Else
                'keep internal counting of ticker symbol
                 lVOL = ws.Cells(lINNERLOOP, 7)
                lVOLtotal = lVOLtotal + lVOL
                
            End If
            
        Next
        DoEvents
        'show percent change as a percentage
        ws.Range("K:K").Select
        Selection.NumberFormat = "0.00%"
        
        'build conditional formatting
        ws.Range("J2").Select
        ws.Range(Selection, Selection.End(xlDown)).Select
        ws.Range(Selection, Selection.End(xlDown)).Select
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
            Formula1:="=0"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 5296274
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="=0"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        ws.Range("A1").Select
        
        'display greatest increase
        ws.Range("P2").Select
        ActiveCell.FormulaR1C1 = "=MAX(RC[-5]:R[9999]C[-5])"
        'greatest decrease
        ws.Range("P3").Select
        ActiveCell.FormulaR1C1 = "=MIN(RC[-5]:R[9999]C[-5])"
        'greatest total volume
        ws.Range("P4").Select
        ActiveCell.FormulaR1C1 = "=MAX(RC[-4]:R[9999]C[-4])"
        'match the ticker symbol
        ws.Range("O2").Select
        ActiveCell.FormulaR1C1 = "=INDEX(C[-6],MATCH(RC[1],C[-4],))"
        ws.Range("O3").Select
        ActiveCell.FormulaR1C1 = "=INDEX(C[-6],MATCH(RC[1],C[-4],))"
        ws.Range("O4").Select
        ActiveCell.FormulaR1C1 = "=INDEX(C[-6],MATCH(RC[1],C[-3],))"
        
        ws.Range("A1").Select
        DoEvents
        
    Next

    Sheets("2018").Select

End Sub

STEP 4) Save the spreadsheet as a macro-enabled spreadsheet.
STEP 5) Close the spreadsheet and re-open; once loaded the messagebox should pop-up

Click OK - the program will run in about one minute per year sheet.



