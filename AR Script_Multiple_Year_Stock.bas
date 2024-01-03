Attribute VB_Name = "Module1"
Sub CalculateYearlyChangeForAllWorksheets()
    Dim ws As Worksheet
    
    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Sheets
    
        ' Call the main subroutine for each worksheet
        CalculateYearlyChange ws
    Next ws
End Sub

Sub CalculateYearlyChange(ws As Worksheet)
    ' Set an initial variable for holding the ticker symbol
    Dim Ticker As String

    ' Set an initial variable for holding the Yearly Change per ticker
    Dim Yearly_Change As Double
    Yearly_Change = 0

    ' Set an initial variable for holding the Percent Change per ticker
    Dim Percent_Change As Double
    Percent_Change = 0

    ' Set an initial variable for holding the Total Volume per ticker
    Dim Total_Volume As LongLong
    Total_Volume = 0

    ' Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2

    ' Variables for tracking greatest % increase, greatest % decrease, and greatest total volume
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim GreatestVolumeTicker As String
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As LongLong

    ' Initialize variables
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestVolume = 0

    ' Loop through all tickers in the current worksheet
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Define the range for yearly changes (assuming they are in column J)
    Dim YearlyChangeRange As Range
    Set YearlyChangeRange = ws.Range("J2:J" & LastRow)

    ' Clear existing conditional formatting rules for yearly changes
    YearlyChangeRange.FormatConditions.Delete

    For i = 2 To LastRow
        ' Check if the ticker changes or if it's the last row
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Or i = LastRow Then
            ' Set the ticker name
            Ticker = ws.Cells(i, 1).Value

            ' Calculate Yearly Change (Last Close Rate - First Open Rate)
            Yearly_Change = ws.Cells(i, 6).Value - ws.Cells(Application.Match(Ticker, ws.Range("A:A"), 0), 3).Value

            ' Calculate Percent Change ((Last Close Rate - First Open Rate) / First Open Rate) * 100
            If ws.Cells(Application.Match(Ticker, ws.Range("A:A"), 0), 3).Value <> 0 Then
                Percent_Change = (Yearly_Change / ws.Cells(Application.Match(Ticker, ws.Range("A:A"), 0), 3).Value) * 100
            Else
                Percent_Change = 0
            End If

            ' Calculate Total Volume
            Total_Volume = Application.SumIf(ws.Range("A:A"), Ticker, ws.Range("G:G"))

            ' Print the ticker symbol in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker

            ' Print the Yearly Change to the Summary Table
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

            ' Print the Percent Change to the Summary Table
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change

            ' Print the Total Volume to the Summary Table
            ws.Range("L" & Summary_Table_Row).Value = Total_Volume

            ' Apply conditional formatting for positive and negative changes
            If Yearly_Change >= 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0) ' Green
            
            Else
                ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0) ' Red
            End If

            ' Update variables for greatest % increase, % decrease, and total volume
            If Percent_Change > GreatestIncrease Then
                GreatestIncrease = Percent_Change
                GreatestIncreaseTicker = Ticker
            End If

            If Percent_Change < GreatestDecrease Then
                GreatestDecrease = Percent_Change
                GreatestDecreaseTicker = Ticker
            End If

            If Total_Volume > GreatestVolume Then
                GreatestVolume = Total_Volume
                GreatestVolumeTicker = Ticker
            End If

            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1

            ' Reset the Yearly Change, Percent Change, and Total Volume
            Yearly_Change = 0
            Percent_Change = 0
            Total_Volume = 0
        End If
    Next i

    ' Output greatest % increase, % decrease, and total volume to separate cells
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("P2").Value = GreatestIncreaseTicker
    ws.Range("Q2").Value = GreatestIncrease & "%"

    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("P3").Value = GreatestDecreaseTicker
    ws.Range("Q3").Value = GreatestDecrease & "%"

    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P4").Value = GreatestVolumeTicker
    ws.Range("Q4").Value = GreatestVolume

' Add Headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
End Sub

