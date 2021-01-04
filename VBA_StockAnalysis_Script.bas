Attribute VB_Name = "Module11"
Sub what_is_this_analysis()

'First I need to set a variable through the worksheets and then to loop throughout the entire workbook
Dim ws As Worksheet

'Loop de Loop
For Each ws In Worksheets
    
    'Columns, where you at? I need a definite starting place
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"

    'Now mah variables, need to include Ticker, Year-Open,Close and Change, Percent of the change, and Total volume
    Dim Ticker As String
    Dim row_data As Long
    Dim Year_Open As Double
    Dim Year_Close As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Volume As Double
    Dim lastrow As Long

    'Start doz values, all should be zeros except where I'm starting to place the data after the headers and how to grab from the bottom up
    row_data = 2
    Year_Open = 0
    Year_Close = 0
    Year_Change = 0
    Percent_Change = 0
    Total_Volume = 0
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'More loops to search through each of the rows
        For i = 2 To lastrow
    
    'I need to start my conditional on opening value
    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
    Year_Open = ws.Cells(i, 3).Value
    End If
            
    'Totaling up aaall da volumes for my total stock volume
        Total_Volume = Total_Volume + ws.Cells(i, 7)
    
            'Tracking and putting a conditional when the ticker changes, and where to put them
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                ws.Cells(row_data, 10).Value = Ticker
                ws.Cells(row_data, 13).Value = Total_Volume
                
                'Then must caLCulate the yearly change
                Year_Close = ws.Cells(i, 6).Value
                
                Year_Change = Year_Close - Year_Open
                ws.Cells(row_data, 11).Value = Year_Change
        
            'Must find percent change and make sure I haz enough decimal points
            If Year_Open = 0 And Year_Close = 0 Then
            Percent_Change = 0
            ws.Cells(row_data, 12).Value = Percent_Change
            ws.Cells(row_data, 12).NumberFormat = "0.00%"
            
            'This will make sure it's not dividing by zero but showing the change starting from zero, terrible sentence
            ElseIf Year_Open = 0 Then
            Dim Percent_New_Stock As String
            Percent_New_Stock = "New Stock"
            ws.Cells(row_data, 12).Value = Percent_Change
                        
            Else
            Percent_Change = Year_Change / Year_Open
            ws.Cells(row_data, 12).Value = Percent_Change
            ws.Cells(row_data, 12).NumberFormat = "0.00%"
            End If
                        
            'conditional formatting to show the ups and downs
            If Year_Change >= 0 Then
            ws.Cells(row_data, 11).Interior.ColorIndex = 4
            
            Else
            ws.Cells(row_data, 11).Interior.ColorIndex = 3
            End If
            
                'Moving and resetting the values for aaallz
                row_data = row_data + 1
                Year_Open = 0
                Year_Close = 0
                Year_Change = 0
                Percent_Change = 0
                Total_Volume = 0

            End If
Next i
Next ws
End Sub



