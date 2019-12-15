Attribute VB_Name = "Module1"
Sub tickermacro()

    'setting variable to hold the worksheet

    'setting a variable to hold the cell ranges for each loop
    Dim cell As Range

    'setting a variable to hold values from the ticker column
    Dim ticker As String

    'setting a variable to hold the sum of the stock volume
    Dim Stock_Volume As Double
    Stock_Volume = 0

    'for tracking each stock in the summary columns
    Dim Sum_Cols As Integer
    
    Dim Year_Change As Double
    
    Dim tickerow As Long
    
    Dim row_start As Long
    
    Dim row_end As Long

    'for holding the values in cells showing the change in percantage
    Dim Percent_Change As Double
    
    Dim open_price As Double
    open_price = 0
    tickerow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row 'should count the number of rows in ticker col

    For Each ws In Worksheets
    
        Sum_Cols = 2
        'looping through all stock metrics
        For i = 2 To tickerow

            'if stock ticker changes then...
            ' this is the end of the row for that ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                'setting ticker
                ticker = ws.Cells(i, 1).Value
            
                'setting change in opening and closing sums
                Year_Change = ws.Cells(i, 6).Value - open_price
                
            If open_price = 0 Then
            
            Percent_Change = 0
            
            Else
            
                'setting percentage change in sums
                Percent_Change = (Year_Change / open_price) * 100
                
            End If
    
                'setting stock totals
                Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
            
                'print ticker to column I
                ws.Range("I" & Sum_Cols).Value = ticker
       
                'print year change to column J
                ws.Range("J" & Sum_Cols).Value = Year_Change
                
                'print change in percentage to column K
                ws.Range("K" & Sum_Cols).Value = Percent_Change
            
                'print stock total to column L
                ws.Range("L" & Sum_Cols).Value = Stock_Volume
                
                'Add 1 to the summary columns
                Sum_Cols = Sum_Cols + 1
                
                'Reset Stock_Volume
                Stock_Volume = 0
                open_price = 0
            'if next cell has the same ticker
            Else
                If open_price = 0 Then
                
                    open_price = ws.Cells(i, 3).Value
                End If
                
                'add to the stock total
                Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
            
            End If
        
        Next i
    
    Next
    
End Sub

Sub CondForm()
 
    Dim changerow As Long
    
    changerow = ActiveSheet.Range("J" & Rows.Count).End(xlUp).Row
    
    For Each ws In Worksheets
    
        For i = 2 To changerow
    
            If ws.Cells(i, 10).Value < 0 Then
 
                ws.Cells(i, 10).Interior.ColorIndex = 3

            Else
            
            ws.Cells(i, 10).Interior.ColorIndex = 4 'green for pos values

            End If
        
        Next i
    
    Next

End Sub
