Sub Button1_Click()

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
   For Each ws In Worksheets

        ' --------------------------------------------
        ' INSERT THE YEAR
        ' --------------------------------------------

        ' Create a Variable to Hold File Name, Last Row, and Year
        'im WorksheetName As String
        Dim i As Long
        Dim volume As Double
        Dim start As Double
        Dim rowcount As Integer
        Dim LastRow As Long
        Dim greatest_increase As Double
        Dim greatest_decrease As Double
        Dim greatest_volume As Double

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Set up column names
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percentage Change"
        ws.Range("L1") = "Total Stock Volume"
        

        volume = 0
        'Initialize variables so we can compare them to results
        greatest_increase = 0
        greatest_decrease = 0
        greatest_volume = 0

        WorksheetName = ws.Name

        'Store first open value
        start = ws.Cells(2, 3).Value

        'Start at the second row
        rowcount = 2

        'For each row
        For i = 2 To LastRow
        
        'Add the volume to the total volume
        volume = volume + ws.Cells(i, 7).Value

            'If we are at the end of the data for the stock
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

                ' Set the ticker for the stock at the next available row in the summary table (rowcount)
                ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value
                
                ' Set the yearly change (start value - current cells close value
                ws.Cells(rowcount, 10).Value = ws.Cells(i, 6).Value - start
                
                'Set the percentage change (yearly change/ day 1 open value)
                ws.Cells(rowcount, 11).Value = ws.Cells(rowcount, 10).Value / start
                
                'Set the next cell to be the volume
                ws.Cells(rowcount, 12).Value = volume
                
                'Check if current percent increase is greater than stored highest, if so set current to greatest
                If ws.Cells(rowcount, 11).Value > greatest_increase Then
                    greatest_increase = ws.Cells(rowcount, 11).Value
                    ws.Range("P2").Value = ws.Cells(i, 1).Value
                End If
                
                'Check if current percent decrease is less than stored lowest, if so set current to greatest_decrease
                If ws.Cells(rowcount, 11).Value < greatest_decrease Then
                    greatest_decrease = ws.Cells(rowcount, 11).Value
                    ws.Range("P3").Value = ws.Cells(i, 1).Value
                End If
                    
                'Check if current volume is greater than stored highest, if so set current to greatest
                If volume > greatest_volume Then
                    greatest_volume = volume
                    ws.Range("P4").Value = ws.Cells(i, 1).Value
                End If
                    
                'Reset volume counter
                volume = 0
                
                'Store opening price of first day of the year for the next stock
                start = ws.Cells(i + 1, 3).Value
                
                'Increment row counter so we put the summary on the next row in summary table
                rowcount = rowcount + 1
            End If
        Next i
        
        
    'Parts of the below code are not written by me, Citiation: https://www.wallstreetmojo.com/vba-conditional-formatting/
        'Definining the variables:
        Dim rng As Range
        Dim condition1 As FormatCondition, condition2 As FormatCondition
        'final cell in the j column
        Dim endcell_j As String

         'final cell in the j column
        Dim endcell_k As String

        'Build strings to put in range
        endcell_j = "J" & CStr(rowcount - 1)
        endcell_k = "K" & CStr(rowcount - 1)

        'Fixing/Setting the range on which conditional formatting is to be desired
        Set rng_j = ws.Range("J2", endcell_j)
        Set rng_k = ws.Range("k2", endcell_k)

        'To delete/clear any existing conditional formatting from the range
        rng_j.FormatConditions.Delete
        rng_k.FormatConditions.Delete
          
        'Set number format to percentage
        rng_k.NumberFormat = "0.00%"

        'Defining and setting the criteria for each conditional format for row J
        Set condition1 = rng_j.FormatConditions.Add(xlCellValue, xlGreater, "=0")
         Set condition2 = rng_j.FormatConditions.Add(xlCellValue, xlLessEqual, "=0")

         'Defining and setting the format to be applied for each condition for row J
         With condition1
          .Interior.ColorIndex = 4
          .Font.Bold = False
         End With

         With condition2
          .Interior.ColorIndex = 3
           .Font.Bold = False
         End With
         
         'Defining and setting the criteria for each conditional format for row K
        Set condition1 = rng_k.FormatConditions.Add(xlCellValue, xlGreater, "=0")
         Set condition2 = rng_k.FormatConditions.Add(xlCellValue, xlLessEqual, "=0")

         'Defining and setting the format to be applied for each condition for row k
         With condition1
          .Interior.ColorIndex = 4
          .Font.Bold = False
         End With

         With condition2
          .Interior.ColorIndex = 3
           .Font.Bold = False
         End With
    'End Citiation
         
    'Greatest increases/decrease table
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("Q2").Value = greatest_increase
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("Q3").Value = greatest_decrease
    ws.Range("Q4").Value = greatest_volume
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
         
    ' --------------------------------------------
    ' FIXES COMPLETE
    ' --------------------------------------------
    Next ws

End Sub