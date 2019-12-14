Sub VBA_Challenge()

'Declarations

    Dim ws As Worksheet

    Dim SummaryRow As Double

    Dim StockName As String

    Dim YearlyChange As Double

    Dim PercentChange As Double

    Dim OpenAmt As Double

    Dim CloseAmt As Double

    Dim TotalVolCount As Double

    Dim lastRow As Long
    
    Dim sheet_name As String
  
    'Instruction to loop through all worksheets

    For Each ws In ThisWorkbook.Worksheets

    ws.Activate

    sheet_name = ActiveSheet.Name

    'Populate Summary Column Headers

    Cells(1, "I").Value = "Stock Name"

    Cells(1, "J").Value = "Yearly Change"

    Cells(1, "K").Value = "Percent Change"

    Cells(1, "L").Value = "Total Stock Volume"

       

    'Initial Values

    SummaryRow = 2

    TotalVolCount = 0

    OpenAmt = Cells(2, 3).Value

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

      

    'Get Stock Name, Opening Value, & 1st vol amount

     For i = 2 To lastRow
      

        If Cells(i + 1, 1).Value = Cells(i, 1).Value Then

            StockName = Cells(i, 1).Value

            Cells(SummaryRow, "I").Value = StockName

            TotalVolCount = TotalVolCount + Cells(i, 7).Value
       

        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
          

            CloseAmt = Cells(i, 6).Value

          
            TotalVolCount = TotalVolCount + Cells(i, 7).Value

            Cells(SummaryRow, "L").Value = TotalVolCount

                       

            'Formula for YearlyChange

            YearlyChange = (CloseAmt - OpenAmt)

            Cells(SummaryRow, "J").Value = YearlyChange

       

            'Condition & Formula for PercentChange
            
            If OpenAmt = 0 Or CloseAmt = 0 Then
            PercentChange = None
            
                Else: PercentChange = (YearlyChange / OpenAmt)

                Cells(SummaryRow, "K").Value = PercentChange

           
            End If
                    

            'Go to next Summary Row

            SummaryRow = SummaryRow + 1
       

            'Reset Volume Counter

            TotalVolCount = 0

            'Reset Opening Value

            OpenAmt = Cells(i + 1, 3).Value
            

            End If

        Next i

         
        'Challenge___________________________________________
        Dim stock_ticker_incr As String

        Dim stock_ticker_decr As String

        Dim stock_ticker_vol As String

        Dim Max_Value_Incr As Double

        Dim Max_Value_Decr As Double

        Dim Max_Value_Vol As Double

        Dim j As Long

        Dim lastChangeSummaryRow As Long

          

        'Populate Row Descriptors

        Cells(2, "O").Value = "Greatest % Increase"

        Cells(3, "O").Value = "Greatest % Decrease"

        Cells(4, "O").Value = "Greatest Total Volume"

   

        'Populate new Column Headers

        Cells(1, "P").Value = "Stock Ticker"

        Cells(1, "Q").Value = "Value"
  

        lastChangeSummaryRow = Cells(Rows.Count, 9).End(xlUp).Row
        

        For j = 2 To lastChangeSummaryRow
       

        'Max % Increase

        Max_Value_Incr = Application.WorksheetFunction.Max(Range("K2:K" & lastChangeSummaryRow))

       

        'Min % Incr

        Max_Value_Decr = Application.WorksheetFunction.Min(Range("K2:K" & lastChangeSummaryRow))

   

        'Greatest Total Volume; get max value from Summary Volume column

        Max_Value_Vol = Application.WorksheetFunction.Max(Range("L2:L" & lastChangeSummaryRow))

        Cells(4, 17).Value = Max_Value_Vol

        
                If Cells(j, 11).Value = Max_Value_Incr Then

                    stock_ticker_incr = Cells(j, 9).Value

  
                    'Paste values

                    Cells(2, 16).Value = stock_ticker_incr

                    Cells(2, 17).Value = Max_Value_Incr

           

                    ElseIf Cells(j, 11).Value = Max_Value_Decr Then

                        stock_ticker_decr = Cells(j, 9).Value

               

                    'Paste values

                    Cells(3, 16).Value = stock_ticker_decr

                    Cells(3, 17).Value = Max_Value_Decr

  

                    ElseIf Cells(j, 12).Value = Max_Value_Vol Then

                        stock_ticker_vol = Cells(j, 9).Value

                   

                        'Paste values

                        Cells(4, 16).Value = stock_ticker_vol
                       

                    End If
      

                Next j
            
            Cells(1, "O").Value = sheet_name
                
        'FORMATTING______________________________________
                 
        Dim lastSummaryRow As Long

    
        lastSummaryRow = Cells(Rows.Count, 10).End(xlUp).Row
   

        For x = 2 To lastSummaryRow


                If Cells(x, 10).Value >= 0 Then

                Cells(x, 10).Interior.ColorIndex = 4
                  

                Else: Cells(x, 10).Interior.ColorIndex = 3
       

            End If
          

        Next x
   

        'Percent Change Columns

        Columns("K:K").Select

        Selection.EntireColumn.Style = "Percent"
      

        Columns("Q:Q").Select

        Selection.EntireColumn.Style = "Percent"
      

        Cells(4, 17).NumberFormat = "0"
                  

        'Make all Headings Bold

        Range("I1:Q1").Select

        Selection.Font.Bold = True
       

        'Autofit all Columns

        Columns("I:Q").Select

        Selection.EntireColumn.AutoFit
        
        'Left-align Year
        Range("O:O").Select
        Selection.HorizontalAlignment = xlLeft
 

    Next ws


End Sub

