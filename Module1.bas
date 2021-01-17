<<<<<<< HEAD
Attribute VB_Name = "Module1"
Sub VBAStockMarketAnalysis():

' Declare and Loop thru Stocks for Year to Create Output
For Each ws In Worksheets

    Dim Ticker As String
    Dim Total_StockVol As Double
    Total_StockVol = 0
    Dim Counter As Double
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Last_Row As Long
    Dim Open_Year As Double
    Dim Close_Year As Double
    Dim Change_Year As Double
    Dim Rollup As Long
    Rollup = 2
    Dim Last_Row_Amount As Long
    Last_Row_Amount = 2
    Dim Percent_Change As Double
    Dim Summary_Row As Long
    Summary_Row = 2
    Dim Greatest_Increase As Double
    Greatest_Increase = 0
    Dim Greatest_Decreae As Double
    Greatest_Decrease = 0
    Dim Biggest_TotalV As Double
    Biggest_TotalV = 0
    

' Assign Column Headers to Sheet
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

' Whats the Last Row
Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To Last_Row

    Total_StockVol = Total_StockVol + ws.Cells(i, 7).Value
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    Ticker = Cells(i, 1).Value
    ws.Range("I" & Summary_Row).Value = Ticker
    ws.Range("L" & Summary_Row).Value = Total_StockVol
 

' Open, Close and Yearly Change Name
        Open_Year = ws.Range("C" & Last_Row_Amount)
        Close_Year = ws.Range("F" & i)
        Change_Year = Close_Year - Open_Year
        ws.Range("J" & Summary_Row).Value = Change_Year
                
      If Open_Year = 0 Then
         Percent_Change = 0
                
                Else
                    Open_Year = ws.Range("C" & Last_Row_Amount)
                    PercentChange = Change_Year / Open_Year
                End If
              
                ws.Range("K" & Summary_Row).NumberFormat = "0.00%"
                ws.Range("K" & Summary_Row).Value = Percent_Change

        If ws.Range("J" & Summary_Row).Value >= 0 Then
            ws.Range("J" & Summary_Row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & Summary_Row).Interior.ColorIndex = 3
                End If
            
              
                Summary_Row = Summary_Row + 1
                PreviousAmount = i + 1
                End If
            Next i
       

            For i = 2 To Last_Row
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                End If

            Next i
       
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            
    
Columns("I").AutoFit
Columns("J").AutoFit
Columns("K").AutoFit
Columns("L").AutoFit

Next ws
End Sub
=======
Attribute VB_Name = "Module1"
Sub VBAStockMarketAnalysis():

' Declare and Loop thru Stocks for Year to Create Output
For Each ws In Worksheets

    Dim Ticker As String
    Dim Total_StockVol As Double
    Total_StockVol = 0
    Dim Counter As Double
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Last_Row As Long
    Dim Open_Year As Double
    Dim Close_Year As Double
    Dim Change_Year As Double
    Dim Rollup As Long
    Rollup = 2
    Dim Last_Row_Amount As Long
    Last_Row_Amount = 2
    Dim Percent_Change As Double
    Dim Summary_Row As Long
    Summary_Row = 2
    Dim Greatest_Increase As Double
    Greatest_Increase = 0
    Dim Greatest_Decreae As Double
    Greatest_Decrease = 0
    Dim Biggest_TotalV As Double
    Biggest_TotalV = 0
    

' Assign Column Headers to Sheet
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

' Whats the Last Row
Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To Last_Row

    Total_StockVol = Total_StockVol + ws.Cells(i, 7).Value
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    Ticker = Cells(i, 1).Value
    ws.Range("I" & Summary_Row).Value = Ticker
    ws.Range("L" & Summary_Row).Value = Total_StockVol
 

' Open, Close and Yearly Change Name
        Open_Year = ws.Range("C" & Last_Row_Amount)
        Close_Year = ws.Range("F" & i)
        Change_Year = Close_Year - Open_Year
        ws.Range("J" & Summary_Row).Value = Change_Year
                
      If Open_Year = 0 Then
         Percent_Change = 0
                
                Else
                    Open_Year = ws.Range("C" & Last_Row_Amount)
                    PercentChange = Change_Year / Open_Year
                End If
              
                ws.Range("K" & Summary_Row).NumberFormat = "0.00%"
                ws.Range("K" & Summary_Row).Value = Percent_Change

        If ws.Range("J" & Summary_Row).Value >= 0 Then
            ws.Range("J" & Summary_Row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & Summary_Row).Interior.ColorIndex = 3
                End If
            
              
                Summary_Row = Summary_Row + 1
                PreviousAmount = i + 1
                End If
            Next i
       

            For i = 2 To Last_Row
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                End If

            Next i
       
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            
    
Columns("I").AutoFit
Columns("J").AutoFit
Columns("K").AutoFit
Columns("L").AutoFit

Next ws
End Sub
>>>>>>> 52c8b5fa00ed81075d49e0e812db2de3950678bc
