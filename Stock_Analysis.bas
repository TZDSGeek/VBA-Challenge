Attribute VB_Name = "Module1"
Sub Stock_Analysis()

'Forloop for worksheets
    For Each ws In Worksheets
        ws.Activate

'Declare the Variables:
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Yearly_Change_Calc As Double
    Dim Percentage_Change As Double
    Dim Percentage_Change_Calc As Double
    Dim Stock_Volume As Long
    Dim Total_Stock_Volume As Double
    Dim Total_Stock_Volume_Var As Double
    Dim Close_Date As Double
    Dim Open_Price As Double
    Dim Stock_Date As Date
    Dim High_Price As Double
    Dim Low_Price As Double

'Set the summary total row and declare variable
    Dim Summary_Total_Row As Integer
    Summary_Total_Row = 2

'Set the total row for the yearly change,Percentage Change, and total Stock volume
    Yearly_Change = 0
    Percentage_Change = 0
    Total_Stock_Volume = 0
    
'Determine the last row

    Last_Row = Cells(Rows.Count, "A").End(xlUp).Row
    
'Set the open price

    Open_Price = Cells(2, "C").Value

    
'Add row summary headers

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

'Create Forloop
    For i = 2 To Last_Row

'Set Ticker and Total Stock Volume
        Ticker = Cells(i, "A").Value
        Total_Stock_Volume_Var = Cells(i, "G").Value

'Create logic to check if the the next cell has the same ticker symbol
        If (Cells(i + 1, "A").Value <> Cells(i, "A").Value) Then



'Aggregate the variables

            Yearly_Change = Cells(i, "F").Value - Open_Price
            Percentage_Change = Yearly_Change / Open_Price * 100
            Total_Stock_Volume = Total_Stock_Volume + Total_Stock_Volume_Var

'Print Ticker, Yearly Change,Percentage Change, and Total Stock Volume

            Range("I" & Summary_Total_Row).Value = Ticker
            Range("J" & Summary_Total_Row).Value = Yearly_Change
            Range("K" & Summary_Total_Row).Value = "%" & Percentage_Change
            Range("L" & Summary_Total_Row).Value = Total_Stock_Volume
            
'Create Conditional formating for the Yearly Change
            
            If (Yearly_Change > 0) Then
            
                Cells(Summary_Total_Row, "J").Interior.ColorIndex = 4
                
            ElseIf (Yearly_Change < 0) Then
            
                Cells(Summary_Total_Row, "J").Interior.ColorIndex = 3
                
            Else
                Cells(Summary_Total_Row, "J").Interior.ColorIndex = 2
                
            End If
            
'Create Conditional formating for the Percentage Change

    If (Percentage_Change > 0) Then
            
                Cells(Summary_Total_Row, "K").Interior.ColorIndex = 4
                
            ElseIf (Percentage_Change < 0) Then
            
                Cells(Summary_Total_Row, "K").Interior.ColorIndex = 3
                
            Else
                Cells(Summary_Total_Row, "K").Interior.ColorIndex = 2
                
            End If
    

'Add one to the summary total row

            Summary_Total_Row = Summary_Total_Row + 1
    
'Reset the Open Price and Total Stock Volume

            Open_Price = Cells(i + 1, "C").Value
            
            Total_Stock_Volume = 0

'If cell immediately following a row has the same ticker...

        Else

'Add the Values again

           Total_Stock_Volume = Total_Stock_Volume + Total_Stock_Volume_Var

        End If

    Next i
    
'--------------------------------------------------------------------------
'Create the summaries
'--------------------------------------------------------------------------

'Create Summary Row Headers

        Range("O2").Value = "Greatest Percent Increase"
        Range("O3").Value = "Greatest Percent Decrease"
        Range("O4").Value = "Greatest Total Stock Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
'Determine the last row of the summary table

        summary_last_row = Cells(Rows.Count, "I").End(xlUp).Row
        
'Greatest increase,Decrease,and total volume Variables

        greatest_percent_increase = Cells(2, "K").Value
        greatest_percent_decrease = Cells(2, "K").Value
        greatest_total_volume = Cells(2, "L").Value
        
'Create forloop for greatest_percent_increase

        For Row = 2 To summary_last_row
        
'Compare to find Greatest_percent_Increase

            If Cells(Row, "K").Value > greatest_percent_increase Then
            
                greatest_percent_increase = Cells(Row, "K").Value
                
                Cells(2, "P").Value = Cells(Row, "I").Value
                
            Else
                
                greatest_percent_increase = greatest_percent_increase
                
            End If
            
'Compare to find the Greatest percent_decrease

            If Cells(Row, "K").Value < greatest_percent_decrease Then
            
                greatest_percent_decrease = Cells(Row, "K").Value
                
                Cells(3, "P").Value = Cells(Row, "I").Value
                
            Else
                
                greatest_percent_decrease = greatest_percent_decrease
                
            End If
        
'Compare to find greatest stock volume

            If Cells(Row, "L").Value > greatest_total_volume Then
            
                greatest_total_volume = Cells(Row, "L").Value
                
                Cells(4, "P").Value = Cells(Row, "I").Value
                
            Else
                
                greatest_total_volume = greatest_total_volume
                
            End If
            
'Print the summary results

Cells(2, "Q").Value = Format(greatest_percent_increase, "Percent")
Cells(3, "Q").Value = Format(greatest_percent_decrease, "Percent")
Cells(4, "Q").Value = Format(greatest_total_volume, "Scientific")
            
        Next Row
    
    Next ws
    
    MsgBox ("Stock Analysis Complete")

End Sub

