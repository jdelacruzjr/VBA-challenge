Sub Analyze_StockMarket()

' Define Variables
' ----------------------------------------------

' Define Variable for Ticker
Dim Ticker As String

' Define Variable for Year Open
Dim year_open As Double

' Define Variable for Year Close
Dim year_close As Double

' Define Variable for Yearly Change
Dim Yearly_Change As Double

' Define Variable for Total Stock Volume
Dim Total_Stock_Volume As Double

' Define Variable for Percent Change
Dim Percent_Change As Double

' Define Variable to reference a row to start
Dim start_data As Integer

' Define Variable for Worksheet to excute code in every Worksheet in the entire Workbook
Dim ws As Worksheet

' Initiate Loop in all Worksheets with one execution
' --------------------------------------------------

For Each ws In Worksheets

    ' Assign Column Headers according to their tasks
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    ' Assign location for the loop to start
    start_data = 2
    previous_i = 1
    Total_Stock_Volume = 0
    
    ' Go to the last row of Column A (1)
        EndRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

        ' Loop through daily stock activity to find Ticker, Yearly Change, Percent Change, and Total Stock Volume
        For i = 2 To EndRow
            
            ' Check if still within same ticker name, if not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ' Set the Ticker Name
            Ticker = ws.Cells(i, 1).Value
            
            ' Initiate Variable to go to the next Ticker alphabetically
            previous_i = previous_i + 1
            
            ' Get Opening Value of First day of year from Column C (3) and Closing Value of Last day of year from Column F (6)
            year_open = ws.Cells(previous_i, 3).Value
            year_close = ws.Cells(i, 6).Value
            
            ' Loop to sum the Total Stock Volume using values found in Column G (7)            
            For j = previous_i To i
            
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(j, 7).Value
                
            Next j
            
            ' Reset the loop at the beginning of a year            
            If year_open = 0 Then
            
                Percent_Change = year_close
                
            Else
                Yearly_Change = year_close - year_open
                
                Percent_Change = Yearly_Change / year_open
                
            End If

         ' --------------------------------------------------         
            ' Print Ticker Name, Yearly Change, Percent Change in the Summary Table
            ws.Cells(start_data, 9).Value = Ticker
            ws.Cells(start_data, 10).Value = Yearly_Change
            ws.Cells(start_data, 11).Value = Percent_Change
            
            ' Use Percentage Format
            ws.Cells(start_data, 11).NumberFormat = "0.00%"
            ws.Cells(start_data, 12).Value = Total_Stock_Volume
            
            ' Add one to the Summary Table Row
            start_data = start_data + 1
            
            ' Reset Variable count
            Total_Stock_Volume = 0
            Yearly_Change = 0
            Percent_Change = 0
            
            ' Move row number to Variable previous_i
            previous_i = i
        
        End If
    
    Next i
    
' The Bonus Summary Table
' --------------------------------------------------
    
    ' Go to Last row of Column K (11)
    kEndRow = ws.Cells(Rows.Count, "K").End(xlUp).Row
    
    ' Define Initial location Variable for bonus summary table values
    Increase = 0
    Decrease = 0
    Greatest = 0
    
        ' Loop to find Max/Min for Percentage Change and Max for Greatest Volume
        For k = 3 To kEndRow
        
            ' Define Previous increment to check
            last_k = k - 1
                        
            ' Define Current row for percentage
            current_k = ws.Cells(k, 11).Value
            
            ' Define Previous row for percentage
            previous_k = ws.Cells(last_k, 11).Value
            
            ' Set Greatest Total Volume row
            volume = ws.Cells(k, 12).Value
            
            ' Previous row to Greatest Volume
            previous_vol = ws.Cells(last_k, 12).Value
            
   ' --------------------------------------------------
            
            ' Find increases for percentage
            If Increase > current_k And Increase > previous_k Then
                
                Increase = Increase
                
            ' Define names for percentage increases                
            ElseIf current_k > Increase And current_k > previous_k Then
                
                Increase = current_k
                
                ' Set name for increase percentage
                increase_name = ws.Cells(k, 9).Value
                
            ElseIf previous_k > Increase And previous_k > current_k Then
            
                Increase = previous_k
                
                ' Set name for previous increase percentage
                increase_name = ws.Cells(last_k, 9).Value
                
            End If
                
       ' --------------------------------------------------
            ' Find decreases for percentage            
            If Decrease < current_k And Decrease < previous_k Then
                                
                Decrease = Decrease
                
            ' Define names for percentage decreases    
            ElseIf current_k < Increase And current_k < previous_k Then
                
                Decrease = current_k

                ' Set name for decrease percentage            
                decrease_name = ws.Cells(k, 9).Value
                
            ElseIf previous_k < Increase And previous_k < current_k Then
            
                Decrease = previous_k

                ' Set name for previous decrease percentage
                decrease_name = ws.Cells(last_k, 9).Value
                
            End If
            
       ' --------------------------------------------------
           ' Find the Greatest Volume          
            If Greatest > volume And Greatest > previous_vol Then
            
                Greatest = Greatest
                
            ' Define name for greatest volume            
            ElseIf volume > Greatest And volume > previous_vol Then
            
                Greatest = volume
                
                ' Set name for greatest volume
                greatest_name = ws.Cells(k, 9).Value
                
            ElseIf previous_vol > Greatest And previous_vol > volume Then
                
                Greatest = previous_vol
                
                ' Set name for previous greatest volume
                greatest_name = ws.Cells(last_k, 9).Value
                
            End If
            
        Next k
  ' --------------------------------------------------
    ' Assign Names to Bonus Summary Table    
    ws.Range("N1").Value = ""
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
    ' Set Values to Bonus Summary Table
    ws.Range("O2").Value = increase_name
    ws.Range("O3").Value = decrease_name
    ws.Range("O4").Value = greatest_name
    ws.Range("P2").Value = Increase
    ws.Range("P3").Value = Decrease
    ws.Range("P4").Value = Greatest
    
    ' Use Percentage format for Greatest Increase and Decrease    
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"

' ----------------------------------------------------
' Conditional Formatting (Colors and Auto-fit Columns)
' ----------------------------------------------------
' Go to the last row of column J (10)

    jEndRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
    
        For j = 2 To jEndRow
            
            ' If Yearly Change is Positive...
            If ws.Cells(j, 10) > 0 Then
            
                ' Fill cells with Green
                ws.Cells(j, 10).Interior.ColorIndex = 4
                
            Else
            
                ' If Yearly Change is Negative, Fill cells with Red
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
            
        Next j
    
' AutoFit Columns based on Column Header Text
ws.Columns("A:Q").AutoFit
    
' Continue to Next Worksheet
Next ws
' --------------------------------------------------

End Sub