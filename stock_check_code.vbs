Sub stock_check():


'setting worksheet variables:
    Dim ws As Worksheet

    ' Starting the loop for each worksheet:
        For Each ws In Worksheets

    'setting the variables in the worksheet:
        Dim ticker As String

        ' Setting the calcuation variables:
            Dim Yearly_Change As Double
            Yearly_Change = 0
    
            Dim total_volume As Double
            total_volume = 0
    
            Dim open_balance As Double
            open_balance = 0

            Dim close_balance As Double
            close_balance = 0
    
            Dim percent_change As Double
            percent_change = 0
          
 
            ' Creating the summary table for calcuations:
                Dim Summary_Table_Row As Integer
                Summary_Table_Row = 2
    
            'set last row:
                lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
                'grab the opening value:
                    open_balance = ws.Cells(2, 3).Value

                ' Loop through all ticker changes:
                    For i = 2 To lastrow

                ' Check if we are still within the same ticker, if it is not...
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                    ' Set the ticker name
                        ticker = ws.Cells(i, 1).Value
                        
                    'grab the closing value:
                        close_balance = ws.Cells(i, 6).Value
                          
                    'calculate the yearly change:
                        Yearly_Change = close_balance - open_balance
                        
                
                    '  percentage change:
                        percent_change = Yearly_Change / open_balance
                        
                        
                'complete the table
                               
                      ' Print the ticker in the Summary Table
                        ws.Range("j" & Summary_Table_Row).Value = ticker
        
                      ' Print the yearly change in the Summary Table
                        ws.Range("k" & Summary_Table_Row).Value = Yearly_Change
                        ws.Range("k" & Summary_Table_Row).NumberFormat = "0.00"
                            ' Add the colour coding for the yearly change change
                                If ws.Range("k" & Summary_Table_Row).Value < 0 Then
                                ws.Range("k" & Summary_Table_Row).Interior.ColorIndex = 3
                            
                                ElseIf ws.Range("k" & Summary_Table_Row).Value >= 0 Then
                                ws.Range("k" & Summary_Table_Row).Interior.ColorIndex = 4
                                
                                End If
                
                    ' Print the percentage in the Summary Table
                        ws.Range("l" & Summary_Table_Row).Value = percent_change
                        ws.Range("l" & Summary_Table_Row).NumberFormat = "0.00%"
                                
                      ' Print the total volume in summary table
                        ws.Range("m" & Summary_Table_Row).Value = total_volume
                      
                      ' Add line to the summary table row
                        Summary_Table_Row = Summary_Table_Row + 1
                      
                      ' Reset the Total
                      total_volume = 0
                
                 
                    ' Create the summary headers
                    
                               ' Add the ticker total Column
                            ws.Range("J1").Value = "ticker"
                            
                            ' Add the yearly change column
                            ws.Range("k1").Value = "yearly change"
                            
                            ' ADD add the total percentage change column
                            ws.Range("L1").Value = "percent change"
                            
                            ' ADD add the total volume column
                            ws.Range("m1").Value = "total volume"
                       
                          
                           
                            'grab the opening value:
                            open_balance = ws.Cells(i + 1, 3).Value
                    
                        Else
                    ' If the cell immediately following a row is the same brand...
                       total_volume = total_volume + ws.Cells(i, 7).Value
                   
                End If

            Next i
   
         Next ws
      
End Sub
