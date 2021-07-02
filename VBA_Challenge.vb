


Sub testing_stockanalyis()

Dim ws As Worksheet
For Each ws In Worksheets


' define all the vraiables for summary

 Dim ticker As String
 Dim row As Long
 Dim summaryrow As Long
 Dim volume As Double
 Dim yrclose As Double
 Dim yropen As Double
 Dim yearly_change As Double
 Dim percent_change As Double
 Dim lastrow As Long
 
 'define variables for to find the highest and lowest percent change(greatest percent increase and decrease) and highest volume
 
 Dim highest_change As Double
 Dim lowest_change As Double
 Dim highest_volume As Double
 Dim highest_changeticker As String
 Dim lowest_changeticker As String
 Dim highest_volumeticker As String
 
 'summary_table
ws.Cells(1, 10).Value = "Ticker"
ws.Cells(1, 11).Value = "Yearly differnce"
ws.Cells(1, 12).Value = "Total volume"
ws.Cells(1, 13).Value = "Percent change"

'summary table for greatest percent increase and decrease and highest total volume
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Values"
ws.Cells(1, 15).Value = "Titles"
ws.Cells(2, 15).Value = "Greatest Percent Increase"
ws.Cells(3, 15).Value = "Greatest Percent Decrease"
ws.Cells(4, 15).Value = "Higehst Total Volume"

 ' Adjust the cell to fit the string
ws.Range("O1:O5").Columns.AutoFit
ws.Range("Q1:Q5").Columns.AutoFit

 
 'assign values for the variables
 highest_change = 0
 lowest_change = 0
 volume_change = 0
 
 
 
 
'assign summaryrow for table to start
 
  summaryrow = 2

'find lastrow of the active worksheet
   lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
 

  'loop through the active worksheet
     
     For row = 2 To lastrow
    'define value for the first year open price
       
       If yropen = 0 Then
          yropen = ws.Cells(2, 3).Value
          End If
          
        
          
        
       'check for the last row of the same ticker
       
         If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value And ws.Cells(row - 1, 1).Value = ws.Cells(row, 1).Value Then
         
          
        ' assign values for the ticker
        
         ticker = ws.Cells(row, 1).Value
         yrclose = ws.Cells(row, 6).Value
         year_change = yrclose - yropen
         percent_change = (year_change / yropen) * 100
         yropen = ws.Cells(row + 1, 3).Value
         
         'add the final value for  total volume
         volume = volume + ws.Cells(row, 7).Value
         
         
         'find the highest total volume and it's ticker here before restting the volume for next ticker
         
         If volume > highest_volume Then
         highest_volume = volume
         highest_volumeticker = ticker
         
        
        End If
         
         
         'fill the summarytable
        
        
          ws.Cells(summaryrow, 10).Value = ticker
          ws.Cells(summaryrow, 12).Value = volume
          ws.Cells(summaryrow, 11).Value = year_change
          ws.Cells(summaryrow, 13).Value = percent_change
          
          
          'conditional formatting for the values in "Yearly differnce"
          
          If ws.Cells(summaryrow, 11) >= 0 Then
          ws.Cells(summaryrow, 11).Interior.Color = RGB(0, 250, 0)
          Else
          ws.Cells(summaryrow, 11).Interior.Color = RGB(250, 0, 0)
           End If
           
          'add new row for the next ticker in the summary table
          
          summaryrow = summaryrow + 1
          
          ' reset the value of the total volume for the next ticker
          
          volume = ws.Cells(row + 1, 7).Value
          
          
          
        
        
        'find the greatest percent increase and decrease
        
        If percent_change > highest_change Then
        highest_change = percent_change
        highest_changeticker = ticker
        
        End If
        
        If percent_change < lowest_change Then
        lowest_change = percent_change
        lowest_changeticker = ticker
        
        End If
         
         
        
        
         'to calculate the total volume
         
        Else
        volume = volume + ws.Cells(row, 7).Value
         
         
         
        
        End If
         
        Next row
         
         
        ' fill the summary table for the highest and lowest percent change and highest volume
        
        ws.Range("P2").Value = highest_changeticker
        ws.Range("Q2").Value = CStr(highest_change & "%")
        ws.Range("P3").Value = lowest_changeticker
        ws.Range("Q3").Value = CStr(lowest_change & "%")
        ws.Range("P4").Value = highest_volumeticker
        ws.Range("Q4").Value = highest_volume
        
        
        Next ws
        
        
        
        End Sub



