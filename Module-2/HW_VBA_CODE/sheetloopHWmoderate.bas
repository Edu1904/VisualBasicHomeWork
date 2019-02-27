Attribute VB_Name = "Module1"
Sub sheetloop():

' Define all variable
Dim ws As Worksheet
Dim lastrow As Long
Dim firstrow As Long
Dim totalstck As Double
Dim ticker As String
Dim summary_tbl_row As Integer
Dim stckopen As Double
Dim stckopen2 As Double
Dim stckclose As Double
Dim stckavg As Double
Dim perctchange As Double

' Use "For" to loop through every sheet

For Each ws In Worksheets

' to activate all the sheets
 ws.Activate
 
' to assing headers for the summary table
 Range("J1").Value = "Ticker"
 Range("M1").Value = "Total Stock Volume"
 Range("K1").Value = "Yearly Change"
 Range("L1").Value = "Percent Change"
 
' determine the last row for column 1, in each sheet
 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'  MsgBox (lastrow)
     
      totalstck = 0
      summary_tbl_row = 2
      stckopen = Cells(2, 3).Value
     
      
'A For to loop through the rows in the sheet
      For i = 2 To lastrow
           
       
'if to find the point where the stock name change
       If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
       
         stckclose = Cells(i, 6).Value
         
         stckavg = stckclose - stckopen
         
          
' to avoid division by 0 in case the initial stock value was 0
           If stckopen > 0 Then
            perctchange = (stckavg * 100) / stckopen
           Else
            perctchange = (stckavg * 100) / 0.01
           End If
         
         totalstck = Cells(i, 7).Value + totalstck
         ticker = Cells(i, 1).Value
         
'         Range("K" & summary_tbl_row).Interior.Color = xlColorIndexNone
         
         
' use 2 "IF" to color cells background base on cell value
         If stckavg < 0 Then

         
         Else
          
         Range("K" & summary_tbl_row).Interior.ColorIndex = 4
         'xlColorIndexNone
     
          End If
      
         If stckavg >= 0 Then

        
         Else
          
         Range("K" & summary_tbl_row).Interior.ColorIndex = 3
         'xlColorIndexNone
      
         
         End If
      
'assign values in the summary table
         Range("J" & summary_tbl_row).Value = ticker
         Range("M" & summary_tbl_row).Value = totalstck
         Range("K" & summary_tbl_row).Value = stckavg
         Range("L" & summary_tbl_row).Value = perctchange
         
        
         
         summary_tbl_row = summary_tbl_row + 1
         
         stckopen = Cells(i + 1, 3).Value
         
       Else
      
         totalstck = Cells(i, 7).Value + totalstck
      
      
       End If
     
     Next i

Next

End Sub

