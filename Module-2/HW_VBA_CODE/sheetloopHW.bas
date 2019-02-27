Attribute VB_Name = "Module1"
Sub sheetloop():

Dim ws As Worksheet
Dim lastrow As Long
Dim totalstck As Double
Dim ticker As String
Dim summary_tbl_row As Integer


For Each ws In Worksheets

 ws.Activate
 Range("J1").Value = "Ticker"
 Range("K1").Value = "Total Stock Volume"
 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'  MsgBox (lastrow)
     
      totalstck = 0
      summary_tbl_row = 2
      
      For i = 2 To lastrow
      
      
       If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      
         totalstck = Cells(i, 7).Value + totalstck
         ticker = Cells(i, 1).Value
      
         Range("J" & summary_tbl_row).Value = ticker
         Range("K" & summary_tbl_row).Value = totalstck
         summary_tbl_row = summary_tbl_row + 1
      
       Else
      
         totalstck = Cells(i, 7).Value + totalstck
      
      
       End If
     
     Next i

Next

End Sub

