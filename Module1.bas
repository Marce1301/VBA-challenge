Attribute VB_Name = "Module1"
Sub Challenge2()

 For Each yr In Worksheets
 
  Dim WorksheetName As String
  LastRow = yr.Cells(Rows.Count, 1).End(xlUp).Row
 
 WorksheetName = yr.Name
 
 Dim Ticker As String
 Dim Dat As Integer
 Dim Op As Double
 Dim Cl As Double
 Dim Yearly_Ch As Double
 Dim Percent_Ch As Double
 Dim Total_Stock_Volume As Double
 
 
Dim Yearly_Change_R As Double
Yearly_Change_R = 2
Dim Percent_Change As Double
Percent_Change_R = 2
Dim O As Integer
O = -250
Total_Stock_Volume = 0

yr.Cells(1, 9).Value = "ticker"
yr.Cells(1, 10).Value = "Yearly Change"
yr.Cells(1, 11).Value = "Percent Change"
yr.Cells(1, 12).Value = "Total Stock Volume"

    For i = 2 To LastRow
 
If yr.Cells(i + 1, 1).Value <> yr.Cells(i, 1).Value Then
 
 Ticker = yr.Cells(i, 1).Value
 Cl = yr.Cells(i, 6).Value
 Op = yr.Cells(i + O, 3).Value
Yearly_Ch = Cl - Op
Percent_Ch = Cl / Op - 1
Total_Stock_Volume = Total_Stock_Volume + yr.Cells(i, 7).Value
 
yr.Range("I" & Yearly_Change_R).Value = Ticker
yr.Range("J" & Yearly_Change_R).Value = Yearly_Ch
yr.Range("K" & Yearly_Change_R).Value = Percent_Ch
yr.Range("l" & Yearly_Change_R).Value = Total_Stock_Volume
    
  Yearly_Change_R = Yearly_Change_R + 1
  
 Total_Stock_Volume = 0
 
 Else
 
 Total_Stock_Volume = Total_Stock_Volume + yr.Cells(i, 7).Value
  
    End If
  
  
  Next i
  
  yr.Columns("i:l").AutoFit
  
' yr.Range("K" & Yearly_Change_R).Select
'Selection.NumberFormat = "0.00%"

  Next yr

  End Sub
  
  Sub Colorformatt()
  
  For Each yr In Worksheets
  
  
  For i = 2 To yr.Cells(Rows.Count, 10).End(xlUp).Row
  
  If yr.Cells(i, 10).Value >= 0 Then
  yr.Cells(i, 10).Interior.ColorIndex = 4
  yr.Cells(i, 11).Interior.ColorIndex = 4
  
  Else
 
  yr.Cells(i, 10).Interior.ColorIndex = 3
yr.Cells(i, 11).Interior.ColorIndex = 3

  End If
   
   Next i
   
   Next yr
   
  End Sub

  
  
  
 
 
 

 
    
    




