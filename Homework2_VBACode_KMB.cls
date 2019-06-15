VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Data()

Dim LastRow As Double
Dim TotVol As Double
Dim j As Integer

Cells(1, 10).Value = "Ticker"
Cells(1, 11).Value = "Total Stock Volume"

LastRow = Range("A1").SpecialCells(XlCellType.xlCellTypeLastCell).Row
j = 0
TotVol = 0

    For i = 2 To LastRow

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
           TotVol = TotVol + Cells(i, 7).Value
           Cells(j + 2, 10).Value = Cells(i, 1).Value
            Cells(j + 2, 11).Value = TotVol
            j = j + 1
            TotVol = 0
         Else
            TotVol = TotVol + Cells(i, 7).Value
            
        End If
        
  Next i
  
End Sub
