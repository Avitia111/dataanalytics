Sub VBA_Homework()
'Define!
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 11).Value = "Change Percent"


'----------------------------------------------

Dim Ticker As String

Dim Group_Index As Double
    Group_Index = 2

Dim Close_Open As Double
    Close_Open = 0
    
Dim Yearly_Change As Double
    Yearly_Change = 0

Dim Percent_change As Double
    Percent_change = 0


Dim Open_Price As Double
    Open_Price = 0

Dim Close_Prrice As Double
    Close_Prrice = 0

Dim Volume As Double
    Volume = 0

Dim Outcome As Integer
    Outcome = 2

' Format
EndingRow = Cells(Rows.Count, 1).End(xlUp).Row

'-------------------------------------------------

For i = 2 To EndingRow

'--------------------------------------------------

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
        Close_Prrice = Cells(i, 6).Value
        Volume = Volume + Cells(i, 7).Value
        Range("I" & Outcome).Value = Ticker
               
        Close_Open = (Close_Prrice - Cells(Group_Index, 3).Value)
            
            If (Cells(Group_Index, 3) <> 0) Then
                    Percent_change = (Close_Open / Cells(Group_Index, 3))
            Else
                Percent_change = 0
            End If
     
        Range("j" & Outcome).Value = Close_Open
        Range("K" & Outcome).Value = Percent_change

    '---------------------------------------------------------------

        If Percent_change < 0 Then
            Range("J" & Outcome).Interior.ColorIndex = 3
            Range("J" & Outcome).NumberFormat = "0.00"
            Range("K" & Outcome).Interior.ColorIndex = 3
            Range("K" & Outcome).NumberFormat = "0.00%"
            
        Else
            Range("J" & Outcome).Interior.ColorIndex = 4
            Range("J" & Outcome).NumberFormat = "0.00"
            Range("K" & Outcome).Interior.ColorIndex = 4
            Range("K" & Outcome).NumberFormat = "0.00%"
        
        End If
               
        Range("L" & Outcome).Value = Volume
        Outcome = Outcome + 1
        Ticker = 0
        Close_Prrice = 0
        Volume = 0
        Group_Index = i + 1
                    
    Else
    '-----------------------------------------------------
        Ticker = Ticker + Cells(i, 1).Value
        Volume = Volume + Cells(i, 7).Value
                        
    End If

    '-------------------
    ' Don't forget Next i to re-itterate
           
  Next i

End Sub
