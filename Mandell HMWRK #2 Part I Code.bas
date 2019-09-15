Attribute VB_Name = "Module1"

'-------------------------------------------------
'Subroutine to sum Total Volume by Ticker (Part I)
'--------------------------------------------------

Sub WallStreetPartOne()

    'Define Variables
    Dim TotalVolume As Double
    Dim Ticker As String
    Dim NextTicker As String
    Dim UniqueTicker As Long
    Dim LastRow As Double
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    TotalVolume = 0
    UniqueTicker = 2
    
    'For loop to calculate Total Volume by Ticker
    For i = 2 To LastRow
        Ticker = Cells(i, 1).Value
        NextTicker = Cells(i + 1, 1).Value
        
        
        If Ticker = NextTicker Then
            TotalVolume = TotalVolume + Cells(i, 7).Value


        ElseIf Ticker <> NextTicker Then
            
            TotalVolume = TotalVolume + Cells(i, 7).Value
            Cells(UniqueTicker, 10).Value = TotalVolume
            Cells(UniqueTicker, 9).Value = Ticker
         
            'Advance Ticker Tracker
            UniqueTicker = UniqueTicker + 1
            
            'Reset Total Volume for Next Ticker
            TotalVolume = Cells(i + 1, 7).Value
            
        End If
        
     Next i
        
     'Formatting of Response Columns
     Cells(1, 9).Value = ("Ticker")
     Cells(1, 10).Value = ("Total Stock Volume")
     Range("J1").ColumnWidth = 16
     Range("I1:J1").Font.Bold = True
          

End Sub



'-----------------------------------------------------
'Subroutine to activate and work though each Worksheet
'-----------------------------------------------------
Sub WS()
    
    Dim WS As Worksheet
    For Each WS In Worksheets
    WS.Activate
        
    Call WallStreetPartOne
    Next WS
    
End Sub




