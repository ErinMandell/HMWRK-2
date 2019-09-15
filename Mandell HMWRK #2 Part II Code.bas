Attribute VB_Name = "Module1"
'-------------------------------------------------
'Subroutine to sum Total Volume by Ticker and Calculate Annual Changes (Part II)
'--------------------------------------------------

Sub WallStreetPartTwo()

    'Define Variables
    Dim TotalVolume As Double
    Dim Ticker As String
    Dim LastRow As Double
    Dim StockOpen As Double
    Dim StockClose As Double
    Dim StockChange As Double
    Dim StockPercentChange As Double
    Dim ChangeAsPercent As String
    Dim SummaryTableRow As Integer
        
    
    'Define Last Row and Set counter values
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    TotalVolume = 0
    UniqueTicker = 2
    SummaryTableRow = 2
    
    
    'For loop to calculate Total Volume and Price Change by Ticker
    For i = 2 To LastRow
        
        'Check to see if we are still in the same stock ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Store Current Ticker Name
            Ticker = Cells(i, 1).Value
            
            'Add Volume Amount to running total
            TotalVolume = TotalVolume + Cells(i, 7).Value
            
            'Print the Ticker Name in the Summary Table
            Range("I" & SummaryTableRow).Value = Ticker
            
            'Print the Total Volume Amount in the Summary Table
            Range("L" & SummaryTableRow).Value = TotalVolume
            
            'Capture Close Price in a variable
            StockClose = Cells(i, 6).Value
            
            'Print Close Price as debug only
            Cells(SummaryTableRow, 15).Value = StockClose
            
            StockChange = (StockClose - StockOpen)
                If StockChange <> 0 Then
                    StockPercentChange = (StockChange / StockOpen)
                Else: StockPercentChange = 0
                End If
                               
            ChangeAsPercent = FormatPercent(StockPercentChange, 2)
                
            Cells(SummaryTableRow, 10).Value = StockChange
            Cells(SummaryTableRow, 11).Value = ChangeAsPercent
                
                If Cells(SummaryTableRow, 10).Value > 0 Then
                   Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
                
                Else
                   Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
                
                End If
            
            'Increment the Summary Table
            SummaryTableRow = SummaryTableRow + 1
            
            'Reset the Total Volume and StockOpen Price
            TotalVolume = 0
            StockOpen = 0
            
         'If the cell immediately following is the same ticker
         Else
         
            'Add to the Total Volume
            TotalVolume = TotalVolume + Cells(i, 7).Value
            
            'Check if Open Price has been captured yet
            If StockOpen = 0 Then
                
                'Capture Open Price
                StockOpen = Cells(i, 3).Value
                
                'Print Open Price as Dbug Only
                Cells(SummaryTableRow, 14).Value = StockOpen
                
            End If

        End If
        
    Next i
        
     'Populate and Format Response Columns and Headers
     Cells(1, 9).Value = ("Ticker")
     Cells(1, 10).Value = ("Yearly Change")
     Cells(1, 11).Value = ("% Change")
     Cells(1, 12).Value = ("Total Stock Volume")
     Range("L1").ColumnWidth = 16
     Range("I1:L1").Font.Bold = True
     
          

End Sub



'-----------------------------------------------------
'Subroutine to activate and work though each Worksheet
'-----------------------------------------------------
Sub WS()
    
    Dim WS As Worksheet
    For Each WS In Worksheets
    WS.Activate
        
    Call WallStreetPartTwo
    Next WS
    
End Sub



