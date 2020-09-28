Attribute VB_Name = "Module1"
Sub AllYearsData()

For Each ws In Worksheets
    ws.Activate
    Call CombinedTicker
Next ws
    
End Sub

