Sub reset()
Dim ws As Worksheet
    For Each ws In Worksheets
    
        ws.Range("I1:P2000").ClearContents
        ws.Range("I1:P2000").ClearFormats
        
    Next ws

End Sub