Public Function GetBackgroundColor(celle As Range) As Long
' tager en reference til en celle, og returnerer baggrundsfarven i den pågældende celle som longværdi

On Error GoTo errorhandler

resultat = celle.Interior.Color

backgroundColor = resultat

errorhandler:
    If Err.Number <> 0 Then
        GetBackgroundColor = CVErr(xlErrNA)
    End If
End Function
