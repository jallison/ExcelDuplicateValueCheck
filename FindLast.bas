Attribute VB_Name = "FindLast"
Function Row(ws As Worksheet) As Long

    On Error Resume Next
    Row = ws.Cells.Find(What:="*", _
                            after:=ws.Range("A1"), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
    
    On Error GoTo 0

End Function

Function Col(ws As Worksheet) As Long

    On Error Resume Next
    Col = ws.Cells.Find(What:="*", _
                            after:=ws.Range("A1"), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
                            
    
    On Error GoTo 0

End Function


