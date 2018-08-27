VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Changes As Range)

    Dim Change As Range
    Dim FoundCells As Range
    Dim FoundCell As Range
    Dim ws As Worksheet
    
    For Each Change In Changes
    
        If Change <> "" Then
        
            For Each ws In ActiveWorkbook.Worksheets
            
                With ws
            
                    If Change.Column = 2 Then
                    
                        Set FoundCells = FindAll(SearchRange:=.Range("B1:B65536"), _
                                            FindWhat:=Change.Value, _
                                            LookIn:=xlValues, _
                                            LookAt:=xlWhole, _
                                            SearchOrder:=xlByColumns, _
                                            MatchCase:=False, _
                                            BeginsWith:=vbNullString, _
                                            EndsWith:=vbNullString, _
                                            BeginEndCompare:=vbTextCompare)
                                            
                        If Not (FoundCells Is Nothing) Then
                        
                            For Each FoundCell In FoundCells
                            
                                If FoundCell.Address <> Change.Address Then
                            
                                    MsgBox ("Ticket# found on sheet: " & ws.Name & " in cell: " & FoundCell.Address(False, False))
                                    
                                End If
                                
                            Next FoundCell
                            
                        End If
                    
                    ElseIf Change.Column = 7 Then
                    
                        Set FoundCells = FindAll(SearchRange:=.Range("G1:G65536"), _
                                            FindWhat:=Change.Value, _
                                            LookIn:=xlValues, _
                                            LookAt:=xlWhole, _
                                            SearchOrder:=xlByColumns, _
                                            MatchCase:=False, _
                                            BeginsWith:=vbNullString, _
                                            EndsWith:=vbNullString, _
                                            BeginEndCompare:=vbTextCompare)
                                            
                        If Not (FoundCells Is Nothing) Then
                        
                            For Each FoundCell In FoundCells
                            
                                If FoundCell.Address <> Change.Address Then
                            
                                    MsgBox ("Pole# found on sheet: " & ws.Name & " in cell: " & FoundCell.Address(False, False))
                                    
                                End If
                                
                            Next FoundCell
                            
                        End If
                    
                    
                    End If
                
                End With
            
            Next
        
        End If
    
    Next

End Sub


