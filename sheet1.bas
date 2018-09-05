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
    Dim ws As Worksheet
    Dim SearchRange As String
    Dim SearchHeader As String
    Dim FoundCells As Range
    Dim FoundCell As Range

    '* Will search for a single or multiple pasted entries
    For Each Change In Changes
    
        '* If that field is blank - no need to search for a duplicate
        If Change <> "" Then
        
            '* Check each sheet for duplicate values - all sheets have the same column layout
            For Each ws In ActiveWorkbook.Worksheets
            
                '* Not all sheets need to be checked
                If Left(ws.Name, 3) <> "WOW" And ws.Name <> "Import" Then
            
                    With ws
                    
                        '* Only search if changes are in column 2 or column 7
                        If Change.Column = 2 Or Change.Column = 4 Then
                        
                            '********************************************
                            '* Set values based on column #             *
                            '*                                          *
                            '* SearchRange is the range of the column   *
                            '* SearchHeader is the column name          *
                            '********************************************
                            If Change.Column = 2 Then
                            
                                SearchRange = "B1:B65536"
                                SearchHeader = "Ticket#"
                            
                            ElseIf Change.Column = 4 Then
                            
                                SearchRange = "D1:D65536"
                                SearchHeader = "Pole#"
                        
                            End If
                            
                            '* Call the FindAll function - supply the range and the value being searched for
                            Set FoundCells = FindAll(SearchRange:=.Range(SearchRange), FindWhat:=Change.Value)
                                                
                            '* If duplicates are found
                            If Not (FoundCells Is Nothing) Then
                            
                                For Each FoundCell In FoundCells
                                    
                                    '* Ignore the cells where values were just entered
                                    If FoundCell.Address <> Change.Address Then
                                
                                        MsgBox (SearchHeader & Change.Value & " found on sheet: " & ws.Name & " in cell: " & FoundCell.Address(False, False))
                                        
                                    End If
                                    
                                Next FoundCell
                                
                            End If
                        
                        End If 'Change.Column = 2 Or Change.Column = 4 Then
                    
                    End With 'ws
                
                End If 'Left(ws.Name, 3) <> "WOW" And ws.Name <> "Import" Then
            
            Next 'Each ws In ActiveWorkbook.Worksheets
        
        End If 'Change <> "" Then
    
    Next 'Each Change In Changes

End Sub

