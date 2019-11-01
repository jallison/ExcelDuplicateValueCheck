VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub CommandButton1_Click()

    Call clean_data("Hills")

End Sub

Private Sub CommandButton3_Click()

    Call clean_data("TECO")

End Sub

Private Function clean_data(ju_name As String)

    Dim ImportSheet As Worksheet    '* will be set to location of import sheet
    Dim StartRow As Integer         '* stores the row number where data starts
    Dim LastRow As Long             '* stores last row used by pasted data
    Dim LastCol As Long             '* stores last col used by pasted data
    Dim ImportRange As Range        '* stores the range that contains data to check for duplicates
    Dim aImport() As Variant        '* array to hold import data
    Dim iLoop As Integer            '* generic loop counter
    Dim ws As Worksheet             '* generic worksheet variable
    Dim FoundCells As Range         '* used to store found duplicates
    Dim iRow As Integer             '* generic row counter

    '* Initialize variables
    Set ImportSheet = ThisWorkbook.Worksheets("Import")
    StartRow = 5
    iRow = StartRow
    iLoop = 1
    
    Performance.TurnOffAll

    With ImportSheet
        
        '* get last row and column used
        LastRow = FindLast.Row(ImportSheet)
        LastCol = FindLast.Col(ImportSheet)
        
        '* Copies pasted data into an array then clears that data
        Set ImportRange = .Range(.Cells(StartRow, 1), .Cells(LastRow, LastCol))
        ReDim ImportArray(1 To LastRow, 1 To LastCol)
        aImport = ImportRange
        ImportRange.Clear
        
        '* Loop thru the array to see which tickets are already in the workbook, if a ticket already exist in TECO or TECO Wait
        '* then we do not want to list it. It is alreay existing work.
        For iLoop = LBound(aImport) To UBound(aImport)
        
            Set ws = ThisWorkbook.Worksheets(ju_name)
                        
            '* Call the FindAll function - supply the range and the value being searched for
            Set FoundCells = FindAll(SearchRange:=ws.Range("B1:B65536"), FindWhat:=aImport(iLoop, 3))
                
            If (FoundCells Is Nothing) Then
            
                Set ws = ThisWorkbook.Worksheets(ju_name & " Wait")
                
                '* Call the FindAll function - supply the range and the value being searched for
                Set FoundCells = FindAll(SearchRange:=ws.Range("B1:B65536"), FindWhat:=aImport(iLoop, 3))
                    
            End If
                            
            If (FoundCells Is Nothing) Then
                        
                Cells(iRow, 1) = aImport(iLoop, 18) + " " + aImport(iLoop, 19)
                Cells(iRow, 2) = aImport(iLoop, 3)
                
                '* This will split 10 digit pole numbers into two five digit numbers
                If Len(aImport(iLoop, 21)) = 10 Then
                
                    Cells(iRow, 4) = Left(aImport(iLoop, 21), 5) + " " + Right(aImport(iLoop, 21), 5)
                    
                Else
                
                    Cells(iRow, 4) = aImport(iLoop, 21)
                    
                End If
                
                Cells(iRow, 5) = aImport(iLoop, 5)
                Cells(iRow, 6) = aImport(iLoop, 8)
                Cells(iRow, 7) = aImport(iLoop, 38)
                Cells(iRow, 11) = aImport(iLoop, 22)
                Cells(iRow, 12) = aImport(iLoop, 23)
                
                iRow = iRow + 1
                        
            End If '*(FoundCells Is Nothing) Then
            
            '* Reset FoundCells
            Set FoundCells = Nothing
            
        Next iLoop '*= 1 To (LastRow - 4)
        
        '* The number of rows has changed so find the new LastRow
        LastRow = FindLast.Row(ImportSheet)
        
        '* If a match is found in the rest of the sheets - list the sheet name and row in the notes column
        '* these tickets need to be researched to see if they have been kicked back
        For iRow = StartRow To LastRow
        
            For Each ws In ActiveWorkbook.Worksheets
            
                If .Cells(iRow, 2).Value <> "" And Left(ws.Name, 3) = Left(ju_name, 3) Then
                
                    Set FoundCells = FindAll(SearchRange:=ws.Range("B1:B65536"), FindWhat:=.Cells(iRow, 2).Value)
                        
                    If Not (FoundCells Is Nothing) Then
                    
                        Cells(iRow, 3) = ws.Name & " " & FoundCells.Address(False, False)
                    
                    End If
                    
                    '* Reset FoundCells
                    Set FoundCells = Nothing
                        
                End If '.Cells(i, 2).Value <> "" And Left(ws.Name, 3) <> "WOW" And ws.Name <> "Import" Then
            
            Next 'Each ws In ActiveWorkbook.Worksheets
            
        Next iRow '= StartRow To LastRow
        
        '* Sort the list by column 3
        Range("A5:L" & LastRow).Sort key1:=Range("C5:C" & LastRow), order1:=xlAscending, Header:=xlNo
        
    End With '*ImportSheet
    
    Performance.TurnOnAll

End Function
