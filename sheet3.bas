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

    Dim ImportSheet As Worksheet    '* will be set to location of import sheet
    Dim ImportStartRow As Integer   '* stores the row number where data starts
    Dim LastRow As Long             '* stores last row used by pasted data
    Dim LastCol As Long             '* stores last col used by pasted data
    Dim ImportRange As Range        '* stores the range that contains data to check for duplicates
    Dim aImport() As Variant        '* array to hold import data
    Dim iLoop As Integer            '* generic loop counter
    Dim ws As Worksheet             '* generic worksheet variable
    Dim FoundCells As Range         '* used to store found duplicates
    Dim Notes As String             '* holds the name of the workshet where duplicate was found
    Dim NewTicket As Boolean        '* sets to true if a duplicate is found
    Dim iRow As Integer             '* generic row counter
    
    '* Initialize variables
    Set ImportSheet = ThisWorkbook.Worksheets("Import")
    ImportStartRow = 5
    iRow = ImportStartRow
    iLoop = 1
    
    With ImportSheet
        
        '* get last row and column used
        LastRow = .UsedRange.Rows.Count
        LastCol = .UsedRange.Columns.Count
        
        '* Copies pasted data into an array then clears that data
        Set ImportRange = .Range(.Cells(ImportStartRow, 1), .Cells(LastRow, LastCol))
        ReDim ImportArray(1 To LastRow, 1 To LastCol)
        aImport = ImportRange
        ImportRange.Clear
        
        '* Loop thru the array to see which tickets are already in the workbook
        Do While iLoop <= (LastRow - 4)
        
            NewTicket = False
            Notes = ""
        
            For Each ws In ActiveWorkbook.Worksheets
            
                '* Exclude some worksheets from being checked
                If Left(ws.Name, 3) <> "WOW" And ws.Name <> "Import" Then
                
                    If aImport(iLoop, 1) <> "" Then
                        
                        '* Call the FindAll function - supply the range and the value being searched for
                        Set FoundCells = FindAll(SearchRange:=ws.Range("B1:B65536"), _
                            FindWhat:=aImport(iLoop, 1), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByColumns, _
                            MatchCase:=False, _
                            BeginsWith:=vbNullString, _
                            EndsWith:=vbNullString, _
                            BeginEndCompare:=vbTextCompare)
                            
                        '* If a match is found in already completed tickets - store the name of the worksheet
                        '* where the match is. These tickets are kick backs.
                        If Not (FoundCells Is Nothing) Then
                        
                            If ws.Name <> "Spectrum" And ws.Name <> "Spectrum Wait" Then
                                
                                NewTicket = True
                                Notes = ws.Name
                            
                            End If
                            
                        Else
                        
                            NewTicket = True
                        
                        End If '*Not (FoundCells Is Nothing) Then
                    
                    End If '*aImport(iLoop, 2) <> "" Then
                    
                End If '*Left(ws.Name, 3) <> "WOW" And ws.Name <> "Import" Then
            
            Next '*Each ws In ActiveWorkbook.Worksheets
            
            '* If the ticket is new, paste it back into the sheet in the format we use
            If NewTicket = True Then
                        
                Cells(iRow, 1) = aImport(iLoop, 14) + " " + aImport(iLoop, 15)
                Cells(iRow, 2) = aImport(iLoop, 1)
                Cells(iRow, 3) = Notes
                
                If Len(aImport(iLoop, 17)) = 10 Then
                
                    Cells(iRow, 4) = Left(aImport(iLoop, 17), 5) + " " + Right(aImport(iLoop, 17), 5)
                    
                Else
                
                    Cells(iRow, 4) = aImport(iLoop, 17)
                    
                End If
                
                Cells(iRow, 5) = aImport(iLoop, 28)
                Cells(iRow, 6) = aImport(iLoop, 34)
                Cells(iRow, 7) = aImport(iLoop, 32)
                Cells(iRow, 11) = aImport(iLoop, 18)
                Cells(iRow, 12) = aImport(iLoop, 19)
                
                iRow = iRow + 1
                        
            End If '*NewTicket = True Then
            
            iLoop = iLoop + 1
            
        Loop '* = 1 To (LastRow - (ImportStartRow - 1))
    
    End With '*ImportSheet


End Sub
