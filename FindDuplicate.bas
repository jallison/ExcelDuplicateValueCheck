Attribute VB_Name = "FindDuplicate"
Option Explicit
Option Compare Text
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modFindAll
' By Chip Peasron, chip@cpearson.com. www.cpearson.com
' Web page for this module: www.cpearson.com/Excel/FindAll.aspx
' 24-October-2007
' Revised 5-January-2010
' This module is described at www.cpearson.com/Excel/FindAll.aspx
' Requires Excel 2000 or later.
'
' This module contains two functions, FindAll and FindAllOnWorksheets that are use
' to find values on a worksheet or multiple worksheets.
'
' FindAll searches a range and returns a range containing the cells in which the
'   searched for text was found. If the string was not found, it returns Nothing.

Function FindAll(SearchRange As Range, _
                FindWhat As Variant, _
               Optional LookIn As XlFindLookIn = xlValues, _
                Optional LookAt As XlLookAt = xlWhole, _
                Optional SearchOrder As XlSearchOrder = xlByRows, _
                Optional MatchCase As Boolean = False, _
                Optional BeginsWith As String = vbNullString, _
                Optional EndsWith As String = vbNullString, _
                Optional BeginEndCompare As VbCompareMethod = vbTextCompare) As Range
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FindAll
' This searches the range specified by SearchRange and returns a Range object
' that contains all the cells in which FindWhat was found. The search parameters to
' this function have the same meaning and effect as they do with the
' Range.Find method. If the value was not found, the function return Nothing. If
' BeginsWith is not an empty string, only those cells that begin with BeginWith
' are included in the result. If EndsWith is not an empty string, only those cells
' that end with EndsWith are included in the result. Note that if a cell contains
' a single word that matches either BeginsWith or EndsWith, it is included in the
' result.  If BeginsWith or EndsWith is not an empty string, the LookAt parameter
' is automatically changed to xlPart. The tests for BeginsWith and EndsWith may be
' case-sensitive by setting BeginEndCompare to vbBinaryCompare. For case-insensitive
' comparisons, set BeginEndCompare to vbTextCompare. If this parameter is omitted,
' it defaults to vbTextCompare. The comparisons for BeginsWith and EndsWith are
' in an OR relationship. That is, if both BeginsWith and EndsWith are provided,
' a match if found if the text begins with BeginsWith OR the text ends with EndsWith.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim FoundCell As Range
Dim FirstFound As Range
Dim LastCell As Range
Dim ResultRange As Range
Dim XLookAt As XlLookAt
Dim Include As Boolean
Dim CompMode As VbCompareMethod
Dim Area As Range
Dim MaxRow As Long
Dim MaxCol As Long
Dim BeginB As Boolean
Dim EndB As Boolean


CompMode = BeginEndCompare
If BeginsWith <> vbNullString Or EndsWith <> vbNullString Then
    XLookAt = xlPart
Else
    XLookAt = LookAt
End If

' this loop in Areas is to find the last cell
' of all the areas. That is, the cell whose row
' and column are greater than or equal to any cell
' in any Area.
For Each Area In SearchRange.Areas
    With Area
        If .Cells(.Cells.Count).Row > MaxRow Then
            MaxRow = .Cells(.Cells.Count).Row
        End If
        If .Cells(.Cells.Count).Column > MaxCol Then
            MaxCol = .Cells(.Cells.Count).Column
        End If
    End With
Next Area
Set LastCell = SearchRange.Worksheet.Cells(MaxRow, MaxCol)


'On Error Resume Next
On Error GoTo 0
Set FoundCell = SearchRange.Find(What:=FindWhat, _
        after:=LastCell, _
        LookIn:=LookIn, _
        LookAt:=XLookAt, _
        SearchOrder:=SearchOrder, _
        MatchCase:=MatchCase)

If Not FoundCell Is Nothing Then
    Set FirstFound = FoundCell
    'Set ResultRange = FoundCell
    'Set FoundCell = SearchRange.FindNext(after:=FoundCell)
    Do Until False ' Loop forever. We'll "Exit Do" when necessary.
        Include = False
        If BeginsWith = vbNullString And EndsWith = vbNullString Then
            Include = True
        Else
            If BeginsWith <> vbNullString Then
                If StrComp(Left(FoundCell.Text, Len(BeginsWith)), BeginsWith, BeginEndCompare) = 0 Then
                    Include = True
                End If
            End If
            If EndsWith <> vbNullString Then
                If StrComp(Right(FoundCell.Text, Len(EndsWith)), EndsWith, BeginEndCompare) = 0 Then
                    Include = True
                End If
            End If
        End If
        If Include = True Then
            If ResultRange Is Nothing Then
                Set ResultRange = FoundCell
            Else
                Set ResultRange = Application.Union(ResultRange, FoundCell)
            End If
        End If
        Set FoundCell = SearchRange.FindNext(after:=FoundCell)
        If (FoundCell Is Nothing) Then
            Exit Do
        End If
        If (FoundCell.Address = FirstFound.Address) Then
            Exit Do
        End If

    Loop
End If
    
Set FindAll = ResultRange

End Function



