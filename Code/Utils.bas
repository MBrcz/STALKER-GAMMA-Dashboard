Attribute VB_Name = "Utils"
Option Explicit
' This sheet contains utility function that were used across many worksheets in the project.

Public Sub SetWorkbookSlicer(ByVal strSlicerName As String, ByVal strParamToTrue As String)
    ' Sets the slicer in the workbook to the demanded by user value. Does not support multiselection!
    
    ' Accepts:
    '   - strSlicerName: name of the slicer that will be set.
    '   - strParamToTrue: name of the parametr that will be set to True, else shall be set to False
    
    ' Returns:
    '   - None
    
    Dim slBox As SlicerCache
    Dim slItem As SlicerItem
    
    Set slBox = ActiveWorkbook.SlicerCaches(strSlicerName)
    
    slBox.ClearManualFilter
    For Each slItem In slBox.SlicerItems
        If slItem.Value <> strParamToTrue Then
            slItem.Selected = False
        End If
    Next slItem
    
End Sub

Public Sub ShowHighlitedShape(ByVal strSignature As String, ByVal strName As String)
    ' Shows the shape to highlight that contains a known name and signature
     
    ' Accepts:
    '   - strSignature: the element in name that shape must have.
    '   - strName: the additional info about shape that will be set.
    ' Return:
    '   - None
    
    
    Dim shpShape As Shape
    Dim boolSelected As Boolean ' Fix for MG and SMG issue (unfortunate)
    
    For Each shpShape In ActiveSheet.Shapes
        If InStr(1, shpShape.Name, strSignature, vbTextCompare) And InStr(1, shpShape.Name, strName, vbTextCompare) And Not boolSelected Then
            shpShape.Line.Visible = msoTrue
            boolSelected = True
        End If
    Next shpShape

End Sub

Public Sub ClearAllHighlitedShapes(ByVal strSignature As String)
    ' Clears all highlited Shapes in the worksheet, that contains the same name signature.
    
    ' Accepts:
    '   - strSignature: the element in name that shape must have in order to be cleared.
    ' Return:
    '   - None
    
    Dim boolSelected As Boolean
    Dim shpShape As Shape
    
    For Each shpShape In ActiveSheet.Shapes
        If InStr(1, shpShape.Name, strSignature, vbTextCompare) Then
            shpShape.Line.Visible = msoFalse
        End If
    Next shpShape

End Sub

Public Sub ZoomToLastVisibleColumnAndRow()
    ' It matches all visible cells to the Worksheet Zoom object.
    ' Intention of this function is to be placed in worksheet_activate event.
    
    ' BEWARE MIGHT BE DANGEROUS WHEN ALL COLUMNS AND ROWS ARE SELECTED.
    ' It is possible to write a mechanism against it, but I do not want to.
    
    Dim visibleRange As Range
    
    Set visibleRange = Cells.SpecialCells(xlCellTypeVisible)
    visibleRange.Select
    
    ActiveWindow.Zoom = True
    Cells(1, 1).Select

End Sub


Public Sub PreventScrollBarFromAutoScrolling()
    ' Prevents scrollbar object from random autoscrolling down to the oblivion.
    
    ' Accepts:
    ' - None
    ' Returns:
    ' - None
    
    Application.Wait (Now + 0.000005)
End Sub

' ------------------------------------------------------------------------
' ------------------------- FUNCTIONS ------------------------------------
' ------------------------------------------------------------------------

Public Function FindPathToRootDir() As String
    ' Finds the path to the local directory in the user machine.
    
    ' Accepts:
    '   - None
    ' Returns:
    '   - str (Path Value)
    
    FindPathToRootDir = ThisWorkbook.Path
    
End Function
Public Function ReadWholeTableContent(wsSheetObject As Worksheet, intStartRow As Integer, intStartColumn As Integer) As Dictionary
    ' Reads the content of the table that starts at intStartRow and intStartColumn.
    ' BEWARE! This implementation requires the pivot to be data base layout and needs to have headers!

    ' Accepts:
    '   - wsSheetObject - the sheet object in which there is a table
    '   - intStartRow - the beginning of the first row in the table
    '   - intStartColumn - the first position of header in table
    
    ' Returns:
    '   - Dictionary - represents the table content.

    Dim intCurrentRow As Integer
    Dim intCurrentColumn As Integer
    Dim intCountElement As Integer

    Dim arrColumnRange As Variant
    Dim strHeaderName As Variant
    Dim dictTableContent As New Dictionary

    intCurrentRow = intStartRow
    intCurrentColumn = intStartColumn

    Do While IsEmpty(wsSheetObject.Cells(intStartRow, intCurrentColumn).Value) = False
        intCurrentRow = intStartRow + 1
        strHeaderName = wsSheetObject.Cells(intStartRow, intCurrentColumn).Value
        arrColumnRange = Array()

        intCountElement = 0
        Do While IsEmpty(wsSheetObject.Cells(intCurrentRow, intCurrentColumn).Value) = False
            ReDim Preserve arrColumnRange(intCountElement)
            arrColumnRange(intCountElement) = wsSheetObject.Cells(intCurrentRow, intCurrentColumn).Value

            intCountElement = intCountElement + 1
            intCurrentRow = intCurrentRow + 1
        Loop

        dictTableContent.Add strHeaderName, arrColumnRange
        intCurrentColumn = intCurrentColumn + 1
    Loop

    If dictTableContent.count > 0 Then
        Set ReadWholeTableContent = dictTableContent
    Else
        Set ReadWholeTableContent = Nothing
        Debug.Print "WARNING! Table content is null!"
    End If
End Function

