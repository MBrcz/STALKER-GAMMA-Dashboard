VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' -----------------------------------------------------------------------------------------
' ------------------------------------------INIT ------------------------------------------
' -----------------------------------------------------------------------------------------

Private wsBackendSheet As Worksheet
Private wsReportSheet As Worksheet


Private Sub SetwsBackendSheet()
    Set wsBackendSheet = Worksheets("lwpn2_Backend")
End Sub

Private Sub SetwsReportSheet()
    Set wsReportSheet = Worksheets("lwpn2_Page")
End Sub

Private Sub Initialize_Sheet()
    
    Call SetwsBackendSheet
    Call SetwsReportSheet
    
End Sub

' ---------------------------------------------------------------------------------------------
' -------------------------------------- METHODS ----------------------------------------------
' ---------------------------------------------------------------------------------------------

Private Function dictReadTableWithOffset() As Dictionary
    Set dictReadTableWithOffset = Utils.ReadWholeTableContent(wsBackendSheet, 30, 18)
End Function

' -------------------------------------------------------------------------------------------
' -------------------------------EVENTS FUNCTIONS -------------------------------------------
' -------------------------------------------------------------------------------------------

Private Sub PlaceImagesInTable()
    ' Places all faction bagdes in the corresponding faction name.
    
    ' Accepts:
    ' - None
    ' Returns:
    ' - None
    
    Dim dictOffsetTableContent As New Dictionary
    Dim arrBadgesNames As Variant
    Dim arrIconsNames() As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Call Initialize_Sheet
    Set dictOffsetTableContent = dictReadTableWithOffset()
    
    arrBadgesNames = dictOffsetTableContent.Item("Faction Name")
    arrIconsNames = GetImagesNamesCorrespondingToBagdes(arrBadgesNames)
    
    Call RemoveImagesThatCointainsText("Picture")
    Call CopyAndPlaceTheImage(arrIconsNames)
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub


Private Sub PlaceNewWeapon()
    ' Places the filter that is based on weapon value in range.
    
    ' Accepts:
    ' - None
    ' Returns:
    ' - None
    
    Dim strFilterToPass As String
    'Dim dictPossibleInputs As New Dictionary
    'Dim arrPossibleInputs As Variant
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    strFilterToPass = ThisWorkbook.Names("lwpn2_WeaponName").RefersToRange.Value
    'Set dictPossibleInputs = Utils.ReadWholeTableContent(wsBackendSheet, 30, 1) ' Read the pivot that contains all possible items to select
    'arrPossibleInputs = dictPossibleInputs.Item("Row Labels")
    
    If strFilterToPass = "" Then
        MsgBox "You cannot pass to a filter empty cell! Select or type some weapon name!", vbCritical, "Cannot find weapon named: None"
    
    ElseIf strFilterToPass <> "" Then
        ' Prevents case sensitivity of input list.
        'strFilterToPass = ScanListWithoutCaseSensitivity(strFilterToPass, arrPossibleInputs)
        Call SetWorkbookModSlicer("Slicer_Weapon_Name1", strFilterToPass)
    End If
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

Private Sub ClearsSupportFilters()
    ' Clears the additional filters in the sheet.
    
    ' Accepts:
    ' - None
    ' Returns:
    ' - None
    
    Dim arrSlicersNames() As Variant
    Dim k As Variant
    
    arrSlicersNames = Array("Slicer_Repair_Kit_Type1", "Slicer_Region_Name2", _
                            "Slicer_Type3", "Slicer_Caliber2")
    
    For Each k In arrSlicersNames
        ActiveWorkbook.SlicerCaches(CStr(k)).ClearManualFilter
    Next k

End Sub

Private Sub OpenWorksheet()
    ' Sets table scrollbar minimum and maximum values and adjust zoom to user screen.
    
    ' Accepts:
    ' - None
    ' Returns:
    ' - None
    
    With lwpn2_Scrollbar
        .Min = ThisWorkbook.Names("lwpn2_scr_min").RefersToRange.Value
        .Max = ThisWorkbook.Names("lwpn2_scr_max").RefersToRange.Value
        .Value = 0
    End With
    
    ThisWorkbook.Names("lwpn2_scr_curr").RefersToRange.Value = 0
    Call ZoomToLastVisibleColumnAndRow
    
End Sub

Private Sub ScrollbarMove()
    ' Method that handles changing the scrollbar position
    
    ' Accepts:
    ' - None
    ' Returns:
    ' - None
    
    If lwpn2_Scrollbar.Value <= ThisWorkbook.Names("lwpn2_scr_max").RefersToRange.Value Then
        ThisWorkbook.Names("lwpn2_scr_curr").RefersToRange.Value = lwpn2_Scrollbar.Value
    Else
        ThisWorkbook.Names("lwpn2_scr_curr").RefersToRange.Value = ThisWorkbook.Names("lwpn2_scr_max").RefersToRange.Value
    End If
    
End Sub

' ----------------------------------------------------------------------------------------------------
' -------------------------------------------- UTILS -------------------------------------------------
' ----------------------------------------------------------------------------------------------------

Private Function ScanListWithoutCaseSensitivity(strInputValue As String, arrListToCheck As Variant) As String
    ' Checks if passed input is in list ignoring case sensitivity. If it is returns the element from array

    ' Accepts:
    ' - strInputValue: the input value that will be searched in list
    ' - arrListToCheck: the list of elements that will be checked
    
    ' Returns:
    ' str element name from list if it exists; else ""
    
    Dim k As Variant
    
    For Each k In arrListToCheck
        If InStr(1, LCase(strInputValue), LCase(k), vbTextCompare) > 0 And Len(strInputValue) = Len(k) Then
            ScanListWithoutCaseSensitivity = CStr(k)
            'Debug.Print (CStr(k))
            Exit Function
        End If
    Next k
    
    ScanListWithoutCaseSensitivity = ""
    
End Function

Private Sub SetWorkbookModSlicer(ByVal strSlicerName As String, ByVal strParamToTrue As String)
    ' Sets the slicer in the workbook to the demanded by user value. Does not support multiselection!
    ' Modification of general function due to unacceptable performance issues. It does not take into consideration possibility of all filters active.
    
    ' Accepts:
    '   - strSlicerName: name of the slicer that will be set.
    '   - strParamToTrue: name of the parametr that will be set to True, else shall be set to False
    
    ' Returns:
    '   - None
    
    Dim slBox As SlicerCache
    Dim slItem As SlicerItem
    
    Set slBox = ActiveWorkbook.SlicerCaches(strSlicerName)
    
    ' Sets proxy filter
    slBox.SlicerItems(1).Selected = True
    For Each slItem In slBox.SlicerItems
        
        ' Searches current selected and turns it off
        If slItem.Value <> strParamToTrue And slItem.Selected <> False And slItem.Name <> slBox.SlicerItems(1).Name Then
            slItem.Selected = False
        ElseIf slItem.Value = strParamToTrue Then
            slItem.Selected = True
        End If
    
    Next slItem
    
    ' Removes proxy filter if it's value demanded filter value is differ than proxy filter
    If strParamToTrue <> slBox.SlicerItems(1).Name Then
        slBox.SlicerItems(1).Selected = False
    End If
    
End Sub

' ----------------------------------------------------------------------------------------------------
' -------------------------------------------- DYNAMIC BADGES ----------------------------------------
' ----------------------------------------------------------------------------------------------------

Private Function GetImagesNamesCorrespondingToBagdes(arrBadgesNames As Variant) As String()
    ' Finds the names of the images that are in the backend sheet and are corresponding to badges names.
    
    ' Accepts:
    ' - arrBadgesNames - the array that contains the names of badges
    ' Returns:
    ' - None
    
    Dim shpIcon As Shape
    Dim arrShapesNames() As String
    Dim k As Variant
    Dim intCount As Integer
    ReDim Preserve arrShapesNames(UBound(arrBadgesNames) + 1)
    
    intCount = 0
    For Each k In arrBadgesNames
        If k = "" Then ' prevents searching image for empty badges.
            Exit For
        End If
        
        For Each shpIcon In ActiveWorkbook.Sheets(wsBackendSheet.Name).Shapes
            If InStr(1, shpIcon.Name, CStr(k), vbTextCompare) Then
                arrShapesNames(intCount) = shpIcon.Name
                intCount = intCount + 1
                Exit For
            End If
        Next shpIcon
    Next k

    GetImagesNamesCorrespondingToBagdes = arrShapesNames

End Function

Private Sub RemoveImagesThatCointainsText(strText As String)
    ' Removes all images in the report sheet that have common text in name
    
    ' Accepts:
    ' - strText - the value of the text that will trigger func to remove element
    ' Returns:
    ' - None
    
    Dim shpImage As Shape
    
    For Each shpImage In wsReportSheet.Shapes
        If InStr(1, shpImage.Name, strText, vbTextCompare) Then
            shpImage.Delete
        End If
    Next shpImage

End Sub

Private Sub CopyAndPlaceTheImage(arrImagesNames As Variant)
    ' Copies the image from backend sheet and places it in the report page in the table.
    
    ' Accepts:
    ' - arrImagesNames - the array that contains the names of the images that will be passed
    ' Return:
    ' - None
    
    Dim varImgName As Variant
    Dim shpSourceImage As Shape
    Dim rngTargetCell As Range
    Dim intTargetCellXPos As Integer
    Dim shpNewImage As Image
    
    On Error GoTo leave ' defends against the size of last image issue, yep you shouldn't use goto but don't have another simple choice
    
    intTargetCellXPos = 10
    For Each varImgName In arrImagesNames
        Set shpSourceImage = ActiveWorkbook.Sheets(wsBackendSheet.Name).Shapes(varImgName)
        Set rngTargetCell = ThisWorkbook.Sheets(wsReportSheet.Name).Range("J" & intTargetCellXPos & ":" & "K" & intTargetCellXPos + 2)
        
        If shpSourceImage.Name = "" Then
            Exit Sub
        End If
        ' This is stupid but it works.
        ' Dunno another simple way to select image after pasting and changing it size (after copy)
        With shpSourceImage
            .LockAspectRatio = msoFalse
            .Width = rngTargetCell.Width
            .Height = rngTargetCell.Height
        End With
        
        Application.Wait (Now + 0.000005)
        shpSourceImage.CopyPicture
        rngTargetCell.PasteSpecial
        
        intTargetCellXPos = intTargetCellXPos + 3
        
        If intTargetCellXPos >= 39 Then
            Debug.Print "PLACED!" ' Something must be wrong with my logic dunno what.
            Exit For
        End If
    Next varImgName

leave:
Debug.Print "Cannot load item due to random paste or copy error! (probably)"
'Debug.Print ("Finish!")
End Sub

' ----------------------------------------------------------------------------------------------
' ------------------------------------- CALLBACKS ----------------------------------------------
' ----------------------------------------------------------------------------------------------
' Switching callbacks to private in order them not to annoy druning binding macros to objects.

Private Sub SearchWeapon_click()
    Call Initialize_Sheet
    Call PlaceNewWeapon
    Call OpenWorksheet
    Call PlaceImagesInTable
End Sub

Private Sub ClearFilters_click()
    Call ClearsSupportFilters
End Sub

Private Sub Worksheet_Activate()
    Call OpenWorksheet
    Call PlaceImagesInTable
End Sub

Private Sub lwpn2_Scrollbar_Change()
    Call ScrollbarMove
    Call PlaceImagesInTable
End Sub

Private Sub lwpn2_TF_click()
    Dim strPrompt As String
    
    strPrompt = lwpn2_Page_Expl_Text()
    MsgBox strPrompt, vbInformation
    
End Sub

