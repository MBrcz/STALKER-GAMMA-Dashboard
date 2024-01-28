VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit
' This module shall contain all functions and events related to inner workings of Weapons all worksheet.

' -----------------------------------------------------------------------------------------
' ------------------------------------------INIT ------------------------------------------
' -----------------------------------------------------------------------------------------

Private strRegionsSlicerName As String
Private strOriginCountrySlicerName As String

Private Sub LetRegionsSlicerName()
    strRegionsSlicerName = "Slicer_Region_Name"
End Sub

Private Sub LetOriginSlicerName()
    strOriginCountrySlicerName = "Slicer_Origin_Country"
End Sub

Private Sub InitializeModule()
    ' Initializes all parameters of the module. USE AT THE BEGINNING OF EVERY FUNCTION HERE.
    
    ' Accepts:
    '   None
    ' Returns:
    '   None
    
    LetRegionsSlicerName
    LetOriginSlicerName

End Sub

' -------------------------------------------------------------------------------------
' --------------------------------------- FUNCTIONS -----------------------------------
' -------------------------------------------------------------------------------------

Private Sub HandleSettingRegionSlicer(ByVal strSetToTrue As String)
    ' This procedure handles the changing of the slicers in the weapons sheet.
    
    ' Accepts:
    '   strSetToTrue - the expression that will be set in filter to True.
    ' Returns:
    '   None
    
    Application.ScreenUpdating = False
    
    ' Initializes the module level variables.
    InitializeModule
    
    ' Clear the Origin Country Filter
    ActiveWorkbook.SlicerCaches(strOriginCountrySlicerName).ClearManualFilter
    
    ' Set the slicer in question.
    Call Utils.SetWorkbookSlicer(strRegionsSlicerName, strSetToTrue)
    
    Call Utils.ClearAllHighlitedShapes("pic_")
    Call Utils.ShowHighlitedShape("pic_", strSetToTrue)
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub ClearAllFilters()
    ' Removes all filters in the current worksheet (ergo restarts slicers)
    ' Accepts:
    '   - None
    ' Returns:
    '   - None
    
    InitializeModule
    ActiveWorkbook.SlicerCaches("Slicer_Region_Name").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Slicer_Origin_Country").ClearManualFilter
    Call Utils.ClearAllHighlitedShapes("pic_")
    
End Sub

' -------------------------------------------------------------------------------------
' -------------------------------------- EVENTS ---------------------------------------
' -------------------------------------------------------------------------------------
' Everything under are the methods that are bound to specified objects in sheet.
' They are private, becouse once bound to objects will be bound forever.
' In case of neccessity of use, just change their encapsulation (from private to public).

Private Sub EuropeFlag_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "European"
    Call HandleSettingRegionSlicer(strSetToTrue)
    
End Sub
 
Private Sub OtherFlag_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "Other"
    Call HandleSettingRegionSlicer(strSetToTrue)
    
End Sub

Private Sub RussianFlag_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "Post-Soviet"
    Call HandleSettingRegionSlicer(strSetToTrue)
    
End Sub

Private Sub USFlag_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "Anglosaxons"
    Call HandleSettingRegionSlicer(strSetToTrue)
        
End Sub

Private Sub ClearFilters_click()
    
    Call ClearAllFilters

End Sub

Private Sub Worksheet_Activate()
    
    Call ZoomToLastVisibleColumnAndRow
    
End Sub

Public Sub cwpn_TF_click()
    Dim strPrompt As String
    
    strPrompt = cwpn_Page_Expl_Text()
    MsgBox strPrompt, vbInformation
    
End Sub

