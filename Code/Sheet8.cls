VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' This module shall contain all functions and events related to inner workings of Weapons all worksheet.

' -----------------------------------------------------------------------------------------
' ------------------------------------------INIT ------------------------------------------
' -----------------------------------------------------------------------------------------

Private strSlicerCaliberName As String
Private strSlicerTypeName As String

Private Sub LetSlicerCaliberName()
    strSlicerCaliberName = "Slicer_Caliber"
End Sub

Private Sub LetSlicerTypeName()
    strSlicerTypeName = "Slicer_Type"
End Sub

Private Sub InitializeModule()
    ' Initializes all parameters of the module. USE AT THE BEGINNING OF EVERY FUNCTION HERE.
    
    ' Accepts:
    '   None
    ' Returns:
    '   None
    
    LetSlicerCaliberName
    LetSlicerTypeName

End Sub

' -------------------------------------------------------------------------------------
' --------------------------------------- FUNCTIONS -----------------------------------
' -------------------------------------------------------------------------------------

Private Sub HandleSettingTypeSlicer(ByVal strSetToTrue As String)
    ' This procedure handles the changing of the slicers in the weapons sheet.
    
    ' Accepts:
    '   strSetToTrue - the expression that will be set in filter to True.
    ' Returns:
    '   None
    
    Application.ScreenUpdating = False
    
    ' Initializes the module level variables.
    InitializeModule
    
    ' Clear the Origin Country Filter
    ActiveWorkbook.SlicerCaches(strSlicerCaliberName).ClearManualFilter
    
    ' Set the slicer in question.
    Call Utils.SetWorkbookSlicer(strSlicerTypeName, strSetToTrue)
    
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
    ActiveWorkbook.SlicerCaches("Slicer_Type").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Slicer_Caliber").ClearManualFilter
    Call Utils.ClearAllHighlitedShapes("pic_")
    
End Sub

' -------------------------------------------------------------------------------------
' -------------------------------------- EVENTS ---------------------------------------
' -------------------------------------------------------------------------------------
' Everything under are the methods that are bound to specified objects in sheet.
' They are private, becouse once bound to objects will be bound forever.
' In case of neccessity of use, just change their encapsulation (from private to public).

Private Sub Shotgun_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "Shotgun"
    Call HandleSettingTypeSlicer(strSetToTrue)
    
End Sub


Private Sub Pistol_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "Pistol"
    Call HandleSettingTypeSlicer(strSetToTrue)
    
End Sub

Private Sub SMG_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "SMG"
    Call HandleSettingTypeSlicer(strSetToTrue)
    
End Sub

Private Sub Sniper_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "Sniper"
    Call HandleSettingTypeSlicer(strSetToTrue)
    
End Sub

Private Sub AssaultRifle_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "Assault Rifle"
    Call HandleSettingTypeSlicer(strSetToTrue)
    
End Sub

Private Sub MG_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "MG"
    Call HandleSettingTypeSlicer(strSetToTrue)
    
End Sub

Private Sub ClearFilters_click()
    
    Call ClearAllFilters

End Sub

Private Sub Worksheet_Activate()
    
    Call ZoomToLastVisibleColumnAndRow
End Sub

Private Sub twpn_TF_click()
    Dim strPrompt As String
    
    strPrompt = twpn_Page_Expl_Text()
    MsgBox strPrompt, vbInformation
    
End Sub

