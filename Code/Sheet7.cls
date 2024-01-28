VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' This module shall contain all functions and events related to inner workings of Weapons all worksheet.

' -----------------------------------------------------------------------------------------
' ------------------------------------------INIT ------------------------------------------
' -----------------------------------------------------------------------------------------

Private strSlicerType1 As String
Private strSlicerFaction As String

Private Sub LetSlicerType1()
    strSlicerType1 = "Slicer_Type1"
End Sub

Private Sub LetSlicerFaction()
    strSlicerFaction = "Slicer_Faction"
End Sub

Private Sub InitializeModule()
    ' Initializes all parameters of the module. USE AT THE BEGINNING OF EVERY FUNCTION HERE.
    
    ' Accepts:
    '   None
    ' Returns:
    '   None
    
    LetSlicerType1
    LetSlicerFaction

End Sub

' -------------------------------------------------------------------------------------
' --------------------------------------- FUNCTIONS -----------------------------------
' -------------------------------------------------------------------------------------

Private Sub HandleSettingSlicer(ByVal strSetToTrue As String, _
                                strShapeSubName As String, _
                                strPrimeSlicerName As String)
    ' This procedure handles the changing of the slicers in the weapons sheet.
    
    ' Accepts:
    '   strSetToTrue - the expression that will be set in filter to True.
    ' Returns:
    '   None
    
    Application.ScreenUpdating = False
    
    ' Initializes the module level variables.
    InitializeModule
    
    ' Set the slicer in question.
    Call Utils.SetWorkbookSlicer(strPrimeSlicerName, strSetToTrue)
    
    Call Utils.ClearAllHighlitedShapes(strShapeSubName)
    Call Utils.ShowHighlitedShape(strShapeSubName, strSetToTrue)
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub ClearAllFilters()
    ' Removes all filters in the current worksheet (ergo restarts slicers)
    ' Accepts:
    '   - None
    ' Returns:
    '   - None
    
    Application.ScreenUpdating = False
    
    InitializeModule
    ActiveWorkbook.SlicerCaches("Slicer_Type1").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Slicer_Faction").ClearManualFilter
    Call Utils.ClearAllHighlitedShapes("arm_")
    
    Application.ScreenUpdating = True
    
End Sub

' -------------------------------------------------------------------------------------
' -------------------------------------- EVENTS ---------------------------------------
' -------------------------------------------------------------------------------------
' Everything under are the methods that are bound to specified objects in sheet.
' They are private, becouse once bound to objects will be bound forever.
' In case of neccessity of use, just change their encapsulation (from private to public).

' -----------------------------------------------
' TYPES
' -----------------------------------------------

Private Sub Exo_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "Exo"
    Call HandleSettingSlicer(strSetToTrue, "arm_T", strSlicerType1)

End Sub

Private Sub Heavy_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "Heavy"
    Call HandleSettingSlicer(strSetToTrue, "arm_T", strSlicerType1)

End Sub

Private Sub Light_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "Light"
    Call HandleSettingSlicer(strSetToTrue, "arm_T", strSlicerType1)

End Sub

Private Sub Medium_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "Medium"
    Call HandleSettingSlicer(strSetToTrue, "arm_T", strSlicerType1)

End Sub

' ---------------------------------
' FACTIONS
' ---------------------------------

Private Sub Bandit_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "Bandit"
    Call HandleSettingSlicer(strSetToTrue, "arm_F", strSlicerFaction)

End Sub

Private Sub Mercenary_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "Mercenary"
    Call HandleSettingSlicer(strSetToTrue, "arm_F", strSlicerFaction)

End Sub

Private Sub Loner_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "Loner"
    Call HandleSettingSlicer(strSetToTrue, "arm_F", strSlicerFaction)

End Sub

Private Sub Duty_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "Duty"
    Call HandleSettingSlicer(strSetToTrue, "arm_F", strSlicerFaction)

End Sub

Private Sub Ecologist_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "Ecologist"
    Call HandleSettingSlicer(strSetToTrue, "arm_F", strSlicerFaction)

End Sub

Private Sub Military_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "Military"
    Call HandleSettingSlicer(strSetToTrue, "arm_F", strSlicerFaction)

End Sub

Private Sub Renegade_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "Renegade"
    Call HandleSettingSlicer(strSetToTrue, "arm_F", strSlicerFaction)

End Sub

Private Sub Sin_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "Sin"
    Call HandleSettingSlicer(strSetToTrue, "arm_F", strSlicerFaction)

End Sub

Private Sub Monolith_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "Monolith"
    Call HandleSettingSlicer(strSetToTrue, "arm_F", strSlicerFaction)

End Sub

Private Sub Unsig_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "UNISG"
    Call HandleSettingSlicer(strSetToTrue, "arm_F", strSlicerFaction)

End Sub


Private Sub Unknown_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "Unknown"
    Call HandleSettingSlicer(strSetToTrue, "arm_F", strSlicerFaction)

End Sub

Private Sub CS_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "Clear Sky"
    Call HandleSettingSlicer(strSetToTrue, "arm_F", strSlicerFaction)
    
End Sub

Private Sub Freedom_click()
    Dim strSetToTrue As String
    
    strSetToTrue = "Freedom"
    Call HandleSettingSlicer(strSetToTrue, "arm_F", strSlicerFaction)

End Sub

' -------------------------------------
' OTHERS
' -------------------------------------

Private Sub ClearFilters_click()
    
    Call ClearAllFilters

End Sub

Private Sub Worksheet_Activate()
    
   Call ZoomToLastVisibleColumnAndRow
End Sub

Private Sub arm_TF_click()
    Dim strPrompt As String
    
    strPrompt = arm_Page_Expl_Text()
    MsgBox strPrompt, vbInformation
    
End Sub


