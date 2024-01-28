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
Private Sub Worksheet_Activate()
    
    Call ZoomToLastVisibleColumnAndRow
End Sub

Private Sub strt_TF_click()
    Dim strPrompt As String
    
    strPrompt = strt_Page_Expl_Text()
    MsgBox strPrompt, vbInformation
    
End Sub
