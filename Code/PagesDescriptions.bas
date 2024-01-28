Attribute VB_Name = "PagesDescriptions"
Option Explicit
' ---------------------------------------------------------------------------------------------------------
' ------------------------------------- SHEETS EXPLANATIONS -----------------------------------------------
' ---------------------------------------------------------------------------------------------------------

Public Function strt_Page_Expl_Text() As String
    ' Text for explaining the starting page in the file.
    
    ' Accepts:
    ' - None
    ' Returns:
    ' - None
    
    Dim strText As String
    
    strText = "     This page contains information about the basics parameters of the project, " & _
              "description of Stalker franchise, basic information about GAMMA mod pack and assumptions of the project." & vbCrLf & vbCrLf & _
              "This report was developed in mind about user using show tabs only in ribbon pannel and dark Excel overall color settings. " & _
              "Therefore, I strongly recommend using such settings." & vbCrLf & vbCrLf & _
              "Special thanks for helpful GAMMA discord users: " & vbCrLf & _
              "1. veerserif - for explaining and finding injected weapons, " & vbCrLf & _
              "2. Talder - for providing and finding missing few other NPCs loadouts."
    
    strt_Page_Expl_Text = strText
    
End Function

Public Function cwpn_Page_Expl_Text() As String
    'Text for explaining the Weapons by regions page in the project.
    
    ' Accepts:
    ' - None
    ' Returns:
    ' - None
    
    Dim strText As String
    
    strText = "   This page contains the information all available weapons in DEBUG MODE by countries and region of the manufacturing. " & _
              "There are four regions of manufacturing: Europe, Anglosaxons, Post-Soviet and Other. " & vbCrLf & vbCrLf & _
              "   It is worth noting that weapon, which has been " & _
              "manufactured by more than one country that and are from different regions, are bound to only one region. For example, theoretically SCAR family " & _
              "weapons were made in collaboration by US and Belgium, then it's bound only to Belgium. Criteria of choosing which country is ARBITRARY." & vbCrLf & vbCrLf & _
              "CONTENT OF THIS FILE IS BASED ON DISCROD FILE, ERGO THERE ARE ALL DEBUG WEAPONS HERE!"
            
    cwpn_Page_Expl_Text = strText
            
End Function

Public Function twpn_Page_Expl_Text() As String
    'Text for explaining the Weapons by types page in the project.
    
    ' Accepts:
    ' - None
    ' Returns:
    ' - None
    
    Dim strText As String
    
    strText = "   This page contains information about weapons by it's types. There are only " & _
              " six weapons types there. Following enumeration shall contain the weapon type and it's content by discord file: " & vbCrLf & _
              "1. Pistols: Pistols; Revolvers; Automatic Pistols; Pistols, " & vbCrLf & _
              "2. Machine Guns: Light Machine Guns; Heavy Machine Guns, " & vbCrLf & _
              "3. Assault Rifles: Assault Rifles, DRM, " & vbCrLf & _
              "4. Submachine guns (SMG): SMG, " & vbCrLf & _
              "5. Shotguns: Shotgun; Double Barrel Shotgun; " & vbCrLf & _
              "6. Snipers: Snipers." & vbCrLf & vbCrLf & _
              "CONTENT OF THIS FILE IS BASED ON DISCROD FILE, ERGO THERE ARE ALL DEBUG WEAPONS HERE!"
              
    twpn_Page_Expl_Text = strText
            
End Function

Public Function arm_Page_Expl_Text() As String
    'Text for explaining the Armors by factions and types page in the project.
    
    ' Accepts:
    ' - None
    ' Returns:
    ' - None
    
    Dim strText As String
    
    strText = "    This page consists the data about the outfits (armors) in the game by factions and types. There are " & _
              "13 factions (12 playable [NOT ZOMBIE] and unknown). Outfits are split by the repair kit they require in order to " & _
              "repair them. There are 4 types (light, medium, heavy, exo)."
    
    arm_Page_Expl_Text = strText
            
End Function

Public Function lwpn2_Page_Expl_Text() As String
    ' Text for explaining the Weapons by Loot Table page
        
    ' Accepts:
    ' - None
    ' Returns:
    ' - None
    
    Dim strText As String
    strText = "    This page tries to answer a very simple question - 'who do I need to farm in order to get X weapon?'" & _
      "There are two types of independent filters there - the slicers straight from Excel is one and Cell $G7:$J9 is the second one." & vbCrLf & vbCrLf & _
      "    The slicers from Excel expand or limit the search size of the second filter. The second filter is a weapon name you wanna check - you can " & _
      "select it using the drop-down menu. After selecting your weapon, you need to press the BIG BUTTON 'Search WEAPON' in order to search the database. " & _
      "After performing those steps, you will see in the middle table of all squads by faction and type that you need to eliminate to get a weapon." & vbCrLf & vbCrLf & _
      "I STRONGLY DO NOT ADVISE TYPING THE WEAPONS NAMES BY HAND DUE TO CASE SENSITIVITY ISSUES!"

    lwpn2_Page_Expl_Text = strText
    
End Function

