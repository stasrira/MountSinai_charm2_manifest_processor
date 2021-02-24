Attribute VB_Name = "mdlCustomMenu"

Public Const cCustomMenuName = "&Custom Menu"
Public Const cWsMenuBar = "Worksheet Menu Bar"

Public Sub LoadCustomMenus()
    Dim cmbBar As CommandBar
    Dim cmbControl As CommandBarControl
    Dim cmbSettings As CommandBarControl
    Dim cmbDBLink As CommandBarControl
    
    'add custom menus to Add-In ribon of Excel
    Set cmbBar = Application.CommandBars(cWsMenuBar)
    
    'check if the custom menu already exists. If it exists, delete it; it will be recreated in the later code
    Dim i As Integer ', boolMenuExists As Boolean
    For i = cmbBar.Controls.Count To 1 Step -1
        If cmbBar.Controls.Item(i).Caption = cCustomMenuName Then
            'boolMenuExists = True
            cmbBar.Controls.Item(i).Delete
            Exit For
        End If
    Next
    
    'create menu bar entries
    Set cmbControl = cmbBar.Controls.Add(Type:=msoControlPopup, Temporary:=True) 'adds a menu item to the Menu Bar
    With cmbControl
        .Caption = cCustomMenuName 'names the menu item
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Import Shipment File For Processing" 'adds a description to the menu item
            .OnAction = "ImportShipmentFile" 'runs the specified macro
            .FaceId = 109 '638 '1098 'assigns an icon to the dropdown
        End With
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Import Study/Demographic File For Processing" 'adds a description to the menu item
            .OnAction = "ImportDemographicFile" 'runs the specified macro
            .FaceId = 109 '638 '1098 'assigns an icon to the dropdown
        End With
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Import Lab Corrected Stats File For Conversion" 'adds a description to the menu item
            .OnAction = "ImportLabStatsFile" 'runs the specified macro
            .FaceId = 109 '638 '1098 'assigns an icon to the dropdown
        End With
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Convert Lab Stats to Shipment Manifest" 'adds a description to the menu item
            .OnAction = "Convert_LabStats_To_Shipment_File" 'runs the specified macro
            .FaceId = 1378 '638 '1098 'assigns an icon to the dropdown
        End With
'        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
'            .Caption = "Duplicate CPT columns as Plasma" 'adds a description to the menu item
'            .OnAction = "DuplicateCPTColumnsAsPlasma" 'runs the specified macro
'            .FaceId = 297 '638 '1098 'assigns an icon to the dropdown
'        End With
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Duplicate Heparin columns as Plasma" 'adds a description to the menu item
            .OnAction = "DuplicateHeparinColumnsAsPlasma" 'runs the specified macro
            .FaceId = 297 '638 '1098 'assigns an icon to the dropdown
        End With
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Validate Currently Selected Specimen column" 'adds a description to the menu item
            .OnAction = "ValidateCurrentSpecimenTimepointColumn" 'runs the specified macro
            .FaceId = 249 '638 '1098 'assigns an icon to the dropdown
        End With
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Validate All SELECTED For Processing Specimen columns" 'adds a description to the menu item
            .OnAction = "ValidateAllSpecimenTimepointColumns" 'runs the specified macro
            .FaceId = 706 '638 '1098 'assigns an icon to the dropdown
        End With
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Process Currently Selected Specimen/Timepoint combination" 'adds a description to the menu item
            .OnAction = "SavePreparedFiles" 'runs the specified macro
            .FaceId = 526 '638 '1098 'assigns an icon to the dropdown
        End With
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Process All SELECTED For Processing Specimen/Timepoint combinations" 'adds a description to the menu item
            .OnAction = "ProcessAllSpecimens" 'runs the specified macro
            .FaceId = 156 '638 '1098 'assigns an icon to the dropdown
        End With
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Refresh All Calculated Data" 'adds a description to the menu item
            .OnAction = "RefreshWorkbookData" 'runs the specified macro
            .FaceId = 37 '638 '1098 'assigns an icon to the dropdown
        End With
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Refresh Database Links" 'adds a description to the menu item
            .OnAction = "RefreshDBConnections" 'runs the specified macro
            .FaceId = 688 '638 '1098 'assigns an icon to the dropdown
        End With
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Help" 'adds a description to the menu item
            .OnAction = "OpenHelpLink" 'runs the specified macro
            .FaceId = 49 '638 '1098 'assigns an icon to the dropdown
        End With
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "About" 'adds a description to the menu item
            .OnAction = "ShowVersionMsg" 'runs the specified macro
            .FaceId = 279 '501 '638 '1098 'assigns an icon to the dropdown
        End With
        
    End With
End Sub


